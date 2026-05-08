#!/usr/bin/env python3
"""
Spidey V2 pre-upload context/token estimator.

Estimates extracted text tokens from prompt + uploaded files before submitting a turn.
This is an approximation, not an exact Claude tokenizer. It is designed to catch risky
large files and identify the largest token contributors.
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import os
import re
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable, List, Optional

SUPPORTED_EXTENSIONS = {".pdf", ".docx", ".xlsx", ".csv", ".txt", ".md"}
DEFAULT_CONTEXT_LIMIT = 1_000_000


@dataclass
class FileEstimate:
    file: str
    extension: str
    size_mb: float
    characters: int
    words: int
    estimated_tokens: int
    risk_band: str
    extraction_notes: str


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def estimate_tokens(text: str) -> int:
    """Rough token estimate. Uses tiktoken if available, otherwise char/word heuristic."""
    text = text or ""
    try:
        import tiktoken  # type: ignore

        # cl100k_base is not Claude's tokenizer, but it is a useful proxy.
        enc = tiktoken.get_encoding("cl100k_base")
        return len(enc.encode(text))
    except Exception:
        chars = len(text)
        words = len(re.findall(r"\S+", text))
        # English technical text often lands between chars/4 and words*1.25-1.45.
        # Use the more conservative/larger value.
        return int(max(chars / 4.0, words * 1.35))


def risk_band(tokens: int, total_limit: int = DEFAULT_CONTEXT_LIMIT) -> str:
    pct = tokens / total_limit
    if pct >= 0.95:
        return "CRITICAL"
    if pct >= 0.85:
        return "LIKELY_OVER_LIMIT_SOON"
    if pct >= 0.70:
        return "CLOSE_TO_LIMIT"
    if pct >= 0.40:
        return "WATCH"
    return "SAFE"


def extract_pdf(path: Path, max_pages: Optional[int] = None) -> tuple[str, str]:
    try:
        import fitz  # PyMuPDF
    except Exception as exc:
        return "", f"PDF skipped: PyMuPDF not installed ({exc})"

    notes = []
    chunks = []
    try:
        with fitz.open(path) as doc:
            page_count = len(doc)
            pages_to_read = page_count if max_pages is None else min(page_count, max_pages)
            for i in range(pages_to_read):
                page = doc.load_page(i)
                chunks.append(page.get_text("text"))
            notes.append(f"PDF pages read: {pages_to_read}/{page_count}")
            if max_pages is not None and page_count > max_pages:
                notes.append("PDF truncated by max-pages option")
    except Exception as exc:
        return "", f"PDF extraction failed: {exc}"
    return clean_text("\n".join(chunks)), "; ".join(notes)


def extract_docx(path: Path) -> tuple[str, str]:
    try:
        import docx  # python-docx
    except Exception as exc:
        return "", f"DOCX skipped: python-docx not installed ({exc})"

    chunks = []
    try:
        d = docx.Document(str(path))
        for p in d.paragraphs:
            if p.text:
                chunks.append(p.text)
        for table in d.tables:
            for row in table.rows:
                chunks.append("\t".join(cell.text for cell in row.cells))
        return clean_text("\n".join(chunks)), f"DOCX paragraphs: {len(d.paragraphs)}, tables: {len(d.tables)}"
    except Exception as exc:
        return "", f"DOCX extraction failed: {exc}"


def extract_xlsx(path: Path, max_cells: Optional[int] = None) -> tuple[str, str]:
    try:
        from openpyxl import load_workbook
    except Exception as exc:
        return "", f"XLSX skipped: openpyxl not installed ({exc})"

    chunks = []
    cell_count = 0
    try:
        wb = load_workbook(filename=path, read_only=True, data_only=False)
        sheet_count = len(wb.sheetnames)
        for ws in wb.worksheets:
            chunks.append(f"\n[SHEET: {ws.title}]\n")
            for row in ws.iter_rows():
                vals = []
                for cell in row:
                    if cell.value is not None:
                        vals.append(str(cell.value))
                        cell_count += 1
                        if max_cells is not None and cell_count >= max_cells:
                            chunks.append("\t".join(vals))
                            return clean_text("\n".join(chunks)), f"XLSX sheets: {sheet_count}; cells read: {cell_count}; truncated by max-cells option"
                if vals:
                    chunks.append("\t".join(vals))
        return clean_text("\n".join(chunks)), f"XLSX sheets: {sheet_count}; non-empty cells read: {cell_count}"
    except Exception as exc:
        return "", f"XLSX extraction failed: {exc}"


def extract_csv(path: Path, max_rows: Optional[int] = None) -> tuple[str, str]:
    chunks = []
    rows = 0
    try:
        with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as f:
            sample = f.read(4096)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample)
            except Exception:
                dialect = csv.excel
            reader = csv.reader(f, dialect)
            for row in reader:
                rows += 1
                chunks.append("\t".join(row))
                if max_rows is not None and rows >= max_rows:
                    return clean_text("\n".join(chunks)), f"CSV rows read: {rows}; truncated by max-rows option"
        return clean_text("\n".join(chunks)), f"CSV rows read: {rows}"
    except Exception as exc:
        return "", f"CSV extraction failed: {exc}"


def extract_text(path: Path) -> tuple[str, str]:
    ext = path.suffix.lower()
    if ext == ".pdf":
        return extract_pdf(path)
    if ext == ".docx":
        return extract_docx(path)
    if ext == ".xlsx":
        return extract_xlsx(path)
    if ext == ".csv":
        return extract_csv(path)
    if ext in {".txt", ".md"}:
        try:
            return clean_text(path.read_text(encoding="utf-8", errors="replace")), "Plain text/Markdown read"
        except Exception as exc:
            return "", f"Text extraction failed: {exc}"
    return "", f"Unsupported extension: {ext}"


def iter_files(paths: Iterable[Path]) -> Iterable[Path]:
    for path in paths:
        if path.is_dir():
            for p in sorted(path.rglob("*")):
                if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS:
                    yield p
        elif path.is_file():
            if path.suffix.lower() == ".zip":
                # For safety, do not auto-extract zip. Tell user to unzip intentionally.
                continue
            if path.suffix.lower() in SUPPORTED_EXTENSIONS:
                yield path


def estimate_file(path: Path, context_limit: int = DEFAULT_CONTEXT_LIMIT) -> FileEstimate:
    text, notes = extract_text(path)
    chars = len(text)
    words = len(re.findall(r"\S+", text))
    tokens = estimate_tokens(text)
    try:
        size_mb = path.stat().st_size / (1024 * 1024)
    except Exception:
        size_mb = 0.0
    return FileEstimate(
        file=str(path),
        extension=path.suffix.lower(),
        size_mb=round(size_mb, 2),
        characters=chars,
        words=words,
        estimated_tokens=tokens,
        risk_band=risk_band(tokens, context_limit),
        extraction_notes=notes,
    )


def estimate_prompt(prompt: str, context_limit: int = DEFAULT_CONTEXT_LIMIT) -> FileEstimate:
    text = prompt or ""
    chars = len(text)
    words = len(re.findall(r"\S+", text))
    tokens = estimate_tokens(text)
    return FileEstimate(
        file="[PROMPT_TEXT]",
        extension="prompt",
        size_mb=0.0,
        characters=chars,
        words=words,
        estimated_tokens=tokens,
        risk_band=risk_band(tokens, context_limit),
        extraction_notes="Prompt text supplied manually",
    )


def summarize(estimates: List[FileEstimate], context_limit: int = DEFAULT_CONTEXT_LIMIT) -> dict:
    total = sum(e.estimated_tokens for e in estimates)
    pct = total / context_limit * 100
    if total >= context_limit:
        status = "OVER_LIMIT"
    elif pct >= 95:
        status = "CRITICAL"
    elif pct >= 85:
        status = "LIKELY_OVER_LIMIT_SOON"
    elif pct >= 70:
        status = "CLOSE_TO_LIMIT"
    else:
        status = "SAFE"
    largest = sorted(estimates, key=lambda e: e.estimated_tokens, reverse=True)[:5]
    return {
        "context_limit": context_limit,
        "total_estimated_tokens": total,
        "percent_of_limit": round(pct, 1),
        "status": status,
        "largest_contributors": [asdict(e) for e in largest],
    }


def print_table(estimates: List[FileEstimate], context_limit: int) -> None:
    estimates_sorted = sorted(estimates, key=lambda e: e.estimated_tokens, reverse=True)
    print("\nEstimated extracted-token budget")
    print("=" * 95)
    print(f"{'Tokens':>10}  {'%Limit':>7}  {'Words':>9}  {'MB':>7}  {'Risk':<22}  File")
    print("-" * 95)
    for e in estimates_sorted:
        pct = e.estimated_tokens / context_limit * 100
        print(f"{e.estimated_tokens:>10,}  {pct:>6.1f}%  {e.words:>9,}  {e.size_mb:>7.2f}  {e.risk_band:<22}  {Path(e.file).name}")
    print("-" * 95)
    summary = summarize(estimates, context_limit)
    print(f"TOTAL: {summary['total_estimated_tokens']:,} / {context_limit:,} tokens ({summary['percent_of_limit']}%)")
    print(f"STATUS: {summary['status']}")
    print("\nNote: This estimates parsed/extracted text tokens, not file size. Claude's exact tokenizer may differ.")


def main() -> None:
    parser = argparse.ArgumentParser(description="Estimate extracted token count for Spidey V2 prompt + files.")
    parser.add_argument("paths", nargs="+", help="Files or folders to scan. Supported: PDF, DOCX, XLSX, CSV, TXT, MD.")
    parser.add_argument("--prompt-file", help="Optional prompt text file to include in the estimate.")
    parser.add_argument("--limit", type=int, default=DEFAULT_CONTEXT_LIMIT, help="Context limit. Default: 1,000,000 tokens.")
    parser.add_argument("--json-out", help="Optional path to write JSON report.")
    args = parser.parse_args()

    estimates: List[FileEstimate] = []
    if args.prompt_file:
        prompt_path = Path(args.prompt_file)
        prompt = prompt_path.read_text(encoding="utf-8", errors="replace")
        estimates.append(estimate_prompt(prompt, args.limit))

    files = list(iter_files(Path(p) for p in args.paths))
    for f in files:
        estimates.append(estimate_file(f, args.limit))

    print_table(estimates, args.limit)

    if args.json_out:
        report = {"summary": summarize(estimates, args.limit), "files": [asdict(e) for e in estimates]}
        Path(args.json_out).write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"\nJSON report written to: {args.json_out}")


if __name__ == "__main__":
    main()

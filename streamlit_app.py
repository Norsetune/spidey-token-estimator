from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

from token_estimator import DEFAULT_CONTEXT_LIMIT, estimate_file, estimate_prompt, summarize

st.set_page_config(page_title="Spidey Context Budget Estimator", layout="wide")

st.title("Spidey V2 Context Budget Estimator")
st.caption("Estimate parsed/extracted text tokens before submitting a turn. This is an approximation, not an exact Claude tokenizer.")

context_limit = st.number_input("Context limit", value=DEFAULT_CONTEXT_LIMIT, step=50_000, min_value=50_000)
prompt_text = st.text_area("Paste prompt text here", height=180, placeholder="Optional: paste the turn prompt to include it in the budget.")
files = st.file_uploader(
    "Upload files to estimate",
    type=["pdf", "docx", "xlsx", "csv", "txt", "md"],
    accept_multiple_files=True,
)

if st.button("Estimate context budget", type="primary"):
    estimates = []
    if prompt_text.strip():
        estimates.append(estimate_prompt(prompt_text, int(context_limit)))

    with tempfile.TemporaryDirectory() as tmpdir:
        for uploaded in files:
            suffix = Path(uploaded.name).suffix
            tmp_path = Path(tmpdir) / uploaded.name
            tmp_path.write_bytes(uploaded.getvalue())
            estimates.append(estimate_file(tmp_path, int(context_limit)))

    if not estimates:
        st.warning("Add a prompt and/or upload files first.")
    else:
        summary = summarize(estimates, int(context_limit))
        st.subheader("Estimated turn budget")
        st.metric("Estimated tokens", f"{summary['total_estimated_tokens']:,}", f"{summary['percent_of_limit']}% of limit")
        st.progress(min(summary["total_estimated_tokens"] / int(context_limit), 1.0))

        status = summary["status"]
        if status == "SAFE":
            st.success("Status: SAFE")
        elif status == "CLOSE_TO_LIMIT":
            st.warning("Status: CLOSE TO LIMIT")
        elif status == "LIKELY_OVER_LIMIT_SOON":
            st.warning("Status: LIKELY OVER LIMIT SOON")
        else:
            st.error(f"Status: {status}")

        rows = []
        for e in sorted(estimates, key=lambda x: x.estimated_tokens, reverse=True):
            rows.append({
                "File": Path(e.file).name,
                "Type": e.extension,
                "Size MB": e.size_mb,
                "Characters": e.characters,
                "Words": e.words,
                "Estimated tokens": e.estimated_tokens,
                "% of limit": round(e.estimated_tokens / int(context_limit) * 100, 1),
                "Risk": e.risk_band,
                "Notes": e.extraction_notes,
            })
        df = pd.DataFrame(rows)
        st.subheader("Per-file estimates")
        st.dataframe(df, use_container_width=True, hide_index=True)

        st.subheader("Largest contributors")
        for item in summary["largest_contributors"]:
            st.write(f"- **{Path(item['file']).name}** — {item['estimated_tokens']:,} tokens")

        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button("Download CSV report", csv, "spidey_context_budget_report.csv", "text/csv")

st.divider()
st.markdown(
    """
**Interpretation:** MB is not the important metric. The important estimate is extracted text tokens after the files are parsed. A small but dense PDF can be riskier than a large image-heavy PDF.

**Recommended bands:** under 700k is usually safer, 700k–850k needs caution, 850k–950k is risky, and above 950k is likely to fail once hidden overhead is included.
"""
)

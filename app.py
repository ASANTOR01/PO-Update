"""
app.py - Streamlit front end for the PO extractor (supports both PO layouts).

Run with:   streamlit run app.py
"""

import io
from datetime import datetime

import pandas as pd
import streamlit as st

from po_core import (
    LAYOUT_BF,
    ParseResult,
    build_dataframe,
    parse_po_pdf,
    to_excel_bytes,
)

st.set_page_config(page_title="PO → Excel Extractor", page_icon="📄", layout="wide")
st.title("📄 Purchase Order → Excel summary")
st.caption("Upload PO PDFs (old WBDCSBC… layout and/or new BF27… layout – detected automatically), "
           "process them, and download one consolidated sheet.")

# --------------------------------------------------------------------------- #
# Sidebar options. These only change how results are *presented/exported*, so
# toggling them does not require re-reading the PDFs.
# --------------------------------------------------------------------------- #
with st.sidebar:
    st.header("Options")
    sort_files = st.checkbox("Sort files by name", value=True,
                             help="The reference sheet lists POs in ascending file-name order.")
    fabrication = st.checkbox("Include 'Fabrication' column", value=True,
                              help="New template column (filled for the BF27… layout, empty for old POs). "
                                   "Untick for the original 14-column layout.")
    pc_formula = st.checkbox("Pc as =SUM formula", value=True,
                             help="Matches the template: Pc = SUM of the size columns. "
                                  "Untick to write the quantity printed on the PDF as a plain number.")
    zero_fill = st.checkbox("Write 0 for blank sizes", value=True,
                            help="Untick to leave blank PDF size cells empty in Excel.")
    lookup_hdrs = st.checkbox("Include lookup headers (right-hand block)", value=True,
                              help="The template's right-hand columns are XLOOKUPs into another workbook. "
                                   "Headers are reproduced; cells are left empty.")

# --------------------------------------------------------------------------- #
# Upload + process
# --------------------------------------------------------------------------- #
uploads = st.file_uploader("Purchase Order PDFs", type=["pdf"], accept_multiple_files=True)

if st.button("Process PDFs", type="primary", disabled=not uploads):
    results: list[ParseResult] = []
    errors: list[str] = []
    progress = st.progress(0.0, text="Reading PDFs…")
    for i, up in enumerate(uploads, start=1):
        try:
            # Copy into BytesIO so pdfplumber gets a seekable stream.
            results.append(parse_po_pdf(io.BytesIO(up.getvalue()), up.name))
        except Exception as exc:   # corrupt / non-PDF / unexpected layout
            errors.append(f"{up.name}: {exc}")
        progress.progress(i / len(uploads), text=f"Processed {i}/{len(uploads)}")
    progress.empty()
    # Streamlit reruns the whole script on every widget interaction (including the
    # download button), so results must live in session_state.
    st.session_state["results"] = results
    st.session_state["errors"] = errors

# --------------------------------------------------------------------------- #
# Results
# --------------------------------------------------------------------------- #
results: list[ParseResult] = st.session_state.get("results", [])
for msg in st.session_state.get("errors", []):
    st.error(f"Could not read {msg}")

if results:
    if sort_files:
        results = sorted(results, key=lambda r: r.file_label)

    summary = pd.DataFrame([{
        "File": r.file_label,
        "Layout": "New (BF27…)" if r.layout == LAYOUT_BF else "Old (WBDCSBC…)",
        "Item rows": len(r.rows),
        "Pieces (parsed)": sum(x["Pc"] for x in r.rows),
        "Pieces (PO TOTAL)": r.printed_total_pcs,
        "Status": "✅ OK" if not r.warnings else f"⚠️ {len(r.warnings)} warning(s)",
    } for r in results])
    st.subheader("Summary")
    st.dataframe(summary, hide_index=True, use_container_width=True)

    for r in results:
        if r.warnings:
            with st.expander(f"⚠️ Warnings – {r.file_label}"):
                for w in r.warnings:
                    st.write("•", w)

    df = build_dataframe(results, blank_sizes_as_zero=zero_fill,
                         include_fabrication=fabrication)
    st.subheader(f"Preview ({len(df):,} rows)")
    st.dataframe(df, hide_index=True, use_container_width=True)

    st.download_button(
        "⬇️ Download Excel",
        data=to_excel_bytes(df, include_lookup_headers=lookup_hdrs, pc_as_formula=pc_formula),
        file_name=f"PO_Summary_{datetime.now():%Y%m%d_%H%M}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
elif not uploads:
    st.info("Upload one or more PO PDFs to get started.")

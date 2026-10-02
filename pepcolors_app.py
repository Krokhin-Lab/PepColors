"""
Program: pepcolors_app.py
Author: Alexandre Prefontaine
Affiliation: University of Manitoba, Manitoba Centre for Proteomics and Systems Biology

Streamlit web application for PepColors.

Run with:
    streamlit run pepcolors_app.py
"""

from pathlib import Path

import streamlit as st

from pepcolors_core import (
    SCHEMES, DEFAULT_SCHEME, read_sequences, excel_bytes, write_html, html_preview, _text_color,
)

PREVIEW_ROWS = 25

st.set_page_config(page_title="PepColors", layout="wide")

st.title("PepColors — residue coloring for peptides")
st.write(
    "Upload a CSV of peptide sequences and download them colored by residue class "
    "(Excel for colored text, HTML for highlights), to help spot patterns in peptides that "
    "deviate from predictions."
)

st.divider()

# --- Options ---
col_a, col_b = st.columns(2)
with col_a:
    scheme = st.selectbox("Coloring scheme", list(SCHEMES), index=list(SCHEMES).index(DEFAULT_SCHEME))
with col_b:
    style = st.radio("Output style", ["Colored text", "Colored highlights"], horizontal=True)
mode = "text" if style == "Colored text" else "highlight"

if mode == "text":
    st.caption("Downloads an Excel file: plain sequence in column A, colored sequence in column B.")
else:
    st.caption("Downloads an HTML page (opens in any browser) with each sequence highlighted "
               "and kept whole. Excel can't highlight single letters inside a cell.")

# White panel so colors read the same in light and dark mode
BOX = ("background:#FFFFFF;color:#222222;border:1px solid #DDDDDD;border-radius:8px;"
       "padding:12px 16px;display:inline-block;min-width:260px")

# --- Legend ---
st.subheader("Color key")
legend = []
for g in SCHEMES[scheme]["groups"]:
    c = g["color"]
    if mode == "text":
        swatch = f'<span style="color:{c};font-weight:bold;font-family:Courier New,monospace">{g["residues"]}</span>'
    else:
        swatch = (f'<span style="background:{c};color:{_text_color(c)};font-weight:bold;'
                  f'font-family:Courier New,monospace;padding:0 4px">{g["residues"]}</span>')
    legend.append(f"<div style='margin:3px 0'>{swatch} &nbsp; {g['label']}</div>")
st.markdown(f"<div style='{BOX}'>{''.join(legend)}</div>", unsafe_allow_html=True)

st.divider()

# --- Input ---
uploaded_file = st.file_uploader("Upload a CSV file (sequences in the first column)", type=["csv"])

if uploaded_file is not None:
    try:
        sequences, skipped = read_sequences(uploaded_file)
    except Exception as e:
        st.error(f"Could not read the file: {e}")
        st.stop()

    if skipped:
        st.warning(f"Skipped {len(skipped)} invalid entries: {', '.join(skipped[:50])}"
                   + (" …" if len(skipped) > 50 else ""))

    if not sequences:
        st.warning("No valid sequences found. Use one-letter codes for the 20 standard amino acids.")
        st.stop()

    st.success(f"{len(sequences)} peptides ready.")

    # --- Download ---
    stem = Path(uploaded_file.name).stem
    if mode == "text":
        st.download_button(
            label="Download colored Excel file",
            data=excel_bytes(sequences, scheme_name=scheme),
            file_name=stem + "_PepColors.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
    else:
        st.download_button(
            label="Download highlighted HTML file",
            data=write_html(sequences, title=f"{stem} — PepColors", scheme_name=scheme).encode("utf-8"),
            file_name=stem + "_PepColors.html",
            mime="text/html",
            type="primary",
        )

    # --- Preview ---
    st.subheader("Preview")
    if len(sequences) > PREVIEW_ROWS:
        st.caption(f"First {PREVIEW_ROWS} of {len(sequences)} peptides.")
    rows = "".join(
        f"<div style='margin:3px 0'>{html_preview(s, mode, scheme)}</div>"
        for s in sequences[:PREVIEW_ROWS]
    )
    st.markdown(f"<div style='{BOX}'>{rows}</div>", unsafe_allow_html=True)

# --- Footer ---
st.divider()
col1, col2 = st.columns([1, 3])
with col1:
    logo = Path(__file__).parent / "assets" / "UM-logo-horizontal-CMYK.jpg"
    if logo.exists():
        st.image(str(logo), width=120)
with col2:
    st.markdown(
        "**Alexandre Préfontaine** — Krokhin Laboratory<br>"
        "Manitoba Centre for Proteomics and Systems Biology<br>"
        "[University of Manitoba](https://umanitoba.ca) · Department of Internal Medicine",
        unsafe_allow_html=True,
    )

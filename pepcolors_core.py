# PepColors — residue coloring for peptide sequences
# Copyright (C) 2026 Oleg V. Krokhin, Alexandre Préfontaine
# University of Manitoba, Manitoba Centre for Proteomics and Systems Biology
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.
"""
Program: pepcolors_core.py
Author: Alexandre Prefontaine
Affiliation: University of Manitoba, Manitoba Centre for Proteomics and Systems Biology

Core module for PepColors: sequence validation, coloring schemes and the
Excel writer. No Streamlit dependency, so it can be used from scripts:

    from pepcolors_core import read_sequences, write_excel
    seqs, skipped = read_sequences("peptides.csv")
    write_excel(seqs, "peptides_colored.xlsx")           # colored text, Excel
    open("peptides_colored.html", "w").write(write_html(seqs))  # highlights, HTML

Adding a coloring scheme
------------------------
A scheme is a dict with a "groups" list. Each group has a label, the residues
it covers and a hex color. Add it to SCHEMES and it appears in the app.
"""

import io
from pathlib import Path

import pandas as pd
import xlsxwriter

# --- Valid amino acid codes ---
VALID_AA = set("ACDEFGHIKLMNPQRSTVWY")

# --- Common header names to exclude from peptide lists ---
KNOWN_HEADERS = {"seq", "sequence", "peptide", "peptides", "pep", "protein"}

# --- Coloring schemes ---
SCHEMES = {
    "Residue class": {
        "groups": [
            {"label": "Aliphatic hydrophobic", "residues": "ILVMA",  "color": "#D62728"},  # red
            {"label": "Aromatic hydrophobic",  "residues": "WFY",    "color": "#C71585"},  # violet red
            {"label": "Basic",                 "residues": "KRH",    "color": "#E6B800"},  # yellow (darkened for legibility as text)
            {"label": "Acidic",                "residues": "ED",     "color": "#2CA02C"},  # green
            {"label": "Proline",               "residues": "P",      "color": "#000000"},  # black
            {"label": "Polar / small",         "residues": "QNSTCG", "color": "#1F5FD8"},  # blue
        ]
    },
}

DEFAULT_SCHEME = "Residue class"
FONT_NAME = "Courier New"   # monospace so residues line up down the column


def residue_colors(scheme_name: str = DEFAULT_SCHEME) -> dict:
    """Return {residue: hex color} for a scheme."""
    return {
        aa: g["color"]
        for g in SCHEMES[scheme_name]["groups"]
        for aa in g["residues"]
    }


def is_valid_sequence(value: str) -> bool:
    """True if value is a non-empty string of the 20 standard residues."""
    cleaned = str(value).strip().upper()
    return bool(cleaned) and cleaned.lower() not in KNOWN_HEADERS and set(cleaned) <= VALID_AA


def read_sequences(source) -> tuple[list[str], list[str]]:
    """
    Read sequences from the first column of a CSV (path or file-like object).
    Returns (valid_sequences, skipped_values). Header rows and blank cells
    are ignored silently; anything else that is not a valid sequence is
    returned in skipped_values so the user can see it.
    """
    df = pd.read_csv(source, header=None, dtype=str, skip_blank_lines=True)
    valid, skipped = [], []
    for raw in df.iloc[:, 0].dropna():
        s = str(raw).strip().upper()
        if not s or s.lower() in KNOWN_HEADERS:
            continue
        (valid if is_valid_sequence(s) else skipped).append(s)
    return valid, skipped


def _text_color(hex_color: str) -> str:
    """Black or white, whichever reads better on the given background."""
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return "#000000" if (0.299 * r + 0.587 * g + 0.114 * b) > 140 else "#FFFFFF"


def _runs(seq: str, colors: dict) -> list[tuple[str, str]]:
    """Split a sequence into (text, color) runs of identical color."""
    runs = []
    for aa in seq:
        c = colors[aa]
        if runs and runs[-1][1] == c:
            runs[-1] = (runs[-1][0] + aa, c)
        else:
            runs.append((aa, c))
    return runs


def write_excel(sequences, target, mode: str = "text", scheme_name: str = DEFAULT_SCHEME):
    """
    Write an .xlsx file with one row per peptide.

    mode="text":      column A plain sequence, column B the same sequence with
                      each residue's font colored (rich text, one cell).
    mode="highlight": column A plain sequence, then one residue per cell from
                      column B onward with the cell filled in the residue's color.
                      (Excel cannot highlight individual characters inside a cell.)

    target may be a file path or a writable binary buffer (e.g. io.BytesIO).
    """
    if mode not in ("text", "highlight"):
        raise ValueError("mode must be 'text' or 'highlight'")
    colors = residue_colors(scheme_name)

    wb = xlsxwriter.Workbook(target, {"in_memory": True})
    ws = wb.add_worksheet("PepColors")

    header = wb.add_format({"bold": True, "bottom": 1})
    plain = wb.add_format({"font_name": FONT_NAME})

    ws.write(0, 0, "Peptide", header)
    ws.set_column(0, 0, max([10] + [len(s) + 2 for s in sequences]))

    if mode == "text":
        ws.write(0, 1, "Colored", header)
        ws.set_column(1, 1, max([10] + [len(s) + 2 for s in sequences]))
        fonts = {c: wb.add_format({"font_name": FONT_NAME, "bold": True, "font_color": c})
                 for c in set(colors.values())}
        for r, seq in enumerate(sequences, start=1):
            ws.write_string(r, 0, seq, plain)
            runs = _runs(seq, colors)
            if len(runs) == 1:                       # write_rich_string needs 2+ fragments
                ws.write_string(r, 1, runs[0][0], fonts[runs[0][1]])
            else:
                parts = []
                for text, c in runs:
                    parts += [fonts[c], text]
                ws.write_rich_string(r, 1, *parts)

    else:  # highlight
        longest = max([1] + [len(s) for s in sequences])
        for i in range(longest):
            ws.write(0, 1 + i, i + 1, header)
        ws.set_column(1, longest, 2.6)
        fills = {c: wb.add_format({"font_name": FONT_NAME, "bold": True, "align": "center",
                                   "bg_color": c, "font_color": _text_color(c)})
                 for c in set(colors.values())}
        for r, seq in enumerate(sequences, start=1):
            ws.write_string(r, 0, seq, plain)
            for i, aa in enumerate(seq):
                ws.write_string(r, 1 + i, aa, fills[colors[aa]])

    ws.freeze_panes(1, 1)
    wb.close()


def excel_bytes(sequences, mode: str = "text", scheme_name: str = DEFAULT_SCHEME) -> bytes:
    """write_excel into memory and return the file contents."""
    buf = io.BytesIO()
    write_excel(sequences, buf, mode=mode, scheme_name=scheme_name)
    return buf.getvalue()


def html_preview(seq: str, mode: str = "text", scheme_name: str = DEFAULT_SCHEME) -> str:
    """HTML snippet showing one colored sequence (used by the Streamlit preview)."""
    colors = residue_colors(scheme_name)
    spans = []
    for aa in seq:
        c = colors[aa]
        if mode == "text":
            spans.append(f'<span style="color:{c};font-weight:bold">{aa}</span>')
        else:
            spans.append(f'<span style="background:{c};color:{_text_color(c)};'
                         f'font-weight:bold;padding:0 1px">{aa}</span>')
    return f'<span style="font-family:{FONT_NAME},monospace;font-size:1.05em">{"".join(spans)}</span>'


def write_html(sequences, title: str = "PepColors", scheme_name: str = DEFAULT_SCHEME) -> str:
    """
    Return a standalone HTML page with every sequence highlighted residue by
    residue, kept whole on one line. Runs of the same color share one span and
    colors live in CSS classes, so tens of thousands of peptides stay small.
    Copy-pasting from the browser into Word/PowerPoint keeps the highlighting.
    """
    from html import escape
    groups = SCHEMES[scheme_name]["groups"]
    cls = {}                       # color -> short class name
    css = []
    for i, g in enumerate(groups):
        c = g["color"]
        if c not in cls:
            cls[c] = f"c{i}"
            css.append(f".c{i}{{background:{c};color:{_text_color(c)}}}")
    colors = residue_colors(scheme_name)

    legend = "".join(
        f'<span class="key"><span class="seq {cls[g["color"]]}">{g["residues"]}</span> {escape(g["label"])}</span>'
        for g in groups
    )
    rows = []
    for n, seq in enumerate(sequences, start=1):
        body = "".join(f'<span class="{cls[c]}">{t}</span>' for t, c in _runs(seq, colors))
        rows.append(f"<tr><td>{n}</td><td class=seq>{body}</td></tr>")

    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{escape(title)}</title>
<style>
body{{font-family:system-ui,-apple-system,Segoe UI,sans-serif;margin:24px;color:#222;background:#fff}}
h1{{font-size:1.3em;margin:0 0 4px}} .meta{{color:#666;font-size:.9em;margin-bottom:12px}}
.legend{{display:flex;flex-wrap:wrap;gap:14px;margin-bottom:16px;font-size:.9em}}
.seq{{font-family:"Courier New",Consolas,monospace;font-weight:bold;letter-spacing:.5px;white-space:nowrap}}
.seq span{{padding:1px 0}}
table{{border-collapse:collapse}} td{{padding:2px 10px 2px 0;vertical-align:middle}}
td:first-child{{color:#999;text-align:right;font-size:.8em;user-select:none}}
{''.join(css)}
@media print{{.seq span{{-webkit-print-color-adjust:exact;print-color-adjust:exact}}}}
</style></head><body>
<h1>{escape(title)}</h1>
<div class="meta">{len(sequences)} peptides · scheme: {escape(scheme_name)}</div>
<div class="legend">{legend}</div>
<table>{''.join(rows)}</table>
</body></html>"""


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        sys.exit("usage: python pepcolors_core.py peptides.csv [text|highlight]")
    src = sys.argv[1]
    mode = sys.argv[2] if len(sys.argv) > 2 else "text"
    seqs, bad = read_sequences(src)
    if mode == "highlight":
        out = Path(src).with_name(Path(src).stem + "_PepColors.html")
        out.write_text(write_html(seqs, title=f"{Path(src).stem} — PepColors"), encoding="utf-8")
    else:
        out = Path(src).with_name(Path(src).stem + "_PepColors.xlsx")
        write_excel(seqs, str(out), mode="text")
    print(f"Wrote {len(seqs)} peptides to {out}" + (f"; skipped {len(bad)}: {', '.join(bad)}" if bad else ""))

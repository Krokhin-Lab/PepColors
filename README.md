# PepColors

Residue coloring for peptide sequences, to help spot patterns in peptides that deviate from retention predictions (e.g. [CHIPs](https://github.com/Krokhin-Lab/CHIPS)).

![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)
![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)

**Launch the interactive app: [PepColors on Streamlit](PASTE_APP_URL_HERE)**

**Authors:** Oleg V. Krokhin, Alexandre Préfontaine  
**Affiliation:** Manitoba Centre for Proteomics and Systems Biology, University of Manitoba  
**License:** GNU General Public License v3.0 — see [LICENSE](LICENSE).

## What it does

Upload a CSV with peptide sequences in the first column and pick an output style:

| Output style | File | What you get |
|---|---|---|
| **Colored text** | Excel (`.xlsx`) | column A plain sequence, column B the same sequence in one cell with each residue's letter colored |
| **Colored highlights** | HTML (`.html`) | every sequence kept whole on one line with each residue highlighted; opens in any browser, and copy-paste into Word or PowerPoint keeps the highlighting |

Excel cannot highlight individual characters inside a cell (only font color can vary within a cell), which is why highlights come as an HTML page.

A header row (`sequence`, `peptide`, …) and blank rows are ignored. Entries containing anything other than the 20 standard one-letter codes are skipped and listed in the app.

## Color key

| Residues | Class | Color |
|---|---|---|
| I L V M A | Aliphatic hydrophobic | Red |
| W F Y | Aromatic hydrophobic | Violet red |
| K R H | Basic | Yellow |
| E D | Acidic | Green |
| P | Proline | Black |
| Q N S T C G | Polar / small | Blue |

Yellow is slightly darkened (`#E6B800`) so it stays readable as text on a white background.

## Running locally

```bash
pip install -r requirements.txt
streamlit run pepcolors_app.py
```

Command line, without the app:

```bash
python pepcolors_core.py peptides.csv            # colored text  -> peptides_PepColors.xlsx
python pepcolors_core.py peptides.csv highlight  # highlights    -> peptides_PepColors.html
```

## Adding a coloring scheme

Schemes live in `SCHEMES` in `pepcolors_core.py`. Each one is a list of groups (label, residues, hex color); every one of the 20 residues must appear in exactly one group. A new entry shows up automatically in the app's scheme selector and legend.

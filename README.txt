# Naming Poetics in Middle High German Epic

This project provides an interactive analysis tool for examining the use of names and naming variants (proper names, antonomasia, and epithets) in Middle High German epic literature. It was developed in the context of the doctoral dissertation **"Naming Poetics in Middle High German Epic"** (Linda Beutel, University of Salzburg).

The goal is to use an XML-TEI encoded corpus and an Excel worksheet to:

- automatically detect new name variants,
- interactively verify and categorize them,
- record their collocational context,
- and export and analyze the structured results.

---

## 🔧 Requirements & Installation

- **Python** ≥ 3.10 required
- Recommended: virtual environment (e.g., `venv`, `conda`)

Install the required dependencies:

```bash
pip install -r requirements.txt
```

*(Note: This file must be created manually if not present — include e.g., `pandas`, `openpyxl`, `plotly` etc.)*

---

## ▶️ Usage

Run the project via the `run.py` entry point. It guides users through an interactive configuration, including optional Excel and TEI file loading:

```bash
python run.py
```

During execution, users can choose between **data collection** and **analysis mode**. Configuration settings are saved per book automatically.

---

## 📁 Project Structure

```bash
.
├── run.py                  # Main script to start the application
├── naming_analysis/        # Central Python package containing:
│   ├── __init__.py             # Package initialization
│   ├── config.py              # Configuration & user prompts
│   ├── controller.py          # Execution control flow
│   ├── analysis.py            # Analysis functionality (wordlists, keywords, visualizations)
│   ├── collection.py          # Interactive data collection & annotation
│   ├── savers.py              # Functions for saving JSON data
│   ├── loaders.py             # Loaders for Excel and TEI data
│   ├── io_utils.py            # JSON read/write utilities
│   ├── tei_utils.py           # TEI-specific XML processing functions
│   ├── shared.py              # Common helpers (e.g., text normalization)
│   ├── types.py               # Type definitions
│   ├── exporter.py            # Export functionality (Excel)
│   ├── validation.py          # Excel column validation
│   └── project_setup.py       # Directory and session setup logic
├── data/
│   ├── template_excel.xlsx           # Excel template, located alongside run.py
│   ├── naming_variants_dict.json     # Predefined naming patterns per book
│   ├── lemma_normalization.json      # Normalization rules for lemma variants
│   ├── ignored_lemmas.json           # Lemmata to ignore (e.g., function words)
│   ├── lemma_categories.json         # Classification: 'a' = name, 'e' = epithet
└── ...
```

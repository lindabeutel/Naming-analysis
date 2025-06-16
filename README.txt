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

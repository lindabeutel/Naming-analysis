# Naming Analysis

## Semi-Automatic Collection and Analysis of Naming Variants in Middle High German Epic

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.21914260.svg)](https://doi.org/10.5281/zenodo.21914260)

------------------------------------------------------------------------

## Visualizations and Tables

Interactive visualizations and tables accompanying the dissertation are available at:
👉 **[lindabeutel.github.io/Naming-analysis](https://lindabeutel.github.io/Naming-analysis/)**

------------------------------------------------------------------------

## Abstract

This repository provides a Python-based research infrastructure for the
semi-automatic collection, normalization, categorization, and analysis
of naming variants in Middle High German texts.

The tool supports the systematic investigation of proper names,
antonomasia, epithets, and figure-specific collocations based on
TEI-encoded texts. Automatic detection is rule-based and consistently
complemented by manual validation.

The program serves as a methodological research tool for structured data
collection and analysis within literary studies.

**Status:** BETA release

------------------------------------------------------------------------

## Project Context

Developed within the dissertation project:

"Naming poetics of Middle High German epic. An analysis of the use of
names, antonomasia, and epithets in selected texts of the
German-speaking Middle Ages."\
(Benennungspoetiken mittelhochdeutscher Epik -- Eine Analyse der
Verwendung von Namen, Antonomasien und Epitheta in ausgewählten Texten
des deutschsprachigen Mittelalters, working title)

PhD project since 2022\
University of Salzburg\
Department of German Studies

The dissertation itself is primarily literary-hermeneutic; this tool
functions as a structured research infrastructure supporting data
collection and analysis.

------------------------------------------------------------------------

## Research Context and Prior Work

Parts of the dataset were supplemented by existing annotations,
particularly work conducted by Nora Ketschik (MHG4SNA), created within
the CRETA/CretAnno environment.

The present infrastructure constitutes an independent tool that extends,
systematizes, and further processes previously automated annotations for
the specific research focus on naming practices.

The CretAnno project is currently no longer active.

------------------------------------------------------------------------

## Terminology

### Naming Variant

A naming variant refers to any lexical form used to refer to a character
within the narrative context. The term includes both proper names and
antonomasia.

In the German data model of this project, the term "Bezeichnung" is used
synonymously with naming variant.

### Antonomasia

Antonomasia is understood as the replacement of a proper name by a
characterizing descriptive expression (cf. Schirren 2009; Drews 2013).

### Epithet

Epithets are understood as attributive additions to a designation
(cf. Gondos 2013).

### Instance Types

Values in the field "Nennende Figur" may carry a marker indicating what
kind of naming instance they are: a collective, an unnamed role figure,
a non-personal source and so on. Unmarked values denote individuals. The
markers, their patterns and the conventions behind them are documented in
`data/instance_types.json`.

------------------------------------------------------------------------

## Language Layers in the System

The system operates across different language layers:

-   CLI interface: English
-   Data structure (Excel, JSON): German
-   Analytical outputs (CSV, visualizations): English

This distinction emerged historically during development and has been
retained for consistency.

------------------------------------------------------------------------

## Data Basis

The repository contains:

-   curated JSON files
-   a single Excel template (`template_excel.xlsx`)
-   normalization resources
-   a naming dictionary

Not included:

-   TEI files from the Middle High German Conceptual Database (MHDBDB)

TEI files should preferably be obtained from the MHDBDB, as the
collection logic is tailored to its TEI structure.

Depending on the specific encoding, adjustments to verse counting 
may be necessary. For example, in works such as _Parzival_, verse 
numbering restarts within each book (e.g., every 29 verses the attribute 
`.//tei:l[@n]` begins again at 1). In such cases, a more robust 
identification via `xml:id` should be used to ensure consistent verse 
tracking across structural divisions.

The datasets contained in the data/ directory are sufficient to 
reproduce all analytical and visualization outputs. No TEI source 
files are required for analysis execution. The dataset is evolving
dynamically during research.

------------------------------------------------------------------------

## System Architecture

    run.py
    ├── controller.py
    ├── Foundation Layer
    │   ├── project_setup.py
    │   ├── config.py
    │   └── project_types.py
    ├── Infrastructure Layer
    │   ├── tei_utils.py
    │   ├── io_utils.py
    │   └── shared.py
    ├── Collection Layer
    │   ├── collection.py
    │   ├── validation.py
    │   └── savers.py
    ├── Analysis Layer
    │   └── analysis.py
    └── Export Layer
        └── exporter.py

------------------------------------------------------------------------

## General Workflow

1.  Initialize a work project for a specific text
2.  Select a TEI file
3.  Create or load an Excel working file
4.  Perform verse-wise collection (naming variants, collocations,
    categorization)
5.  Persist results in JSON files
6.  Conduct analytical processing
7.  Export results (CSV, HTML visualizations)

------------------------------------------------------------------------

## Installation

Recommended: Python 3.13 in a virtual environment.

    python -m venv venv
    venv\Scripts\activate
    pip install -r requirements.txt
    python run.py

The program was developed and tested exclusively under Windows.

Execution should be performed from the project root directory.

------------------------------------------------------------------------

## Minimal Startup Workflow

1.  Download a TEI file (preferably from the MHDBDB).
2.  Run `run.py`.
3.  Enter the name of the work.
4.  Optionally extend the naming dictionary.
5.  Create a new Excel file from the provided template.
6.  Select the TEI file.
7.  Start "Collect new data."
8.  Select collection modes.

------------------------------------------------------------------------

## Progress Mechanism

For each work, a `progress_<work>.json` file is created.

This file records the last processed verse per module.

The progress file can be manually edited to reset verse counters.

During each iteration, existing entries in result JSON files
(`missing_naming_variants.json`, `categorization.json`,
`collocations.json`) are checked to prevent duplicates.

**Important:** Manual edits must be saved before restarting a collection
loop; otherwise, they may be overwritten.

------------------------------------------------------------------------

## Matching Logic

Automatic detection relies on:

-   regex-based word-boundary matching
-   normalization (lowercasing, graphemic harmonization)
-   comparison against a naming dictionary
-   string-based plausibility checks (substring and token-set
    comparisons)

Normalization is primarily informed by:

-   Mittelhochdeutsches Handwörterbuch by Matthias Lexer
-   frequent variants documented in the MHDBDB

All automatically detected entries require manual confirmation.

------------------------------------------------------------------------

## Methodological Framework

The workflow is entirely rule-based and dictionary-driven.

No machine learning models are implemented or trained.

The project does not aim at computational innovation but at providing a
transparent, research-oriented infrastructure for philological analysis.

------------------------------------------------------------------------

## Analytical Functions

The analysis module includes:

-   Wordlists
-   Keyword analysis (inspired by AntConc; Anthony 2024)
-   Figure profiles
-   Collocation analysis
-   Plotly-based visualizations

------------------------------------------------------------------------

## Generalisability

While this repository contains curated Middle High German data, the
underlying workflow is in principle language-independent and adaptable
to other TEI-based corpora.

------------------------------------------------------------------------

## Citation

Beutel-Thurow, L. (2026). Naming-analysis (Version v0.2.1-beta) [Computer software]. https://doi.org/10.5281/zenodo.21914260

------------------------------------------------------------------------

## License

Code, JSON files, the Excel template, and generated visualizations are
licensed under:

CC BY-NC-SA 4.0

TEI files from the MHDBDB are not part of this repository and remain
subject to their respective license (also CC BY-NC-SA 4.0).

------------------------------------------------------------------------

## FAIR Principles

The project aims to support:

-   transparency of data processing
-   reproducibility of analysis
-   disclosure of normalization and matching logic
-   versioning via Zenodo

------------------------------------------------------------------------

## Roadmap

Planned for the main release:

-   implementation of a centralized validation layer

------------------------------------------------------------------------

## References

Anthony, Laurence. 2024. AntConc (Version 4.3.1).\
Drews, Lydia. 2013. "Antonomasie."\
Gondos, Lisa. 2013. "Epitheton."\
Institut für maschinelle Sprachverarbeitung. 2022. "Tools und Demos \|
CRETA."\
Ketschik, Nora. 2023. MHG4SNA.\
Lexer, Matthias. Mittelhochdeutsches Handwörterbuch.\
Middle High German Conceptual Database (MHDBDB). University of
Salzburg.\
Schirren, Thomas. 2009. "Tropes in Classical Rhetoric."\
Schmid, Helmut. 2019. "Deep Learning-Based Morphological Taggers and
Lemmatizers."

# Doc ID Parser

A command-line Python tool for extracting document IDs from `.docx` and `.txt` files using customizable regex-based prefix matching.

> **Document IDs** refer to unique identifiers assigned to documents, commonly used in legal and litigation contexts.

---

## Features

- Scans single files or entire folders & sub-folders
- Supports `.docx` and `.txt` file formats
- Allows flexible matching for whitespace, line breaks, suffixes, and poor OCR
- Generates two output reports:
  - `Results.csv`: all matching document IDs with source info
  - `Hit_Report.csv`: summary of prefix counts by source file
- Retains the order of document IDs by which they appear in each source file

---

## Files Included

- `docid_parser.py` – Main script
- `README.md` – This file
- `pyproject.toml` – Poetry project configuration
- `poetry.lock` – Exact dependency versions
- `.gitignore` - Git ignore file
- `documentation.docx` - In-depth documentation on functionality and use

---

## Requirements

- Python 3.10+
- [Poetry](https://python-poetry.org/docs/#installation)

---

## Quick Start Guide

### 1. Clone or Download the Repository

*Option A: Using Git

```bash
git clone https://github.com/luca-kk/docid_parser.git
cd docid_parser
```

*Option B: Manual Download

- Download the Zip from GitHub repo
- Extract it
- Open a terminal inside the extracted folder

### 2. Install Dependencies with Poetry and Run

Located inside the downloaded repository, run:

```bash
poetry install --no-root
poetry run python docid_parser.py
```

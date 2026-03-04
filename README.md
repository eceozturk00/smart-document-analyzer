# Smart Document Analyzer

A Python tool that analyzes **Word (.docx)** and **PDF (.pdf)** documents, extracts keywords, and exports results to **Excel** and **JSON**.

## Features
- Supports **DOCX** and **PDF**
- Detects headings/sections (best-effort)
- Extracts **top keywords** using basic NLP
- Exports:
  - `output.xlsx` (structured table)
  - `output.json` (chunks + stats + keywords)

## Install
```bash
python analyzer.py --input input.docx --excel output.xlsx --json output.json

<p align="center">
  <img src="https://img.shields.io/badge/SCQF-Level%206-2F5496?style=for-the-badge&logo=education&logoColor=white" alt="SCQF Level 6"/>
  <img src="https://img.shields.io/badge/Units-5%20of%205%20Complete-28a745?style=for-the-badge&logo=checkmarx&logoColor=white" alt="5/5 Complete"/>
  <img src="https://img.shields.io/badge/Built%20With-Python%20%2B%20Office%20COM-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python + Office COM"/>
</p>

# SCQF Level 6 — Foundation Diploma in Business & IT

> Full coursework repository for the **JAIN College SCQF Level 6 Foundation Diploma** — 5 units covering Business Management, Finance, HR, IT Skills, and Contemporary Business Issues.

---

## Table of Contents

- [Overview](#overview)
- [Units](#units)
- [Repository Structure](#repository-structure)
- [Final Submissions](#final-submissions)
- [Support Pack](#support-pack)
- [Tech Stack](#tech-stack)
- [Build Pipeline](#build-pipeline)
- [Scripts Reference](#scripts-reference)
- [Licence](#licence)

---

## Overview

This repository contains **all coursework, build scripts, and supporting materials** for completing the SCQF Level 6 Foundation Diploma. The entire submission pipeline — from content generation through Office document creation to final PDF export — is automated using Python with Microsoft Office COM automation.

| Attribute | Detail |
|-----------|--------|
| **Programme** | SCQF Level 6 Foundation Diploma in Business & IT |
| **Institution** | JAIN College (via KareTech Solutions) |
| **Batch** | 2025 |
| **Units** | 5 (3 written reports + 2 IT practical) |
| **Total Marks** | 440 across all units |

---

## Units

| Code | Unit Title | Marks | Type | Status |
|------|-----------|-------|------|--------|
| `J229 76` | Understanding Business | 100 | Written Report → PDF | ✅ Complete |
| `F1FJ 12` | Spreadsheets and Databases | 70 | Excel + Access + Report → PDF | ✅ Complete |
| `F1FE 12` | Word Processing and Presenting | 70 | Word + PowerPoint + Report → PDF | ✅ Complete |
| `J22A 76` | Management of People and Finance | 100 | Written Report → PDF | ✅ Complete |
| `HE9E 46` | Contemporary Business Issues | 100 | Written Report → PDF | ✅ Complete |

---

## Repository Structure

```
SCQF-L6-Course-Materials/
│
├── 📄 README.md                          # This file
├── 📄 .gitignore
│
├── 📁 Assessment_Papers/                  # Official assessment briefs (PDF)
│   ├── Assessment_F1FE_12_Word_Processing.pdf
│   ├── Assessment_F1FJ_12_Spreadsheet_Database.pdf
│   ├── HE9E 46_Contemporary Business Issues.pdf
│   ├── J229 76_Understanding Business._Assessment.pdf
│   ├── J22A 76_Management of People and Finance_Assessment.pdf
│   └── Reference Sample Assessment SCQF Level 7.pdf
│
├── 📁 SCQF_L6_FINAL_SUBMISSION/           # ⭐ FINAL OUTPUT
│   ├── pdf_submissions/                   # 5 submission-ready PDFs
│   │   ├── J22976_Understanding_Business.pdf
│   │   ├── F1FJ12_Spreadsheet_Database.pdf
│   │   ├── F1FE12_Word_Presentation.pdf
│   │   ├── J22A76_Management_People_Finance.pdf
│   │   └── HE9E46_Contemporary_Business_Issues.pdf
│   ├── source_files/                      # DOCX, XLSM, ACCDB, PPTX sources
│   │   ├── F1FJ12_Spreadsheet_Database/
│   │   ├── F1FE12_Word_Presentation/
│   │   ├── HE9E46_Contemporary_Business_Issues/
│   │   ├── J22976_Understanding_Business/
│   │   └── J22A76_Management_People_Finance/
│   └── evidence_screenshots/              # Excel/Access/PPT screenshots
│
├── 📁 SCQF_L6_SUPPORT_PACK/               # 📚 Student Reference Pack
│   ├── README_START_HERE.txt
│   ├── F1FJ12_Spreadsheet_Database/       # Sample data + feature guides
│   ├── F1FE12_Word_Presentation/          # Templates + mail merge guide
│   ├── J22976_Understanding_Business/     # Essay outline + table templates
│   ├── J22A76_Management_People_Finance/  # Structure template + Maslow diagram
│   └── HE9E46_Contemporary_Business_Issues/ # Lifecycle + strategy templates
│
└── 📁 scripts/                            # Python build & refinement scripts
    ├── build_unit_a.py                    # J229 76 report builder
    ├── build_unit_b_*.py                  # F1FJ 12 Excel/Access/Report builders
    ├── build_unit_c_*.py                  # F1FE 12 Word/PPT/Report builders
    ├── build_unit_d.py                    # J22A 76 report builder
    ├── build_unit_e.py                    # HE9E 46 report builder
    ├── refine_unit_*.py                   # Post-build refinement scripts
    ├── qa_check.py                        # Quality assurance validation
    └── build_support_*.py                 # Support pack generators
```

---

## Final Submissions

All 5 submission-ready PDFs in `SCQF_L6_FINAL_SUBMISSION/pdf_submissions/`:

| File | Size | Pages | Content |
|------|------|-------|---------|
| `J22976_Understanding_Business.pdf` | 375 KB | 21 | Comparative business analysis, PESTEC, stakeholders |
| `F1FJ12_Spreadsheet_Database.pdf` | 1,320 KB | 20 | Excel features + Access database + evidence screenshots |
| `F1FE12_Word_Presentation.pdf` | 337 KB | 19 | Conference pack + PowerPoint + evaluation |
| `J22A76_Management_People_Finance.pdf` | 248 KB | 26 | HRM, Maslow, industrial action, finance ratios |
| `HE9E46_Contemporary_Business_Issues.pdf` | 239 KB | 27 | SME analysis, lifecycle, business strategies |

Each PDF includes:
- Cover page with student details
- Declaration of Originality
- Automatic Table of Contents
- Harvard referencing (in-text + reference list)
- Mapping table linking tasks to marking criteria
- Consistent footer on every page

---

## Support Pack

The `SCQF_L6_SUPPORT_PACK/` folder contains **plagiarism-safe reference materials** for fellow SCQF Level 6 students:

| Resource | Description |
|----------|-------------|
| Sample datasets (`.xlsx`) | Generic sales and training data for practice |
| Blank templates (`.docx`) | Section headings + placeholder text only |
| Step-by-step guides (`.pdf`) | How to use Excel/Access features (no analysis) |
| Screenshot checklists (`.txt`) | What evidence to capture for each unit |
| Maslow blank diagram (`.png`) | Empty pyramid template for student use |
| PowerPoint template (`.pptx`) | 5-slide skeleton with layout instructions |

> **Important:** The support pack contains NO completed answers, NO evaluation paragraphs, and NO content that matches the actual submissions. All datasets use different dummy data. All templates contain only headings and `[placeholder]` text.

---

## Tech Stack

The entire build pipeline is automated with Python and Microsoft Office COM:

| Technology | Purpose |
|-----------|---------|
| ![Python](https://img.shields.io/badge/Python-3.13-3776AB?logo=python&logoColor=white) | Build scripts, automation orchestration |
| ![pywin32](https://img.shields.io/badge/pywin32-311-blue) | COM automation bridge to Office applications |
| ![python-docx](https://img.shields.io/badge/python--docx-1.2-blue) | DOCX content generation and manipulation |
| ![openpyxl](https://img.shields.io/badge/openpyxl-3.1-blue) | Excel workbook creation |
| ![matplotlib](https://img.shields.io/badge/matplotlib-3.10-blue) | Chart/diagram generation for evidence |
| ![Pillow](https://img.shields.io/badge/Pillow-11.3-blue) | Image processing for screenshots |
| ![Word COM](https://img.shields.io/badge/Word-16.0-2B579A?logo=microsoftword&logoColor=white) | TOC updates, mail merge, PDF export |
| ![Excel COM](https://img.shields.io/badge/Excel-16.0-217346?logo=microsoftexcel&logoColor=white) | Pivot tables, VBA macros, charts, slicers |
| ![Access COM](https://img.shields.io/badge/Access-16.0-A4373A?logo=microsoftaccess&logoColor=white) | Tables, relationships, forms, queries |
| ![PowerPoint COM](https://img.shields.io/badge/PowerPoint-16.0-B7472A?logo=microsoftpowerpoint&logoColor=white) | Slides, action buttons, kiosk mode |
| ![PyMuPDF](https://img.shields.io/badge/PyMuPDF-1.24-orange) | PDF page rendering for visual QA |

---

## Build Pipeline

```
┌──────────────────────────────────────────────────────────────┐
│                    BUILD PIPELINE                             │
├──────────────────────────────────────────────────────────────┤
│                                                              │
│  1. GENERATE     python scripts/build_unit_*.py              │
│     ↓            Creates DOCX/XLSM/ACCDB/PPTX source files  │
│                                                              │
│  2. SCREENSHOT   COM automation captures evidence PNGs       │
│     ↓            Excel formulas, Access forms, PPT slides    │
│                                                              │
│  3. ASSEMBLE     python-docx embeds screenshots + text       │
│     ↓            into structured report DOCX files           │
│                                                              │
│  4. REFINE       python scripts/refine_unit_*.py             │
│     ↓            AI risk reduction, tone adjustments,        │
│                  transition variation, marks removal          │
│                                                              │
│  5. QA           python scripts/qa_check.py                  │
│     ↓            Validates footers, TOC, declarations,       │
│                  references, mapping tables                   │
│                                                              │
│  6. EXPORT       Word COM → ExportAsFixedFormat(path, 17)    │
│                  Produces final submission PDFs               │
│                                                              │
└──────────────────────────────────────────────────────────────┘
```

---

## Scripts Reference

### Build Scripts
| Script | Unit | Purpose |
|--------|------|---------|
| `build_unit_a.py` | J229 76 | Generates Understanding Business report |
| `build_unit_b_excel.py` | F1FJ 12 | Creates Excel workbook with all features |
| `build_unit_b_access.py` | F1FJ 12 | Creates Access database with tables/forms/queries |
| `build_unit_b_report.py` | F1FJ 12 | Assembles report with embedded screenshots |
| `build_unit_c_word.py` | F1FE 12 | Creates conference Welcome Pack document |
| `build_unit_c_merge_and_ppt.py` | F1FE 12 | Mail merge + PowerPoint creation |
| `build_unit_c_report.py` | F1FE 12 | Assembles report with slide exports |
| `build_unit_d.py` | J22A 76 | Generates Management of People & Finance report |
| `build_unit_e.py` | HE9E 46 | Generates Contemporary Business Issues report |

### Refinement Scripts
| Script | Purpose |
|--------|---------|
| `refine_unit_a.py` | 13 changes — transitions, reflective phrases, evaluation |
| `refine_unit_d.py` | 40 changes — transitions, paragraph splits, ratio limits |
| `refine_unit_e.py` | 20 changes — transitions, reflective, lifecycle linking |
| `refine_units_bc.py` | 65 changes — IT unit language, naming, marks removal |

### Quality Assurance
| Script | Purpose |
|--------|---------|
| `qa_check.py` | Validates all 5 PDFs against submission checklist |

---

## Licence

This repository is for **educational reference only**. The assessment briefs are the property of SQA/JAIN College. The completed submissions are original student work and must not be copied or submitted by others.

---

<p align="center">
  <sub>Built with Python + Office COM automation | SCQF Level 6 Foundation Diploma | 2026</sub>
</p>

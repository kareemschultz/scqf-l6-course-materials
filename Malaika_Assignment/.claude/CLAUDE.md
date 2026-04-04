# Malaika_Assignment — Claude Code Config

## Student
- **Name**: Malaika Abdul-Jabar
- **ID**: 320122505
- **Email**: malaika.abduljabar@my.open.uwi.edu
- **Course**: MGMT268 — Human Resource Management (UWI Global Campus)

## Current Assignment
MGMT268 Assessment #1 Individual Discussion
- **Question**: "Why do we need the Behavioural Approach when managing Human Resources?"
- **Due**: Sunday, 5 April 2026, 11:59 PM
- **Word limit**: max 3,000 words | APA citations | Topics 1–4
- **Status**: COMPLETE — submit `FINAL_SUBMISSION/Malaika_MGMT268_Assessment1_FINAL.pdf`

## Folder Structure

```
Malaika_Assignment/
├── FINAL_SUBMISSION/         ← Submit from here
│   ├── Malaika_MGMT268_Assessment1_FINAL.pdf      ← SUBMIT THIS
│   ├── Malaika_MGMT268_Assessment1_FINAL.docx
│   ├── Malaika_MGMT268_Assessment1_WITH_VIDEO.pdf
│   └── Malaika_MGMT268_Assessment1_WITH_VIDEO.docx
├── 01_assignment/            ← Figures + Excalidraw source
├── 02_course_materials/      ← UWI scraped content
├── 03_research/
├── 04_animation/hr_animation/ ← Manim animation
└── 05_build_scripts/         ← All build scripts
```

## Deliverables
| File | Description | Status |
|------|-------------|--------|
| `FINAL_SUBMISSION/Malaika_MGMT268_Assessment1_FINAL.docx` | Illustration version | Done |
| `FINAL_SUBMISSION/Malaika_MGMT268_Assessment1_FINAL.pdf` | **Submit this** (713 KB) | Done |
| `FINAL_SUBMISSION/Malaika_MGMT268_Assessment1_WITH_VIDEO.docx` | With animation reference | Done |
| `FINAL_SUBMISSION/Malaika_MGMT268_Assessment1_WITH_VIDEO.pdf` | PDF of video version (713 KB) | Done |
| `04_animation/hr_animation/media/.../HRMDialogueScene.mp4` | Companion animation | Done |

## Figures in Document
| # | File (in `01_assignment/`) | Caption |
|---|---------------------------|---------|
| 1 | `hackman_oldham_model.png` | Hackman & Oldham's JCM (1976) |
| 2 | `comparison_table.png` | Scientific Management vs Behavioural Approach |
| 3 | `gns_diagram.png` | JCM with Growth Need Strength moderator |
| 4 | `hr_dialogue_illustration.png` | Excalidraw dialogue scene (Manager vs HR Manager) |
| 5 | `animation_frame.png` | Storyboard frame from companion animation |

**Image widths**: 5.9" max (text area = 6.0" — page 8.5", margins 1.25" each side)

## Rebuild Sequence
```bash
# From Malaika_Assignment/ directory:
python 05_build_scripts/make_diagrams_extra.py
python 05_build_scripts/generate_hr_dialogue.py
python 05_build_scripts/excalidraw_render/render_excalidraw.py \
    01_assignment/hr_dialogue.excalidraw \
    --output 01_assignment/hr_dialogue_illustration.png

python 05_build_scripts/build_essay_v2.py          # close Word first
python 05_build_scripts/build_essay_video.py

# PDF export:
powershell -ExecutionPolicy Bypass -File 05_build_scripts/convert_pdf.ps1
```

## PDF Export
Python win32com late-binding fails on this system. Use PowerShell COM via `05_build_scripts/convert_pdf.ps1` — reads/writes from `FINAL_SUBMISSION/`.

## Skills Available
- `.claude/skills/excalidraw-diagram/` — generate + render Excalidraw diagrams to PNG
- Parent repo `.claude/skills/` — same skills available at repo level

## Reusable SCQF Scripts (from `../scripts/`)
| Script | What to reuse |
|--------|--------------|
| `refine_unit_d/e.py` | `replace_in_paragraph()`, `append_to_paragraph()` — AI-detection reduction |
| `refine_units_bc.py` | `replace_in_runs()` with case-insensitive option |
| `build_unit_d.py` | `apply_footer()`, `add_body()`, `add_bullet()`, table styling helpers, PDF COM export |
| `qa_check.py` | `check_pdf_via_word()` validation pattern (adapt for APA essay structure) |

**Adaptation notes**: Change citation style Harvard → APA, line spacing → 2.0 for APA, footer text for Malaika's info.

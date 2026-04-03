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

## Deliverables
| File | Description | Status |
|------|-------------|--------|
| `Malaika_MGMT268_Assessment1_FINAL.docx` | Illustration version (1.4 MB) | Done |
| `Malaika_MGMT268_Assessment1_FINAL.pdf` | **Submit this** (736 KB) | Done |
| `Malaika_MGMT268_Assessment1_WITH_VIDEO.docx` | With animation reference | Done |
| `Malaika_MGMT268_Assessment1_WITH_VIDEO.pdf` | PDF of video version | Done |
| `hr_animation/media/.../HRMDialogueScene.mp4` | Companion animation 9.3 MB | Done |

## Figures in Document
- Figure 1: `hackman_oldham_model.png` — Hackman-Oldham JCM
- Figure 2: `comparison_table.png` — Rational vs Behavioural comparison
- Figure 3: `gns_diagram.png` — JCM with Growth Need Strength moderator
- Figure 4: `hr_dialogue_illustration.png` — Excalidraw dialogue scene
- Figure 5: `animation_frame.png` — Storyboard frame from animation

## Skills Available
- `.claude/skills/excalidraw-diagram/` — generate + render Excalidraw diagrams
- Parent folder `.claude/skills/` — same skills shared at repo level

## Rebuild Sequence
```bash
python make_diagram.py                    # Figure 1
python make_diagrams_extra.py             # Figures 2, 3, 5
python generate_hr_dialogue.py            # hr_dialogue.excalidraw
python excalidraw_render/render_excalidraw.py hr_dialogue.excalidraw --output hr_dialogue_illustration.png
python build_essay_v2.py                  # Word doc (close Word first)
# PDF via PowerShell:
powershell -ExecutionPolicy Bypass -File convert_pdf.ps1
```

# SCQF L6 Course Materials — Claude Code Config

## Student
- **Name**: Kareem Nurw Jason Schultz
- **ID**: 252IFCBR0596
- **Programme**: JAIN College SCQF Level 6 Foundation Diploma in Business & IT

## Repository Layout
```
SCQF-L6-Course-Materials/
├── SCQF_L6_FINAL_SUBMISSION/   # Final PDFs for 5 units (COMPLETE)
├── SCQF_L6_SUPPORT_PACK/       # Student reference materials
├── Malaika_Assignment/         # MGMT268 guest assignment (Malaika Abdul-Jabar)
└── .claude/                    # This config + shared skills
    ├── CLAUDE.md               # This file
    ├── skills/
    │   └── excalidraw-diagram/ # Excalidraw diagram → PNG skill
    └── plugins/                # Claude Code plugin configs
```

## Shared Skills

### excalidraw-diagram
Generates `.excalidraw` JSON files and renders them to PNG using Playwright.
- **Invoke**: ask Claude to "create an Excalidraw diagram of ..."
- **Render**: `python .claude/skills/excalidraw-diagram/references/render_excalidraw.py <file.excalidraw> --output <out.png>`
- See `.claude/skills/excalidraw-diagram/SKILL.md` for full instructions

## Tech Stack
- Python 3.13+ | pywin32 311 | python-docx 1.2.0 | matplotlib 3.10.6
- Playwright (Node + Python) for browser automation and diagram rendering
- Microsoft Office 16.0 COM automation for PDF export
- Excalidraw via esm.sh for hand-drawn diagram rendering

## PDF Export (use PowerShell — COM works best this way)
```powershell
powershell -ExecutionPolicy Bypass -Command "
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $doc = $word.Documents.Open('C:\full\path\to\file.docx', $false, $true)
  $doc.SaveAs([ref]'C:\full\path\to\file.pdf', [ref]17)
  $doc.Close([ref]$false)
  $word.Quit()
"
```

## Footer Format (SCQF units)
`252IFCBR0596 | Kareem Nurw Jason Schultz | [Unit Code] | [Unit Title] | Page X`

---
name: excalidraw-diagram
description: Generate hand-drawn style Excalidraw diagrams as .excalidraw files and render them to PNG for embedding in Word/PDF documents. Use for illustrations, concept maps, dialogue scenes, flow diagrams, and academic figures.
---

# Excalidraw Diagram Generator

Generates `.excalidraw` JSON files and renders them to PNG using Playwright.

## Workflow

1. Design the diagram elements as a Python generator script
2. Output a valid `.excalidraw` JSON file (type, version, elements, appState, files)
3. Render to PNG: `python .claude/skills/excalidraw-diagram/references/render_excalidraw.py <file.excalidraw> --output <out.png>`
4. Embed PNG in Word document via python-docx

## Critical Rules (raw .excalidraw format)

- **NO `label` property** on shapes — use two-element labeling:
  - Shape: `"boundElements": [{"type": "text", "id": "t1"}]`
  - Text: `"containerId": "shape_id"`, `"textAlign": "center"`, `"verticalAlignment": "middle"`
- Use `"roughness": 1` for hand-drawn look, `0` for clean edges
- Use `"roundness": {"type": 3}` for rounded rectangles
- Arrow `points` are `[[0,0],[dx,dy]]` relative to element x,y

## Color Palette

| Role | Background | Stroke |
|------|-----------|--------|
| Manager/Blue | `#dbe4ff` | `#4a9eed` |
| HR/Green | `#d3f9d8` | `#22c55e` |
| Highlight | `#fff3bf` | `#f59e0b` |
| Header | `#2C3E50` | `#2C3E50` |

## Render Setup

```bash
# First time only — Playwright Chromium is already installed
cd .claude/skills/excalidraw-diagram/references
python render_excalidraw.py ../../../my_diagram.excalidraw --output ../../../my_diagram.png
```

## Example Usage

```
Create an Excalidraw dialogue illustration showing the Manager vs HR Manager debate
about Rational vs Behavioural approaches in HRM, referencing Hackman-Oldham and
the Hawthorne Studies. Save as hr_dialogue.excalidraw and render to PNG.
```

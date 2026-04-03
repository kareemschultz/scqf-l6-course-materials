# HRM Dialogue Animation

An animated educational video illustrating the difference between the
**Rational Approach** and the **Behavioral Approach** to Human Resource
Management (HRM). Two characters — a Manager and an HR Manager — have
an 11-line dialogue, punctuated by a mid-video comparison chart.

Built with **Manim Community Edition** and **manim-voiceover + gTTS** —
no paid API keys required.

---

## Video Structure

| Segment | Duration (approx.) |
|---|---|
| Intro — title + characters walk in | ~10 s |
| Dialogue lines 1–6 | ~60 s |
| Midpoint T-chart (Rational vs Behavioral) | ~20 s |
| Dialogue lines 7–11 | ~55 s |
| Outro — conclusion + End of Lesson | ~20 s |
| **Total** | **~3–5 min** |

---

## Prerequisites

### System dependencies

| Tool | Install |
|---|---|
| Python 3.10+ | [python.org](https://python.org) |
| ffmpeg | `sudo apt install ffmpeg` (Linux) / `brew install ffmpeg` (macOS) |
| LaTeX (optional) | `sudo apt install texlive-full` — only needed if you add LaTeX math |

Verify ffmpeg is on your PATH:
```bash
ffmpeg -version
```

### Python packages

```bash
pip install -r requirements.txt
```

This installs:
- `manim>=0.18.0` — animation engine
- `manim-voiceover[gtts]>=0.3.7` — TTS integration with gTTS

> **Internet required on first render** — gTTS downloads audio via Google
> Translate. Subsequent renders use cached files in `media/voiceovers/`.

---

## Rendering

### One-command render (recommended)

```bash
bash render.sh          # 1080p60 — production quality
bash render.sh low      # 480p15  — fast preview (~3x faster)
bash render.sh medium   # 720p30  — medium quality
```

### Manual render commands

```bash
# High quality (1920×1080, 60 fps)
manim -pqh main.py HRMDialogueScene

# Low quality preview (854×480, 15 fps) — much faster
manim -pql main.py HRMDialogueScene

# Without auto-play
manim -qh main.py HRMDialogueScene
```

Output is saved to `media/videos/main/<resolution>/HRMDialogueScene.mp4`.

---

## Project Structure

```
hr_animation/
├── main.py          # Scene logic — HRMDialogueScene class
├── characters.py    # CorporateCharacter VMobject + speech bubble helper
├── requirements.txt # Python dependencies
├── render.sh        # One-command render script
└── README.md        # This file
```

---

## Customization Guide

### Swap TTS voices

Each character uses a regional gTTS accent. Change them in `main.py`:

```python
def _uk_service():
    return GTTSService(lang="en", tld="co.uk")   # British accent

def _us_service():
    return GTTSService(lang="en", tld="com")      # US accent
```

Other `tld` options: `"com.au"` (Australian), `"co.in"` (Indian), `"ca"` (Canadian).

To use a completely different language:
```python
GTTSService(lang="fr", tld="fr")   # French
GTTSService(lang="de", tld="de")   # German
```

### Change character colors

In `main.py`, edit the constants near the top:

```python
MANAGER_COLOR = DARK_BLUE   # Any Manim color or hex string, e.g. "#2E4057"
HR_COLOR = TEAL             # e.g. "#1B998B"
```

### Add more dialogue lines

In `_dialogue_section_two()` (or create a new section method), call:

```python
self.set_speech_service(_uk_service())   # or _us_service()
self.speak_line(
    character=self.manager,           # or self.hr_manager
    voiceover_text="Your dialogue here.",
    bubble_side="right",              # "right" for manager, "left" for HR manager
    bubble_anchor_x=0.0,
)
```

### Adjust speech bubble font size

In `speak_line()`, change `font_size=19`:

```python
bubble = create_speech_bubble(
    text=voiceover_text,
    side=bubble_side,
    font_size=19,       # ← increase for larger text
    ...
)
```

### Change background color

```python
BG_COLOR = "#F0F0F0"   # Light grey — change to any hex color
```

---

## Troubleshooting

### `ModuleNotFoundError: No module named 'manim'`
```bash
pip install manim
```

### `ModuleNotFoundError: No module named 'manim_voiceover'`
```bash
pip install "manim-voiceover[gtts]"
```

### `FileNotFoundError: ffmpeg not found`
Install ffmpeg and ensure it is on your system PATH:
```bash
# Linux (Debian/Ubuntu)
sudo apt-get install ffmpeg

# macOS
brew install ffmpeg

# Windows — download from https://ffmpeg.org/download.html
# and add the bin/ folder to your PATH
```

### gTTS network error during render
```
gtts.tts.gTTSError: Failed to connect
```
You need an internet connection for the first render. Check your connection, then re-run. Audio files are cached so the error won't repeat for already-downloaded lines.

### Render is very slow
Use low quality for development iterations:
```bash
bash render.sh low
```
Only switch to `high` for the final production render.

### `cairo` / `pango` error on Linux
```bash
sudo apt-get install libcairo2-dev libpango1.0-dev
```

### Speech bubble text overflows
Decrease `font_size` in `speak_line()` or shorten the dialogue text in the script. The `wrap_text(text, width=38)` call in `characters.py` controls line length.

---

## License

Educational use only. gTTS relies on Google Translate's public TTS endpoint.

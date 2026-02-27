"""Create a basic 5-slide presentation template for the support pack."""
import subprocess, time, os

subprocess.run(['taskkill', '/F', '/IM', 'POWERPNT.EXE'], capture_output=True)
time.sleep(1)

import pythoncom
import win32com.client

BASE = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    'SCQF_L6_SUPPORT_PACK', 'F1FE12_Word_Presentation'
)
OUT_PATH = os.path.join(BASE, 'Basic_Presentation_Template.pptx')

pythoncom.CoInitialize()
ppt = win32com.client.Dispatch('PowerPoint.Application')
ppt.Visible = True

pres = ppt.Presentations.Add()

# Helper to add a text box
def add_textbox(slide, left, top, width, height, text, size=14, bold=False, color=None):
    tb = slide.Shapes.AddTextbox(1, left, top, width, height)  # 1 = msoTextOrientationHorizontal
    tb.TextFrame.TextRange.Text = text
    tb.TextFrame.TextRange.Font.Size = size
    tb.TextFrame.TextRange.Font.Bold = bold
    if color:
        tb.TextFrame.TextRange.Font.Color.RGB = color
    tb.TextFrame.WordWrap = True
    return tb

# Measurements in points (1 inch = 72pt)
W = 720  # slide width
H = 540  # slide height

# ── Slide 1: Title/Welcome ──
s1 = pres.Slides.Add(1, 12)  # ppLayoutBlank
add_textbox(s1, 50, 60, 620, 60, '[Your Conference Name]', size=32, bold=True, color=0x964F2F)
add_textbox(s1, 50, 140, 620, 40, '[Subtitle / Tagline]', size=20, color=0x666666)
add_textbox(s1, 50, 220, 620, 40, '[Venue] | [Date]', size=16, color=0x888888)
add_textbox(s1, 50, 320, 620, 80,
    '[Add action buttons here for: Agenda, Venue Map, Wi-Fi Info]\n'
    '[Insert audio file set to auto-play on this slide]',
    size=12, color=0xAAAAAA)
add_textbox(s1, 50, 440, 620, 40,
    'TEMPLATE NOTE: Add a Home button on every slide for navigation.',
    size=10, color=0xCC0000)

# ── Slide 2: Conference Highlights ──
s2 = pres.Slides.Add(2, 12)
add_textbox(s2, 50, 30, 620, 40, 'Conference Highlights', size=28, bold=True)
add_textbox(s2, 50, 90, 620, 300,
    '[Insert a video file here]\n\n'
    'Video settings:\n'
    '- Playback > Start: Automatically\n'
    '- Optionally trim to show key moments\n\n'
    '[Add a Home action button in the bottom-right corner]',
    size=14, color=0x666666)

# ── Slide 3: Agenda ──
s3 = pres.Slides.Add(3, 12)
add_textbox(s3, 50, 30, 620, 40, 'Conference Agenda', size=28, bold=True)
add_textbox(s3, 50, 90, 620, 350,
    '[Paste your agenda content here]\n\n'
    'Time     Session\n'
    '09:00    [Session 1]\n'
    '10:00    [Session 2]\n'
    '11:00    [Break]\n'
    '11:30    [Session 3]\n'
    '12:30    [Lunch]\n'
    '13:30    [Session 4]\n'
    '15:00    [Closing]\n\n'
    '[Add a Home action button]',
    size=13, color=0x666666)

# ── Slide 4: Venue / Wi-Fi ──
s4 = pres.Slides.Add(4, 12)
add_textbox(s4, 50, 30, 620, 40, 'Venue & Wi-Fi Information', size=28, bold=True)
add_textbox(s4, 50, 90, 620, 200,
    '[Insert venue details, directions, parking info]\n\n'
    'Wi-Fi Network: [network name]\n'
    'Password: [password]\n\n'
    '[Add a Home action button]',
    size=14, color=0x666666)

# ── Slide 5: Thank You / Copyright ──
s5 = pres.Slides.Add(5, 12)
add_textbox(s5, 50, 60, 620, 60, 'Thank You', size=36, bold=True, color=0x964F2F)
add_textbox(s5, 50, 160, 620, 80,
    '[Your conference name] | [Date]\n'
    '[Organiser name / website]',
    size=16, color=0x666666)
add_textbox(s5, 50, 320, 620, 120,
    'Copyright Citation:\n'
    '[Include copyright attribution for any images, videos, or music used]\n\n'
    'Example: Background image by [Author] from [Source], licensed under '
    'Creative Commons CC BY 4.0.',
    size=11, color=0x888888)
add_textbox(s5, 50, 460, 620, 40,
    'TEMPLATE NOTE: Set kiosk mode via Slide Show > Set Up Slide Show > '
    'Browsed at a kiosk.',
    size=10, color=0xCC0000)

# Save
pres.SaveAs(os.path.abspath(OUT_PATH))
pres.Close()
ppt.Quit()
pythoncom.CoUninitialize()

print(f'Created {OUT_PATH}')

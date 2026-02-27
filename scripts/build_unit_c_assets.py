"""
UNIT C: F1FE 12 - Create source assets:
1. Conference_Text.txt
2. Delegate_List.xlsx
3. TechSummit_Logo.png (programmatic)
4. Audio file (WAV)
5. Video file (basic MP4 placeholder)
"""
import os
import struct
import wave
import array
import math
from PIL import Image, ImageDraw, ImageFont
from openpyxl import Workbook

BASE = r"C:\Users\admin\Documents\SCQF-L6-Course-Materials\SCQF_L6_FINAL_SUBMISSION"
SRC = os.path.join(BASE, "source_files", "F1FE12_Word_Presentation")
os.makedirs(SRC, exist_ok=True)


def create_conference_text():
    """Create Conference_Text.txt with welcome pack content"""
    text = """TechSummit 2026 - Digital Innovation Conference
Welcome Pack

Welcome to TechSummit 2026

We are delighted to welcome you to the TechSummit 2026 Digital Innovation Conference, hosted at the Edinburgh International Conference Centre (EICC). This year's conference brings together over 500 technology professionals, industry leaders, and innovators from across the United Kingdom and beyond.

About the Conference

TechSummit 2026 is now in its fifth year and has established itself as Scotland's premier technology conference. Our mission is to connect professionals, share cutting-edge knowledge, and inspire the next generation of digital innovation. This year's theme, "Transforming Tomorrow Through Technology," reflects our commitment to exploring how emerging technologies can create positive change across industries and communities.

Keynote Speakers

Our distinguished lineup of keynote speakers includes:

Dr Sarah Chen, Chief Technology Officer at InnovateTech Solutions, will discuss the future of artificial intelligence in healthcare. Professor James MacLeod from the University of Edinburgh will present his groundbreaking research on quantum computing applications. Fiona Stewart, Managing Director of Digital Scotland, will share insights on Scotland's digital transformation strategy. Raj Patel, Head of Innovation at GlobalTech Partners, will explore sustainable technology and green computing initiatives.

Conference Venues and Facilities

The main conference hall is located on Level 2 of the EICC. Breakout sessions will be held in the Morrison and Pentland Rooms on Level 3. The exhibition area, featuring over 40 technology vendors, is situated in the main atrium on the ground floor. Refreshments will be served in the Cromdale Suite throughout the day.

Networking Opportunities

We encourage all delegates to take advantage of the networking opportunities available. The Welcome Reception takes place on the evening of Day 1 in the Castle Suite, with stunning views of Edinburgh Castle. Speed networking sessions are scheduled during the lunch breaks on both days.

Important Information

Registration opens at 08:00 each morning in the main foyer. Please ensure you wear your delegate badge at all times for security purposes. Free Wi-Fi is available throughout the venue using the network name TECHSUMMIT2026 with the password Innovation2026. Emergency exits are clearly marked throughout the building. First aid stations are located at the reception desk on each floor.

We hope you enjoy TechSummit 2026 and find the conference both informative and inspiring.

The TechSummit Organising Committee
"""
    filepath = os.path.join(SRC, "Conference_Text.txt")
    with open(filepath, "w", encoding="utf-8") as f:
        f.write(text)
    print(f"[OK] Conference_Text.txt")


def create_delegate_list():
    """Create Delegate_List.xlsx with 10 delegates"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Delegates"

    headers = ["Delegate ID", "Title", "First Name", "Last Name", "Organisation", "Role", "Email"]
    ws.append(headers)

    delegates = [
        ["TS-001", "Dr", "Sarah", "Chen", "InnovateTech Solutions", "CTO", "s.chen@innovatetech.com"],
        ["TS-002", "Prof", "James", "MacLeod", "University of Edinburgh", "Professor", "j.macleod@ed.ac.uk"],
        ["TS-003", "Ms", "Fiona", "Stewart", "Digital Scotland", "Managing Director", "f.stewart@digitalscot.gov.uk"],
        ["TS-004", "Mr", "Raj", "Patel", "GlobalTech Partners", "Head of Innovation", "r.patel@globaltech.com"],
        ["TS-005", "Mrs", "Emma", "Campbell", "ScotBank PLC", "IT Director", "e.campbell@scotbank.co.uk"],
        ["TS-006", "Mr", "David", "Morrison", "CloudNine Systems", "Solutions Architect", "d.morrison@cloudnine.io"],
        ["TS-007", "Dr", "Anna", "Kowalski", "NHS Scotland", "Digital Health Lead", "a.kowalski@nhs.scot"],
        ["TS-008", "Mr", "Michael", "Fraser", "EnergyTech Ltd", "Software Engineer", "m.fraser@energytech.co.uk"],
        ["TS-009", "Ms", "Laura", "Henderson", "DataFlow Analytics", "Data Scientist", "l.henderson@dataflow.com"],
        ["TS-010", "Mr", "Thomas", "Sinclair", "CyberGuard UK", "Security Consultant", "t.sinclair@cyberguard.co.uk"],
    ]

    for d in delegates:
        ws.append(d)

    # Format
    for cell in ws[1]:
        cell.font = cell.font.copy(bold=True)

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 2

    filepath = os.path.join(SRC, "Delegate_List.xlsx")
    wb.save(filepath)
    print(f"[OK] Delegate_List.xlsx")


def create_logo():
    """Create TechSummit logo programmatically"""
    img = Image.new('RGB', (600, 200), '#003366')
    draw = ImageDraw.Draw(img)

    # Try to use a nice font, fallback to default
    try:
        font_large = ImageFont.truetype("arial.ttf", 48)
        font_small = ImageFont.truetype("arial.ttf", 20)
    except:
        font_large = ImageFont.load_default()
        font_small = ImageFont.load_default()

    # Draw text
    draw.text((30, 30), "TechSummit", fill='#FFFFFF', font=font_large)
    draw.text((30, 100), "2026", fill='#4FC3F7', font=font_large)
    draw.text((30, 160), "Digital Innovation Conference", fill='#B0BEC5', font=font_small)

    # Decorative element - circuit-like lines
    draw.line([(400, 50), (580, 50)], fill='#4FC3F7', width=2)
    draw.line([(580, 50), (580, 150)], fill='#4FC3F7', width=2)
    draw.line([(400, 150), (580, 150)], fill='#4FC3F7', width=2)
    draw.ellipse([(390, 45), (410, 65)], fill='#4FC3F7')
    draw.ellipse([(390, 145), (410, 165)], fill='#4FC3F7')
    draw.ellipse([(570, 95), (590, 115)], fill='#4FC3F7')

    filepath = os.path.join(SRC, "TechSummit_Logo.png")
    img.save(filepath, "PNG")
    print(f"[OK] TechSummit_Logo.png")


def create_audio():
    """Create a simple WAV audio file with a tone (voiceover placeholder)"""
    sample_rate = 44100
    duration = 5  # seconds
    frequency = 440  # Hz (A4 note)

    num_samples = sample_rate * duration
    samples = array.array('h')

    for i in range(num_samples):
        t = i / sample_rate
        # Fade in/out
        envelope = min(t / 0.5, 1.0) * min((duration - t) / 0.5, 1.0)
        # Mix two frequencies for a richer sound
        value = int(16000 * envelope * (
            0.6 * math.sin(2 * math.pi * frequency * t) +
            0.4 * math.sin(2 * math.pi * (frequency * 1.5) * t)
        ))
        samples.append(max(-32768, min(32767, value)))

    filepath = os.path.join(SRC, "welcome_audio.wav")
    with wave.open(filepath, 'w') as wav:
        wav.setnchannels(1)
        wav.setsampwidth(2)
        wav.setframerate(sample_rate)
        wav.writeframes(samples.tobytes())

    print(f"[OK] welcome_audio.wav ({duration}s)")


def create_video():
    """Create a basic AVI video file (simpler than MP4 - no codec issues)"""
    # Create frames as images first, then use ffmpeg if available
    # For simplicity, create a series of PNG frames
    frames_dir = os.path.join(SRC, "video_frames")
    os.makedirs(frames_dir, exist_ok=True)

    width, height = 640, 480
    fps = 10
    duration = 5

    for i in range(fps * duration):
        img = Image.new('RGB', (width, height), '#003366')
        draw = ImageDraw.Draw(img)

        try:
            font = ImageFont.truetype("arial.ttf", 36)
            font_small = ImageFont.truetype("arial.ttf", 20)
        except:
            font = ImageFont.load_default()
            font_small = font

        # Animated text
        y_offset = int(10 * math.sin(i / 5))
        draw.text((120, 150 + y_offset), "TechSummit 2026", fill='white', font=font)
        draw.text((140, 220), "Welcome to Edinburgh", fill='#4FC3F7', font=font_small)
        draw.text((180, 260), f"Starting Soon...", fill='#90CAF9', font=font_small)

        # Progress bar
        progress = i / (fps * duration)
        draw.rectangle([(100, 350), (540, 370)], outline='white', width=1)
        draw.rectangle([(100, 350), (100 + int(440 * progress), 370)], fill='#4FC3F7')

        img.save(os.path.join(frames_dir, f"frame_{i:04d}.png"))

    # Try to create video with ffmpeg
    import subprocess
    video_path = os.path.join(SRC, "welcome_video.mp4")
    try:
        result = subprocess.run([
            'ffmpeg', '-y', '-framerate', str(fps),
            '-i', os.path.join(frames_dir, 'frame_%04d.png'),
            '-c:v', 'libx264', '-pix_fmt', 'yuv420p',
            video_path
        ], capture_output=True, text=True, timeout=30)
        if result.returncode == 0:
            print(f"[OK] welcome_video.mp4")
        else:
            print(f"[WARN] ffmpeg failed: {result.stderr[:200]}")
            # Save just the first frame as a placeholder
            print(f"[OK] Video frames saved in {frames_dir} (ffmpeg not available for MP4)")
    except FileNotFoundError:
        print(f"[INFO] ffmpeg not available - video frames saved in {frames_dir}")
        # Create a WMV using Windows built-in tools or just use frames
        print(f"[OK] Will use frame images as video evidence")


if __name__ == "__main__":
    create_conference_text()
    create_delegate_list()
    create_logo()
    create_audio()
    create_video()
    print("\n[DONE] All Unit C assets created")

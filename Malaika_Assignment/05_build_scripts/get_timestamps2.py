"""Get YouTube video timestamps using python -m yt_dlp"""
import subprocess, json, re
from pathlib import Path

VIDEOS = {
    "Topic1_Evolution_of_HRM":   "https://www.youtube.com/watch?v=Kxc8KceOb14",
    "Topic4_Job_Analysis":       "https://www.youtube.com/watch?v=oas5n1nFHQQ",
    "Topic4_Job_Design":         "https://www.youtube.com/watch?v=uUG-Z5sg2UM",
}
KEYWORDS = ["shovel","drive system","hands","cogs","human condition","scientific management",
            "behavioral","job design","motivation","autonomy","obsolete","enrichment","boredom"]
OUT_DIR = Path(__file__).parent / "video_transcripts"
OUT_DIR.mkdir(exist_ok=True)

def ts(secs):
    s=int(float(secs)); m,s2=divmod(s,60); h,m2=divmod(m,60)
    return f"{h}:{m2:02d}:{s2:02d}" if h else f"{m2}:{s2:02d}"

results = {}
for name, url in VIDEOS.items():
    print(f"\n{name}")
    files = list(OUT_DIR.glob(f"{name}*.vtt"))
    if not files:
        subprocess.run(
            ["python","-m","yt_dlp","--write-auto-sub","--sub-lang","en",
             "--skip-download","--sub-format","vtt","-o",str(OUT_DIR/f"{name}.%(ext)s"),url],
            capture_output=True, timeout=60
        )
        files = list(OUT_DIR.glob(f"{name}*.vtt"))

    if not files:
        print("  No VTT found"); results[name]={"url":url,"hits":[]}; continue

    raw = files[0].read_text(encoding='utf-8', errors='replace')
    hits = []
    for block in raw.split('\n\n'):
        lines = [l.strip() for l in block.strip().split('\n') if l.strip()]
        if not lines or '-->' not in lines[0]: continue
        text_lines = [l for l in lines[1:] if not re.match(r'^[\d:.,\s\->]+$', l)]
        text = ' '.join(text_lines).lower()
        start_raw = lines[0].split('-->')[0].strip()
        parts = re.split(r'[:.,]', start_raw)
        try:
            nums = [int(x) for x in parts if x.strip().isdigit()]
            if len(nums) >= 3: secs = nums[-3]*3600 + nums[-2]*60 + nums[-1]
            elif len(nums) == 2: secs = nums[0]*60 + nums[1]
            else: secs = nums[0] if nums else 0
        except: secs = 0

        for kw in KEYWORDS:
            if kw in text:
                hits.append({"kw":kw,"ts":ts(secs),"secs":secs,
                             "text":' '.join(text_lines)[:150],
                             "url_ts":f"{url}&t={secs}s"})
                break

    results[name] = {"url":url,"hits":hits}
    for h in hits[:8]:
        print(f"  [{h['ts']}] ({h['kw']}) {h['text'][:80]}")

out = OUT_DIR/"timestamp_results.json"
out.write_text(json.dumps(results, indent=2))
print(f"\nSaved: {out}")

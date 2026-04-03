"""
Downloads YouTube transcripts with timestamps for course videos.
Searches for key quotes to get exact timestamps.
"""
import subprocess, json, re
from pathlib import Path

VIDEOS = {
    "Topic1_Evolution_of_HRM":               "https://www.youtube.com/watch?v=Kxc8KceOb14",
    "Topic1_Transformation_Personnel_HRM":   "https://www.youtube.com/watch?v=8ReX2poQyJ0",
    "Topic4_Job_Analysis":                   "https://www.youtube.com/watch?v=oas5n1nFHQQ",
    "Topic4_Job_Design":                     "https://www.youtube.com/watch?v=uUG-Z5sg2UM",
}

KEYWORDS = [
    "shovel", "drive system", "hands", "cogs", "human conditions",
    "scientific management", "taylorism", "behavioral", "job design",
    "motivation", "meaningfulness", "autonomy", "enrichment", "obsolete"
]

OUT_DIR = Path(__file__).parent / "video_transcripts"
OUT_DIR.mkdir(exist_ok=True)

def seconds_to_ts(s):
    s = int(float(s))
    m, sec = divmod(s, 60)
    h, m = divmod(m, 60)
    if h:
        return f"{h}:{m:02d}:{sec:02d}"
    return f"{m}:{sec:02d}"

results = {}

for name, url in VIDEOS.items():
    print(f"\nFetching transcript: {name}")
    try:
        r = subprocess.run(
            ["yt-dlp", "--write-auto-sub", "--sub-lang", "en",
             "--skip-download", "--sub-format", "json3",
             "-o", str(OUT_DIR / f"{name}.%(ext)s"), url],
            capture_output=True, text=True, timeout=60
        )

        # Find the downloaded json3 file
        files = list(OUT_DIR.glob(f"{name}*.json3"))
        if not files:
            # try vtt
            r2 = subprocess.run(
                ["yt-dlp", "--write-auto-sub", "--sub-lang", "en",
                 "--skip-download", "--sub-format", "vtt",
                 "-o", str(OUT_DIR / f"{name}.%(ext)s"), url],
                capture_output=True, text=True, timeout=60
            )
            files = list(OUT_DIR.glob(f"{name}*.vtt"))

        if files:
            with open(files[0], encoding='utf-8', errors='replace') as f:
                raw = f.read()

            # Parse VTT format
            hits = []
            if files[0].suffix == '.vtt':
                blocks = raw.split('\n\n')
                for block in blocks:
                    lines = block.strip().split('\n')
                    if len(lines) >= 2 and '-->' in lines[0]:
                        ts_line = lines[0]
                        text = ' '.join(lines[1:]).lower()
                        for kw in KEYWORDS:
                            if kw in text:
                                start = ts_line.split('-->')[0].strip()
                                # Convert VTT timestamp to seconds
                                parts = start.replace(',','.').split(':')
                                if len(parts) == 3:
                                    secs = int(parts[0])*3600 + int(parts[1])*60 + float(parts[2])
                                elif len(parts) == 2:
                                    secs = int(parts[0])*60 + float(parts[1])
                                else:
                                    secs = float(parts[0])
                                hits.append({
                                    "keyword": kw,
                                    "timestamp": seconds_to_ts(secs),
                                    "seconds": int(secs),
                                    "text": ' '.join(lines[1:])[:200]
                                })
                                break

            results[name] = {"url": url, "hits": hits[:20]}
            print(f"  Found {len(hits)} keyword hits")
            for h in hits[:5]:
                print(f"  [{h['timestamp']}] ({h['keyword']}) {h['text'][:80]}")
        else:
            print(f"  No transcript file found")
            results[name] = {"url": url, "hits": [], "error": "No transcript"}

    except Exception as e:
        print(f"  Error: {e}")
        results[name] = {"url": url, "hits": [], "error": str(e)}

with open(OUT_DIR / "timestamp_results.json", "w") as f:
    json.dump(results, f, indent=2)

print("\n\nSaved timestamp_results.json")

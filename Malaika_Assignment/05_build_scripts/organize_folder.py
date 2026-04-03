"""Organises the Malaika_Assignment folder into a clean structure."""
import shutil, os
from pathlib import Path

BASE = Path(__file__).parent

# Target structure
DIRS = {
    "01_assignment":        [],
    "02_course_materials/pdfs":    [],
    "02_course_materials/notes":   [],
    "03_research/notebooklm":      [],
    "03_research/transcripts":     [],
    "04_animation":                [],
    "05_build_scripts":            [],
}
for d in DIRS:
    (BASE / d).mkdir(parents=True, exist_ok=True)

def mv(src, dst_dir):
    src = BASE / src
    if src.exists():
        dst = BASE / dst_dir / src.name
        if not dst.exists():
            shutil.move(str(src), str(dst))
            print(f"  Moved: {src.name} -> {dst_dir}/")
        else:
            print(f"  Skip (exists): {src.name}")

# 01 Assignment deliverables
mv("Malaika_MGMT268_Assessment1_FINAL.docx", "01_assignment")
mv("Malaika_MGMT268_Assessment1.docx",       "01_assignment")
mv("hackman_oldham_model.png",               "01_assignment")

# 02 Course materials
for f in (BASE / "course_pdfs").glob("*.pdf"):
    dst = BASE / "02_course_materials/pdfs" / f.name
    if not dst.exists():
        shutil.copy2(str(f), str(dst))
for f in (BASE / "course_content").glob("*"):
    if f.is_file():
        dst = BASE / "02_course_materials/notes" / f.name
        if not dst.exists():
            shutil.copy2(str(f), str(dst))

# 03 Research
for f in (BASE / "video_transcripts").glob("*"):
    if f.is_file():
        mv(f"video_transcripts/{f.name}", "03_research/transcripts")

# 04 Animation
if (BASE / "hr_animation").exists():
    dst = BASE / "04_animation" / "hr_animation"
    if not dst.exists():
        shutil.copytree(str(BASE / "hr_animation"), str(dst))
        print("  Copied: hr_animation -> 04_animation/")

# 05 Build scripts
for fname in ["build_essay.py","build_essay_v2.py","make_diagram.py",
              "full_pipeline.js","scrape_uwi.js","uwi_login.js",
              "download_all.js","create_notebook.py","get_timestamps2.py",
              "save_session.py","check_login.js","debug_login.js",
              "get_yt_links.js","get_yt_links2.js","extract_chrome_cookies.py",
              "notebooklm_login.py","notebooklm_setup.js"]:
    mv(fname, "05_build_scripts")

# Clean up temp screenshots
for f in BASE.glob("*.png"):
    if "screenshot" in f.name.lower() or f.name.startswith("step") or f.name.startswith("login"):
        f.unlink()
        print(f"  Deleted: {f.name}")

print("\nFolder organisation complete.")
print("\nFinal structure:")
for item in sorted(BASE.iterdir()):
    if item.is_dir() and not item.name.startswith('.') and item.name != "node_modules":
        count = sum(1 for _ in item.rglob("*") if _.is_file())
        print(f"  {item.name}/  ({count} files)")
    elif item.is_file() and item.name not in ["package.json","package-lock.json"]:
        print(f"  {item.name}")

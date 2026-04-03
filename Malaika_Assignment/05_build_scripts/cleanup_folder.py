"""Final cleanup: restructure Malaika_Assignment like SCQF_L6_FINAL_SUBMISSION."""
import shutil, os
from pathlib import Path

BASE = Path(__file__).parent

def mv(src, dst_dir, rename=None):
    src = BASE / src
    if not src.exists():
        return
    dst_dir_path = BASE / dst_dir
    dst_dir_path.mkdir(parents=True, exist_ok=True)
    dst = dst_dir_path / (rename or src.name)
    if dst.exists():
        print(f"  Skip (exists): {dst.relative_to(BASE)}")
        return
    shutil.move(str(src), str(dst))
    print(f"  Moved: {src.name}  ->  {dst_dir}/")

def rmdir_if_empty(path):
    p = BASE / path
    if p.exists() and p.is_dir():
        files = list(p.rglob("*"))
        if not files:
            p.rmdir()
            print(f"  Removed empty dir: {path}")

# ── 1. FINAL_SUBMISSION  ───────────────────────────────────────────────────
#    Mirror of SCQF_L6_FINAL_SUBMISSION — submission-ready files only
for fname in [
    "Malaika_MGMT268_Assessment1_FINAL.pdf",
    "Malaika_MGMT268_Assessment1_WITH_VIDEO.pdf",
    "Malaika_MGMT268_Assessment1_FINAL.docx",
    "Malaika_MGMT268_Assessment1_WITH_VIDEO.docx",
]:
    mv(fname, "FINAL_SUBMISSION")

# Also grab from 01_assignment if already moved there
for fname in [
    "01_assignment/Malaika_MGMT268_Assessment1_FINAL.docx",
    "01_assignment/Malaika_MGMT268_Assessment1.docx",
]:
    mv(fname, "01_assignment")   # keep in 01 too, but primary copy in FINAL_SUBMISSION

# ── 2. 01_assignment — figures and excalidraw  ────────────────────────────
for f in [
    "comparison_table.png",
    "gns_diagram.png",
    "hr_dialogue_illustration.png",
    "animation_frame.png",
    "hr_dialogue.excalidraw",
]:
    mv(f, "01_assignment")

# ── 3. 05_build_scripts — loose scripts and tools  ───────────────────────
for f in [
    "build_essay_video.py",
    "export_to_pdf.py",
    "generate_hr_dialogue.py",
    "make_diagrams_extra.py",
    "organize_folder.py",
    "get_timestamps.py",
    "get_student_info.js",
    "convert_pdf.ps1",
    "convert_to_pdf.vbs",
    "pipeline_summary.json",
    "student_info.json",
    "package.json",
    "package-lock.json",
]:
    mv(f, "05_build_scripts")

# Move excalidraw_render/ into 05_build_scripts/
excal = BASE / "excalidraw_render"
if excal.exists():
    dst = BASE / "05_build_scripts/excalidraw_render"
    if not dst.exists():
        shutil.move(str(excal), str(dst))
        print("  Moved: excalidraw_render/  ->  05_build_scripts/")

# ── 4. Remove duplicate source dirs at root (content already in 02/04)  ──
for dup_dir in ["course_content", "course_pdfs", "video_transcripts"]:
    d = BASE / dup_dir
    if d.exists():
        # Only remove if 02_course_materials already has the content
        target_notes = BASE / "02_course_materials/notes"
        target_pdfs  = BASE / "02_course_materials/pdfs"
        if target_notes.exists() and any(target_notes.iterdir()):
            shutil.rmtree(str(d))
            print(f"  Removed duplicate: {dup_dir}/")

# Remove duplicate hr_animation at root (already in 04_animation/)
dup_anim = BASE / "hr_animation"
if dup_anim.exists():
    anim_in_04 = BASE / "04_animation/hr_animation"
    if anim_in_04.exists():
        shutil.rmtree(str(dup_anim))
        print("  Removed duplicate: hr_animation/ (kept 04_animation/hr_animation/)")

# ── Print final structure  ────────────────────────────────────────────────
print("\nFinal structure:")
for item in sorted(BASE.iterdir()):
    skip = {"node_modules", "hr_animation_remotion", ".venv", "__pycache__"}
    if item.name.startswith(".") or item.name in skip:
        continue
    if item.is_dir():
        count = sum(1 for _ in item.rglob("*")
                    if _.is_file() and "node_modules" not in str(_))
        print(f"  {item.name}/  ({count} files)")
    else:
        size = item.stat().st_size
        print(f"  {item.name}  ({size//1024} KB)")

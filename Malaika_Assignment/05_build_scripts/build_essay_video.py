"""
Creates Malaika_MGMT268_Assessment1_WITH_VIDEO.docx
Copies the FINAL.docx and embeds HRMDialogueScene.mp4 as an OLE object
in the Appendix B section using COM automation.
"""
import shutil, time
from pathlib import Path

SRC   = Path(__file__).parent / "Malaika_MGMT268_Assessment1_FINAL.docx"
DEST  = Path(__file__).parent / "Malaika_MGMT268_Assessment1_WITH_VIDEO.docx"
VIDEO = Path(__file__).parent / "hr_animation/media/videos/main/1080p60/HRMDialogueScene.mp4"

if not SRC.exists():
    print(f"ERROR: Run build_essay_v2.py first to create {SRC.name}")
    raise SystemExit(1)

if not VIDEO.exists():
    print(f"ERROR: Video not found: {VIDEO}")
    raise SystemExit(1)

# Copy the illustration version as base
shutil.copy2(str(SRC), str(DEST))
print(f"Copied base: {DEST.name}")

try:
    import win32com.client
    import pywintypes

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    doc = word.Documents.Open(str(DEST.resolve()))

    # Find the "Appendix B" paragraph and position cursor after the
    # animation frame image -- we'll append the OLE object at document end
    # then add a label.
    rng = doc.Content
    rng.Collapse(0)  # collapse to end

    # Add a paragraph break and label
    rng.InsertParagraphAfter()
    rng.Collapse(0)
    rng.InsertAfter("\nClick the icon below to play the companion animation:")
    rng.InsertParagraphAfter()
    rng.Collapse(0)

    # Insert the video as an OLE Package object (displayable icon)
    try:
        shape = rng.InlineShapes.AddOLEObject(
            FileName=str(VIDEO.resolve()),
            LinkToFile=False,
            DisplayAsIcon=True,
            IconLabel="HRMDialogueScene.mp4  -- Click to Play"
        )
        print("OLE object inserted successfully.")
    except Exception as ole_err:
        print(f"OLE insertion note: {ole_err}")
        # Fallback: insert as hyperlink reference
        rng.InsertAfter(f"[Video file: {VIDEO.name}  -- play with any media player]")

    doc.Save()
    doc.Close()
    word.Quit()
    print(f"Saved: {DEST}")

except Exception as e:
    print(f"COM automation error: {e}")
    print("Video version saved as copy of illustration version (no embedded video).")
    print(f"  --> Manually embed {VIDEO.name} in Appendix B if needed.")

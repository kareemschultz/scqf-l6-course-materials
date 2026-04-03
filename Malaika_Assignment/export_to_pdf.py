"""Export Word docs to PDF using early-bound Word COM automation."""
import win32com.client
from pathlib import Path

BASE = Path(__file__).parent

DOCS = [
    BASE / "Malaika_MGMT268_Assessment1_FINAL.docx",
    BASE / "Malaika_MGMT268_Assessment1_WITH_VIDEO.docx",
]

# EnsureDispatch generates proper type stubs -- fixes late-binding issues
import time, pythoncom

word = win32com.client.gencache.EnsureDispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0   # wdAlertsNone -- suppress all dialogs

for docx in DOCS:
    if not docx.exists():
        print(f"  Skip (not found): {docx.name}")
        continue
    pdf_out = docx.with_suffix(".pdf")
    try:
        doc = word.Documents.Open(
            str(docx.resolve()),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
        time.sleep(1)   # let Word fully load
        # pump COM messages to clear any pending dialogs
        pythoncom.PumpWaitingMessages()

        doc.SaveAs2(str(pdf_out.resolve()), FileFormat=17)
        doc.Close(False)
        size = pdf_out.stat().st_size // 1024
        print(f"  PDF: {pdf_out.name}  ({size} KB)")
    except Exception as e:
        print(f"  ERROR exporting {docx.name}: {e}")

word.Quit()
print("Done.")

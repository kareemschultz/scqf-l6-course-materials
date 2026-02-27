"""
UNIT C: F1FE 12 - Take screenshots of Word doc, PowerPoint, and mail merge
Uses window-specific capture positioned on primary monitor
"""
import win32com.client
import win32gui
import win32con
import time
import os
import pythoncom
import subprocess
from PIL import ImageGrab

BASE = r"C:\Users\admin\Documents\SCQF-L6-Course-Materials\SCQF_L6_FINAL_SUBMISSION"
SRC = os.path.join(BASE, "source_files", "F1FE12_Word_Presentation")
SS_DIR = os.path.join(BASE, "evidence_screenshots")

DOCX_PATH = os.path.join(SRC, "TechSummit2026_WelcomePack.docx")
BADGES_PATH = os.path.join(SRC, "TechSummit2026_Badges.docx")
PPTX_PATH = os.path.join(SRC, "TechSummit2026_DigitalSignage.pptx")


def find_window(partial_title):
    result = []
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            t = win32gui.GetWindowText(hwnd)
            if partial_title.lower() in t.lower():
                result.append(hwnd)
    win32gui.EnumWindows(cb, None)
    return result[0] if result else None


def capture(hwnd, filename, delay=0.8):
    time.sleep(delay)
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.MoveWindow(hwnd, 0, 0, 1600, 1000, True)
        time.sleep(0.2)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.3)
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        time.sleep(0.5)
    except:
        pass
    rect = win32gui.GetWindowRect(hwnd)
    img = ImageGrab.grab(bbox=rect)
    path = os.path.join(SS_DIR, filename)
    img.save(path, "PNG")
    print(f"  Screenshot: {filename}")


def take_word_screenshots():
    pythoncom.CoInitialize()
    subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'], capture_output=True)
    time.sleep(2)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    word.DisplayAlerts = 0

    try:
        # Open Welcome Pack
        doc = word.Documents.Open(DOCX_PATH)
        time.sleep(2)

        hwnd = find_window("Word")
        if hwnd:
            # Page 1 - TOC
            doc.ActiveWindow.ActivePane.SmallScroll(Down=0)
            capture(hwnd, "unitc_word_01_toc.png")

            # Page 2 - Welcome with logo
            word.Selection.GoTo(1, 2, 2)  # wdGoToPage, wdGoToAbsolute, page 2
            time.sleep(0.5)
            capture(hwnd, "unitc_word_02_welcome_logo.png")

            # Page 3 - Agenda table
            word.Selection.GoTo(1, 2, 3)
            time.sleep(0.5)
            capture(hwnd, "unitc_word_03_agenda_table.png")

            # Page 4 - Important info
            word.Selection.GoTo(1, 2, 4)
            time.sleep(0.5)
            capture(hwnd, "unitc_word_04_important_info.png")

            # Show styles pane
            try:
                word.ActiveDocument.ActiveWindow.StyleAreaWidth = 100
                time.sleep(0.5)
                capture(hwnd, "unitc_word_05_styles_pane.png")
                word.ActiveDocument.ActiveWindow.StyleAreaWidth = 0
            except:
                pass

        doc.Close(False)

        # Open Badges
        badge_doc = word.Documents.Open(BADGES_PATH)
        time.sleep(1)
        hwnd = find_window("Word")
        if hwnd:
            capture(hwnd, "unitc_word_06_mail_merge_badges.png")

            # Show page 2 badge
            try:
                word.Selection.GoTo(1, 2, 2)
                time.sleep(0.5)
                capture(hwnd, "unitc_word_07_badge_page2.png")
            except:
                pass

        badge_doc.Close(False)
        print("[OK] Word screenshots complete")

    except Exception as e:
        print(f"[ERROR] Word screenshots: {e}")
    finally:
        word.Visible = False
        word.Quit()
        pythoncom.CoUninitialize()


def take_ppt_screenshots():
    pythoncom.CoInitialize()
    subprocess.run(['taskkill', '/F', '/IM', 'POWERPNT.EXE'], capture_output=True)
    time.sleep(2)

    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True

    try:
        pres = ppt.Presentations.Open(PPTX_PATH, WithWindow=True)
        time.sleep(2)

        hwnd = find_window("PowerPoint")

        # Normal view - take screenshots of each slide
        for i in range(1, 6):
            try:
                ppt.ActiveWindow.View.GotoSlide(i)
                time.sleep(0.5)
                capture(hwnd, f"unitc_ppt_{i:02d}_slide{i}.png")
            except:
                pass

        # Export slides as images directly (more reliable)
        for i in range(1, 6):
            try:
                slide_path = os.path.join(SS_DIR, f"unitc_ppt_slide{i}_export.png")
                pres.Slides(i).Export(slide_path, "PNG", 1280, 720)
                print(f"  Slide {i} exported")
            except Exception as e:
                print(f"  Slide {i} export: {e}")

        # Switch to slide sorter view for overview
        try:
            ppt.ActiveWindow.ViewType = 4  # ppViewSlideSorter
            time.sleep(1)
            capture(hwnd, "unitc_ppt_06_slide_sorter.png")
        except:
            pass

        pres.Close()
        print("[OK] PowerPoint screenshots complete")

    except Exception as e:
        print(f"[ERROR] PPT screenshots: {e}")
    finally:
        try:
            ppt.Quit()
        except:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    take_word_screenshots()
    time.sleep(3)
    take_ppt_screenshots()

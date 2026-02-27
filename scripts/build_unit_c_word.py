"""
UNIT C: F1FE 12 - Word Processing & Presenting
Part 1: Conference Welcome Pack DOCX with TOC, mail merge, logo, agenda table
"""
import win32com.client
import os
import time
import pythoncom
import subprocess

BASE = r"C:\Users\admin\Documents\SCQF-L6-Course-Materials\SCQF_L6_FINAL_SUBMISSION"
SRC = os.path.join(BASE, "source_files", "F1FE12_Word_Presentation")
SS_DIR = os.path.join(BASE, "evidence_screenshots")

CONFERENCE_TEXT = os.path.join(SRC, "Conference_Text.txt")
DELEGATE_LIST = os.path.join(SRC, "Delegate_List.xlsx")
LOGO_PATH = os.path.join(SRC, "TechSummit_Logo.png")
DOCX_PATH = os.path.join(SRC, "TechSummit2026_WelcomePack.docx")
BADGES_PATH = os.path.join(SRC, "TechSummit2026_Badges.docx")


def create_welcome_pack():
    """Create the 4-page Conference Welcome Pack using Word COM"""
    pythoncom.CoInitialize()

    subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'], capture_output=True)
    time.sleep(2)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone

    try:
        doc = word.Documents.Add()

        # 1 cm = 28.35 points
        CM = 28.35

        # --- Page Setup ---
        doc.PageSetup.TopMargin = 2.5 * CM
        doc.PageSetup.BottomMargin = 2.5 * CM
        doc.PageSetup.LeftMargin = 2.5 * CM
        doc.PageSetup.RightMargin = 2.5 * CM

        # --- Header: "TechSummit 2026 - Conference Welcome Pack" ---
        header = doc.Sections(1).Headers(1)  # wdHeaderFooterPrimary
        header.Range.Text = "TechSummit 2026 - Conference Welcome Pack"
        header.Range.Font.Size = 9
        header.Range.Font.Color = 0x663300  # Dark blue (BGR)
        header.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter

        # --- Footer ---
        footer = doc.Sections(1).Footers(1)
        footer.Range.Text = "252IFCBR0596 | Kareem Nurw Jason Schultz | F1FE 12 | Using Software Application Packages"
        footer.Range.Font.Size = 8
        footer.Range.ParagraphFormat.Alignment = 1

        # Add page number to footer
        footer.Range.InsertAfter(" | Page ")
        footer.Range.Fields.Add(footer.Range.Characters.Last, 33)  # wdFieldPage

        rng = doc.Content

        # ============================================
        # PAGE 1: Table of Contents
        # ============================================
        rng.Text = "Table of Contents"
        rng.Style = doc.Styles("Heading 1")
        rng.InsertParagraphAfter()

        rng = doc.Content
        rng.Collapse(0)  # wdCollapseEnd

        # Insert TOC field
        doc.TablesOfContents.Add(
            Range=rng,
            UseHeadingStyles=True,
            UpperHeadingLevel=1,
            LowerHeadingLevel=3,
            UseFields=False
        )

        # Add page break
        rng = doc.Content
        rng.Collapse(0)
        rng.InsertBreak(7)  # wdPageBreak

        # ============================================
        # PAGE 2: Welcome & About (with Logo)
        # ============================================
        rng = doc.Content
        rng.Collapse(0)

        # Insert logo with tight wrapping
        rng.Text = "\n"
        rng.Collapse(0)

        logo_shape = doc.Shapes.AddPicture(
            FileName=LOGO_PATH,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0,
            Top=0,
            Width=( CM *8),
            Height=( CM *2.7)
        )
        logo_shape.WrapFormat.Type = 2  # wdWrapTight
        logo_shape.Left = 0
        logo_shape.Top = ( CM *0.5)

        rng = doc.Content
        rng.Collapse(0)

        # Welcome heading
        para = doc.Paragraphs.Add()
        para.Range.Text = "Welcome to TechSummit 2026"
        para.Style = doc.Styles("Heading 1")
        para.Range.InsertParagraphAfter()

        # Welcome text
        welcome_text = (
            "We are delighted to welcome you to the TechSummit 2026 Digital Innovation Conference, "
            "hosted at the Edinburgh International Conference Centre (EICC). This year's conference "
            "brings together over 500 technology professionals, industry leaders, and innovators from "
            "across the United Kingdom and beyond.\n\n"
            "TechSummit 2026 is now in its fifth year and has established itself as Scotland's premier "
            "technology conference. Our mission is to connect professionals, share cutting-edge knowledge, "
            "and inspire the next generation of digital innovation. This year's theme, 'Transforming "
            "Tomorrow Through Technology,' reflects our commitment to exploring how emerging technologies "
            "can create positive change across industries and communities."
        )

        rng = doc.Content
        rng.Collapse(0)
        para = doc.Paragraphs.Add()
        para.Range.Text = welcome_text
        para.Style = doc.Styles("Normal")
        para.Range.Font.Size = 11

        # About heading
        rng = doc.Content
        rng.Collapse(0)
        para = doc.Paragraphs.Add()
        para.Range.Text = "Keynote Speakers"
        para.Style = doc.Styles("Heading 1")

        speakers_text = (
            "Dr Sarah Chen, Chief Technology Officer at InnovateTech Solutions, will discuss the future of "
            "artificial intelligence in healthcare. Professor James MacLeod from the University of Edinburgh "
            "will present his groundbreaking research on quantum computing applications. Fiona Stewart, "
            "Managing Director of Digital Scotland, will share insights on Scotland's digital transformation "
            "strategy. Raj Patel, Head of Innovation at GlobalTech Partners, will explore sustainable "
            "technology and green computing initiatives."
        )

        rng = doc.Content
        rng.Collapse(0)
        para = doc.Paragraphs.Add()
        para.Range.Text = speakers_text
        para.Style = doc.Styles("Normal")

        # Page break
        rng = doc.Content
        rng.Collapse(0)
        rng.InsertBreak(7)

        # ============================================
        # PAGE 3: Agenda Table
        # ============================================
        rng = doc.Content
        rng.Collapse(0)
        para = doc.Paragraphs.Add()
        para.Range.Text = "Conference Agenda - Day 1"
        para.Style = doc.Styles("Heading 1")

        # Create agenda table (3cm time column, 10cm description column)
        rng = doc.Content
        rng.Collapse(0)
        rng.InsertParagraphAfter()
        rng = doc.Content
        rng.Collapse(0)

        table = doc.Tables.Add(
            Range=rng,
            NumRows=10,
            NumColumns=2
        )

        # Set column widths (3cm and 10cm)
        table.Columns(1).Width = ( CM *3)
        table.Columns(2).Width = ( CM *10)

        # Header row
        table.Cell(1, 1).Range.Text = "Time"
        table.Cell(1, 2).Range.Text = "Session"

        # Shade header
        table.Rows(1).Range.Font.Bold = True
        table.Rows(1).Range.Font.Color = 16777215  # White
        table.Rows(1).Shading.BackgroundPatternColor = 10040064  # Dark blue (BGR: 0x993300 -> needs correct)

        # Agenda items
        agenda = [
            ("08:00", "Registration and Welcome Coffee"),
            ("09:00", "Opening Ceremony and Welcome Address"),
            ("09:30", "Keynote: Dr Sarah Chen - AI in Healthcare"),
            ("10:30", "Coffee Break and Networking"),
            ("11:00", "Keynote: Prof James MacLeod - Quantum Computing"),
            ("12:00", "Panel Discussion: Future of Scottish Tech"),
            ("13:00", "Lunch and Exhibition Tour"),
            ("14:00", "Breakout Sessions (Choice of 4 tracks)"),
            ("16:00", "Keynote: Fiona Stewart - Digital Scotland Strategy"),
        ]

        for i, (time_str, session) in enumerate(agenda, 2):
            table.Cell(i, 1).Range.Text = time_str
            table.Cell(i, 2).Range.Text = session

        # Borders
        table.Borders.Enable = True

        # Page break
        rng = doc.Content
        rng.Collapse(0)
        rng.InsertBreak(7)

        # ============================================
        # PAGE 4: Important Information & Venue
        # ============================================
        rng = doc.Content
        rng.Collapse(0)
        para = doc.Paragraphs.Add()
        para.Range.Text = "Venue and Important Information"
        para.Style = doc.Styles("Heading 1")

        venue_text = (
            "The main conference hall is located on Level 2 of the EICC. Breakout sessions will be held "
            "in the Morrison and Pentland Rooms on Level 3. The exhibition area, featuring over 40 "
            "technology vendors, is situated in the main atrium on the ground floor. Refreshments will "
            "be served in the Cromdale Suite throughout the day."
        )

        rng = doc.Content
        rng.Collapse(0)
        para = doc.Paragraphs.Add()
        para.Range.Text = venue_text

        # Networking heading
        rng = doc.Content
        rng.Collapse(0)
        para = doc.Paragraphs.Add()
        para.Range.Text = "Networking Opportunities"
        para.Style = doc.Styles("Heading 1")

        networking_text = (
            "We encourage all delegates to take advantage of the networking opportunities available. "
            "The Welcome Reception takes place on the evening of Day 1 in the Castle Suite, with "
            "stunning views of Edinburgh Castle. Speed networking sessions are scheduled during the "
            "lunch breaks on both days."
        )

        rng = doc.Content
        rng.Collapse(0)
        para = doc.Paragraphs.Add()
        para.Range.Text = networking_text

        # Important info heading
        rng = doc.Content
        rng.Collapse(0)
        para = doc.Paragraphs.Add()
        para.Range.Text = "Important Information"
        para.Style = doc.Styles("Heading 1")

        info_text = (
            "Registration opens at 08:00 each morning in the main foyer. Please ensure you wear your "
            "delegate badge at all times for security purposes. Free Wi-Fi is available throughout the "
            "venue using the network name TECHSUMMIT2026 with the password Innovation2026. Emergency "
            "exits are clearly marked throughout the building. First aid stations are located at the "
            "reception desk on each floor.\n\n"
            "We hope you enjoy TechSummit 2026 and find the conference both informative and inspiring.\n\n"
            "The TechSummit Organising Committee"
        )

        rng = doc.Content
        rng.Collapse(0)
        para = doc.Paragraphs.Add()
        para.Range.Text = info_text

        # Update TOC
        try:
            doc.TablesOfContents(1).Update()
        except:
            pass

        # Save
        if os.path.exists(DOCX_PATH):
            os.remove(DOCX_PATH)
        doc.SaveAs2(DOCX_PATH, 12)  # wdFormatXMLDocument
        print(f"[OK] Welcome Pack saved: {DOCX_PATH}")

        doc.Close(False)

        # ============================================
        # MAIL MERGE: ID Badges from Delegate_List
        # ============================================
        print("Creating mail merge badges...")
        badge_doc = word.Documents.Add()

        # Set up badge template
        badge_doc.PageSetup.TopMargin = ( CM *1)
        badge_doc.PageSetup.BottomMargin = ( CM *1)

        rng = badge_doc.Content
        rng.Text = "TechSummit 2026 - Delegate Badge\n\n"
        rng.Font.Size = 16
        rng.Font.Bold = True

        rng = badge_doc.Content
        rng.Collapse(0)

        # Connect to data source
        badge_doc.MailMerge.OpenDataSource(
            Name=DELEGATE_LIST,
            Connection="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + DELEGATE_LIST,
            SQLStatement="SELECT * FROM [Delegates$]"
        )

        # Insert merge fields
        mm = badge_doc.MailMerge
        rng = badge_doc.Content
        rng.Collapse(0)

        rng.InsertAfter("Delegate ID: ")
        rng.Collapse(0)
        mm.Fields.Add(badge_doc.Content.Characters.Last, "Delegate_ID")

        rng = badge_doc.Content
        rng.Collapse(0)
        rng.InsertAfter("\n")
        rng.Collapse(0)
        mm.Fields.Add(badge_doc.Content.Characters.Last, "Title")

        rng = badge_doc.Content
        rng.Collapse(0)
        rng.InsertAfter(" ")
        rng.Collapse(0)
        mm.Fields.Add(badge_doc.Content.Characters.Last, "First_Name")

        rng = badge_doc.Content
        rng.Collapse(0)
        rng.InsertAfter(" ")
        rng.Collapse(0)
        mm.Fields.Add(badge_doc.Content.Characters.Last, "Last_Name")

        rng = badge_doc.Content
        rng.Collapse(0)
        rng.InsertAfter("\n")
        rng.Collapse(0)
        mm.Fields.Add(badge_doc.Content.Characters.Last, "Organisation")

        rng = badge_doc.Content
        rng.Collapse(0)
        rng.InsertAfter("\n")
        rng.Collapse(0)
        mm.Fields.Add(badge_doc.Content.Characters.Last, "Role")

        # Execute merge
        try:
            mm.Destination = 0  # wdSendToNewDocument
            mm.Execute()
            time.sleep(1)

            # Save merged document
            merged_doc = word.ActiveDocument
            if os.path.exists(BADGES_PATH):
                os.remove(BADGES_PATH)
            merged_doc.SaveAs2(BADGES_PATH, 12)
            merged_doc.Close(False)
            print(f"[OK] Merged badges saved: {BADGES_PATH}")
        except Exception as e:
            print(f"[WARN] Mail merge execute: {e}")
            # Save template anyway
            if os.path.exists(BADGES_PATH):
                os.remove(BADGES_PATH)
            badge_doc.SaveAs2(BADGES_PATH, 12)
            print(f"[OK] Badge template saved (merge fields present)")

        badge_doc.Close(False)

        print("[OK] Word documents complete")

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback; traceback.print_exc()
    finally:
        try:
            word.Quit()
        except:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    create_welcome_pack()

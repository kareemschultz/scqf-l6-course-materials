#!/usr/bin/env python3
"""
Build SCQF L6 Final Assignment Documents
- Opens actual JAIN College .docx templates
- Fills student info
- Injects humanized markdown content with proper formatting
- Adds footers with page numbers
- Saves and converts to PDF
"""

import os
import re
import copy
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import subprocess

# ============================================================
# CONFIGURATION
# ============================================================

STUDENT_NAME = "KAREEM SCHULTZ"
STUDENT_USN = "252IFCBR0596"
COLLEGE = "JAIN College of Engineering and Technology (JCET), Bangalore"

TEMPLATE_DIR = "/home/karetech/clawd/temp/SCQF-L6-Course-Materials/Assessment_Templates/"
OUTPUT_DIR = "/home/karetech/clawd/temp/SCQF_Final_v2/"

ASSIGNMENTS = [
    {
        "template": "J229 76_Understanding Business._Student Assessment Template.docx",
        "content": "/home/karetech/clawd/temp/SCQF-L6-Course-Materials/My_Assignments/J229_76_Understanding_Business/Drafts/J229_76_Draft_v2_Humanized.md",
        "output": "252IFCBR0596_KareemSchultz_J229_76_Understanding_Business.docx",
        "unit_code": "J229 76",
        "unit_title": "Understanding Business",
        "heading_para_idx": 55,  # Heading 1 with assignment title
    },
    {
        "template": "J22A 76_Management of People and Finance_Student_Assessment Template.docx",
        "content": "/home/karetech/clawd/temp/SCQF-L6-Course-Materials/My_Assignments/J22A_76_Management_People_Finance/Drafts/J22A_76_Draft_v2_Humanized.md",
        "output": "252IFCBR0596_KareemSchultz_J22A_76_Management_People_Finance.docx",
        "unit_code": "J22A 76",
        "unit_title": "Management of People and Finance",
        "heading_para_idx": 53,  # Normal style para with title; 54 is empty Heading 1
    },
    {
        "template": "HE9E 46_Contemporary Business Issues_Student Assessment Template.docx",
        "content": "/home/karetech/clawd/temp/SCQF-L6-Course-Materials/My_Assignments/HE9E_46_Contemporary_Business/Drafts/HE9E_46_Draft_v2_Humanized.md",
        "output": "252IFCBR0596_KareemSchultz_HE9E_46_Contemporary_Business_Issues.docx",
        "unit_code": "HE9E 46",
        "unit_title": "Contemporary Business Issues",
        "heading_para_idx": 51,  # Heading 1 with assignment title
    },
]

FONT_NAME = "Times New Roman"
BODY_SIZE = Pt(12)
TABLE_SIZE = Pt(11)
LINE_SPACING = 1.5

os.makedirs(OUTPUT_DIR, exist_ok=True)


# ============================================================
# MARKDOWN PARSER
# ============================================================

def parse_markdown(filepath):
    """Parse markdown file into structured content blocks."""
    with open(filepath, "r", encoding="utf-8") as f:
        lines = f.readlines()

    # Skip metadata header (Student, USN, Note lines, blank lines at top, title lines)
    content_started = False
    skip_until_content = True
    blocks = []
    i = 0

    while i < len(lines):
        line = lines[i].rstrip("\n")

        # Skip metadata at the top
        if skip_until_content:
            stripped = line.strip()
            if stripped == "":
                i += 1
                continue
            if stripped.startswith("# J2") or stripped.startswith("# HE9") or stripped.startswith("# Assignment Draft"):
                i += 1
                continue
            if stripped.startswith("**Student:**") or stripped.startswith("**USN:**") or stripped.startswith("**Note:**"):
                i += 1
                continue
            if stripped == "---":
                i += 1
                continue
            skip_until_content = False

        line_stripped = line.strip()

        # Skip horizontal rules
        if line_stripped == "---":
            i += 1
            continue

        # Skip final checklist section
        if line_stripped.startswith("**FINAL CHECKLIST:**"):
            break

        # Skip checklist items
        if line_stripped.startswith("- [ ]") or line_stripped.startswith("- [x]"):
            i += 1
            continue

        # Empty line
        if line_stripped == "":
            i += 1
            continue

        # Headings
        if line_stripped.startswith("# "):
            blocks.append({"type": "heading1", "text": line_stripped[2:].strip()})
            i += 1
            continue
        if line_stripped.startswith("## "):
            blocks.append({"type": "heading2", "text": line_stripped[3:].strip()})
            i += 1
            continue
        if line_stripped.startswith("### "):
            blocks.append({"type": "heading3", "text": line_stripped[4:].strip()})
            i += 1
            continue

        # Numbered list items (1. 2. etc)
        num_match = re.match(r'^(\d+)\.\s+(.+)', line_stripped)
        if num_match:
            blocks.append({"type": "numbered", "text": line_stripped})
            i += 1
            continue

        # Bullet items
        if line_stripped.startswith("- "):
            blocks.append({"type": "bullet", "text": line_stripped[2:].strip()})
            i += 1
            continue

        # Table - collect all table rows
        if line_stripped.startswith("|"):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                table_lines.append(lines[i].strip())
                i += 1
            blocks.append({"type": "table", "lines": table_lines})
            continue

        # Regular paragraph - collect continuation lines
        para_text = line_stripped
        i += 1
        while i < len(lines):
            next_line = lines[i].strip()
            # Continue paragraph if it's not empty and not a special line
            if (next_line and
                not next_line.startswith("#") and
                not next_line.startswith("- ") and
                not next_line.startswith("|") and
                not next_line == "---" and
                not re.match(r'^\d+\.\s+', next_line)):
                para_text += " " + next_line
                i += 1
            else:
                break

        blocks.append({"type": "paragraph", "text": para_text})

    return blocks


def parse_inline_formatting(text):
    """Parse inline markdown formatting into runs: [(text, bold, italic)]"""
    runs = []
    # Pattern to match **bold**, *italic*, or normal text
    # Handle **bold** first, then *italic*
    pattern = re.compile(r'(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*)')

    pos = 0
    for match in pattern.finditer(text):
        # Add normal text before match
        if match.start() > pos:
            normal_text = text[pos:match.start()]
            if normal_text:
                runs.append((normal_text, False, False))

        if match.group(2):  # ***bold italic***
            runs.append((match.group(2), True, True))
        elif match.group(3):  # **bold**
            runs.append((match.group(3), True, False))
        elif match.group(4):  # *italic*
            runs.append((match.group(4), False, True))

        pos = match.end()

    # Add remaining text
    if pos < len(text):
        remaining = text[pos:]
        if remaining:
            runs.append((remaining, False, False))

    if not runs:
        runs.append((text, False, False))

    return runs


def parse_table_lines(table_lines):
    """Parse markdown table into header and rows."""
    if len(table_lines) < 2:
        return None, None

    def split_row(line):
        # Remove leading/trailing pipes and split
        cells = line.strip("|").split("|")
        return [c.strip() for c in cells]

    header = split_row(table_lines[0])

    # Check if second line is separator
    separator = table_lines[1]
    if re.match(r'^[\|\s\-:]+$', separator):
        data_start = 2
    else:
        data_start = 1

    rows = []
    for line in table_lines[data_start:]:
        rows.append(split_row(line))

    return header, rows


# ============================================================
# DOCUMENT BUILDER
# ============================================================

def set_run_font(run, font_name=FONT_NAME, size=BODY_SIZE, bold=None, italic=None):
    """Set font properties on a run."""
    run.font.name = font_name
    run.font.size = size
    # Set East Asian font name too
    r = run._element
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = parse_xml(f'<w:rPr {nsdecls("w")}></w:rPr>')
        r.insert(0, rPr)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>')
        rPr.insert(0, rFonts)
    else:
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:cs'), font_name)

    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic


def set_paragraph_spacing(paragraph, line_spacing=LINE_SPACING, space_after=Pt(6)):
    """Set paragraph spacing."""
    pf = paragraph.paragraph_format
    pf.line_spacing = line_spacing
    if space_after is not None:
        pf.space_after = space_after


def add_formatted_paragraph(doc, text, style_name="Normal", font_size=BODY_SIZE,
                           alignment=None, is_reference=False):
    """Add a paragraph with inline formatting parsed from markdown."""
    p = doc.add_paragraph(style=style_name)

    if alignment:
        p.alignment = alignment

    runs = parse_inline_formatting(text)
    for run_text, is_bold, is_italic in runs:
        run = p.add_run(run_text)
        set_run_font(run, size=font_size, bold=is_bold if is_bold else None,
                     italic=is_italic if is_italic else None)

    set_paragraph_spacing(p)

    if is_reference:
        # Hanging indent for references
        pf = p.paragraph_format
        pf.left_indent = Cm(1.27)  # 0.5 inch
        pf.first_line_indent = Cm(-1.27)  # Negative for hanging

    return p


def add_bullet_paragraph(doc, text, style_name="List Paragraph"):
    """Add a bullet point paragraph."""
    p = doc.add_paragraph(style=style_name)

    # Set bullet formatting via numbering
    pPr = p._element.get_or_add_pPr()
    numPr = parse_xml(
        f'<w:numPr {nsdecls("w")}>'
        f'  <w:ilvl w:val="0"/>'
        f'  <w:numId w:val="1"/>'
        f'</w:numPr>'
    )
    # Actually, let's just use the List Paragraph style and add a bullet char
    # Clear and re-add with bullet
    p.clear()

    runs = parse_inline_formatting(text)
    for run_text, is_bold, is_italic in runs:
        run = p.add_run(run_text)
        set_run_font(run, size=BODY_SIZE, bold=is_bold if is_bold else None,
                     italic=is_italic if is_italic else None)

    set_paragraph_spacing(p, space_after=Pt(3))

    # Set left indent for bullet appearance
    pf = p.paragraph_format
    pf.left_indent = Cm(1.27)

    return p


def add_numbered_paragraph(doc, text):
    """Add a numbered list paragraph."""
    p = doc.add_paragraph(style="List Paragraph")

    runs = parse_inline_formatting(text)
    for run_text, is_bold, is_italic in runs:
        run = p.add_run(run_text)
        set_run_font(run, size=BODY_SIZE, bold=is_bold if is_bold else None,
                     italic=is_italic if is_italic else None)

    set_paragraph_spacing(p, space_after=Pt(3))
    pf = p.paragraph_format
    pf.left_indent = Cm(1.27)

    return p


def add_table_to_doc(doc, header, rows):
    """Add a formatted table to the document."""
    num_cols = len(header)
    table = doc.add_table(rows=1 + len(rows), cols=num_cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header row
    for j, cell_text in enumerate(header):
        cell = table.cell(0, j)
        cell.text = ""
        p = cell.paragraphs[0]
        runs = parse_inline_formatting(cell_text)
        for run_text, is_bold, is_italic in runs:
            run = p.add_run(run_text)
            set_run_font(run, size=TABLE_SIZE, bold=True)

        # Header shading
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="D9E2F3" w:val="clear"/>')
        cell._element.get_or_add_tcPr().append(shading)

    # Data rows
    for i, row_data in enumerate(rows):
        for j, cell_text in enumerate(row_data):
            if j < num_cols:
                cell = table.cell(i + 1, j)
                cell.text = ""
                p = cell.paragraphs[0]
                runs = parse_inline_formatting(cell_text)
                for run_text, is_bold, is_italic in runs:
                    run = p.add_run(run_text)
                    set_run_font(run, size=TABLE_SIZE,
                                bold=is_bold if is_bold else None,
                                italic=is_italic if is_italic else None)

    # Add spacing after table
    doc.add_paragraph("")

    return table


def add_footer(doc, unit_code, unit_title):
    """Add footer: StudentUSN_Name_UnitCode_UnitTitle_PageNumber (spec format)."""
    # Format: 252IFCBR0596_KareemSchultz_J229 76_Understanding Business_<page>
    footer_prefix = f"{STUDENT_USN}_KareemSchultz_{unit_code}_{unit_title}_"

    for section in doc.sections:
        footer = section.footer
        footer.is_linked_to_previous = False

        # Clear existing footer
        for p in footer.paragraphs:
            p.clear()

        # Use the first paragraph or add one
        if footer.paragraphs:
            p = footer.paragraphs[0]
        else:
            p = footer.add_paragraph()

        p.style = doc.styles['Footer']
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Add footer prefix text
        run = p.add_run(footer_prefix)
        set_run_font(run, size=Pt(9))

        # Add PAGE field inline (no tab separation — part of the same string)
        fldChar1 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
        run_page = p.add_run()
        run_page._element.append(fldChar1)

        instrText = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
        run_instr = p.add_run()
        run_instr._element.append(instrText)

        fldChar2 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
        run_end = p.add_run()
        run_end._element.append(fldChar2)


def fill_student_info(doc):
    """Fill in Table 2 (Student Info) with USN and College."""
    if len(doc.tables) >= 3:
        table = doc.tables[2]
        # Row 1: Student USN
        if len(table.rows) > 1:
            cell = table.rows[1].cells[0]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(f"Student USN: {STUDENT_USN}")
            set_run_font(run, size=Pt(11))

        # Row 2: College Name & Site
        if len(table.rows) > 2:
            cell = table.rows[2].cells[0]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(f"College Name & Site: {COLLEGE}")
            set_run_font(run, size=Pt(11))


def inject_content(doc, blocks, assignment_info):
    """Inject parsed content blocks into the document after existing content."""
    in_references = False

    for block in blocks:
        btype = block["type"]

        if btype == "heading1":
            text = block["text"]
            # Check if this is the References heading
            if text.strip().lower() == "references":
                in_references = True
            p = doc.add_paragraph(style="Heading 1")
            run = p.add_run(text)
            set_run_font(run, size=Pt(14))
            set_paragraph_spacing(p, space_after=Pt(6))

        elif btype == "heading2":
            p = doc.add_paragraph(style="Heading 2")
            run = p.add_run(block["text"])
            set_run_font(run, size=Pt(13))
            set_paragraph_spacing(p, space_after=Pt(6))

        elif btype == "heading3":
            p = doc.add_paragraph(style="Heading 3")
            run = p.add_run(block["text"])
            set_run_font(run, size=Pt(12), bold=True)
            set_paragraph_spacing(p, space_after=Pt(4))

        elif btype == "paragraph":
            text = block["text"]
            if in_references:
                add_formatted_paragraph(doc, text, is_reference=True)
            else:
                add_formatted_paragraph(doc, text)

        elif btype == "bullet":
            add_bullet_paragraph(doc, block["text"])

        elif btype == "numbered":
            add_numbered_paragraph(doc, block["text"])

        elif btype == "table":
            header, rows = parse_table_lines(block["lines"])
            if header and rows:
                add_table_to_doc(doc, header, rows)


def fix_j22a_heading(doc):
    """Fix J22A template's weird heading - para 53 Normal, para 54 empty Heading 1."""
    # Make paragraph 53 into Heading 1
    if len(doc.paragraphs) > 54:
        p53 = doc.paragraphs[53]
        p54 = doc.paragraphs[54]

        # Set p53 to Heading 1 style
        p53.style = doc.styles['Heading 1']

        # Remove the empty Heading 1 paragraph (p54)
        p54_elem = p54._element
        p54_elem.getparent().remove(p54_elem)


def convert_to_pdf(docx_path, pdf_path):
    """Convert docx to PDF using mammoth -> HTML -> Chrome headless."""
    import mammoth

    html_path = docx_path.replace('.docx', '.html')

    # Convert docx to HTML
    with open(docx_path, "rb") as f:
        result = mammoth.convert_to_html(f)
        html = result.value

    # Wrap in proper HTML with styling
    full_html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
    @page {{
        size: A4;
        margin: 2.54cm;
    }}
    body {{
        font-family: 'Times New Roman', Times, serif;
        font-size: 12pt;
        line-height: 1.5;
        color: #000;
        max-width: 210mm;
        margin: 0 auto;
        padding: 20px;
    }}
    h1 {{
        font-size: 14pt;
        font-weight: bold;
        margin-top: 18pt;
        margin-bottom: 6pt;
        page-break-after: avoid;
    }}
    h2 {{
        font-size: 13pt;
        font-weight: bold;
        margin-top: 14pt;
        margin-bottom: 6pt;
        page-break-after: avoid;
    }}
    h3 {{
        font-size: 12pt;
        font-weight: bold;
        margin-top: 12pt;
        margin-bottom: 4pt;
        page-break-after: avoid;
    }}
    p {{
        margin-bottom: 6pt;
        text-align: justify;
    }}
    table {{
        border-collapse: collapse;
        width: 100%;
        margin: 12pt 0;
        font-size: 11pt;
    }}
    th, td {{
        border: 1px solid #000;
        padding: 6px 8px;
        text-align: left;
    }}
    th {{
        background-color: #D9E2F3;
        font-weight: bold;
    }}
    ul, ol {{
        margin-left: 1.27cm;
        margin-bottom: 6pt;
    }}
    li {{
        margin-bottom: 3pt;
    }}
</style>
</head>
<body>
{html}
</body>
</html>"""

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(full_html)

    # Convert HTML to PDF using Chrome headless
    chrome_path = "/usr/bin/google-chrome"
    cmd = [
        chrome_path,
        "--headless",
        f"--print-to-pdf={pdf_path}",
        "--print-to-pdf-no-header",
        "--no-sandbox",
        "--disable-gpu",
        "--disable-software-rasterizer",
        html_path
    ]

    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    if result.returncode != 0:
        print(f"  Chrome PDF conversion warning: {result.stderr[:200]}")

    # Cleanup HTML
    if os.path.exists(html_path):
        os.remove(html_path)

    return os.path.exists(pdf_path)


# ============================================================
# MAIN
# ============================================================

def process_assignment(assignment):
    """Process a single assignment."""
    template_path = os.path.join(TEMPLATE_DIR, assignment["template"])
    output_path = os.path.join(OUTPUT_DIR, assignment["output"])
    content_path = assignment["content"]

    print(f"\n{'='*60}")
    print(f"Processing: {assignment['unit_code']} - {assignment['unit_title']}")
    print(f"{'='*60}")

    # 1. Open template
    print("  Opening template...")
    doc = Document(template_path)

    # 2. Fill student info
    print("  Filling student info...")
    fill_student_info(doc)

    # 3. Fix J22A heading quirk
    if "J22A" in assignment["template"]:
        print("  Fixing J22A heading...")
        fix_j22a_heading(doc)

    # 4. Parse markdown content
    print("  Parsing markdown content...")
    blocks = parse_markdown(content_path)
    print(f"    Found {len(blocks)} content blocks")

    # Count by type
    type_counts = {}
    for b in blocks:
        t = b["type"]
        type_counts[t] = type_counts.get(t, 0) + 1
    for t, c in sorted(type_counts.items()):
        print(f"      {t}: {c}")

    # 5. Inject content
    print("  Injecting content...")
    inject_content(doc, blocks, assignment)

    # 6. Add footer
    print("  Adding footer...")
    add_footer(doc, assignment["unit_code"], assignment["unit_title"])

    # 7. Save
    print(f"  Saving to {output_path}...")
    doc.save(output_path)

    # Check file size
    size_kb = os.path.getsize(output_path) / 1024
    print(f"  File size: {size_kb:.1f} KB")

    # 8. Verify
    print("  Verifying...")
    verify_doc = Document(output_path)
    print(f"    Paragraphs: {len(verify_doc.paragraphs)}")
    print(f"    Tables: {len(verify_doc.tables)}")

    # Check student info
    if len(verify_doc.tables) >= 3:
        tbl = verify_doc.tables[2]
        usn_text = tbl.rows[1].cells[0].text
        college_text = tbl.rows[2].cells[0].text
        print(f"    USN: {usn_text}")
        print(f"    College: {college_text[:50]}...")

    # Check footer
    for section in verify_doc.sections:
        footer_text = ""
        for p in section.footer.paragraphs:
            footer_text += p.text
        if footer_text:
            print(f"    Footer: {footer_text[:60]}...")

    return output_path


def main():
    print("SCQF L6 Assignment Builder")
    print("=" * 60)

    docx_files = []
    for assignment in ASSIGNMENTS:
        output_path = process_assignment(assignment)
        docx_files.append(output_path)

    # Convert to PDF
    print(f"\n{'='*60}")
    print("Converting to PDF...")
    print(f"{'='*60}")

    for docx_path in docx_files:
        pdf_path = docx_path.replace('.docx', '.pdf')
        print(f"  Converting {os.path.basename(docx_path)}...")
        success = convert_to_pdf(docx_path, pdf_path)
        if success:
            pdf_size = os.path.getsize(pdf_path) / 1024
            print(f"    PDF created: {pdf_size:.1f} KB")
        else:
            print(f"    PDF creation failed!")

    # Final summary
    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")
    print(f"Output directory: {OUTPUT_DIR}")
    for f in sorted(os.listdir(OUTPUT_DIR)):
        fpath = os.path.join(OUTPUT_DIR, f)
        size = os.path.getsize(fpath) / 1024
        print(f"  {f} ({size:.1f} KB)")

    print("\nDone!")


if __name__ == "__main__":
    main()

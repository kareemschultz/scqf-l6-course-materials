"""
QA Gate: Verify all 5 PDFs meet submission requirements
"""
import win32com.client
import os
import time
import pythoncom
import subprocess

BASE = r"C:\Users\admin\Documents\SCQF-L6-Course-Materials\SCQF_L6_FINAL_SUBMISSION"
PDF_DIR = os.path.join(BASE, "pdf_submissions")
QA_DIR = os.path.join(BASE, "qa_reports")
os.makedirs(QA_DIR, exist_ok=True)

STUDENT_ID = "252IFCBR0596"
STUDENT_NAME = "Kareem Nurw Jason Schultz"

# Expected PDFs
EXPECTED = {
    "J22976_Understanding_Business.pdf": {
        "unit_code": "J229 76",
        "unit_title": "Understanding Business",
        "marks": 100,
        "sections": ["1.1.1", "1.1.2", "1.1.3", "1.2.1", "1.3.1", "1.3.2", "2.1", "2.2", "2.3"],
    },
    "F1FJ12_Spreadsheet_Database.pdf": {
        "unit_code": "F1FJ 12",
        "unit_title": "Using Software Application Packages",
        "marks": 70,
        "sections": ["1.1.1", "1.1.2", "1.1.3", "1.2.1", "1.2.2", "1.2.3", "1.3.1", "1.3.2",
                      "2.1", "2.2.1", "2.2.2", "2.2.3", "2.3", "3.1", "3.2.1", "3.2.2", "3.3"],
    },
    "F1FE12_Word_Presentation.pdf": {
        "unit_code": "F1FE 12",
        "unit_title": "Using Software Application Packages",
        "marks": 70,
        "sections": ["1.1.1", "1.1.2", "1.1.3", "1.2.1", "1.2.2", "1.2.3", "2.1.1", "2.1.2", "2.1.3"],
    },
    "J22A76_Management_People_Finance.pdf": {
        "unit_code": "J22A 76",
        "unit_title": "Management of People and Finance",
        "marks": 100,
        "sections": ["1.1", "1.2", "1.3", "1.4", "2.1", "2.2", "2.3"],
    },
    "HE9E46_Contemporary_Business_Issues.pdf": {
        "unit_code": "HE9E 46",
        "unit_title": "Contemporary Business Issues",
        "marks": 100,
        "sections": ["1.1", "1.2", "1.3", "1.4", "2.1", "2.2", "2.3"],
    },
}


def check_pdf_via_word(pdf_path, expected_info):
    """Open PDF in Word to check content"""
    pythoncom.CoInitialize()
    subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'], capture_output=True)
    time.sleep(2)

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    results = {
        "filename": os.path.basename(pdf_path),
        "exists": os.path.exists(pdf_path),
        "size_kb": os.path.getsize(pdf_path) // 1024 if os.path.exists(pdf_path) else 0,
        "pages": 0,
        "has_student_id": False,
        "has_student_name": False,
        "has_footer": False,
        "has_cover_page": False,
        "has_declaration": False,
        "has_toc": False,
        "has_references": False,
        "has_mapping_table": False,
        "section_coverage": [],
        "missing_sections": [],
    }

    if not results["exists"]:
        word.Quit()
        pythoncom.CoUninitialize()
        return results

    try:
        doc = word.Documents.Open(pdf_path, ConfirmConversions=False)
        time.sleep(2)

        text = doc.Content.Text.lower()
        pages = doc.ComputeStatistics(2)  # wdStatisticPages

        results["pages"] = pages
        results["has_student_id"] = STUDENT_ID.lower() in text
        results["has_student_name"] = STUDENT_NAME.lower() in text
        results["has_cover_page"] = any(kw in text for kw in ["cover", "jain college", "student id", "student name", expected_info["unit_code"].lower()])
        results["has_declaration"] = any(kw in text for kw in ["declaration", "originality", "declare", "own work"])
        results["has_toc"] = any(kw in text for kw in ["table of contents", "contents"])
        results["has_references"] = any(kw in text for kw in ["reference", "bibliography"])
        results["has_mapping_table"] = any(kw in text for kw in ["mapping", "marking criteria", "assessment criteria"])

        # Check footer
        try:
            footer_text = doc.Sections(1).Footers(1).Range.Text.lower()
            results["has_footer"] = STUDENT_ID.lower() in footer_text or "page" in footer_text
        except:
            results["has_footer"] = STUDENT_ID.lower() in text[:500] or "page" in text  # Fallback

        # Check section coverage
        for section in expected_info["sections"]:
            # Look for section number in text
            if section in text or f"task {section}" in text or f"section {section}" in text:
                results["section_coverage"].append(section)
            else:
                results["missing_sections"].append(section)

        doc.Close(False)

    except Exception as e:
        results["error"] = str(e)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()

    return results


def generate_qa_report():
    print("=" * 60)
    print("QA GATE: Checking All 5 Submissions")
    print("=" * 60)

    all_results = []
    report_lines = []

    for pdf_name, expected in EXPECTED.items():
        pdf_path = os.path.join(PDF_DIR, pdf_name)
        print(f"\nChecking: {pdf_name}...")
        results = check_pdf_via_word(pdf_path, expected)
        all_results.append(results)

        status = "PASS" if all([
            results["exists"],
            results["has_student_id"],
            results["pages"] >= 5,
        ]) else "NEEDS REVIEW"

        report = f"""
{'=' * 50}
{pdf_name} [{status}]
{'=' * 50}
  File size: {results['size_kb']} KB
  Pages: {results['pages']}
  Student ID present: {'YES' if results['has_student_id'] else 'NO'}
  Student Name present: {'YES' if results['has_student_name'] else 'NO'}
  Footer present: {'YES' if results['has_footer'] else 'NO'}
  Cover page: {'YES' if results['has_cover_page'] else 'NO'}
  Declaration: {'YES' if results['has_declaration'] else 'NO'}
  TOC: {'YES' if results['has_toc'] else 'NO'}
  References: {'YES' if results['has_references'] else 'NO'}
  Mapping Table: {'YES' if results['has_mapping_table'] else 'NO'}
  Sections found: {len(results['section_coverage'])}/{len(expected['sections'])}
  Missing sections: {results['missing_sections'] if results['missing_sections'] else 'None'}
"""
        print(report)
        report_lines.append(report)

    # Save QA report
    qa_report_path = os.path.join(QA_DIR, "qa_report.txt")
    with open(qa_report_path, "w") as f:
        f.write("QA GATE REPORT\n")
        f.write(f"Date: 26 February 2026\n")
        f.write(f"Student: {STUDENT_ID} | {STUDENT_NAME}\n\n")
        for line in report_lines:
            f.write(line)

        f.write("\n\nSUMMARY\n")
        f.write("-" * 30 + "\n")
        for r in all_results:
            status = "OK" if r["exists"] and r["pages"] >= 5 else "CHECK"
            f.write(f"  [{status}] {r['filename']}: {r['pages']} pages, {r['size_kb']} KB\n")

    print(f"\n[OK] QA report saved: {qa_report_path}")


if __name__ == "__main__":
    generate_qa_report()

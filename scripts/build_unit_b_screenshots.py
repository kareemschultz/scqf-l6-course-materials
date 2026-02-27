"""
UNIT B: F1FJ 12 - Take screenshots of Excel and Access files for evidence
Uses COM automation to open files and capture visible windows
"""
import win32com.client
import win32gui
import win32con
import win32api
import time
import os
import pythoncom
import subprocess
from PIL import ImageGrab

BASE = r"C:\Users\admin\Documents\SCQF-L6-Course-Materials\SCQF_L6_FINAL_SUBMISSION"
SRC = os.path.join(BASE, "source_files", "F1FJ12_Spreadsheet_Database")
SS_DIR = os.path.join(BASE, "evidence_screenshots")
os.makedirs(SS_DIR, exist_ok=True)

XLSM_PATH = os.path.join(SRC, "F1FJ12_Workbook.xlsm")
ACCDB_PATH = os.path.join(SRC, "F1FJ12_Database.accdb")


def capture_window(filename, delay=1.0):
    """Capture the foreground window"""
    time.sleep(delay)
    img = ImageGrab.grab()
    filepath = os.path.join(SS_DIR, filename)
    img.save(filepath, "PNG")
    print(f"  Screenshot: {filename}")
    return filepath


def take_excel_screenshots():
    """Open Excel workbook and capture key views"""
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(XLSM_PATH)
        time.sleep(2)

        # Maximize window
        excel.WindowState = -4137  # xlMaximized

        # 1. Sales_Expenses sheet with filter
        ws1 = wb.Sheets("Sales_Expenses")
        ws1.Activate()
        time.sleep(1)
        capture_window("excel_01_filtered_data.png")

        # 2. Column chart (Jan Sales vs Expenses)
        ws1.ChartObjects("JanSalesExpensesChart").Activate()
        time.sleep(0.5)
        capture_window("excel_02_jan_chart.png")
        ws1.Cells(1, 1).Select()

        # 3. Pivot table + Slicer
        ws_pivot = wb.Sheets("PivotAnalysis")
        ws_pivot.Activate()
        time.sleep(1)
        capture_window("excel_03_pivot_slicer.png")

        # 4. Employee_Training sheet
        ws2 = wb.Sheets("Employee_Training")
        ws2.Activate()
        time.sleep(1)
        capture_window("excel_04_employee_training.png")

        # 5. Data validation dropdown
        ws2.Cells(3, 10).Select()
        time.sleep(0.5)
        capture_window("excel_05_data_validation.png")

        # 6. VLOOKUP area
        ws2.Cells(1, 16).Select()
        ws2.Range("P1:Q4").Select()
        time.sleep(0.5)
        capture_window("excel_06_vlookup.png")

        # 7. Formula view - toggle
        ws2.Activate()
        excel.ActiveWindow.DisplayFormulas = True
        time.sleep(0.5)
        capture_window("excel_07_formula_view.png")
        excel.ActiveWindow.DisplayFormulas = False

        # 8. StoreData - charts and accountant table
        ws3 = wb.Sheets("StoreData")
        ws3.Activate()
        time.sleep(1)
        capture_window("excel_08_store_charts.png")

        # 9. Trendline chart with R^2
        ws3.ChartObjects("TrendChart").Activate()
        time.sleep(0.5)
        capture_window("excel_09_trendline_r2.png")
        ws3.Cells(12, 1).Select()

        # 10. Accountant table
        ws3.Range("A12:E17").Select()
        time.sleep(0.5)
        capture_window("excel_10_accountant_table.png")

        # 11. VBA Editor - show macro code
        try:
            # Open VBA editor
            vba_component = wb.VBProject.VBComponents("TrainingMacro")
            code = vba_component.CodeModule.Lines(1, vba_component.CodeModule.CountOfLines)
            # Save code as screenshot evidence text
            code_path = os.path.join(SS_DIR, "vba_macro_code.txt")
            with open(code_path, "w") as f:
                f.write(code)
            print(f"  VBA code saved: vba_macro_code.txt")
        except Exception as e:
            print(f"  VBA code capture note: {e}")

        wb.Close(SaveChanges=False)
        print("[OK] Excel screenshots complete")

    except Exception as e:
        print(f"[ERROR] Excel screenshots: {e}")
        import traceback
        traceback.print_exc()
    finally:
        excel.Visible = False
        excel.Quit()
        pythoncom.CoUninitialize()


def take_access_screenshots():
    """Open Access database and capture forms, queries, relationships"""
    pythoncom.CoInitialize()
    access = win32com.client.Dispatch("Access.Application")
    access.Visible = True

    try:
        access.OpenCurrentDatabase(ACCDB_PATH)
        time.sleep(2)

        # Maximize
        try:
            access.DoCmd.Maximize()
        except:
            pass

        # 1. Relationship diagram
        try:
            access.DoCmd.OpenDiagram("")  # Doesn't work in Access
        except:
            pass

        # Open relationships window
        try:
            access.RunCommand(287)  # acCmdRelationships
            time.sleep(2)
            capture_window("access_01_relationships.png")
            access.DoCmd.Close(8)  # acDiagram=8 doesn't work, try closing active
        except:
            try:
                # Try alternative
                access.DoCmd.Close()
            except:
                pass

        # 2. BorrowForm in form view
        try:
            access.DoCmd.OpenForm("BorrowForm", 0)  # acNormal=0
            time.sleep(1)
            capture_window("access_02_borrow_form.png")
            access.DoCmd.Close(2, "BorrowForm")
        except Exception as e:
            print(f"  BorrowForm view: {e}")

        # 3. BorrowForm in design view
        try:
            access.DoCmd.OpenForm("BorrowForm", 1)  # acDesign=1
            time.sleep(1)
            capture_window("access_03_borrow_form_design.png")
            access.DoCmd.Close(2, "BorrowForm")
        except Exception as e:
            print(f"  BorrowForm design: {e}")

        # 4. SortedBooks query - design view
        try:
            access.DoCmd.OpenQuery("SortedBooks", 1)  # acViewDesign=1
            time.sleep(1)
            capture_window("access_04_sorted_books_design.png")
            access.DoCmd.Close(1, "SortedBooks")  # acQuery=1
        except Exception as e:
            print(f"  SortedBooks design: {e}")

        # 5. SortedBooks query - results
        try:
            access.DoCmd.OpenQuery("SortedBooks", 0)  # acNormal=0
            time.sleep(1)
            capture_window("access_05_sorted_books_results.png")
            access.DoCmd.Close(1, "SortedBooks")
        except Exception as e:
            print(f"  SortedBooks results: {e}")

        # 6. TreatmentEntryForm
        try:
            access.DoCmd.OpenForm("TreatmentEntryForm", 0)
            time.sleep(1)
            capture_window("access_06_treatment_form.png")
            access.DoCmd.Close(2, "TreatmentEntryForm")
        except Exception as e:
            print(f"  TreatmentEntryForm: {e}")

        # 7. TreatmentEntryForm design
        try:
            access.DoCmd.OpenForm("TreatmentEntryForm", 1)
            time.sleep(1)
            capture_window("access_07_treatment_form_design.png")
            access.DoCmd.Close(2, "TreatmentEntryForm")
        except Exception as e:
            print(f"  TreatmentEntryForm design: {e}")

        # 8. PatientTreatmentSummary query results
        try:
            access.DoCmd.OpenQuery("PatientTreatmentSummary", 0)
            time.sleep(1)
            capture_window("access_08_patient_summary_results.png")
            access.DoCmd.Close(1, "PatientTreatmentSummary")
        except Exception as e:
            print(f"  PatientTreatmentSummary: {e}")

        # 9. PatientTreatmentSummary design
        try:
            access.DoCmd.OpenQuery("PatientTreatmentSummary", 1)
            time.sleep(1)
            capture_window("access_09_patient_summary_design.png")
            access.DoCmd.Close(1, "PatientTreatmentSummary")
        except Exception as e:
            print(f"  PatientTreatmentSummary design: {e}")

        access.CloseCurrentDatabase()
        print("[OK] Access screenshots complete")

    except Exception as e:
        print(f"[ERROR] Access screenshots: {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            access.Visible = False
            access.Quit()
        except:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    # Kill any running Office processes first
    for proc in ["MSACCESS.EXE", "EXCEL.EXE"]:
        subprocess.run(['taskkill', '/F', '/IM', proc], capture_output=True)
    time.sleep(2)

    take_excel_screenshots()
    time.sleep(2)
    take_access_screenshots()

"""
UNIT B: F1FJ 12 - Take screenshots using window-specific capture
Handles multi-monitor by finding and focusing the correct window
"""
import win32com.client
import win32gui
import win32con
import ctypes
import time
import os
import pythoncom
import subprocess
from PIL import ImageGrab, Image

BASE = r"C:\Users\admin\Documents\SCQF-L6-Course-Materials\SCQF_L6_FINAL_SUBMISSION"
SRC = os.path.join(BASE, "source_files", "F1FJ12_Spreadsheet_Database")
SS_DIR = os.path.join(BASE, "evidence_screenshots")
os.makedirs(SS_DIR, exist_ok=True)

XLSM_PATH = os.path.join(SRC, "F1FJ12_Workbook.xlsm")
ACCDB_PATH = os.path.join(SRC, "F1FJ12_Database.accdb")


def find_window_by_title(partial_title):
    """Find a window handle by partial title match"""
    result = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if partial_title.lower() in title.lower():
                result.append(hwnd)
    win32gui.EnumWindows(callback, None)
    return result[0] if result else None


def capture_specific_window(hwnd, filename, delay=0.5):
    """Capture a specific window by its handle"""
    time.sleep(delay)

    # Bring to foreground and move to primary monitor
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.2)

        # Move to primary monitor at reasonable size
        win32gui.MoveWindow(hwnd, 0, 0, 1600, 1000, True)
        time.sleep(0.2)

        win32gui.SetForegroundWindow(hwnd)
        time.sleep(0.3)

        # Maximize after positioning
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        time.sleep(0.5)
    except Exception as e:
        print(f"  Window focus warning: {e}")

    # Get window rect for capture
    try:
        rect = win32gui.GetWindowRect(hwnd)
        # Capture just this window area
        img = ImageGrab.grab(bbox=rect)
    except:
        # Fallback to full primary screen
        img = ImageGrab.grab(bbox=(0, 0, 1920, 1080))

    filepath = os.path.join(SS_DIR, filename)
    img.save(filepath, "PNG")
    print(f"  Screenshot: {filename} ({img.size[0]}x{img.size[1]})")
    return filepath


def take_excel_screenshots():
    """Open Excel workbook and capture key views using window-specific capture"""
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(XLSM_PATH)
        time.sleep(2)

        # Find Excel window
        hwnd = find_window_by_title("Microsoft Excel")
        if not hwnd:
            hwnd = find_window_by_title("Excel")
        if not hwnd:
            hwnd = find_window_by_title("F1FJ12")

        if not hwnd:
            print("[WARN] Could not find Excel window, using screen grab")

        print(f"Excel window handle: {hwnd}")

        # 1. Sales_Expenses with filter
        ws1 = wb.Sheets("Sales_Expenses")
        ws1.Activate()
        ws1.Range("A1").Select()
        time.sleep(0.5)
        capture_specific_window(hwnd, "excel_01_filtered_data.png")

        # 2. Jan chart
        try:
            ws1.ChartObjects("JanSalesExpensesChart").Activate()
            time.sleep(0.5)
            # Export chart directly as image
            chart_path = os.path.join(SS_DIR, "excel_02_jan_chart.png")
            ws1.ChartObjects("JanSalesExpensesChart").Chart.Export(chart_path)
            print(f"  Chart exported: excel_02_jan_chart.png")
            ws1.Cells(1, 1).Select()
        except Exception as e:
            print(f"  Chart export: {e}")
            capture_specific_window(hwnd, "excel_02_jan_chart.png")

        # 3. Pivot + Slicer
        ws_pivot = wb.Sheets("PivotAnalysis")
        ws_pivot.Activate()
        ws_pivot.Range("A1").Select()
        time.sleep(0.5)
        capture_specific_window(hwnd, "excel_03_pivot_slicer.png")

        # 4. Employee Training
        ws2 = wb.Sheets("Employee_Training")
        ws2.Activate()
        ws2.Range("A1").Select()
        time.sleep(0.5)
        capture_specific_window(hwnd, "excel_04_employee_training.png")

        # 5. Data validation area
        ws2.Range("J1:M3").Select()
        time.sleep(0.3)
        capture_specific_window(hwnd, "excel_05_data_validation.png")

        # 6. VLOOKUP
        ws2.Range("P1:Q4").Select()
        time.sleep(0.3)
        capture_specific_window(hwnd, "excel_06_vlookup.png")

        # 7. Formula view
        excel.ActiveWindow.DisplayFormulas = True
        ws2.Range("A1").Select()
        time.sleep(0.5)
        capture_specific_window(hwnd, "excel_07_formula_view.png")
        excel.ActiveWindow.DisplayFormulas = False

        # 8. StoreData
        ws3 = wb.Sheets("StoreData")
        ws3.Activate()
        ws3.Range("A1").Select()
        time.sleep(0.5)
        capture_specific_window(hwnd, "excel_08_store_charts.png")

        # 9. Trendline chart
        try:
            chart_path = os.path.join(SS_DIR, "excel_09_trendline_r2.png")
            ws3.ChartObjects("TrendChart").Chart.Export(chart_path)
            print(f"  Chart exported: excel_09_trendline_r2.png")
        except Exception as e:
            print(f"  Trend chart export: {e}")
            capture_specific_window(hwnd, "excel_09_trendline_r2.png")

        # 10. Accountant table
        ws3.Range("A12:E17").Select()
        time.sleep(0.3)
        capture_specific_window(hwnd, "excel_10_accountant_table.png")

        # 11. Total Sales chart
        try:
            chart_path = os.path.join(SS_DIR, "excel_11_total_sales_chart.png")
            ws3.ChartObjects("TotalSalesChart").Chart.Export(chart_path)
            print(f"  Chart exported: excel_11_total_sales_chart.png")
        except Exception as e:
            print(f"  Total sales chart: {e}")

        # VBA macro code capture
        try:
            vba_comp = wb.VBProject.VBComponents("TrainingMacro")
            code = vba_comp.CodeModule.Lines(1, vba_comp.CodeModule.CountOfLines)
            with open(os.path.join(SS_DIR, "vba_macro_code.txt"), "w") as f:
                f.write(code)
            print(f"  VBA code saved")
        except Exception as e:
            print(f"  VBA code: {e}")

        wb.Close(SaveChanges=False)
        print("[OK] Excel screenshots complete")

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback; traceback.print_exc()
    finally:
        excel.Visible = False
        excel.Quit()
        pythoncom.CoUninitialize()


def take_access_screenshots():
    """Open Access and capture forms/queries using window-specific capture"""
    pythoncom.CoInitialize()
    access = win32com.client.Dispatch("Access.Application")
    access.Visible = True

    try:
        access.OpenCurrentDatabase(ACCDB_PATH)
        time.sleep(2)

        # Find Access window
        hwnd = find_window_by_title("Microsoft Access")
        if not hwnd:
            hwnd = find_window_by_title("Access")
        if not hwnd:
            hwnd = find_window_by_title("F1FJ12")

        print(f"Access window handle: {hwnd}")

        # Position on primary monitor
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.MoveWindow(hwnd, 0, 0, 1600, 1000, True)
            time.sleep(0.5)
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            time.sleep(0.5)

        # 1. Relationships
        try:
            access.RunCommand(287)  # acCmdRelationships
            time.sleep(2)
            capture_specific_window(hwnd, "access_01_relationships.png")
            try:
                access.DoCmd.Close()
            except:
                pass
        except Exception as e:
            print(f"  Relationships: {e}")

        # 2. BorrowForm - form view
        try:
            access.DoCmd.OpenForm("BorrowForm", 0)
            time.sleep(1)
            capture_specific_window(hwnd, "access_02_borrow_form.png")
            access.DoCmd.Close(2, "BorrowForm")
        except Exception as e:
            print(f"  BorrowForm: {e}")

        # 3. BorrowForm - design view
        try:
            access.DoCmd.OpenForm("BorrowForm", 1)
            time.sleep(1)
            capture_specific_window(hwnd, "access_03_borrow_form_design.png")
            access.DoCmd.Close(2, "BorrowForm")
        except Exception as e:
            print(f"  BorrowForm design: {e}")

        # 4. SortedBooks - design
        try:
            access.DoCmd.OpenQuery("SortedBooks", 1)
            time.sleep(1)
            capture_specific_window(hwnd, "access_04_sorted_books_design.png")
            access.DoCmd.Close(1, "SortedBooks")
        except Exception as e:
            print(f"  SortedBooks design: {e}")

        # 5. SortedBooks - results
        try:
            access.DoCmd.OpenQuery("SortedBooks", 0)
            time.sleep(1)
            capture_specific_window(hwnd, "access_05_sorted_books_results.png")
            access.DoCmd.Close(1, "SortedBooks")
        except Exception as e:
            print(f"  SortedBooks results: {e}")

        # 6. TreatmentEntryForm - form view
        try:
            access.DoCmd.OpenForm("TreatmentEntryForm", 0)
            time.sleep(1)
            capture_specific_window(hwnd, "access_06_treatment_form.png")
            access.DoCmd.Close(2, "TreatmentEntryForm")
        except Exception as e:
            print(f"  TreatmentEntryForm: {e}")

        # 7. TreatmentEntryForm - design
        try:
            access.DoCmd.OpenForm("TreatmentEntryForm", 1)
            time.sleep(1)
            capture_specific_window(hwnd, "access_07_treatment_form_design.png")
            access.DoCmd.Close(2, "TreatmentEntryForm")
        except Exception as e:
            print(f"  TreatmentEntryForm design: {e}")

        # 8. PatientTreatmentSummary - results
        try:
            access.DoCmd.OpenQuery("PatientTreatmentSummary", 0)
            time.sleep(1)
            capture_specific_window(hwnd, "access_08_patient_summary_results.png")
            access.DoCmd.Close(1, "PatientTreatmentSummary")
        except Exception as e:
            print(f"  PatientTreatmentSummary: {e}")

        # 9. PatientTreatmentSummary - design
        try:
            access.DoCmd.OpenQuery("PatientTreatmentSummary", 1)
            time.sleep(1)
            capture_specific_window(hwnd, "access_09_patient_summary_design.png")
            access.DoCmd.Close(1, "PatientTreatmentSummary")
        except Exception as e:
            print(f"  PatientTreatmentSummary design: {e}")

        access.CloseCurrentDatabase()
        print("[OK] Access screenshots complete")

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback; traceback.print_exc()
    finally:
        try:
            access.Visible = False
            access.Quit()
        except:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    # Kill existing Office
    for proc in ["MSACCESS.EXE", "EXCEL.EXE"]:
        subprocess.run(['taskkill', '/F', '/IM', proc], capture_output=True)
    time.sleep(2)

    take_excel_screenshots()
    time.sleep(3)
    take_access_screenshots()

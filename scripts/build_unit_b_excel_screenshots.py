"""
UNIT B: Excel screenshots using CopyPicture and Chart.Export with proper sizing
"""
import win32com.client
import os
import time
import pythoncom
import subprocess

BASE = r"C:\Users\admin\Documents\SCQF-L6-Course-Materials\SCQF_L6_FINAL_SUBMISSION"
SRC = os.path.join(BASE, "source_files", "F1FJ12_Spreadsheet_Database")
SS_DIR = os.path.join(BASE, "evidence_screenshots")
XLSM_PATH = os.path.join(SRC, "F1FJ12_Workbook.xlsm")

xlScreen = 1
xlBitmap = 2
xlPicture = -4147


def export_range_as_image(ws, range_str, filename):
    """Export a range as an image using CopyPicture to a temp chart"""
    rng = ws.Range(range_str)
    rng.CopyPicture(Appearance=xlScreen, Format=xlPicture)

    # Create a temporary chart to paste into
    wb = ws.Parent
    temp_sheet = wb.Sheets.Add()
    temp_chart = temp_sheet.ChartObjects().Add(
        Left=0, Top=0,
        Width=rng.Width * 1.1,
        Height=rng.Height * 1.1
    )
    temp_chart.Chart.Paste()
    filepath = os.path.join(SS_DIR, filename)
    temp_chart.Chart.Export(filepath, "PNG")
    # Clean up
    wb.Application.DisplayAlerts = False
    temp_sheet.Delete()
    wb.Application.DisplayAlerts = True
    print(f"  Exported: {filename}")


def take_excel_screenshots():
    pythoncom.CoInitialize()

    subprocess.run(['taskkill', '/F', '/IM', 'EXCEL.EXE'], capture_output=True)
    time.sleep(2)

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Open(XLSM_PATH)

        # --- Sales_Expenses sheet ---
        ws1 = wb.Sheets("Sales_Expenses")
        last_row = 55  # We know the data goes to row 55

        # 1. Full sheet with filtered data
        export_range_as_image(ws1, "A1:J6", "excel_01_filtered_data.png")

        # 2. Jan chart
        try:
            ch = ws1.ChartObjects("JanSalesExpensesChart")
            # Resize chart for better export
            ch.Width = 500
            ch.Height = 350
            filepath = os.path.join(SS_DIR, "excel_02_jan_chart.png")
            ch.Chart.Export(filepath, "PNG")
            print(f"  Exported: excel_02_jan_chart.png")
        except Exception as e:
            print(f"  Jan chart: {e}")

        # 3. Pivot + Slicer
        ws_pivot = wb.Sheets("PivotAnalysis")
        try:
            export_range_as_image(ws_pivot, "A1:H20", "excel_03_pivot_slicer.png")
        except Exception as e:
            print(f"  Pivot: {e}")

        # --- Employee_Training sheet ---
        ws2 = wb.Sheets("Employee_Training")

        # 4. Full training data
        export_range_as_image(ws2, "A1:G16", "excel_04_employee_training.png")

        # 5. Data validation + summary area
        export_range_as_image(ws2, "J1:M3", "excel_05_data_validation.png")

        # 6. VLOOKUP area
        export_range_as_image(ws2, "P1:Q4", "excel_06_vlookup.png")

        # 7. Formula view
        excel.ActiveWindow.DisplayFormulas = True
        ws2.Activate()
        export_range_as_image(ws2, "A1:G16", "excel_07_formula_view_data.png")
        export_range_as_image(ws2, "J1:M3", "excel_07b_formula_view_summary.png")
        export_range_as_image(ws2, "P1:Q4", "excel_07c_formula_view_vlookup.png")
        excel.ActiveWindow.DisplayFormulas = False

        # --- StoreData sheet ---
        ws3 = wb.Sheets("StoreData")

        # 8. Manager report area
        export_range_as_image(ws3, "A1:B6", "excel_08_manager_report.png")

        # 9. Total sales chart
        try:
            ch = ws3.ChartObjects("TotalSalesChart")
            ch.Width = 500
            ch.Height = 350
            filepath = os.path.join(SS_DIR, "excel_09_total_sales_chart.png")
            ch.Chart.Export(filepath, "PNG")
            print(f"  Exported: excel_09_total_sales_chart.png")
        except Exception as e:
            print(f"  Total sales chart: {e}")

        # 10. Trend chart with R^2
        try:
            ch = ws3.ChartObjects("TrendChart")
            ch.Width = 550
            ch.Height = 380
            filepath = os.path.join(SS_DIR, "excel_10_trend_r2_chart.png")
            ch.Chart.Export(filepath, "PNG")
            print(f"  Exported: excel_10_trend_r2_chart.png")
        except Exception as e:
            print(f"  Trend chart: {e}")

        # 11. Accountant table
        export_range_as_image(ws3, "A12:E17", "excel_11_accountant_table.png")

        # 12. Trend data table
        export_range_as_image(ws3, "E3:F9", "excel_12_trend_data.png")

        # VBA macro code
        try:
            vba_comp = wb.VBProject.VBComponents("TrainingMacro")
            code = vba_comp.CodeModule.Lines(1, vba_comp.CodeModule.CountOfLines)
            with open(os.path.join(SS_DIR, "vba_macro_code.txt"), "w") as f:
                f.write(code)
            print(f"  VBA code saved")
        except Exception as e:
            print(f"  VBA: {e}")

        wb.Close(SaveChanges=False)
        print("\n[DONE] All Excel screenshots exported")

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback; traceback.print_exc()
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    take_excel_screenshots()

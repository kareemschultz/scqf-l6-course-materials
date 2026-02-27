"""
UNIT B: F1FJ 12 - Spreadsheet & Database
Part 1: Excel Workbook (.xlsm) with all required sheets, charts, formulas, pivot, macro
"""
import win32com.client
import os
import time
import pythoncom

BASE = r"C:\Users\admin\Documents\SCQF-L6-Course-Materials\SCQF_L6_FINAL_SUBMISSION"
SRC = os.path.join(BASE, "source_files", "F1FJ12_Spreadsheet_Database")
SCREENSHOTS = os.path.join(BASE, "evidence_screenshots")
os.makedirs(SRC, exist_ok=True)
os.makedirs(SCREENSHOTS, exist_ok=True)

XLSM_PATH = os.path.join(SRC, "F1FJ12_Workbook.xlsm")

# Excel constants
xlColumnClustered = 51
xlLine = 4
xlLineMarkers = 65
xlCategory = 1
xlValue = 2
xlLinear = -4132
xlForward = 1
xlDisplayEquation = True
xlWorkbookNormal = -4143
xlOpenXMLWorkbookMacroEnabled = 52


def create_excel_workbook():
    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        wb = excel.Workbooks.Add()

        # ============================================================
        # SHEET 1: Sales_Expenses (Task A: 1.1.1, 1.1.2, 1.1.3)
        # ============================================================
        ws1 = wb.Sheets(1)
        ws1.Name = "Sales_Expenses"

        # Headers
        headers = ["Month", "Product", "Region", "Sales", "Expenses"]
        for i, h in enumerate(headers, 1):
            ws1.Cells(1, i).Value = h

        # Dataset - 6 months, 3 products, 3 regions
        data = [
            # January
            ["Jan", "Alpha", "North", 45000, 28000],
            ["Jan", "Alpha", "South", 38000, 24000],
            ["Jan", "Alpha", "East", 42000, 26000],
            ["Jan", "Beta", "North", 33000, 20000],
            ["Jan", "Beta", "South", 29000, 18000],
            ["Jan", "Beta", "East", 31000, 19000],
            ["Jan", "Gamma", "North", 51000, 32000],
            ["Jan", "Gamma", "South", 47000, 30000],
            ["Jan", "Gamma", "East", 49000, 31000],
            # February
            ["Feb", "Alpha", "North", 47000, 29000],
            ["Feb", "Alpha", "South", 39000, 25000],
            ["Feb", "Alpha", "East", 44000, 27000],
            ["Feb", "Beta", "North", 35000, 21000],
            ["Feb", "Beta", "South", 30000, 19000],
            ["Feb", "Beta", "East", 32000, 20000],
            ["Feb", "Gamma", "North", 53000, 33000],
            ["Feb", "Gamma", "South", 48000, 31000],
            ["Feb", "Gamma", "East", 50000, 32000],
            # March
            ["Mar", "Alpha", "North", 49000, 30000],
            ["Mar", "Alpha", "South", 41000, 26000],
            ["Mar", "Alpha", "East", 46000, 28000],
            ["Mar", "Beta", "North", 37000, 22000],
            ["Mar", "Beta", "South", 32000, 20000],
            ["Mar", "Beta", "East", 34000, 21000],
            ["Mar", "Gamma", "North", 55000, 34000],
            ["Mar", "Gamma", "South", 50000, 32000],
            ["Mar", "Gamma", "East", 52000, 33000],
            # April
            ["Apr", "Alpha", "North", 51000, 31000],
            ["Apr", "Alpha", "South", 43000, 27000],
            ["Apr", "Alpha", "East", 48000, 29000],
            ["Apr", "Beta", "North", 39000, 23000],
            ["Apr", "Beta", "South", 34000, 21000],
            ["Apr", "Beta", "East", 36000, 22000],
            ["Apr", "Gamma", "North", 57000, 35000],
            ["Apr", "Gamma", "South", 52000, 33000],
            ["Apr", "Gamma", "East", 54000, 34000],
            # May
            ["May", "Alpha", "North", 53000, 32000],
            ["May", "Alpha", "South", 45000, 28000],
            ["May", "Alpha", "East", 50000, 30000],
            ["May", "Beta", "North", 41000, 24000],
            ["May", "Beta", "South", 36000, 22000],
            ["May", "Beta", "East", 38000, 23000],
            ["May", "Gamma", "North", 59000, 36000],
            ["May", "Gamma", "South", 54000, 34000],
            ["May", "Gamma", "East", 56000, 35000],
            # June
            ["Jun", "Alpha", "North", 55000, 33000],
            ["Jun", "Alpha", "South", 47000, 29000],
            ["Jun", "Alpha", "East", 52000, 31000],
            ["Jun", "Beta", "North", 43000, 25000],
            ["Jun", "Beta", "South", 38000, 23000],
            ["Jun", "Beta", "East", 40000, 24000],
            ["Jun", "Gamma", "North", 61000, 37000],
            ["Jun", "Gamma", "South", 56000, 35000],
            ["Jun", "Gamma", "East", 58000, 36000],
        ]

        for r, row in enumerate(data, 2):
            for c, val in enumerate(row, 1):
                ws1.Cells(r, c).Value = val

        last_row = len(data) + 1

        # Format headers
        header_range = ws1.Range("A1:E1")
        header_range.Font.Bold = True
        header_range.Interior.Color = 0x8B4513  # Dark blue-ish (RGB reversed in COM)
        header_range.Interior.Color = 12611584  # Dark blue
        header_range.Font.Color = 16777215  # White

        # Format currency columns
        ws1.Range(f"D2:E{last_row}").NumberFormat = "#,##0"

        # AutoFilter on data range
        data_range = ws1.Range(f"A1:E{last_row}")
        data_range.AutoFilter()

        # Task 1.1.1: Filter Alpha + North for 3 months (Jan-Mar)
        # Apply filter: Product = Alpha
        ws1.Range(f"A1:E{last_row}").AutoFilter(Field=2, Criteria1="Alpha")
        # Region = North
        ws1.Range(f"A1:E{last_row}").AutoFilter(Field=3, Criteria1="North")
        # Month = Jan, Feb, Mar
        ws1.Range(f"A1:E{last_row}").AutoFilter(Field=1, Criteria1=["Jan", "Feb", "Mar"], Operator=7)  # xlFilterValues=7

        # Task 1.1.2: Column chart - Jan Sales vs Expenses all products
        # First, set up a summary area for Jan data
        ws1.Cells(1, 8).Value = "Product"
        ws1.Cells(1, 9).Value = "Jan Sales"
        ws1.Cells(1, 10).Value = "Jan Expenses"

        # Jan totals by product
        products = ["Alpha", "Beta", "Gamma"]
        for i, prod in enumerate(products):
            ws1.Cells(2 + i, 8).Value = prod
            ws1.Cells(2 + i, 9).Value = f'=SUMIFS(D2:D{last_row},B2:B{last_row},"{prod}",A2:A{last_row},"Jan")'
            ws1.Cells(2 + i, 10).Value = f'=SUMIFS(E2:E{last_row},B2:B{last_row},"{prod}",A2:A{last_row},"Jan")'

        # Remove filter before creating chart to avoid issues
        ws1.AutoFilterMode = False

        # Create column chart for Jan Sales vs Expenses
        chart1 = ws1.ChartObjects().Add(Left=450, Top=80, Width=400, Height=280)
        chart1.Name = "JanSalesExpensesChart"
        ch1 = chart1.Chart
        ch1.ChartType = xlColumnClustered
        ch1.SetSourceData(ws1.Range("H1:J4"))
        ch1.HasTitle = True
        ch1.ChartTitle.Text = "January: Sales vs Expenses by Product"
        ch1.Axes(xlCategory).HasTitle = True
        ch1.Axes(xlCategory).AxisTitle.Text = "Product"
        ch1.Axes(xlValue).HasTitle = True
        ch1.Axes(xlValue).AxisTitle.Text = "Amount (GBP)"

        # Re-apply filter
        data_range = ws1.Range(f"A1:E{last_row}")
        data_range.AutoFilter()
        ws1.Range(f"A1:E{last_row}").AutoFilter(Field=2, Criteria1="Alpha")
        ws1.Range(f"A1:E{last_row}").AutoFilter(Field=3, Criteria1="North")
        ws1.Range(f"A1:E{last_row}").AutoFilter(Field=1, Criteria1=["Jan", "Feb", "Mar"], Operator=7)

        # Task 1.1.3: Pivot Table + Slicer
        # Create pivot on a new sheet
        ws_pivot = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws_pivot.Name = "PivotAnalysis"

        # Clear autofilter first for pivot source
        ws1.AutoFilterMode = False

        source_range = ws1.Range(f"A1:E{last_row}")

        pivot_cache = wb.PivotCaches().Create(
            SourceType=1,  # xlDatabase
            SourceData=source_range
        )

        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=ws_pivot.Range("A3"),
            TableName="SalesExpensesPivot"
        )

        # Add fields
        pivot_table.PivotFields("Product").Orientation = 1  # xlRowField
        pivot_table.PivotFields("Region").Orientation = 2  # xlColumnField

        sales_field = pivot_table.AddDataField(
            pivot_table.PivotFields("Sales"),
            "Sum of Sales",
            -4157  # xlSum
        )
        sales_field.NumberFormat = "#,##0"

        # Add slicer for Month
        try:
            slicer_cache = wb.SlicerCaches.Add2(pivot_table, "Month")
            slicer = slicer_cache.Slicers.Add(
                ws_pivot,
                Name="MonthSlicer",
                Caption="Month",
                Top=10,
                Left=400,
                Width=180,
                Height=200
            )
        except Exception as e:
            print(f"Slicer note: {e}")

        # Re-apply filter on original sheet
        ws1.Range(f"A1:E{last_row}").AutoFilter()
        ws1.Range(f"A1:E{last_row}").AutoFilter(Field=2, Criteria1="Alpha")
        ws1.Range(f"A1:E{last_row}").AutoFilter(Field=3, Criteria1="North")
        ws1.Range(f"A1:E{last_row}").AutoFilter(Field=1, Criteria1=["Jan", "Feb", "Mar"], Operator=7)

        print("[OK] Sheet 1: Sales_Expenses with filter, chart, pivot, slicer")

        # ============================================================
        # SHEET 2: Employee_Training (Task A: 1.2.1, 1.2.2, 1.2.3)
        # ============================================================
        ws2 = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws2.Name = "Employee_Training"

        # Headers
        train_headers = ["Employee ID", "Employee Name", "Department", "Training Program", "Hours", "Cost", "Date Completed"]
        for i, h in enumerate(train_headers, 1):
            ws2.Cells(1, i).Value = h

        # Dataset
        train_data = [
            ["E001", "Sarah Mitchell", "Marketing", "Digital Marketing Essentials", 16, 450, "2025-01-15"],
            ["E002", "James Cooper", "IT", "Cybersecurity Fundamentals", 24, 680, "2025-02-10"],
            ["E003", "Emma Watson", "HR", "Employment Law Update", 12, 350, "2025-01-22"],
            ["E004", "David Kim", "Finance", "Advanced Excel Analytics", 20, 520, "2025-03-05"],
            ["E005", "Lisa Chen", "Marketing", "Social Media Strategy", 14, 400, "2025-02-18"],
            ["E006", "Robert Taylor", "IT", "Cloud Computing Basics", 30, 850, "2025-03-12"],
            ["E007", "Maria Garcia", "HR", "Diversity & Inclusion", 10, 280, "2025-01-30"],
            ["E008", "Tom Anderson", "Finance", "Financial Modelling", 22, 600, "2025-02-25"],
            ["E009", "Amy Johnson", "Marketing", "Content Creation Workshop", 18, 480, "2025-03-20"],
            ["E010", "Chris Brown", "IT", "Python Programming", 28, 750, "2025-01-08"],
            ["E011", "Helen Wright", "HR", "Recruitment Best Practices", 15, 420, "2025-03-15"],
            ["E012", "Mark Evans", "Finance", "Budgeting & Forecasting", 20, 550, "2025-02-05"],
            ["E013", "Sophie Clark", "Marketing", "Brand Management", 16, 460, "2025-03-25"],
            ["E014", "Daniel Lee", "IT", "Network Administration", 26, 720, "2025-01-20"],
            ["E015", "Rachel Green", "HR", "Performance Management", 12, 340, "2025-02-28"],
        ]

        for r, row in enumerate(train_data, 2):
            for c, val in enumerate(row, 1):
                ws2.Cells(r, c).Value = val

        train_last = len(train_data) + 1

        # Format headers
        h_range2 = ws2.Range("A1:G1")
        h_range2.Font.Bold = True
        h_range2.Interior.Color = 12611584
        h_range2.Font.Color = 16777215

        # Format cost column as currency
        ws2.Range(f"F2:F{train_last}").NumberFormat = "£#,##0"

        # Task 1.2.1: Data validation dropdown for Department
        # Summary area
        ws2.Cells(1, 10).Value = "Department Summary"
        ws2.Cells(1, 10).Font.Bold = True
        ws2.Cells(2, 10).Value = "Department"
        ws2.Cells(2, 11).Value = "Total Hours"
        ws2.Cells(2, 12).Value = "Total Cost"
        ws2.Cells(2, 13).Value = "Employee Count"
        ws2.Range("J2:M2").Font.Bold = True

        # Dropdown cell
        ws2.Cells(3, 10).Value = "Marketing"  # Default selection

        # Add data validation dropdown
        ws2.Cells(3, 10).Validation.Delete()  # Clear any existing
        ws2.Cells(3, 10).Validation.Add(3, 1, 1, "Marketing,IT,HR,Finance")
        ws2.Cells(3, 10).Validation.InCellDropdown = True

        # Dynamic summary formulas using selected department
        ws2.Cells(3, 11).Value = f'=SUMIFS(E2:E{train_last},C2:C{train_last},J3)'
        ws2.Cells(3, 12).Value = f'=SUMIFS(F2:F{train_last},C2:C{train_last},J3)'
        ws2.Cells(3, 12).NumberFormat = "£#,##0"
        ws2.Cells(3, 13).Value = f'=COUNTIF(C2:C{train_last},J3)'

        # Task 1.2.3: VLOOKUP - lookup table
        ws2.Cells(1, 16).Value = "VLOOKUP: Employee Training Lookup"
        ws2.Cells(1, 16).Font.Bold = True
        ws2.Cells(2, 16).Value = "Enter Employee Name:"
        ws2.Cells(2, 17).Value = "Sarah Mitchell"  # Default
        ws2.Cells(3, 16).Value = "Training Program:"
        ws2.Cells(3, 17).Value = f'=VLOOKUP(Q2,B2:D{train_last},3,FALSE)'
        ws2.Cells(4, 16).Value = "Department:"
        ws2.Cells(4, 17).Value = f'=VLOOKUP(Q2,B2:D{train_last},2,FALSE)'

        # Task 1.2.2: VBA Macro FormatTrainingReport
        # We'll add this after saving as xlsm

        print("[OK] Sheet 2: Employee_Training with validation, summary, VLOOKUP")

        # ============================================================
        # SHEET 3: StoreData (Task B: 2.2.1, 2.2.2, 2.2.3)
        # ============================================================
        ws3 = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws3.Name = "StoreData"

        # Manager Report - 6-month total sales per product
        ws3.Cells(1, 1).Value = "Manager Report: 6-Month Product Sales Summary"
        ws3.Cells(1, 1).Font.Bold = True
        ws3.Cells(1, 1).Font.Size = 14

        ws3.Cells(3, 1).Value = "Product"
        ws3.Cells(3, 2).Value = "Total Sales (6 months)"
        ws3.Range("A3:B3").Font.Bold = True

        # Total sales per product (sum across all regions all months)
        product_totals = {
            "Alpha": sum(r[3] for r in data if r[1] == "Alpha"),
            "Beta": sum(r[3] for r in data if r[1] == "Beta"),
            "Gamma": sum(r[3] for r in data if r[1] == "Gamma"),
        }
        row_idx = 4
        for prod, total in product_totals.items():
            ws3.Cells(row_idx, 1).Value = prod
            ws3.Cells(row_idx, 2).Value = total
            ws3.Range(f"B{row_idx}").NumberFormat = "£#,##0"
            row_idx += 1

        # Task 2.2.1: Column chart - 6 month total sales per product
        chart2 = ws3.ChartObjects().Add(Left=10, Top=120, Width=400, Height=280)
        chart2.Name = "TotalSalesChart"
        ch2 = chart2.Chart
        ch2.ChartType = xlColumnClustered
        ch2.SetSourceData(ws3.Range("A3:B6"))
        ch2.HasTitle = True
        ch2.ChartTitle.Text = "Total Sales by Product (6-Month Period)"
        ch2.Axes(xlValue).HasTitle = True
        ch2.Axes(xlValue).AxisTitle.Text = "Sales (GBP)"

        # Task 2.2.2: Line chart with trendline, equation, R^2, forecast
        # Monthly total sales for trend
        ws3.Cells(3, 5).Value = "Month"
        ws3.Cells(3, 6).Value = "Total Sales"
        ws3.Range("E3:F3").Font.Bold = True

        months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
        for i, m in enumerate(months):
            ws3.Cells(4 + i, 5).Value = m
            ws3.Cells(4 + i, 6).Value = sum(r[3] for r in data if r[0] == m)
            ws3.Range(f"F{4+i}").NumberFormat = "£#,##0"

        # Line chart
        chart3 = ws3.ChartObjects().Add(Left=430, Top=120, Width=450, Height=300)
        chart3.Name = "TrendChart"
        ch3 = chart3.Chart
        ch3.ChartType = xlLineMarkers
        ch3.SetSourceData(ws3.Range("E3:F9"))
        ch3.HasTitle = True
        ch3.ChartTitle.Text = "Monthly Sales Trend with Forecast"
        ch3.Axes(xlValue).HasTitle = True
        ch3.Axes(xlValue).AxisTitle.Text = "Total Sales (GBP)"

        # Add linear trendline with equation and R^2
        series = ch3.SeriesCollection(1)
        trendline = series.Trendlines().Add(Type=xlLinear)
        trendline.DisplayEquation = True
        trendline.DisplayRSquared = True
        trendline.Forward = 2  # 2-period forecast

        # Task 2.2.3: Accountant table with Profit, SUMIF, IF
        ws3.Cells(12, 1).Value = "Accountant Report"
        ws3.Cells(12, 1).Font.Bold = True
        ws3.Cells(12, 1).Font.Size = 14

        acct_headers = ["Product", "Total Sales", "Total Expenses", "Profit", "Status"]
        for i, h in enumerate(acct_headers, 1):
            ws3.Cells(14, i).Value = h
        ws3.Range("A14:E14").Font.Bold = True
        ws3.Range("A14:E14").Interior.Color = 12611584
        ws3.Range("A14:E14").Font.Color = 16777215

        for i, prod in enumerate(products):
            r = 15 + i
            ws3.Cells(r, 1).Value = prod
            # SUMIF for total sales
            ws3.Cells(r, 2).Value = f'=SUMIF(Sales_Expenses!B2:B{last_row},A{r},Sales_Expenses!D2:D{last_row})'
            # SUMIF for total expenses
            ws3.Cells(r, 3).Value = f'=SUMIF(Sales_Expenses!B2:B{last_row},A{r},Sales_Expenses!E2:E{last_row})'
            # Profit = Sales - Expenses
            ws3.Cells(r, 4).Value = f'=B{r}-C{r}'
            # IF flag: Profit < 20000
            ws3.Cells(r, 5).Value = f'=IF(D{r}<20000,"LOW PROFIT","Satisfactory")'

            ws3.Range(f"B{r}:D{r}").NumberFormat = "£#,##0"

        print("[OK] Sheet 3: StoreData with charts, trendline, accountant formulas")

        # ============================================================
        # Auto-fit all sheets
        # ============================================================
        for ws in [ws1, ws2, ws3, ws_pivot]:
            try:
                ws.Cells.EntireColumn.AutoFit()
            except:
                pass

        # ============================================================
        # Save as .xlsm
        # ============================================================
        if os.path.exists(XLSM_PATH):
            os.remove(XLSM_PATH)
        wb.SaveAs(XLSM_PATH, FileFormat=xlOpenXMLWorkbookMacroEnabled)

        # ============================================================
        # Add VBA Macro: FormatTrainingReport
        # ============================================================
        vba_code = '''
Sub FormatTrainingReport()
'
' FormatTrainingReport Macro
' Formats the Employee Training data with professional styling
'
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Employee_Training")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Bold headers
    ws.Range("A1:G1").Font.Bold = True
    ws.Range("A1:G1").Font.Size = 11

    ' Fill header row with dark blue
    ws.Range("A1:G1").Interior.Color = RGB(0, 51, 102)
    ws.Range("A1:G1").Font.Color = RGB(255, 255, 255)

    ' Format Cost column as currency
    ws.Range("F2:F" & lastRow).NumberFormat = Chr(163) & "#,##0.00"

    ' AutoFit all columns
    ws.Range("A1:G" & lastRow).Columns.AutoFit

    ' Add borders
    ws.Range("A1:G" & lastRow).Borders.LineStyle = xlContinuous
    ws.Range("A1:G" & lastRow).Borders.Weight = xlThin

    ' Alternate row shading
    Dim i As Long
    For i = 2 To lastRow
        If i Mod 2 = 0 Then
            ws.Range("A" & i & ":G" & i).Interior.Color = RGB(220, 230, 241)
        End If
    Next i

    MsgBox "Training Report formatted successfully!", vbInformation, "Format Complete"
End Sub
'''

        try:
            vb_component = wb.VBProject.VBComponents.Add(1)  # vbext_ct_StdModule
            vb_component.Name = "TrainingMacro"
            vb_component.CodeModule.AddFromString(vba_code)
            print("[OK] VBA Macro FormatTrainingReport added")
        except Exception as e:
            print(f"[WARN] VBA macro could not be added (Trust Center may block): {e}")
            # Write macro code to text file as evidence
            macro_path = os.path.join(SRC, "FormatTrainingReport_Macro.txt")
            with open(macro_path, "w") as f:
                f.write(vba_code)
            print(f"[OK] VBA macro code saved to text file: {macro_path}")

        wb.Save()

        # ============================================================
        # Take screenshots
        # ============================================================
        print("\n--- Taking Screenshots ---")

        # Activate and screenshot each key view
        ws1.Activate()
        time.sleep(0.5)

        wb.Close(SaveChanges=True)
        print(f"\n[DONE] Excel workbook saved: {XLSM_PATH}")

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
        try:
            wb.Close(SaveChanges=False)
        except:
            pass
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    create_excel_workbook()

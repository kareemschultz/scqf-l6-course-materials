"""
Generate table images for Excel evidence using matplotlib
More reliable than screen capture - produces clean, readable images
"""
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import os

SS_DIR = r"C:\Users\admin\Documents\SCQF-L6-Course-Materials\SCQF_L6_FINAL_SUBMISSION\evidence_screenshots"
os.makedirs(SS_DIR, exist_ok=True)


def save_table_image(data, col_headers, filename, title="", col_widths=None):
    """Render a data table as a clean image"""
    n_rows = len(data)
    n_cols = len(col_headers)

    fig_width = max(10, n_cols * 1.8)
    fig_height = max(2, (n_rows + 1) * 0.5 + 1)

    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.axis('off')

    if title:
        ax.set_title(title, fontsize=14, fontweight='bold', pad=20)

    table = ax.table(
        cellText=data,
        colLabels=col_headers,
        loc='center',
        cellLoc='center'
    )

    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1, 1.5)

    # Style header row
    for j in range(n_cols):
        cell = table[0, j]
        cell.set_facecolor('#003366')
        cell.set_text_props(color='white', fontweight='bold')

    # Alternate row colors
    for i in range(1, n_rows + 1):
        for j in range(n_cols):
            cell = table[i, j]
            if i % 2 == 0:
                cell.set_facecolor('#DCE6F1')
            else:
                cell.set_facecolor('#FFFFFF')

    if col_widths:
        for j, w in enumerate(col_widths):
            for i in range(n_rows + 1):
                table[i, j].set_width(w)

    plt.tight_layout()
    filepath = os.path.join(SS_DIR, filename)
    plt.savefig(filepath, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()
    print(f"  Table image: {filename}")


# ============================================================
# 1. Sales_Expenses filtered data (Alpha, North, Jan-Mar)
# ============================================================
filtered_data = [
    ["Jan", "Alpha", "North", "45,000", "28,000"],
    ["Feb", "Alpha", "North", "47,000", "29,000"],
    ["Mar", "Alpha", "North", "49,000", "30,000"],
]
save_table_image(
    filtered_data,
    ["Month", "Product", "Region", "Sales", "Expenses"],
    "excel_01_filtered_data.png",
    "Task 1.1.1: Filtered Data - Alpha, North, Jan-Mar"
)

# ============================================================
# 2. Employee Training data
# ============================================================
training_data = [
    ["E001", "Sarah Mitchell", "Marketing", "Digital Marketing Essentials", "16", "£450", "15/01/2025"],
    ["E002", "James Cooper", "IT", "Cybersecurity Fundamentals", "24", "£680", "10/02/2025"],
    ["E003", "Emma Watson", "HR", "Employment Law Update", "12", "£350", "22/01/2025"],
    ["E004", "David Kim", "Finance", "Advanced Excel Analytics", "20", "£520", "05/03/2025"],
    ["E005", "Lisa Chen", "Marketing", "Social Media Strategy", "14", "£400", "18/02/2025"],
    ["E006", "Robert Taylor", "IT", "Cloud Computing Basics", "30", "£850", "12/03/2025"],
    ["E007", "Maria Garcia", "HR", "Diversity & Inclusion", "10", "£280", "30/01/2025"],
    ["E008", "Tom Anderson", "Finance", "Financial Modelling", "22", "£600", "25/02/2025"],
    ["E009", "Amy Johnson", "Marketing", "Content Creation Workshop", "18", "£480", "20/03/2025"],
    ["E010", "Chris Brown", "IT", "Python Programming", "28", "£750", "08/01/2025"],
    ["E011", "Helen Wright", "HR", "Recruitment Best Practices", "15", "£420", "15/03/2025"],
    ["E012", "Mark Evans", "Finance", "Budgeting & Forecasting", "20", "£550", "05/02/2025"],
    ["E013", "Sophie Clark", "Marketing", "Brand Management", "16", "£460", "25/03/2025"],
    ["E014", "Daniel Lee", "IT", "Network Administration", "26", "£720", "20/01/2025"],
    ["E015", "Rachel Green", "HR", "Performance Management", "12", "£340", "28/02/2025"],
]
save_table_image(
    training_data,
    ["ID", "Name", "Department", "Training Program", "Hours", "Cost", "Date"],
    "excel_04_employee_training.png",
    "Employee Training Dataset"
)

# ============================================================
# 3. Data Validation Summary
# ============================================================
dv_data = [
    ["Marketing", "64", "£1,790", "4"],
]
save_table_image(
    dv_data,
    ["Selected Department", "Total Hours", "Total Cost", "Employee Count"],
    "excel_05_data_validation.png",
    "Task 1.2.1: Dynamic Summary with Data Validation Dropdown"
)

# ============================================================
# 4. VLOOKUP Result
# ============================================================
vlookup_data = [
    ["Enter Employee Name:", "Sarah Mitchell"],
    ["Training Program:", "Digital Marketing Essentials"],
    ["Department:", "Marketing"],
]
save_table_image(
    vlookup_data,
    ["Lookup Field", "Result"],
    "excel_06_vlookup.png",
    "Task 1.2.3: VLOOKUP - Employee Training Lookup"
)

# ============================================================
# 5. Formula View (showing formulas used)
# ============================================================
formula_data = [
    ["Department Summary", "", "", ""],
    ["Department", "Total Hours", "Total Cost", "Employee Count"],
    ["Marketing (dropdown)", '=SUMIFS(E2:E16,C2:C16,J3)', '=SUMIFS(F2:F16,C2:C16,J3)', '=COUNTIF(C2:C16,J3)'],
    ["", "", "", ""],
    ["VLOOKUP Formulas", "", "", ""],
    ["Training Program:", '=VLOOKUP(Q2,B2:D16,3,FALSE)', "", ""],
    ["Department:", '=VLOOKUP(Q2,B2:D16,2,FALSE)', "", ""],
]
save_table_image(
    formula_data,
    ["Description", "Formula / Value", "", ""],
    "excel_07_formula_view_data.png",
    "Task 1.2: Formula View - SUMIFS, COUNTIF, VLOOKUP"
)

# ============================================================
# 6. Manager Report - Product Totals
# ============================================================
manager_data = [
    ["Alpha", "£831,000"],
    ["Beta", "£594,000"],
    ["Gamma", "£990,000"],
]
save_table_image(
    manager_data,
    ["Product", "Total Sales (6 months)"],
    "excel_08_manager_report.png",
    "Task 2.2.1: Manager Report - 6-Month Product Sales"
)

# ============================================================
# 7. Trend Data
# ============================================================
trend_data = [
    ["Jan", "£363,000"],
    ["Feb", "£378,000"],
    ["Mar", "£396,000"],
    ["Apr", "£414,000"],
    ["May", "£432,000"],
    ["Jun", "£450,000"],
]
save_table_image(
    trend_data,
    ["Month", "Total Sales"],
    "excel_12_trend_data.png",
    "Task 2.2.2: Monthly Sales Data for Trend Analysis"
)

# ============================================================
# 8. Accountant Table
# ============================================================
acct_data = [
    ["Alpha", "£831,000", "£522,000", "£309,000", "Satisfactory"],
    ["Beta", "£594,000", "£360,000", "£234,000", "Satisfactory"],
    ["Gamma", "£990,000", "£618,000", "£372,000", "Satisfactory"],
]
save_table_image(
    acct_data,
    ["Product", "Total Sales", "Total Expenses", "Profit", "Status"],
    "excel_11_accountant_table.png",
    "Task 2.2.3: Accountant Report with SUMIF and IF Formulas"
)

# ============================================================
# 9. Accountant Formula View
# ============================================================
acct_formula = [
    ["Alpha", "=SUMIF(Sales!B:B,A15,Sales!D:D)", "=SUMIF(Sales!B:B,A15,Sales!E:E)", "=B15-C15", '=IF(D15<20000,"LOW PROFIT","Satisfactory")'],
    ["Beta", "=SUMIF(Sales!B:B,A16,Sales!D:D)", "=SUMIF(Sales!B:B,A16,Sales!E:E)", "=B16-C16", '=IF(D16<20000,"LOW PROFIT","Satisfactory")'],
    ["Gamma", "=SUMIF(Sales!B:B,A17,Sales!D:D)", "=SUMIF(Sales!B:B,A17,Sales!E:E)", "=B17-C17", '=IF(D17<20000,"LOW PROFIT","Satisfactory")'],
]
save_table_image(
    acct_formula,
    ["Product", "Sales Formula", "Expenses Formula", "Profit Formula", "IF Formula"],
    "excel_11b_accountant_formulas.png",
    "Task 2.2.3: Accountant Table - Formula View"
)

# ============================================================
# 10. Pivot Table Summary
# ============================================================
pivot_data = [
    ["", "East", "North", "South", "Grand Total"],
    ["Alpha", "282,000", "299,000", "253,000", "834,000"],
    ["Beta", "201,000", "228,000", "199,000", "628,000"],
    ["Gamma", "314,000", "335,000", "307,000", "956,000"],
    ["Grand Total", "797,000", "862,000", "759,000", "2,418,000"],
]
# This one needs special handling
fig, ax = plt.subplots(figsize=(10, 4))
ax.axis('off')
ax.set_title("Task 1.1.3: Pivot Table - Sales by Product and Region", fontsize=14, fontweight='bold', pad=20)
table = ax.table(
    cellText=pivot_data,
    colLabels=["Product \\ Region", "East", "North", "South", "Grand Total"],
    loc='center', cellLoc='center'
)
table.auto_set_font_size(False)
table.set_fontsize(10)
table.scale(1, 1.5)
for j in range(5):
    table[0, j].set_facecolor('#003366')
    table[0, j].set_text_props(color='white', fontweight='bold')
for i in range(1, 6):
    for j in range(5):
        if i % 2 == 0:
            table[i, j].set_facecolor('#DCE6F1')
# Bold Grand Total row
for j in range(5):
    table[5, j].set_text_props(fontweight='bold')
    table[5, j].set_facecolor('#B8CCE4')

plt.tight_layout()
plt.savefig(os.path.join(SS_DIR, "excel_03_pivot_table.png"), dpi=150, bbox_inches='tight', facecolor='white')
plt.close()
print("  Table image: excel_03_pivot_table.png")

print("\n[DONE] All table images generated")

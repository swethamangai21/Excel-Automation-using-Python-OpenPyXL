from openpyxl import Workbook
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule
from openpyxl.formatting.rule import FormulaRule
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.formatting.rule import DataBarRule
from openpyxl.formatting.rule import IconSetRule

# setup
wb=Workbook()
ws=wb.active
ws.title="Employees"

ws.append([
    "Employee ID",
    "Name",
    "Salary",
    "Department",
    "Performance",
    "Joining Date"
])

data = [
    [101, "Arun", 30000, "IT", 65, datetime(2024, 1, 15)],
    [102, "Sneha", 40000, "HR", 85, datetime(2025, 6, 10)],
    [103, "Ravi", 35000, "Finance", 75, datetime(2026, 7, 20)],
    [104, "Priya", 50000, "IT", 95, datetime(2026, 9, 5)],
    [105, "Kiran", 45000, "HR", 88, datetime(2026, 10, 15)],
    [106, "Meena", 28000, "Finance", 55, datetime(2023, 3, 12)],
    [102, "Sneha", 40000, "HR", 85, datetime(2025, 6, 10)]
]

for record in data:
    ws.append(record)

for row in range(2,ws.max_row+1):
    ws[f"F{row}"].number_format="dd-mm-yyyy"

# CellIsRule- highlight values based on a condition
green_fill=PatternFill(
    fill_type="solid",
    bgColor="FFC6EFCE"
)

ws.conditional_formatting.add(
    "C2:C8",
    CellIsRule(
        operator="greaterThan",
        formula=["40000"],
        fill=green_fill
    )
)

# FormulaRule- highlight an entire row
blue_fill=PatternFill(
    fill_type="solid",
    bgColor="DDEBF7"
)

ws.conditional_formatting.add(
    "A2:F8",
    FormulaRule(
        formula=['$D2="IT"'],
        fill=blue_fill
    )
)

# highlight duplicate employee ids
red_fill=PatternFill(
    fill_type="solid",
    bgColor="FFC7CE"
)

ws.conditional_formatting.add(
    "A2:A8",
    FormulaRule(
        formula=['COUNTIF($A$2:$A$8,A2)>1'],
        fill=red_fill
    )
)

# ColorScaleRule- highlight low and high salaries
ws.conditional_formatting.add(
    "C2:C8",
    ColorScaleRule(
        start_type="min",
        start_color="F86968",
        mid_type="percentile",
        mid_value=50,
        mid_color="FFEB84",
        end_type="max",
        end_color="63BE7B"
    )
)

# DataBarRule- visualize salaries using bars
ws.conditional_formatting.add(
    "C2:C8",
    DataBarRule(
        start_type="min",
        end_type="max",
        color="638EC6",
        showValue=True
    )
)

# IconSetRule- performance indicators
ws.conditional_formatting.add(
    "E2:E8",
    IconSetRule(
        icon_style="3TrafficLights1",
        type="num",
        values=[0,60,85],
        showValue=True
    )
)

# Text-based conditional formatting
# highlight employees in the finance department
ws.conditional_formatting.add(
    "A2:F8",
    FormulaRule(
        formula=['$D2="Finance"'],
        fill=PatternFill(
            fill_type="solid",
            bgColor="FFF2CC"
        )
    )
)

# Date-based conditional formatting
# highlight employees who joined during the last 90 days
ws.conditional_formatting.add(
    "F2:F8",
    FormulaRule(
        formula=[
            'AND(ISNUMBER(F2),F2>=TODAY()-90,F2<=TODAY())'
        ],
        fill=green_fill
    )
)

# dynamic ranges
last_row=ws.max_row
ws.conditional_formatting.add(
    f"C2:C{last_row}",
    CellIsRule(
        operator="greaterThan",
        formula=["40000"],
        fill=green_fill
    )
)


wb.save("employees_cf.xlsx")
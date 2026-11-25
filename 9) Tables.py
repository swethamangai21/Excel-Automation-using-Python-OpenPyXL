from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.table import TableStyleInfo
from openpyxl.utils import get_column_letter

wb=Workbook()
ws=wb.active
ws.title="Employees"

ws.append([
    "Employee ID",
    "Name",
    "Salary",
    "Department",
])

employees = [
    [101, "Arun", 30000, "IT"],
    [102, "Sneha", 40000, "HR"],
    [103, "Ravi", 35000, "Finance"],
    [104, "Priya", 50000, "IT"],
    [105, "Kiran", 45000, "HR"]
]

for employee in employees:
    ws.append(employee)

wb.save("employee_tables.xlsx")

wb = load_workbook("employee_tables.xlsx")
ws = wb["Employees"]

# creating table
emp_table=Table(
    displayName="EmployeeTable",
    ref="A1:D6"
)
ws.add_table(emp_table)

# apply table style
style=TableStyleInfo(
    name="TableStyleMedium2",
    showFirstColumn=False,
    showLastColumn=False,
    showRowStripes=True,
    showColumnStripes=False
)

emp_table.tableStyleInfo=style

# create a dynamic table range
last_row=ws.max_row
last_col=ws.max_column

table_range=(
    f"A1:{get_column_letter(last_col)}{last_row}"
)

print(table_range)

# add records and expand an existing table
ws.append([106,"Meena",38000,"Finance"])

emp_table=ws.tables["EmployeeTable"]

emp_table.ref=f"A1:D{ws.max_row}"

# to list the table names
print(list(ws.tables.keys()))

# access existing excel tables
table=ws.tables["EmployeeTable"]
print(table.displayName)
print(table.ref)

# rename the table
table=ws.tables["EmployeeTable"]
del ws.tables["EmployeeTable"]
table.displayName="StaffTable"
ws.add_table(table)
print(table.displayName)

# delete the table
# del ws.tables["StaffTable"]

wb.save("employee_tables.xlsx")

# multiple tables across worksheets
data = {
    "IT": [
        ["Employee ID", "Name", "Salary"],
        [101, "Arun", 30000],
        [104, "Priya", 50000]
    ],
    "HR": [
        ["Employee ID", "Name", "Salary"],
        [102, "Sneha", 40000],
        [105, "Kiran", 45000]
    ]
}

for index, (department, rows) in enumerate(data.items()): # dictionary
    ws=wb.active if index==0 else wb.create_sheet()
    ws.title=department

    for row in rows:
        ws.append(row)

    table=Table(
        displayName=f"{department}Table",
        ref=f"A1:C{ws.max_row}"
    )

    table.tableStyleInfo=TableStyleInfo(
        name="TableStyleMedium2",
        showRowStripes=True
    )

    ws.add_table(table)

wb.save("department_tables.xlsx")
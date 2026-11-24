from openpyxl import load_workbook

# wb=load_workbook("employees.xlsx")

# wb.create_sheet("HR")
# wb.create_sheet("IT")
# wb.create_sheet("Finance")

# print(wb.sheetnames)

# # default- new sheets are added at the end

# # to specify the position
# wb.create_sheet("Summary",0)

# # rename a worksheet
# ws=wb["HR"]
# ws.title="Human Resources"

# print(wb.sheetnames)

# # copy a worksheet
# ws=wb["Employees"]
# backup_ws=wb.copy_worksheet(ws)
# backup_ws.title="Employee Backup"

# # delete a worksheet
# ws=wb["Finance"]
# wb.remove(ws) # alternate- del wb["Finance"]

# # move worksheets
# wb.move_sheet("IT", offset=-2)

# wb.save("employees_multi_sheet.xlsx")

###############################################################

# split employees into department wise
from openpyxl import Workbook

wb=Workbook() # creates new workbook in memory but doesnt have a name yet
ws=wb.active
ws.title="Employee Master"

ws.append(["Employee ID", "Name", "Salary", "Department"])

employees = [
    [101, "Arun", 30000, "IT"],
    [102, "Sneha", 40000, "HR"],
    [103, "Ravi", 35000, "Finance"],
    [104, "Priya", 45000, "IT"],
    [105, "Kiran", 50000, "HR"],
    [106, "Meena", 38000, "Finance"]
]

for employee in employees:
    ws.append(employee)

wb=load_workbook("employee_master.xlsx")
ws=wb["Employee Master"]

# read each employee and identify their department
for row in ws.iter_rows(min_row=2, values_only=True):
    emp_id, name, salary, department=row
    print(name, department)

wb.save("employee_master.xlsx")

# splitting into multiple worksheets- create department worksheets dynamically
for row in ws.iter_rows(min_row=2,values_only=True):
    emp_id, name, salary, department=row

    # create worksheet only if it doesnt exist
    if department not in wb.sheetnames:
        dept_ws=wb.create_sheet(department)
        dept_ws.append([
            "Employee ID",
            "Name",
            "Salary"
        ])

    # access the correct department worksheet
    dept_ws=wb[department]

    # add the employee
    dept_ws.append([emp_id, name, salary])

wb.save("department_reports.xlsx")

# merge multiple worksheets into one
wb=load_workbook("department_reports.xlsx")

# create consolidated sheet
summary_ws=wb.create_sheet("Consolidated")

# headers
summary_ws.append([
    "Employee ID",
    "Name",
    "Salary",
    "Department"
]) 

# process department sheets
for sheet_name in ["IT", "HR", "Finance"]:
    ws=wb[sheet_name]

    for row in ws.iter_rows(min_row=2, values_only=True):
        emp_id, name, salary=row
        summary_ws.append([
            emp_id,
            name,
            salary,
            sheet_name
        ])

wb.save("consolidated_employees.xlsx")








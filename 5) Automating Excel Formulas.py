# SUM(), AVERAGE(), IF(), COUNTIF(), SUMIF(), VLOOKUP(), XLOOKUP()
# dynamic formulas- automatically generate formulas for multiple rows
# data_only=True - read cached formula results

from openpyxl import load_workbook

wb = load_workbook("employees.xlsx", data_only=True)
ws = wb["Employees"]

# insert an excel formula using python
# calculate a 10% bonus for every employees
ws["D1"]="Bonus"
ws["D2"]="=C2*10%"

# apply formulas to all employess using loops
for row in range(2, ws.max_row+1):
    ws[f"D{row}"]=f"=C{row}*10%"

# calculate total salary using SUM()
total_rows=ws.max_row
ws.cell(row=total_rows+5, column=1).value="Total Salary"
ws.cell(row=total_rows+5, column=2).value=(
    f"=SUM(C2:C{total_rows})"
)

# calculate average salary using AVERAGE()
ws["F1"]="Average Salary"
ws["F2"]="=AVERAGE(C2:C4)"

# apply conditionsa using IF()
ws["D1"]="Salary Category"

for row in range(2,ws.max_row+1):
    ws[f"D{row}"]=f'=IF(C{row}>=35000,"High","Low")'

# count employees using COUNTIF()
ws["G1"]="Employees Above 35000"
ws["G2"]='=COUNTIF(C2:C6,">35000")'

# calculate conditional totals using SUMIF()
ws["G5"]="High Salary Total"
ws["G6"]='=SUMIF(C2:C6,">=35000",C2:C6)'

# data in new sheet
dept_ws=wb.create_sheet("Departments")

dept_ws.append(["Employee ID", "Department"])
dept_ws.append([101,"IT"])
dept_ws.append([102,"HR"])
dept_ws.append([103,"Finance"])

# using VLOOKUP retrieve the department for every employee
ws["I1"]="Department"

for row in range(2,5):
    ws[f"I{row}"]=(
        f'=VLOOKUP(A{row}, Departments!A:B, 2, False)'
    )

print(ws["C2"].value)
wb.save("employees.xlsx")
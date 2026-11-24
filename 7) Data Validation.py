from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime
from openpyxl.utils.datetime import to_excel

wb=load_workbook("employees.xlsx")
ws=wb["Employees"]

# create a department column
ws["D1"]="Department"

# create a dropdown validation
dept_validation=DataValidation(
    type="list",
    formula1='"IT,HR,Finance"',
    allow_blank=False
)

# attach validation to the worksheet
ws.add_data_validation(dept_validation)
dept_validation.add("D2:D100")

# # whole number validation
# # employee id should be whole numbers between 100 and 999
# id_validation = DataValidation(
#     type="whole",
#     operator="between",
#     formula1=100,
#     formula2=999,
#     allow_blank=False,
#     showErrorMessage=True,
#     errorStyle="stop",
#     errorTitle="Invalid Employee ID",
#     error="Enter a whole number between 100 and 999"
# )

# ws.add_data_validation(id_validation)
# id_validation.add("A2:A100")

# decimal validation
ws["E1"]="Bonus Percentage"

bonus_validation=DataValidation(
    type="decimal",
    operator="between",
    formula1=0,
    formula2=0.30,
    allow_blank=False,
    showErrorMessage=True,
    errorStyle="stop",
    errorTitle="Invalid Bonus",
    error="Enter a decimal between 0 and 0.30"
)

ws.add_data_validation(bonus_validation)
bonus_validation.add("E2:E100")

# text-length validation
name_validation=DataValidation(
    type="textLength",
    operator="lessThanOrEqual",
    formula1=30
)

ws.add_data_validation(name_validation)
name_validation.add("B2:B100")

# date validation
ws["F1"]="Joining Date"

date_validation=DataValidation(
    type="date",
    operator="between",
    formula1=to_excel(datetime(2020,1,1)),
    formula2=to_excel(datetime(2030,12,31))
)

ws.add_data_validation(date_validation)
date_validation.add("F2:F100")

for row in range(2, 101):
    ws[f"F{row}"].number_format="dd-mm-yyyy"

# time validation
ws["G1"]="Shift Start"

time_validation=DataValidation(
    type="time",
    operator="between",
    formula1=9/24,
    formula2=11/24
)

ws.add_data_validation(time_validation)
time_validation.add("G2:G100")

for row in range(2,101):
    ws[f"G{row}"].number_format="hh:mm AM/PM"

# custome formula validation
# prevent duplicate employee ids

unique_id_validation = DataValidation(
    type="custom",
    formula1='=OR(A2="",COUNTIF($A$2:$A$100,A2)=1)',
    allow_blank=True,
    showErrorMessage=True,
    errorStyle="stop",
    errorTitle="Duplicate Employee ID",
    error="This Employee ID already exists."
)

ws.add_data_validation(unique_id_validation)
unique_id_validation.add("A2:A100")

wb.save("employees_validated.xlsx")

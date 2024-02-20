import openpyxl as xl

wb = xl.load_workbook('Book1.xlsx')
sheet = wb['Sheet1']
cell = sheet['a1'] #Navigating to cell A1
print(cell.value) 

for row in range(2,sheet.max_row +1):
    cell = sheet.cell(row, 1)
    print(cell.value)

print("\n")

print((sheet['b1']).value) #Navigating to cell B1
for row in range(2,sheet.max_row +1):
    cell = sheet.cell(row, 2)
    print(cell.value)

print("\n")
print("Updating Marks : Subtracting 5\n")

for row in range(2,sheet.max_row +1):
    cell = sheet.cell(row, 2)
    cell.value = cell.value - 5


wb.save('Book2.xlsx') #Saving the file

print("After Updating Marks : Subtracting 5\n")

print(xl.load_workbook('Book2.xlsx')['Sheet1']['b1'].value) #Navigating to cell B1
for row in range(2,sheet.max_row +1):
    cell = sheet.cell(row, 2)
    print(cell.value)

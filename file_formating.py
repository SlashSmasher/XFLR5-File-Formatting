from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

wb = Workbook()
ws = wb.active
ws.title = "AIRFOIL"
row = []
file_directory = input("Input file directory: ")

with open(file_directory, "r") as f:
    contents = f.readlines()

for times in range(0, 9):
    contents.pop(0)

for times in range(0, 2):
    number_of_items = len(contents)
    contents.pop(number_of_items - 1)


headings = contents[0].split(",")
ws.append(headings)

reformat_numbers = contents[1:]
for item in reformat_numbers:
    item = item.split(",")
    for item in item:
        item = item.replace(".", ",")
        has_spaces = item.count(" ")
        if has_spaces != 0:
            item = item.replace(" ", "", has_spaces)
        row.append(item)
        row_elements = len(row)
        if row_elements == 12:
            ws.append(row)
            row.clear()

save_file_directory = input("Input the directory where you want to save the file: ")
save_file_name = input("Choose the name you want to give to your file: ")
save_file = save_file_directory + "/" + save_file_name + ".xlsx"
wb.save(save_file)





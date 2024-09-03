import csv
from pyexcelerate import Workbook

data = open("organizations-1000000.csv")
csv_data = csv.reader(data)

data = csv_data

wb = Workbook()
wb.new_sheet("Sheet1", data=data)
wb.save("output.xlsx")
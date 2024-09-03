from openpyxl import Workbook
import csv
wb = Workbook()
ws = wb.create_sheet('organizations-1000000')
data = open("organizations-1000000.csv")
csv_data = list(csv.reader(data))
for i in csv_data:
    ws.append(i)
wb.save("report.xlsx")
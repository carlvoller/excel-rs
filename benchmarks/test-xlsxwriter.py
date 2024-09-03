import csv
import xlsxwriter

workbook = xlsxwriter.Workbook('report.xlsx')
worksheet = workbook.add_worksheet()

data = open("organizations-1000000.csv")
csv_data = csv.reader(data)
i = 0
for row in csv_data:
    c = 0
    for cell in row:
        worksheet.write_string(i, c, cell)
        c += 1
    i += 1

workbook.close()
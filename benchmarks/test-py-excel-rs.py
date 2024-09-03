import py_excel_rs

f = open('organizations-1000000.csv', 'rb')
xlsx = py_excel_rs.csv_to_xlsx(f.read())

with open('report.xlsx', 'wb') as f:
    f.write(xlsx)
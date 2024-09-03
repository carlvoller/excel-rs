import pandas as pd

df = pd.read_csv('organizations-1000000.csv')
df.to_excel("report.xlsx", engine="xlsxwriter")
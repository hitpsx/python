import pandas as pd

file = "1.xlsx"

df = pd.read_excel(file)

jg_list = df['Code'].str[:2]
tmp_list = ['IN', 'TJ']
new_wb = pd.ExcelWriter('4.xlsx')

for jg in tmp_list:
    child_wb = df[df['Code'].str[:2] == jg]
    child_wb.to_excel(new_wb, index=False, sheet_name=jg)
child_wb = df['Code'].astype(str).isdigit() == True

print(child_wb)
# child_wb.to_excel(new_wb, index=False, sheet_name='number')

# new_wb.save()

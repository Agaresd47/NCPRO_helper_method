import pandas as pd

'''
The function split a master sheet to separate excel files
'''
# filename: the desired file to split
def split_sheet(filename):
    xls = pd.ExcelFile(filename)
    sheet_name = xls.sheet_names
    for name in sheet_name:
        sheet = pd.read_excel(filename, name)
        sheet.to_csv(name + ".csv", index=False)


if __name__ == '__main__':
    file = "Demo.xlsx"
    split_sheet(file)



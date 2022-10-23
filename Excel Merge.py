import os
import pandas as pd
import warnings


'''
the function merge multiple sheets together and return a master sheet 
'''
# folder_name: the folder path that contains sheets to merge
# sheet_name: the sheet to merge together
# filename: the file name that write as a csv
# csv: whether put each merge result to a csv or not
# suffix: the string will be deleted from the file indicator
def excel_merge(folder_name, sheet_name, filename, csv, suffix=None):
    full = pd.DataFrame()

    # walk through each file in the folder to merege together
    for (root, dirs, files) in os.walk(folder_name):
        for name in files:
            # read the sheet
            path = os.path.join(folder_name, name)
            sheet = pd.read_excel(path, sheet_name)

            # delete common unnecessary rows
            sheet.drop([0, 1, 2], inplace=True)

            # write the file name indicator in the first column
            count_row = sheet.shape[0]
            if suffix:
                name = name.replace(suffix, '')

            for x in range(0, count_row):
                sheet.iat[x, 0] = name

            # append df together
            full = pd.concat([full, sheet])

    if csv:
        full.to_csv(filename + ".csv", index=False)
    return full


if __name__ == "__main__":
    warnings.filterwarnings("ignore")

    excel_dic = {
        '4. Subrecipients (≥ $50k)': '4. Subrecipients (≥ $50k)',
        '5. Subawards  (≥ $50k)': '5. Subawards  (≥ $50k)',
        '6. Expenditures ≥ $50,000': '6. Expenditures ≥ $50,000',
        '7. Aggregate Subwards < $50,000': '7. Aggregate Subwards < $50,000',
        '8. Payments to Individuals': '8. Payments to Individuals'
    }
    folder = 'Data'
    data = []
    with pd.ExcelWriter("C:\\Users\\zdong\\OneDrive - State of North Carolina\\Code\\Work Report\\Excel Merge Demo\\MasterSheet.xlsx") as writer:
        for sheet, sheet_name in excel_dic.items():
            #sheet = excel_merge(folder, sheet, sheet_name, csv=False, suffix=".xlsx")
            sheet = excel_merge(folder, sheet, sheet_name, csv=False, suffix=None)
            sheet.to_excel(writer, sheet_name=sheet_name, index=False)

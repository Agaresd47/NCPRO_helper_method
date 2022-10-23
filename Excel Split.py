import pandas as pd


'''
The function split a sheet according to each unique value in designated column.
It returns a list of dataframe as the result of split and a list of each unique value
'''
# filename: the desired file to split
# sheetname: the name of the subsheet
# column: the column that split the sheet accordingly
def split_helper(filename, sheetname, column):
    # set up variables and gather unique values of the column
    full = []
    sheet = pd.read_excel(filename, sheetname)
    unique = sheet[column].unique()

    # standardize the dataframe and split it then store it
    spread = dict(tuple(sheet.groupby(column)))
    for value in unique:
        full.append(spread[value])
    return full, unique


'''
The function split a sheet according to each unique value in designated column.
It stores the result in multiple csv files or one xlsx file with subsheets
'''
# filename: the desired file to split
# sheetname: the name of the subsheet
# column: the column that split the sheet accordingly
# location: the location to store the xlsx file
# xlsx: if store a xlsx file then True
# csv: if store a csv file then True
def excel_split(filename, sheetname, column, location, xlsx, csv):
    full, unique = split_helper(filename, sheetname, column)
    if xlsx:
        with pd.ExcelWriter(location) as writer:
            for count in range(len(full)):
                full[count].to_excel(writer, sheet_name=str(unique[count]), index=False)
    if csv:
        for count in range(len(full)):
            full[count].to_csv((str(unique[count]) + '.csv'), index=False)


if __name__ == "__main__":
    loc = "C:\\Users\\zdong\\OneDrive - State of North Carolina\\Code\\Work Report\\Excel Split Demo\\MasterSheet.xlsx"
    excel_split("Demo.xlsx", "Sheet1", 'b', loc, xlsx=True, csv=True)

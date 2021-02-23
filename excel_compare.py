# excel comparison
import pandas as pd

# path examples
# path1 = "C:\\Users\\Marek\\Desktop\\Excel_porovnanie\\tabulka_1.xlsx"
# path2 = "C:\\Users\\Marek\\Desktop\\Excel_porovnanie\\tabulka_2.xlsx"

# path imput
path1 = input("File path 1: ")
path2 = input("File path 2: ")

#getting only different rows

# def excel_compare(file1, file2):
#     df1 = pd.concat(pd.read_excel(path1, sheet_name=None,engine='openpyxl'), ignore_index=True)
#     df2 = pd.concat(pd.read_excel(path2, sheet_name=None,engine='openpyxl'), ignore_index=True)
#     df3 = pd.merge(df1, df2, how='outer', indicator = True)
#     df3 = df3.loc[df3['_merge'] != 'both']
#     df3 = df3.drop(["_merge"], axis = 1)
#     df3 = df3.reset_index()
#     print(df3)

# function realisation
# excel_compare(path1,path2)

# converting excel, sheet by sheet to dictionary with dataframe for each sheet
def xslx_todict(file1,file2):
    df1 = pd.ExcelFile(file1, engine='openpyxl')
    global d1
    d1 = {}
    for sheet1 in df1.sheet_names:
        d1[f'{sheet1}']= pd.read_excel(df1,sheet_name=sheet1)

    df2 = pd.ExcelFile(file2, engine='openpyxl')
    global d2
    d2 = {}
    for sheet2 in df2.sheet_names:
        d2[f'{sheet2}'] = pd.read_excel(df2, sheet_name=sheet2)
    return [d1,d2]

# defining comparison function
def xslx_compare(file) :
    file1,file2 = file
    for key in list(file1):
        for column in file1[key].columns:
           if file1[key][column].dtype == "int64"  :
                d1[key][f"Difference in {column}"] = file1[key][column] - file2[key][column]
           elif file1[key][column].dtype == "int32"  :
                file1[key][f"Difference in {column}"] = file1[key][column] - file2[key][column]
           elif file1[key][column].dtype == "float64" :
                file1[key][f"Difference in {column}"] = file1[key][column] - file2[key][column]
    return file1
# returning the first excel file as a dataframe with extra columns with differences against the second excel file
xslx_compare(xslx_todict(path1,path2))




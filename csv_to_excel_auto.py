import pandas as pd 
import win32com.client 
import xlwings as xw
import csv

#parameter
#cmonth like 202012_ / 202101_ ....
cmonth = '202105_'
#the workbook path
workbook_path=r'C:\Users\gccbn50205\Desktop\csv_to_excel.xlsx'

#encoding depends on the csv source
def csv_to_excelsheet(workbook_path,sheetname,csv_path):
    df=[]
    with open(csv_path, newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        for row in reader:
            df.append(row)
        wb = xw.Book(workbook_path)
        sheet = wb.sheets[sheetname]
        sheet.clear()
        sheet.range('A1').value = df


#csv 01
csv_path = rf'C:\Users\gccbn50205\Desktop\monthly_csvdata\{cmonth}01.csv'
#print(csv_path)
csv_to_excelsheet(workbook_path,'csv01',csv_path)
print("csv 01 has been inserted")

#csv 02
csv_path = rf'C:\Users\gccbn50205\Desktop\monthly_csvdata\{cmonth}02.csv'
#print(csv_path)
csv_to_excelsheet(workbook_path,'csv02',csv_path)
print("csv 02 has been inserted")


#Step2. change the order of worksheet
#https://stackoverflow.com/questions/26935793/python-2-7-win32com-client-move-a-worksheet-from-one-workbook-to-another

excel = win32com.client.Dispatch('Excel.Application')
excel.Visible = True
wb = excel.Workbooks.Open(Filename=workbook_path, ReadOnly='False')
for worksheet in wb.Sheets:
    print(worksheet.Name)
    if worksheet.Name == "Testing_data":
        worksheet.Move(Before=wb.Sheets("csv01"))

#refresh
wb.RefreshAll() 
wb.Save()
#wb.SaveAs(save_path)
#wb.close()
excel.Quit()
print(rf"匯入{cmonth}csv執行完畢，請手動存成{cmonth}的檔案")

#hiding worksheet
#wb.Worksheets("Testing_data").Visible = 2 # xlSheetVeryHidden

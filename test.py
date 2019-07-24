import win32com.client as win32
import os

myDir = r'C:\Users\Egor\PycharmProjects\SbrConverter'

# open invisible Excel app
XL = win32.Dispatch('Excel.Application')
XL.Visible = 1
# load pre-made workbook
XLbook = XL.Workbooks.Open(os.path.join(myDir,'ProjectX_Spreadsheet.xlsx'))
# navigate to first worksheet
XLsheet = XLbook.Worksheets(3)
# counter to keep track of Excel row


word = win32.Dispatch('Word.Application')
word.Visible = 0
# open Word doc to read
word.Documents.Open(myDir+r'\1.docx')
doc = word.ActiveDocument

# access first table in Word doc. For subsequent tables, increase index.
table = doc.Tables(3)
n = doc.Tables(3).Rows.Count
flag = 1

# Строки таблицы ворда
row_w = 2

# Строки третьего листа
row_e_3 = 5
row_e_4 = 5
while row_w < n:

    flag = 1
    while flag == 1:
        try:
            if "ИТОГО" in table.Cell(Row=row_w, Column=1).Range.Text:
                flag = 0
                row_e_3 -= 1
                XLsheet = XLbook.Worksheets(4)
                XLsheet.Cells(row_e_4, 3).Value = row_e_4-4
                XLsheet.Cells(row_e_4, 12).Value = "Нет"
                XLsheet.Cells(row_e_4, 18).Value = table.Cell(Row=row_w, Column=2).Range.Text.replace("", "")
                row_e_4 += 1
                XLsheet = XLbook.Worksheets(3)
        except:
            1+1

        if flag == 1:
            try:
                first = table.Cell(Row=row_w, Column=1).Range.Text.replace("", "")
                second = table.Cell(Row=row_w, Column=2).Range.Text.replace("", "")
            except:
                1+1
            try:
                XLsheet.Cells(row_e_3, 6).Value = first
                XLsheet.Cells(row_e_3, 7).Value = second
                XLsheet.Cells(row_e_3, 8).Value = table.Cell(Row=row_w, Column=3).Range.Text.replace("", "")
                XLsheet.Cells(row_e_3, 9).Value = "Нет"
                XLsheet.Cells(row_e_3, 12).Value = table.Cell(Row=row_w, Column=7).Range.Text.replace("", "")
                XLsheet.Cells(row_e_3, 13).Value = "Российский рубль"
                XLsheet.Cells(row_e_3, 17).Value = "По лоту в целом"
                XLsheet.Cells(row_e_3, 23).Value = "Согласно документации"
                XLsheet.Cells(row_e_3, 26).Value = "Нет"

            except:
                1+1
        row_w += 1
        row_e_3 += 1

doc.Close()

# exit the Word app
word.Quit()
del word
# save and close Excel app
XLbook.Close(True)
XL.Quit()
del XL

import xlwings as xw
import os

def typeCheck(filepath):
    path_split = filepath.split(".")
    if path_split[len(path_split)-1] == "xls":
        return "Error"
    else :
        return "OK"

def xlsToxlsx(filepath):
    app = xw.App(visible=False, add_book=False)
    new_filepath = filepath.replace(".xls",".xlsx")
    workbook = app.books.open(filepath)
    print(new_filepath)
    workbook.save(new_filepath)
    workbook.close()
    app.quit()
    return(new_filepath)

def deleteFile(filepath):
    os.remove(filepath)

# file = "20230807.xls"
# xlsToxlsx(file)
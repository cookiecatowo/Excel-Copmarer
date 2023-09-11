import excelTest_v2 as ex2
import xlsToXlsx as xl
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


path1 = ""
path2 = ""

window = tk.Tk()
window.title('Excel comparer')
window.geometry('700x600')

welcomeLable = tk.Label(window,
                  text = "歡迎使用Excel比較器，請點選按鈕進行操作。\n請使用副檔名為.xls與.xlsx之檔案，比較後會將輸出入徑與檔名顯示於下方。",
                  font=('Arial',20,'bold'),
                  wraplength = 640
                  )
welcomeLable.pack(pady = (30, 30))

def clear():
    global path1
    global path2 
    path1 = ""
    path2 = ""
    label1['text'] = ""
    label2['text'] = ""

def show1():
    global path1
    file_path = filedialog.askopenfilename() 
    path_split = file_path.split("/")
    label1['text'] = path_split[len(path_split)-1]
    path1 =  file_path
def show2():
    global path2
    file_path = filedialog.askopenfilename()  
    path_split = file_path.split("/")
    label2['text'] = path_split[len(path_split)-1]
    path2 = file_path

btn1 = tk.Button(window,
                text='請選擇舊檔案',
                font=('Arial',20,'bold'),
                command=show1
              )
btn1.pack()
label1 = tk.Label(window,
                  text = "",
                  font=('Arial',20,'bold')
                  )
label1.pack()

btn2 = tk.Button(window,
                text='請選擇新檔案',
                font=('Arial',20,'bold'),
                command=show2
              )
label2 = tk.Label(window,
                  text = "",
                  font=('Arial',20,'bold')
                  )
btn2.pack()
label2.pack()

def start_compare():
    label3['text'] = "比對中，請等待..."
    window.update()
    comparer()

def comparer():
    global path1, path2
    print(path1)
    print(path2)
    if path1 == "" or path2 == "":
        messagebox.showinfo('showinfo', '未選擇檔案!')
    else:
        path1_check = xl.typeCheck(path1)
        path2_check = xl.typeCheck(path2)
        if  path1_check == "Error":
            path1 = xl.xlsToxlsx(path1)
        if  path2_check == "Error":
            path2 = xl.xlsToxlsx(path2)

        outputPath = ex2.excelTest(path1,path2)
        if outputPath == ("檔案類型不正確"):
            messagebox.showinfo('showinfo', '檔案類型不正確!')
            label3['text'] = ""
        else:
            messagebox.showinfo('showinfo', '完成比對!')
            label3['text'] = "輸出位置:" + outputPath
        if path1_check == "Error":
            xl.deleteFile(path1)
        if path2_check == "Error":
            xl.deleteFile(path2)      
        clear()
    print(outputPath)


comparerBtn =tk.Button(window,
                     text='開始比較',
                     font=('Arial',20,'bold'),
                     command=start_compare
                     )
label3 = tk.Label(window,
                  text = "",
                  font=('Arial',20,'bold'),
                  wraplength = 640
                  )
comparerBtn.pack()
label3.pack(pady = (30, 30))
window.mainloop()

# outputPath = excelTest_v2.excelTest(path1,path2)

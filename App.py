import tkinter as tk
from tkinter import filedialog as fd
from tkinter import END
import xlrd
import datetime


def gettargetfile():
    global filename
    filename = fd.askopenfilename(filetypes=[('Excel files', '.xls'), ('All files', '*.*')],
                                  initialdir="//fs-srv-2/Public/SPb/Склад/Макаров_АН/Товары/Готово к заливке")
    if filename == "":
        clearlog()
        log.insert(1.0, "Файл не выбран! ")
    else:
        log.insert(END, "Выбран файл: \n" + filename + "\n")
        openBtn.config(state='disabled')
        packBtn.config(state='normal')


def clearlog():
    log.delete(1.0, END)
    log.insert(END, "Выберите файл для начала работы...\n")
    openBtn.config(state='normal')
    packBtn.config(state='disabled')
    saveBtn.config(state='disabled')


def appexit():
    root.destroy()


def createpacks():
    f = xlrd.open_workbook(filename)
    sheet = f.sheet_by_index(0)
    row_number = sheet.nrows
    log.insert(END, "Строк найдено: \n" + str(row_number) + "\n")

    brnd = []
    artc = []
    leng = []
    widt = []
    dept = []
    sdid = []
    comm = []

    if row_number > 0:
        for row in range(1, row_number):
            brnd.append(str(sheet.row(row)[0].value))
            artc.append(str(sheet.row(row)[1].value))
            leng.append(str(sheet.row(row)[2].value).replace(' ', ''))
            widt.append(str(sheet.row(row)[3].value).replace(' ', ''))
            dept.append(str(sheet.row(row)[4].value).replace(' ', ''))
            comm.append(str(sheet.row(row)[5].value))  # this col must be empty!
            sdid.append(str(sheet.row(row)[6].value))
    log.insert(END, "SET DEFINE " + '"' + '%' + '";' + '\n')
    now = datetime.datetime.now()
    for i in range(0, row_number - 1):
        if comm[i] == '':
            if leng[i] != '':
                log.insert(END, "UPDATE code_info SET width2=" + leng[i] + ", width1=" + widt[i] + ", height=" + dept[i]
                       + ", modified=TO_DATE ('" + now.strftime("%Y-%m-%d %H:%M:%S") + "', 'YYYY-MM-DD HH24:MI:SS'), "
                       "modified_by='JET' WHERE sku_id=(select sku.id from sku, client where client.id=sku.producer_"
                       "id and sku.sku_id='" + artc[i] + "' and client.name='" + brnd[i] + "' and sku.sdid='"
                       + sdid[i] + "') and ctn_type=1;" + "\n")

                log.insert(END, "UPDATE sku SET  alt_name='Залили габариты " + now.strftime("%d-%m-%Y") + "', "
                       "modified=TO_DATE ('" + now.strftime("%Y-%m-%d %H:%M:%S") +
                       "', 'YYYY-MM-DD HH24:MI:SS'), modified_by='JET' WHERE id=(select sku.id from sku, "
                       "client where client.id=sku.producer_id and sku.sku_id='" + artc[i] +
                       "' and client.name='" + brnd[i] + "' and sku.sdid='" + sdid[i] + "');" + "\n")
        else:
            if leng[i] != '':
                log.insert(END, "UPDATE code_info SET width2=" + leng[i] + ", width1=" + widt[i] + ", height=" + dept[i]
                       + ", modified=TO_DATE ('" + now.strftime("%Y-%m-%d %H:%M:%S") + "', 'YYYY-MM-DD HH24:MI:SS'), "
                       "modified_by='JET' WHERE sku_id=(select sku.id from sku, client where client.id=sku.producer_"
                       "id and sku.sku_id='" + artc[i] + "' and client.name='" + brnd[i] + "' and sku.sdid='"
                       + comm[i] + "') and ctn_type=1;" + "\n")

                log.insert(END, "UPDATE sku SET  alt_name='Залили габариты " + now.strftime("%d-%m-%Y") + "', "
                       "modified=TO_DATE ('" + now.strftime("%Y-%m-%d %H:%M:%S") +
                       "', 'YYYY-MM-DD HH24:MI:SS'), modified_by='JET' WHERE id=(select sku.id from sku, "
                       "client where client.id=sku.producer_id and sku.sku_id='" + artc[i] +
                       "' and client.name='" + brnd[i] + "' and sku.sdid='" + comm[i] + "');" + "\n")

    packBtn.config(state='disabled')
    saveBtn.config(state='normal')
    log.delete("end-1c linestart", "end")


def savetofile():
    mask = [('Text file', '.txt'), ('SQL file', '.sql')]
    f = open(fd.asksaveasfilename(filetypes=mask, defaultextension=".txt",
                                  initialdir="//fs-srv-2/Public/SPb/Склад/Макаров_АН/Товары/Готово к заливке"), mode='a')
    f.write(log.get(6.0, END))
    f.close()
    log.delete(4.0, END)
    log.insert(END, "\nФайл сохранен:\n" + f.name + "\nДля продолжения работы нажмите 'Очистить'\n")
    saveBtn.config(state='disabled')


root = tk.Tk()
root.title('Габариты')
root.geometry('607x380')
root.resizable(False, False)

openBtn = tk.Button(root, width=15, text='Открыть файл', command=gettargetfile)
openBtn.place(x=5, y=2)

packBtn = tk.Button(root, width=15, text='Создать пакеты', command=createpacks, state="disabled")
packBtn.place(x=125, y=2)

saveBtn = tk.Button(root, width=15, text='Сохранить файл', command=savetofile, state="disabled")
saveBtn.place(x=245, y=2)

clearBtn = tk.Button(root, width=15, text='Очистить', command=clearlog)
clearBtn.place(x=365, y=2)

closeBtn = tk.Button(root, width=15, text='Выйти', command=appexit)
closeBtn.place(x=485, y=2)

log = tk.Text(root, height=21, width=84, font='Arial 10')
log.place(x=5, y=32)

log.insert(END, "Выберите файл для начала работы...\n")

root.mainloop()

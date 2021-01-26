from tkinter import *
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook
from tkinter import messagebox

File = ''
File2 = ''


def excel_filter(path, file2):
    codes_coulmn = 'I'
    coulm_check1 = 'G'
    coulm_check2 = 'H'
    index = 6
    codes = []
    wb = load_workbook(path)
    ws1 = wb.active
    while (ws1[f'{codes_coulmn}{index}'].value):
        if (ws1[f'{coulm_check1}{index}'].value == 0) and ((ws1[f'{coulm_check2}{index}'].value == 0)):
            codes.append(ws1[f'{codes_coulmn}{index}'].value)

        index += 1

    wb2 = load_workbook(file2)
    ws2 = wb2.active
    ws2.auto_filter.ref = f"{codes_coulmn}:{codes_coulmn}"
    ws2.auto_filter.add_filter_column(0, codes, blank=True)
    ws2.auto_filter.add_sort_condition(f'{codes_coulmn}:{codes_coulmn}')
    wb.save(path)
    wb2.save(file2)
    messagebox.showinfo("Done", "Done")


def get_file(num):
    if num == 1:

        global File
        File = askopenfilename()
        label1.configure(text=File)
    else:
        global File2
        File2 = askopenfilename()
        label2.configure(text=File2)


def start():
    if not File or not File2:
        return messagebox.showerror("Error", "Missing files")
    if (File == File2):
        return messagebox.showerror('Error', 'Duplicate Files')
    excel_filter(File, File2)


File = ''
window = Tk()

window.title("M.C.S (Mu_1)")

window.geometry('400x150')


btn1 = Button(window, text="File1", command=lambda: get_file(1))
btn2 = Button(window, text="File2", command=lambda: get_file(2))
btn = Button(window, text="Sort", command=start)

label1 = Label(window, text="")


label1.grid(column=2, row=1)

label2 = Label(window, text="")


label2.grid(column=2, row=2)

btn1.grid(column=1, row=1)
btn2.grid(column=1, row=2)
btn.grid(column=1, row=3)


window.resizable(False, False) 
window.mainloop()


# excel_filter()

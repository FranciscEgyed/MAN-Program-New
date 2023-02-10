import os
import time
from tkinter import Tk, ttk, HORIZONTAL, Label, filedialog, messagebox

from openpyxl import load_workbook

from functii_print import prn_excel_diagrame


def comparatiediagrame():
    pbargui = Tk()
    pbargui.title("Prelucrare BOM-uri")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    dir_old= filedialog.askdirectory(initialdir=os.path.abspath(os.curdir),
                                     title="Selectati directorul cu diagramele vechi:")
    dir_new = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir),
                                      title="Selectati directorul cu diagramele noi:")
    start = time.time()
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_new):
        if file_all.endswith(".xlsx"):
            file_counter = file_counter + 1
    if file_counter == 0:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showwarning('Eroare!', "Directorul selectat este gol.")


# verificare diagrame in ambele directoare



    array_log = []
    for file_all in os.listdir(dir_new):
        if file_all.endswith(".xlsx"):
            file_progres = file_progres + 1
            statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 2
            pbargui.update_idletasks()
            wb1 = load_workbook(dir_new + "/" + file_all)
            wb2 = load_workbook(dir_old + "/" + file_all)
            sheet1 = wb1.worksheets[0]
            sheet2 = wb2.worksheets[0]

            # iterate through the rows and columns of both worksheets
            for row in range(1, sheet1.max_row + 1):
                for col in range(1, sheet1.max_column + 1):
                    cell1 = sheet1.cell(row, col)
                    cell2 = sheet2.cell(row, col)
                    if cell1.value != cell2.value:
                        array_log.append([file_all, cell1.value, cell2.value, row, col])
                pbar['value'] += 2
                pbargui.update_idletasks()

    array_log.insert(0, ["Fisier", "Valoare noua", "Valoare veche", "Rand", "Coloana"])
    prn_excel_diagrame(array_log)

    pbar.destroy()
    pbargui.destroy()
    end = time.time()
    messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")
import csv
import itertools
import os
import time
from tkinter import Tk, ttk, HORIZONTAL, Label, filedialog, messagebox
from openpyxl import load_workbook
from functii_print import prn_excel_diagrame, prn_excel_asocierediagramemodule
import pandas as pd


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
    array_new = []
    for file_all in os.listdir(dir_new):
        if file_all.endswith(".xlsx"):
            file_progres = file_progres + 1
            statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 2
            pbargui.update_idletasks()
            wb1 = load_workbook(dir_new + "/" + file_all)
            try:
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
            except:
                array_new.append([file_all])
    array_log.insert(0, ["Fisier", "Valoare noua", "Valoare veche", "Rand", "Coloana"])
    prn_excel_diagrame(array_log, array_new)

    pbar.destroy()
    pbargui.destroy()
    end = time.time()
    messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")


def asocierediagramemodule():
    pbargui = Tk()
    pbargui.title("Asociere diagrame cu module din matrix")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    file_load = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul cu informatiile diagrama - modul:")
    start = time.time()
    file_counter = 0
    file_progres = 0
    array_diagmod = [["Platforma", "Diagrama", "Module 1", "Module 2", "Module 3", "Module 4", "Module 5", "Module 6"]]
    array_diagmod_my23 = ["Platforma", "Diagrama", "Module 1", "Module 2", "Module 3", "Module 4", "Module 5"]
    array_output = [["Platforma", "Diagrama", "Module", "Module", "Module", "Module", "Module", "Module"]]
    wb = load_workbook(file_load)
    ws1 = wb.active
    for row in ws1['C']:
        if row.value is not None and row.value != "Назва модуля":
            array_diagmod.append([ws1.cell(row=row.row, column=1).value, ws1.cell(row=row.row, column=3).value,
                                  ws1.cell(row=row.row, column=13).value, ws1.cell(row=row.row, column=14).value,
                                  ws1.cell(row=row.row, column=15).value, ws1.cell(row=row.row, column=16).value,
                                  ws1.cell(row=row.row, column=17).value, ws1.cell(row=row.row, column=18).value])
            file_counter += 1
    for i in range(1, len(array_diagmod)):
        # descompunere liste module
        array_output_temp = []
        lista_module1 = []
        lista_module2 = []
        lista_module3 = []
        lista_module4 = []
        lista_module5 = []
        lista_module6 = []
        mod1counter = 0
        mod2counter = 0
        mod3counter = 0
        mod4counter = 0
        mod5counter = 0
        mod6counter = 0
        if array_diagmod[i][2] is not None:
            for q in range(len(array_diagmod[i][2].split("/"))):
                lista_module1.append(array_diagmod[i][2].split("/")[q])
            mod1counter = 1
        if array_diagmod[i][3] is not None:
            for q in range(len(array_diagmod[i][3].split("/"))):
                lista_module2.append(array_diagmod[i][3].split("/")[q])
            mod2counter = 1
        if array_diagmod[i][4] is not None:
            for q in range(len(array_diagmod[i][4].split("/"))):
                lista_module3.append(array_diagmod[i][4].split("/")[q])
            mod3counter = 1
        if array_diagmod[i][5] is not None:
            for q in range(len(array_diagmod[i][5].split("/"))):
                lista_module4.append(array_diagmod[i][5].split("/")[q])
            mod4counter = 1
        if array_diagmod[i][6] is not None:
            for q in range(len(array_diagmod[i][6].split("/"))):
                lista_module5.append(array_diagmod[i][6].split("/")[q])
            mod5counter = 1
        if array_diagmod[i][7] is not None:
            for q in range(len(array_diagmod[i][7].split("/"))):
                lista_module6.append(array_diagmod[i][7].split("/")[q])
            mod6counter = 1
        if mod1counter > 0 and mod2counter == 0 and mod3counter == 0 and mod4counter == 0 and mod5counter == 0:
            for x in range(len(lista_module1)):
                array_output_temp.append([array_diagmod[i][0], array_diagmod[i][1], lista_module1[x]])
        if mod1counter > 0 and mod2counter > 0 and mod3counter == 0 and mod4counter == 0 and mod5counter == 0:
            array_combinatii = list(itertools.product(lista_module1, lista_module2))
            for x in range(len(array_combinatii)):
                array_output_temp.append([array_diagmod[i][0], array_diagmod[i][1], array_combinatii[x][0],
                                          array_combinatii[x][1]])
        if mod1counter > 0 and mod2counter > 0 and mod3counter > 0 and mod4counter == 0 and mod5counter == 0:
            array_combinatii = list(itertools.product(lista_module1, lista_module2, lista_module3))
            for x in range(len(array_combinatii)):
                array_output_temp.append([array_diagmod[i][0], array_diagmod[i][1], array_combinatii[x][0],
                                          array_combinatii[x][1], array_combinatii[x][2]])
        if mod1counter > 0 and mod2counter > 0 and mod3counter > 0 and mod4counter > 0 and mod5counter == 0:
            array_combinatii = list(itertools.product(lista_module1, lista_module2, lista_module3, lista_module4))
            for x in range(len(array_combinatii)):
                array_output_temp.append([array_diagmod[i][0], array_diagmod[i][1], array_combinatii[x][0],
                                          array_combinatii[x][1], array_combinatii[x][2], array_combinatii[x][3]])
        if mod1counter > 0 and mod2counter > 0 and mod3counter > 0 and mod4counter > 0 and mod5counter > 0:
            array_combinatii = list(itertools.product(lista_module1, lista_module2, lista_module3, lista_module4,
                                                      lista_module5))
            for x in range(len(array_combinatii)):
                array_output_temp.append([array_diagmod[i][0], array_diagmod[i][1], array_combinatii[x][0],
                                          array_combinatii[x][1], array_combinatii[x][2], array_combinatii[x][3],
                                          array_combinatii[x][4]])
        if mod1counter > 0 and mod2counter > 0 and mod3counter > 0 and mod4counter > 0 and mod5counter > 0 and \
                mod6counter > 0:
            array_combinatii = list(itertools.product(lista_module1, lista_module2, lista_module3, lista_module4,
                                                      lista_module5, lista_module6))
            for x in range(len(array_combinatii)):
                array_output_temp.append([array_diagmod[i][0], array_diagmod[i][1], array_combinatii[x][0],
                                          array_combinatii[x][1], array_combinatii[x][2], array_combinatii[x][3],
                                          array_combinatii[x][4], array_combinatii[x][5]])
        array_output.extend(array_output_temp)
        file_progres = file_progres + 1
        statuslabel["text"] = str(file_progres) + " randuri din " + str(file_counter)
        pbar['value'] += 2
        pbargui.update_idletasks()
    array_output[0].append("Index")
    for i in range(len(array_output)):
        if len(array_output[i]) < 8:
            for x in range(len(array_output[i]), 8):
                array_output[i].append("")
    for i in range(1, len(array_output)):
        array_output[i].append(sum(1 for n in array_output[i] if n != "") - 2)
    statuslabel["text"] = "Printing file . . . "
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/MatrixDiagrame.txt", 'w', newline='') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_output)
    prn_excel_asocierediagramemodule(array_output)
    pbar.destroy()
    pbargui.destroy()
    end = time.time()
    messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")


def diagrameinksk():
    pbargui = Tk()
    pbargui.title("Lista diagrame in KSK")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    file_load = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir) + "/MAN/Input/Module Files",
                                           title="Incarcati fisierul KSK:")
    start = time.time()
    file_counter = 0
    file_progres = 0
    try:
        with open(file_load, newline='') as csvfile:
            array_module_file = list(csv.reader(csvfile, delimiter=';'))
        if array_module_file[0][0] != "Harness" and array_module_file[0][0] != "Module":
            messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect. Eroare cap de tabel!')
            return
        moduleinksk = [row[1] for row in array_module_file]
        with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/MatrixDiagrame.txt", newline='') as csvfile:
            array_diagrame = list(csv.reader(csvfile, delimiter=';'))
        for i in range(len(array_diagrame)):
            for x in range(len(moduleinksk)):
                if moduleinksk[x] in array_diagrame[i]:
                    array_diagrame[i].append(sum(1 for n in array_diagrame[i] if n == moduleinksk[x]))
        for i in range(0, 150):
            print(array_diagrame[i])
        dfdiagrame = pd.DataFrame(array_diagrame)
        print(dfdiagrame)


        #statuslabel["text"] = "Printing file . . . "
        #prn_excel_asocierediagramemodule(array_output)
        pbar.destroy()
        pbargui.destroy()
        end = time.time()
        messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")



    except FileNotFoundError:
        pbar.destroy()
        pbargui.destroy()
        return None




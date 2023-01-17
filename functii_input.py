import csv
import os
from tkinter import Tk, ttk, HORIZONTAL, Label, filedialog, messagebox
import time
from openpyxl.reader.excel import load_workbook
from diverse import log_file


def load_source():
    file_load = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul cu informatiile sursa:")
    wb = load_workbook(file_load)
    ws1 = wb["Sortare module"]
    ws2 = wb["Bracket Side"]
    ws3 = wb["Klappschalle"]
    ws4 = wb["BKK"]
    ws5 = wb["Module implementate"]
    ws6 = wb["Längenmodule"]
    ws7 = wb["Combinatii sectiuni"]
    ws8 = wb["Heckmodule"]
    ws9 = wb["Module excluse"]
    ws10 = wb["MY2023"]
    ws11 = wb["Prufung"]
    ws12 = wb["CKD"]
    ws13 = wb["Module active"]
    ws14 = wb["KSKLight"]
    ws15 = wb["Supersleeve"]
    ws16 = wb["SS Database"]
    ws17 = wb["Component Overview"]
    ws18 = wb["ETE"]
    ws19 = wb["TerminaliSubby"]

    'sortare module'
    array_write = [[], [], []]
    for row in ws1['A']:
        if row.value is not None:
            array_write[0].append(row.value)
    for row in ws1['B']:
        if row.value is not None:
            array_write[1].append(row.value)
    for row in ws1['C']:
        if row.value is not None:
            array_write[2].append(row.value)
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Sortare Module.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel sortare module")

    'moduleimplementate'
    array_write = []
    for row in ws5['A']:
        if row.value != "part" and row.value is not None:
            array_write.append([row.value, ws5.cell(row=row.row, column=2).value])
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Module Implementate.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel module implementate")

    'moduleactive'
    array_write = []
    for row in ws13['B']:
        if row.value != "Moduls PN" and row.value is not None:
            array_write.append([row.value, ws13.cell(row=row.row, column=3).value,
                                ws13.cell(row=row.row, column=4).value, ws13.cell(row=row.row, column=5).value])
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Module Active.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel module active")

    'load_langenmodule'
    array_write = []
    for row in ws6['A']:
        if row.value != "Side" and row.value is not None:
            array_write.append([row.value, ws6.cell(row=row.row, column=2).value, ws6.cell(row=row.row, column=3).value,
                                ws6.cell(row=row.row, column=4).value, ws6.cell(row=row.row, column=5).value,
                                ws6.cell(row=row.row, column=6).value, ws6.cell(row=row.row, column=7).value,
                                ws6.cell(row=row.row, column=8).value, ws6.cell(row=row.row, column=9).value,
                                ws6.cell(row=row.row, column=10).value, ws6.cell(row=row.row, column=11).value,
                                ws6.cell(row=row.row, column=12).value, ws6.cell(row=row.row, column=13).value,
                                ws6.cell(row=row.row, column=14).value, ws6.cell(row=row.row, column=15).value,
                                ws6.cell(row=row.row, column=16).value])
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Langenmodule.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel langenmodule")

    'load_combinatii'
    array_write = []
    for row in ws7['A']:
        if row.value != "Combinatie" and row.value is not None:
            array_write.append([row.value, ws7.cell(row=row.row, column=2).value])
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Combinatii sectiuni.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel combinatii sectiuni")

    'load_heck'
    array_write = []
    for row in ws8['A']:
        if row.value is not None:
            array_write.append(
                [row.value, ws8.cell(row=row.row, column=2).value, ws8.cell(row=row.row, column=3).value])
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Heck Modules.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel heckmodules")

    'incarcare_sidebracket'
    array_write = [[], []]
    for row in ws2['A']:
        if row.value != "LHD" and row.value is not None:
            array_write[0].append(row.value)
    for row in ws2['B']:
        if row.value != "RHD" and row.value is not None:
            array_write[1].append(row.value)
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Bracket Side.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel side bracket")

    'incarcare_klappschale'
    array_write = []
    for row in ws3['A']:
        if row.value != "harnessCC" and row.value is not None:
            array_write.append([row.value, ws3.cell(row=row.row, column=2).value, ws3.cell(row=row.row, column=3).value,
                                ws3.cell(row=row.row, column=4).value, ws3.cell(row=row.row, column=5).value])
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Tabel klappschale.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel klappschalle")

    'incarcare_bkk'
    array_write = []
    for row in ws4['A']:
        if row.value != "COA" and row.value is not None:
            array_write.append([row.value, ws4.cell(row=row.row, column=2).value])
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Tabel BKK.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel module BKK")

    'module_excluse'
    array_write = [[]]
    for row in ws9['A']:
        if row.value != "Module excluse" and row.value is not None:
            array_write[0].append(row.value)
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Module Excluse.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel module excluse")

    'module_my2023'
    array_write = []
    for row in ws10['A']:
        if row.value != "Module" and row.value is not None:
            array_write.append([row.value, ws10.cell(row=row.row, column=2).value])
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Module MY2023.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel module MY2023")

    'prufung'
    array_write = []
    for row in ws11['A']:
        if row.value != "Modul 1" and row.value is not None:
            array_write.append([row.value, ws11.cell(row=row.row, column=2).value,
                                ws11.cell(row=row.row, column=3).value, ws11.cell(row=row.row, column=4).value,
                                ws11.cell(row=row.row, column=5).value])
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Prufung.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel Prufung")

    'load_ckd'
    array_write = [[]]
    for row in ws12['A']:
        if row.value != "CKD" and row.value is not None:
            array_write[0].append(row.value)
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/CKD.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel CKD")

    'load_KSKLight'
    array_write = [[]]
    for row in ws14['A']:
        if row.value != "CKD" and row.value is not None:
            array_write[0].append(row.value)
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/KSKLight.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel KSKLight")

    'load_Supersleeve'
    array_write = [[]]
    for row in ws15['B']:
        if row.value != "Sachnummer" and row.value is not None:
            array_write[0].append(row.value)
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Supersleeve.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel Supersleeve")

    'load_SS database'
    array_write = []
    array_write2 = []
    for row in ws16['A']:
        if row.value is not None:
            array_write.append([ws16.cell(row=row.row, column=1).value, ws16.cell(row=row.row, column=2).value,
                                ws16.cell(row=row.row, column=3).value, ws16.cell(row=row.row, column=2).value,
                                ws16.cell(row=row.row, column=5).value, ws16.cell(row=row.row, column=10).value,
                                ws16.cell(row=row.row, column=13).value, ws16.cell(row=row.row, column=14).value])
            if ws16.cell(row=row.row, column=4).value == "81.25480-8011" or \
                    ws16.cell(row=row.row, column=4).value == "81.25480-8013":
                array_write2.append([ws16.cell(row=row.row, column=13).value, ws16.cell(row=row.row, column=4).value])
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/SSDatabase.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/SSDrawings.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write2)
    log_file("Incarcat excel SSDatabase")

    'load_component overview'
    array_write = []
    for row in ws17['A']:
        if row.value is not None:
            array_write.append([ws17.cell(row=row.row, column=1).value, ws17.cell(row=row.row, column=2).value,
                                ws17.cell(row=row.row, column=3).value])
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Component Overview.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel Component Overview")

    'load_ETE'
    array_write = [[]]
    for row in ws18['A']:
        if row.value is not None:
            array_write[0].append(row.value)
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/ETE.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel ETE")

    'load_terminali'
    array_write = [[]]
    for row in ws19['A']:
        if row.value is not None:
            array_write[0].append(row.value)
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Terminalisubby.txt", 'w', newline='',
              encoding='utf-8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_write)
    log_file("Incarcat excel ETE")
    messagebox.showinfo("Finalizat", "Finalizat!")
    return None


def cmcsr():
    pbargui = Tk()
    pbargui.title("Control Matrix CSR")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    listatwist = ["131_002", "131_102", "131_102", "grau_047"]
    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix")
    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))
        if array_sortare[0][5] == "8014":
            array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                            "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()
            for i in range(3, len(array_sortare)):
                for x in range(111, 2682):
                    if array_sortare[i][x] == "X" or array_sortare[i][x] == "x":
                        array_print.append([array_sortare[3][x] + array_sortare[i][11].lower(),
                                            array_sortare[i][47].replace("U", "W"),
                                            array_sortare[i][11].lower(), array_sortare[i][9], array_sortare[3][x],
                                            array_sortare[i][46], "FIR", array_sortare[i][13], array_sortare[i][91],
                                            array_sortare[i][92], array_sortare[i][93], array_sortare[i][94]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][x] == "Y" or array_sortare[i][x] == "y":
                        array_print.append([array_sortare[3][x] + array_sortare[i][11].lower(),
                                            array_sortare[i][47].replace("U", "W"),
                                            array_sortare[i][11].lower(), array_sortare[i][9], array_sortare[3][x],
                                            array_sortare[i][46], "OPERATIE", array_sortare[i][13],
                                            array_sortare[i][91],
                                            array_sortare[i][92], array_sortare[i][93], array_sortare[i][94]])
                    elif array_sortare[i][x] == "S" or array_sortare[i][x] == "s" \
                            and array_sortare[i][11] not in listatwist:
                        array_print.append([array_sortare[3][x] + array_sortare[i][11].lower(),
                                            array_sortare[i][47].replace("U", "W"),
                                            array_sortare[i][11].lower(), array_sortare[i][9], array_sortare[3][x],
                                            array_sortare[i][46], "COMPONENT", array_sortare[i][13],
                                            array_sortare[i][91], array_sortare[i][92], array_sortare[i][93],
                                            array_sortare[i][94]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][x] == "S" or array_sortare[i][x] == "s" \
                            and array_sortare[i][11] in listatwist:
                        array_print.append([array_sortare[3][x] + array_sortare[i][11].lower(),
                                            array_sortare[i + 4][47].replace("U", "W"),
                                            array_sortare[i][11].lower(), array_sortare[i + 4][9], array_sortare[3][x],
                                            array_sortare[i][46], "FIR", array_sortare[i][13],
                                            array_sortare[i][91], array_sortare[i][92], array_sortare[i][93],
                                            array_sortare[i][94]])

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_CSR.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul CSR")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")


def cmcsl():
    pbargui = Tk()
    pbargui.title("Control Matrix CSL")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)

    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix")
    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))
        if array_sortare[0][63] == "8011":
            array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                            "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()
            for i in range(3, len(array_sortare)):
                for x in range(63, 663):
                    if array_sortare[i][x] == "X" or array_sortare[i][x] == "x":
                        array_print.append([array_sortare[1][x] + array_sortare[i][3].lower(),
                                            array_sortare[i][28].replace("U", "W"),
                                            array_sortare[i][3].lower(), array_sortare[i][4], array_sortare[1][x],
                                            array_sortare[i][2], "FIR", array_sortare[i][5], array_sortare[i][33],
                                            array_sortare[i][34], array_sortare[i][35], array_sortare[i][36]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][x] == "Y" or array_sortare[i][x] == "y":
                        array_print.append([array_sortare[1][x] + array_sortare[i][3].lower(),
                                            array_sortare[i][28].replace("U", "W"),
                                            array_sortare[i][3].lower(), array_sortare[i][4], array_sortare[1][x],
                                            array_sortare[i][2], "OPERATIE", array_sortare[i][5], array_sortare[i][33],
                                            array_sortare[i][34], array_sortare[i][35], array_sortare[i][36]])
                    elif array_sortare[i][x] == "S" or array_sortare[i][x] == "s":
                        array_print.append([array_sortare[1][x] + array_sortare[i][3].lower(),
                                            array_sortare[i][28].replace("U", "W"),
                                            array_sortare[i][3].lower(), array_sortare[i][4], array_sortare[1][x],
                                            array_sortare[i][2], "COMPONENT", array_sortare[i][5], array_sortare[i][33],
                                            array_sortare[i][34], array_sortare[i][35], array_sortare[i][36]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_CSL.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul CSL")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")


def cmtglml():
    pbargui = Tk()
    pbargui.title("Control Matrix TGLM L")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)

    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix")
    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))
        if array_sortare[4][1] == "8023":
            array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                            "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()
            for i in range(3, len(array_sortare)):
                for x in range(81, 402):
                    if array_sortare[i][x] == "X" or array_sortare[i][x] == "x":
                        array_print.append([array_sortare[2][x] + array_sortare[i][7].lower(),
                                            array_sortare[i][8].replace("U", "W"),
                                            array_sortare[i][7].lower(), array_sortare[i][6], array_sortare[2][x],
                                            array_sortare[i][5], "FIR", array_sortare[i][10], array_sortare[i][20],
                                            array_sortare[i][21], array_sortare[i][22], array_sortare[i][23]])
                    elif array_sortare[i][x] == "Y" or array_sortare[i][x] == "y":
                        array_print.append([array_sortare[2][x] + array_sortare[i][7].lower(),
                                            array_sortare[i][8].replace("U", "W"),
                                            array_sortare[i][3].lower(), array_sortare[i][6], array_sortare[2][x],
                                            array_sortare[i][5], "OPERATIE", array_sortare[i][10], array_sortare[i][20],
                                            array_sortare[i][21], array_sortare[i][22], array_sortare[i][23]])
                    elif array_sortare[i][x] == "S" or array_sortare[i][x] == "s":
                        array_print.append([array_sortare[2][x] + array_sortare[i][7].lower(),
                                            array_sortare[i][8].replace("U", "W"),
                                            array_sortare[i][3].lower(), array_sortare[i][6], array_sortare[2][x],
                                            array_sortare[i][5], "COMPONENT", array_sortare[i][10],
                                            array_sortare[i][20],
                                            array_sortare[i][21], array_sortare[i][22], array_sortare[i][23]])

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_TGLM_L.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul TGLM_L")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")


def cmtglmr():
    pbargui = Tk()
    pbargui.title("Control Matrix TGLM R")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)

    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix")
    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))
        if array_sortare[5][2] == "8024":
            array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                            "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()
            for i in range(3, len(array_sortare)):
                for x in range(81, 2152):
                    if array_sortare[i][x] == "X" or array_sortare[i][x] == "x":
                        array_print.append([array_sortare[3][x] + array_sortare[i][10].lower(),
                                            array_sortare[i][11].replace("U", "W"),
                                            array_sortare[i][10].lower(), array_sortare[i][9], array_sortare[3][x],
                                            array_sortare[i][8], "FIR", array_sortare[i][13], array_sortare[i][23],
                                            array_sortare[i][24], array_sortare[i][25], array_sortare[i][26]])
                    elif array_sortare[i][x] == "Y" or array_sortare[i][x] == "y":
                        array_print.append([array_sortare[3][x] + array_sortare[i][10].lower(),
                                            array_sortare[i][11].replace("U", "W"),
                                            array_sortare[i][10].lower(), array_sortare[i][9], array_sortare[3][x],
                                            array_sortare[i][8], "OPERATIE", array_sortare[i][13], array_sortare[i][23],
                                            array_sortare[i][24], array_sortare[i][25], array_sortare[i][26]])
                    elif array_sortare[i][x] == "S" or array_sortare[i][x] == "s":
                        array_print.append([array_sortare[3][x] + array_sortare[i][10].lower(),
                                            array_sortare[i][11].replace("U", "W"),
                                            array_sortare[i][10].lower(), array_sortare[i][9], array_sortare[3][x],
                                            array_sortare[i][8], "COMPONENT", array_sortare[i][13],
                                            array_sortare[i][23],
                                            array_sortare[i][24], array_sortare[i][25], array_sortare[i][26]])

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_TGLM_R.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul TGLM_R")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")


def cm4axell():
    pbargui = Tk()
    pbargui.title("Control Matrix 4AXEL L")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)

    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix")
    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))
        if array_sortare[3][1] == "AXL":
            array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                            "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()
            for i in range(2, len(array_sortare)):
                for x in range(78, 260):
                    if array_sortare[i][x] == "X" or array_sortare[i][x] == "x":
                        array_print.append([array_sortare[1][x] + array_sortare[i][12].lower(),
                                            array_sortare[i][9].replace("U", "W"),
                                            array_sortare[i][12].lower(), array_sortare[i][13], array_sortare[1][x],
                                            array_sortare[i][8], "FIR", array_sortare[i][14], array_sortare[i][38],
                                            array_sortare[i][39], array_sortare[i][40], array_sortare[i][41]])
                    elif array_sortare[i][x] == "Y" or array_sortare[i][x] == "y":
                        array_print.append([array_sortare[1][x] + array_sortare[i][12].lower(),
                                            array_sortare[i][9].replace("U", "W"),
                                            array_sortare[i][12].lower(), array_sortare[i][13], array_sortare[1][x],
                                            array_sortare[i][8], "OEPRATIE", array_sortare[i][14], array_sortare[i][38],
                                            array_sortare[i][39], array_sortare[i][40], array_sortare[i][41]])
                    elif array_sortare[i][x] == "S" or array_sortare[i][x] == "s":
                        array_print.append([array_sortare[1][x] + array_sortare[i][12].lower(),
                                            array_sortare[i][9].replace("U", "W"),
                                            array_sortare[i][12].lower(), array_sortare[i][13], array_sortare[1][x],
                                            array_sortare[i][8], "COMPONENT", array_sortare[i][14],
                                            array_sortare[i][38],
                                            array_sortare[i][39], array_sortare[i][40], array_sortare[i][41]])

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_4AXEL_L.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul 4AXEL L")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")


def cm4axelr():
    pbargui = Tk()
    pbargui.title("Control Matrix 4AXEL R")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)

    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix")
    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))
        if array_sortare[2][79] == "8026":
            array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                            "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()
            for i in range(2, len(array_sortare)):
                for x in range(79, 632):
                    if array_sortare[i][x] == "X" or array_sortare[i][x] == "x":
                        array_print.append([array_sortare[1][x] + array_sortare[i][11].lower(),
                                            array_sortare[i][8].replace("U", "W"),
                                            array_sortare[i][11].lower(), array_sortare[i][10], array_sortare[1][x],
                                            array_sortare[i][7], "FIR", array_sortare[i][12], array_sortare[i][37],
                                            array_sortare[i][38], array_sortare[i][39], array_sortare[i][40]])
                    elif array_sortare[i][x] == "Y" or array_sortare[i][x] == "y":
                        array_print.append([array_sortare[1][x] + array_sortare[i][11].lower(),
                                            array_sortare[i][8].replace("U", "W"),
                                            array_sortare[i][11].lower(), array_sortare[i][10], array_sortare[1][x],
                                            array_sortare[i][7], "OPERATIE", array_sortare[i][12], array_sortare[i][37],
                                            array_sortare[i][38], array_sortare[i][39], array_sortare[i][40]])
                    elif array_sortare[i][x] == "S" or array_sortare[i][x] == "s":
                        array_print.append([array_sortare[1][x] + array_sortare[i][11].lower(),
                                            array_sortare[i][8].replace("U", "W"),
                                            array_sortare[i][11].lower(), array_sortare[i][10], array_sortare[1][x],
                                            array_sortare[i][7], "COMPONENT", array_sortare[i][12],
                                            array_sortare[i][37],
                                            array_sortare[i][38], array_sortare[i][39], array_sortare[i][40]])

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_4AXEL_R.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul 4AXEL R")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")


def cmss():
    pbargui = Tk()
    pbargui.title("Control Matrix Super Sleeve")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    timelabel = Label(pbargui, text="Time . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    timelabel.grid(row=2, column=2)

    file_load = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul Control Matrix Super Sleeve:")
    wb = load_workbook(file_load)
    statuslabel["text"] = "Loading excel file . . . "
    pbargui.update_idletasks()
    ws1 = wb["BASICK_Module"]
    ws2 = wb["Matrix SS_location"]
    file_counter = len(ws1['A'])
    file_progres = 0
    array_write = []
    start0 = time.time()
    for row in ws1['A']:
        if row.value is not None:
            array_write.append([ws1.cell(row=row.row, column=1).value, ws1.cell(row=row.row, column=2).value,
                                ws1.cell(row=row.row, column=3).value])
            end0 = time.time()
            file_progres += 1
            pbar['value'] += 2
            statuslabel["text"] = "Inregistrari parcurse " + str(file_progres) + "/" + str(file_counter)
            timelabel["text"] = "Estimated time to complete : " + \
                                str(((file_counter * 0.00095) - (end0 - start0)) / 60)[:5] + " minutes."
            pbargui.update_idletasks()

    array_write[0].extend([ws2.cell(row=2, column=1).value, ws2.cell(row=2, column=6).value,
                           ws2.cell(row=2, column=8).value, ws2.cell(row=2, column=9).value,
                           ws2.cell(row=2, column=11).value, ws2.cell(row=2, column=12).value,
                           ws2.cell(row=2, column=13).value, ws2.cell(row=2, column=14).value,
                           ws2.cell(row=2, column=2).value, "Nume", "Lungime mm"])
    # file_counter = len(ws2['AA']) * len(array_write) + len(ws2['AC']) * len(array_write)
    file_counter = 200
    print(file_counter)
    file_progres = 0
    start1 = time.time()
    for i in range(0, 200):  # len(array_write)):
        start2 = time.time()
        for row in ws2['AA']:
            if row.value is not None and row.value == array_write[i][0]:
                array_write[i].extend([ws2.cell(row=row.row, column=1).value, ws2.cell(row=row.row, column=6).value,
                                       ws2.cell(row=row.row, column=8).value, ws2.cell(row=row.row, column=9).value,
                                       ws2.cell(row=row.row, column=11).value, ws2.cell(row=row.row, column=12).value,
                                       ws2.cell(row=row.row, column=13).value, ws2.cell(row=row.row, column=14).value,
                                       ws2.cell(row=row.row, column=2).value, ws2.cell(row=row.row, column=10).value])
        for row in ws2['AC']:
            if row.value is not None and row.value == array_write[i][0]:
                array_write[i].extend([ws2.cell(row=row.row, column=1).value, ws2.cell(row=row.row, column=6).value,
                                       ws2.cell(row=row.row, column=8).value, ws2.cell(row=row.row, column=9).value,
                                       ws2.cell(row=row.row, column=11).value, ws2.cell(row=row.row, column=12).value,
                                       ws2.cell(row=row.row, column=13).value, ws2.cell(row=row.row, column=14).value,
                                       ws2.cell(row=row.row, column=2).value, ws2.cell(row=row.row, column=10).value])
        end1 = time.time()
        print(end1 - start2)
        file_progres += 2
        pbar['value'] += 2
        statuslabel["text"] = "Inregistrari parcurse " + str(file_progres) + "/" + str(file_counter)
        timelabel["text"] = "Estimated time to complete : " + \
                            str((file_counter * 0.035) - (end1 - start1))[:5] + " minutes."
        pbargui.update_idletasks()
        end3 = time.time()
        print(end3 - start1)


def cmcsrnew():
    pbargui = Tk()
    pbargui.title("Control Matrix CSR")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    listatwist = ["131_002", "131_102", "131_102", "grau_047"]
    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix:")
    array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                    "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
    idx_conector = []
    idx_pin = []
    idx_module = []
    idx_knname = ""
    idx_realname = ""
    idx_leadset = ""
    idx_kanbanag = ""
    idx_pnmaterial = ""

    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))

        if "8014" in array_sortare[0]:
            for i, j in enumerate(array_sortare[5]):
                if j == 'Kurzname':
                    idx_conector.append(i)
                elif j == 'Pin':
                    idx_pin.append(i)
                elif j == 'KN Name':
                    idx_knname = i
                elif j == 'Real Name':
                    idx_realname = i
                elif j == 'Ledset':
                    idx_leadset = i
                elif j == 'KANBAN-AG':
                    idx_kanbanag = i
                elif j == 'PN':
                    idx_pnmaterial = i
            for i, j in enumerate(array_sortare[3]):
                if len(j) == 13:
                    idx_module.append(i)

            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()

            for i in range(6, len(array_sortare)):
                for x in range(0, len(idx_module)):
                    if array_sortare[i][idx_module[x]] == "X" or array_sortare[i][idx_module[x]] == "x":
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "FIR",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "Y" or array_sortare[i][idx_module[x]] == "y":
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "OPERATIE",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "S" or array_sortare[i][idx_module[x]] == "s" \
                            and array_sortare[i][idx_realname] not in listatwist:
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "COMPONENT",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "S" or array_sortare[i][idx_module[x]] == "s" \
                            and array_sortare[i][idx_realname] in listatwist:
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname],
                                            array_sortare[i][idx_knname].lower(), array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "FIR",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_CSR.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul CSR")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")


def cmcslnew():
    pbargui = Tk()
    pbargui.title("Control Matrix CSL")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    listatwist = ["131_002", "131_102", "131_102", "grau_047"]
    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix:")
    array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                    "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
    idx_conector = []
    idx_pin = []
    idx_module = []
    idx_knname = ""
    idx_realname = ""
    idx_leadset = ""
    idx_kanbanag = ""
    idx_pnmaterial = ""

    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))

        if "8011" in array_sortare[0]:
            for i, j in enumerate(array_sortare[2]):
                if j == 'von' or j == 'nach':
                    idx_conector.append(i)
                elif j == 'Pin':
                    idx_pin.append(i)
                elif j == 'Kanban name':
                    idx_knname = i
                elif j == 'REAL NAME':
                    idx_realname = i
                elif j == 'Leadset':
                    idx_leadset = i
                elif j == 'Actual Kanban-AG':
                    idx_kanbanag = i
                elif j == 'FORS PN':
                    idx_pnmaterial = i
            for i, j in enumerate(array_sortare[1]):
                if len(j) == 13:
                    idx_module.append(i)
            print(idx_knname)
            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()

            for i in range(6, len(array_sortare)):
                for x in range(0, len(idx_module)):
                    if array_sortare[i][idx_module[x]] == "X" or array_sortare[i][idx_module[x]] == "x":
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "FIR",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "Y" or array_sortare[i][idx_module[x]] == "y":
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "OPERATIE",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "S" or array_sortare[i][idx_module[x]] == "s" \
                            and array_sortare[i][idx_realname] not in listatwist:
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "COMPONENT",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "S" or array_sortare[i][idx_module[x]] == "s" \
                            and array_sortare[i][idx_realname] in listatwist:
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname],
                                            array_sortare[i][idx_knname].lower(), array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "FIR",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_CSL.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul CSL")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")


def cmtglmlnew():
    pbargui = Tk()
    pbargui.title("Control Matrix TGLM L")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    listatwist = ["131_002", "131_102", "131_102", "grau_047"]
    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix:")
    array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                    "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
    idx_conector = []
    idx_pin = []
    idx_module = []
    idx_knname = ""
    idx_realname = ""
    idx_leadset = ""
    idx_kanbanag = ""
    idx_pnmaterial = ""

    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))

        if "8023" in array_sortare[1]:
            for i, j in enumerate(array_sortare[3]):
                if j == 'Kurzname1' or j == 'Kurzname2':
                    idx_conector.append(i)
                elif j == 'Pin1' or j == 'Pin2':
                    idx_pin.append(i)
                elif j == 'KANBAN Name':
                    idx_knname = i
                elif j == 'Ltg-Nr.':
                    idx_realname = i
                elif j == 'Leadset':
                    idx_leadset = i
                elif j == 'COD':
                    idx_kanbanag = i
                elif 'Wires' in j:
                    idx_pnmaterial = i
            print(array_sortare[2])
            for i, j in enumerate(array_sortare[2]):
                if len(j) == 13:
                    idx_module.append(i)
            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()

            for i in range(6, len(array_sortare)):
                for x in range(0, len(idx_module)):
                    if array_sortare[i][idx_module[x]] == "X" or array_sortare[i][idx_module[x]] == "x":
                        array_print.append([array_sortare[2][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "FIR",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "Y" or array_sortare[i][idx_module[x]] == "y":
                        array_print.append([array_sortare[2][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "OPERATIE",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "S" or array_sortare[i][idx_module[x]] == "s" \
                            and array_sortare[i][idx_realname] not in listatwist:
                        array_print.append([array_sortare[2][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "COMPONENT",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "S" or array_sortare[i][idx_module[x]] == "s" \
                            and array_sortare[i][idx_realname] in listatwist:
                        array_print.append([array_sortare[2][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname],
                                            array_sortare[i][idx_knname].lower(), array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "FIR",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_TGLM_L.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul TGLM L")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")


def cmtglmrnew():
    pbargui = Tk()
    pbargui.title("Control Matrix TGLM R")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    listatwist = ["131_002", "131_102", "131_102", "grau_047"]
    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix:")
    array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                    "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
    idx_conector = []
    idx_pin = []
    idx_module = []
    idx_knname = ""
    idx_realname = ""
    idx_leadset = ""
    idx_kanbanag = ""
    idx_pnmaterial = ""

    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))

        if "8024" in array_sortare[5]:
            for i, j in enumerate(array_sortare[4]):
                if j == 'Kurzname1' or j == 'Kurzname2':
                    idx_conector.append(i)
                elif j == 'Pin1' or j == 'Pin2':
                    idx_pin.append(i)
                elif j == 'KANBAN Name':
                    idx_knname = i
                elif j == 'Ltg-Nr.':
                    idx_realname = i
                elif j == 'Leadset':
                    idx_leadset = i
                elif j == 'COD':
                    idx_kanbanag = i
                elif 'Wires' in  j:
                    idx_pnmaterial = i
            for i, j in enumerate(array_sortare[3]):
                if len(j) == 13:
                    idx_module.append(i)
            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()

            for i in range(6, len(array_sortare)):
                for x in range(0, len(idx_module)):
                    if array_sortare[i][idx_module[x]] == "X" or array_sortare[i][idx_module[x]] == "x":
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "FIR",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "Y" or array_sortare[i][idx_module[x]] == "y":
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "OPERATIE",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "S" or array_sortare[i][idx_module[x]] == "s" \
                            and array_sortare[i][idx_realname] not in listatwist:
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "COMPONENT",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "S" or array_sortare[i][idx_module[x]] == "s" \
                            and array_sortare[i][idx_realname] in listatwist:
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname],
                                            array_sortare[i][idx_knname].lower(), array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "FIR",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_TGLM_R.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul TGLM R")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")


def cm4axellnew():
    pbargui = Tk()
    pbargui.title("Control Matrix 4AXEL L")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    listatwist = ["131_002", "131_102", "131_102", "grau_047"]
    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix:")
    array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                    "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
    idx_conector = []
    idx_pin = []
    idx_module = []
    idx_knname = ""
    idx_realname = ""
    idx_leadset = ""
    idx_kanbanag = ""
    idx_pnmaterial = ""

    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))

        if "AXL" in array_sortare[3]:
            for i, j in enumerate(array_sortare[2]):
                if j == 'Xcode_1' or j == 'Xcode_2':
                    idx_conector.append(i)
                elif j == 'Cavity_1' or j == 'Cavity_2':
                    idx_pin.append(i)
                elif 'KANBAN' in j:
                    idx_knname = i
                elif j == 'WireNo':
                    idx_realname = i
                elif j == 'Leadset':
                    idx_leadset = i
                elif j == 'COD':
                    idx_kanbanag = i
                elif j == "PN":
                    idx_pnmaterial = i
            for i, j in enumerate(array_sortare[1]):
                if len(j) == 13:
                    idx_module.append(i)
            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()

            for i in range(6, len(array_sortare)):
                for x in range(0, len(idx_module)):
                    if array_sortare[i][idx_module[x]] == "X" or array_sortare[i][idx_module[x]] == "x":
                        array_print.append([array_sortare[1][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "FIR",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "Y" or array_sortare[i][idx_module[x]] == "y":
                        array_print.append([array_sortare[1][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "OPERATIE",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "S" or array_sortare[i][idx_module[x]] == "s" \
                            and array_sortare[i][idx_realname] not in listatwist:
                        array_print.append([array_sortare[1][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "COMPONENT",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "S" or array_sortare[i][idx_module[x]] == "s" \
                            and array_sortare[i][idx_realname] in listatwist:
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname],
                                            array_sortare[i][idx_knname].lower(), array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "FIR",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_4AXEL_L.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul 4AXEL L")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")


def cm4axelrnew():
    pbargui = Tk()
    pbargui.title("Control Matrix 4AXEL R")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    listatwist = ["131_002", "131_102", "131_102", "grau_047"]
    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix:")
    array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                    "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
    idx_conector = []
    idx_pin = []
    idx_module = []
    idx_knname = ""
    idx_realname = ""
    idx_leadset = ""
    idx_kanbanag = ""
    idx_pnmaterial = ""

    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))

        if "8026" in array_sortare[2]:
            for i, j in enumerate(array_sortare[3]):
                if j == 'Xcode_1' or j == 'Xcode_2':
                    idx_conector.append(i)
                elif j == 'Cavity_1' or j == 'Cavity_2':
                    idx_pin.append(i)
                elif 'KANBAN' in j:
                    idx_knname = i
                elif j == 'WireNo':
                    idx_realname = i
                elif j == 'LeadSet':
                    idx_leadset = i
                elif j == 'COD':
                    idx_kanbanag = i
                elif j == "PN":
                    idx_pnmaterial = i
            for i, j in enumerate(array_sortare[1]):
                if len(j) == 13:
                    idx_module.append(i)
            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()
            print(idx_realname)
            for i in range(6, len(array_sortare)):
                for x in range(0, len(idx_module)):
                    if array_sortare[i][idx_module[x]] == "X" or array_sortare[i][idx_module[x]] == "x":
                        array_print.append([array_sortare[1][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "FIR",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "Y" or array_sortare[i][idx_module[x]] == "y":
                        array_print.append([array_sortare[1][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "OPERATIE",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "S" or array_sortare[i][idx_module[x]] == "s" \
                            and array_sortare[i][idx_realname] not in listatwist:
                        array_print.append([array_sortare[1][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_knname], array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "COMPONENT",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][idx_module[x]] == "S" or array_sortare[i][idx_module[x]] == "s" \
                            and array_sortare[i][idx_realname] in listatwist:
                        array_print.append([array_sortare[3][idx_module[x]] + array_sortare[i][idx_realname].lower(),
                                            array_sortare[i][idx_kanbanag].replace("23U", "23W"),
                                            array_sortare[i][idx_realname],
                                            array_sortare[i][idx_knname].lower(), array_sortare[3][idx_module[x]],
                                            array_sortare[i][idx_leadset], "FIR",
                                            array_sortare[i][idx_pnmaterial], array_sortare[i][idx_conector[0]],
                                            array_sortare[i][idx_pin[0]], array_sortare[i][idx_conector[1]],
                                            array_sortare[i][idx_pin[1]]])

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_4AXEL_R.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul 4AXEL R")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")
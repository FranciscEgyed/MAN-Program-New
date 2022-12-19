import datetime
import os
import time
from collections import Counter
from tkinter import messagebox, Tk, ttk, Label, HORIZONTAL, filedialog
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook
from diverse import log_file
from functii_print import prn_excel_variatii


def extragere_lungimi_ksk():
    pbargui = Tk()
    pbargui.title("Extragere lungimi KSK")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    save_time = datetime.datetime.now().strftime("%d.%m.%Y_%H.%M.%S")
    array_lungimi = [["KSK", "Dwg No.", "Module", "Ltg-Nr.", "Leitung", "Farbe", "Quer.", "Kurzname", "Pin", "Lange"]]
    counter = 0
    dir_salvare = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir),
                                          title="Selectati directorul pentru salvare")
    dir_prelucrare = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/",
                                             title="Selectati directorul cu fisiere:")

    start = time.time()
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_prelucrare):
        if file_all.endswith(".xlsx") and not file_all.startswith("BOM"):
            file_counter = file_counter + 1
    for file_all in os.listdir(dir_prelucrare):
        try:
            if file_all.endswith(".xlsx") and not file_all.startswith("BOM"):
                counter = counter + 1
                wb = load_workbook(dir_prelucrare + "/" + file_all)
                ws0 = wb.worksheets[0]
                ws1 = wb.worksheets[1]
                ws2 = wb.worksheets[2]
                file_progres = file_progres + 1
                statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
                pbar['value'] += 2
                pbargui.update_idletasks()
                for row in ws1['A']:
                    if row.value != "Harness":
                        array_lungimi.append([ws0.cell(row=2, column=1).value, ws1.cell(row=row.row, column=1).value,
                                              ws1.cell(row=row.row, column=2).value,
                                              ws1.cell(row=row.row, column=3).value,
                                              ws1.cell(row=row.row, column=4).value,
                                              ws1.cell(row=row.row, column=5).value,
                                              ws1.cell(row=row.row, column=6).value,
                                              ws1.cell(row=row.row, column=7).value,
                                              ws1.cell(row=row.row, column=8).value,
                                              ws1.cell(row=row.row, column=9).value])
                for row in ws2['A']:
                    if row.value != "Harness":
                        array_lungimi.append([ws0.cell(row=2, column=1).value, ws2.cell(row=row.row, column=1).value,
                                              ws2.cell(row=row.row, column=2).value,
                                              ws2.cell(row=row.row, column=3).value,
                                              ws2.cell(row=row.row, column=4).value,
                                              ws2.cell(row=row.row, column=5).value,
                                              ws2.cell(row=row.row, column=6).value,
                                              ws2.cell(row=row.row, column=7).value,
                                              ws2.cell(row=row.row, column=8).value,
                                              ws2.cell(row=row.row, column=9).value])
        except PermissionError:
            messagebox.showerror('Eroare scriere', "Fisierul " + file_all + "este read-only!")
            quit()

    wbs = Workbook()
    wss1 = wbs.active
    for i in range(len(array_lungimi)):
        for x in range(len(array_lungimi[i])):
            try:
                if "E-" in array_lungimi[i][x]:
                    try:
                        wss1.cell(column=x + 1, row=i + 1, value=str(array_lungimi[i][x]))
                    except:
                        wss1.cell(column=x + 1, row=i + 1, value=str(array_lungimi[i][x]))
                else:
                    try:
                        wss1.cell(column=x + 1, row=i + 1, value=float(array_lungimi[i][x]))
                    except:
                        wss1.cell(column=x + 1, row=i + 1, value=array_lungimi[i][x])
            except TypeError:
                wss1.cell(column=x + 1, row=i + 1, value=array_lungimi[i][x])
    if dir_salvare == "":
        try:
            wbs.save(
                os.path.abspath(os.curdir) + "/MAN/Output/Report Files/Export Lungimi " + save_time + ".xlsx")
            log_file("Creat Export Lungimi.xlsx")
        except PermissionError:
            log_file("Eroare scriere. Nu am salvat Export Lungimi " + save_time + ".xlsx")
            messagebox.showerror('Eroare scriere', "Fisierul este read-only!")
            quit()
    else:
        try:
            wbs.save(dir_salvare + "/Export Lungimi " + save_time + ".xlsx")
            log_file("Creat Export Lungimi.xlsx")
        except PermissionError:
            log_file("Eroare scriere. Nu am salvat Export Lungimi.xlsx")
            messagebox.showerror('Eroare scriere', "Fisierul este read-only!")
            quit()
    pbar.destroy()
    pbargui.destroy()
    end = time.time()
    messagebox.showinfo('Finalizat!', 'Prelucrate ' + str(counter) + " fisiere in " + str(end - start)[:6] + " secunde.")


def extragere_bom_ksk():
    pbargui = Tk()
    pbargui.title("Extragere BOM KSK")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    save_time = datetime.datetime.now().strftime("%d.%m.%Y_%H.%M.%S")
    array_bom = [["KSK", "Module", "Quantity", "Bezei", "VOBES-ID", "Benennung", "Verwendung", "Verwendung",
                  "Kurzname", "xy", "Teilenummer", "Vorzugsteil", "TAB-Nummer", "Referenzteil", "Farbe",
                  "E-Komponente", "E-Komponente Part-Nr.", "Einh."]]
    cc = []
    counter = 0
    dir_salvare = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir),
                                          title="Selectati directorul pentru salvare")
    dir_prelucrare = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/",
                                          title="Selectati directorul cu fisiere:")
    start = time.time()
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_prelucrare):
        if file_all.endswith(".xlsx") and file_all.startswith("BOM"):
            file_counter = file_counter + 1
    for file_all in os.listdir(dir_prelucrare):
        try:
            if file_all.endswith(".xlsx") and file_all.startswith("BOM"):
                counter = counter + 1
                wb = load_workbook(dir_prelucrare + "/" + file_all)
                ws0 = wb.worksheets[0]
                ws1 = wb.worksheets[1]
                file_progres = file_progres + 1
                statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
                pbar['value'] += 2
                pbargui.update_idletasks()
                for row in ws1['A']:
                    if row.value != "Module" and ws1.cell(row=row.row, column=2).value != 0:
                        array_bom.append([ws0.cell(row=2, column=1).value, ws1.cell(row=row.row, column=1).value,
                                          ws1.cell(row=row.row, column=2).value, ws1.cell(row=row.row, column=3).value,
                                          ws1.cell(row=row.row, column=4).value, ws1.cell(row=row.row, column=5).value,
                                          ws1.cell(row=row.row, column=6).value, ws1.cell(row=row.row, column=7).value,
                                          ws1.cell(row=row.row, column=8).value, ws1.cell(row=row.row, column=9).value,
                                          ws1.cell(row=row.row, column=10).value, ws1.cell(row=row.row, column=11).value,
                                          ws1.cell(row=row.row, column=12).value, ws1.cell(row=row.row, column=13).value,
                                          ws1.cell(row=row.row, column=14).value, ws1.cell(row=row.row, column=15).value,
                                          ws1.cell(row=row.row, column=16).value, ws1.cell(row=row.row, column=17).value])
        except PermissionError:
            messagebox.showerror('Eroare scriere', "Fisierul " + file_all + "este read-only!")
            quit()
    for i in range(1, len(array_bom)):
        ksk = ""
        partno = ""
        conectorno = ""
        ksk = array_bom[i][0]
        for x in range(3, len(array_bom[i])):
            if len(str(array_bom[i][x])) == 13 and str(array_bom[i][x]).find(".") == 2 and \
                    str(array_bom[i][x]).find("-") == 8:
                partno = array_bom[i][x]
                break
            else:
                partno = "-"
            for q in range(3, len(array_bom[i])):
                if str(array_bom[i][q]).startswith("X"):
                    conectorno = array_bom[i][q]
                    break
                else:
                    conectorno = "-"
        array_bom[i].append(ksk + partno + conectorno)
        cc.append(ksk + partno + conectorno)
    concatenare = list(Counter(cc).items())
    for i in range(1, len(array_bom)):
        for y in concatenare:
            if array_bom[i][18] == y[0]:
                array_bom[i].append(y[1])
    wbs = Workbook()
    wss1 = wbs.active
    for i in range(len(array_bom)):
        for x in range(len(array_bom[i])):
            try:
                if "E-" in array_bom[i][x]:
                    try:
                        wss1.cell(column=x + 1, row=i + 1, value=str(array_bom[i][x]))
                    except:
                        wss1.cell(column=x + 1, row=i + 1, value=str(array_bom[i][x]))
                else:
                    try:
                        wss1.cell(column=x + 1, row=i + 1, value=float(array_bom[i][x]))
                    except:
                        wss1.cell(column=x + 1, row=i + 1, value=array_bom[i][x])
            except TypeError:
                wss1.cell(column=x + 1, row=i + 1, value=array_bom[i][x])
    wss1.cell(column=19, row=1, value="Concatenare Ksk+Part+Conector")
    wss1.cell(column=20, row=1, value="Count Concatenare")
    pbar['value'] += 2
    pbargui.update_idletasks()
    if dir_salvare == "":
        try:
            wbs.save(
                os.path.abspath(os.curdir) + "/MAN/Output/Report Files/Export BOMs " + save_time + ".xlsx")
            log_file("Creat Export BOMs.xlsx")
        except PermissionError:
            log_file("Eroare scriere. Nu am salvat Export BOMs " + save_time + ".xlsx")
            messagebox.showerror('Eroare scriere', "Fisierul este read-only!")
            quit()
    else:
        try:
            wbs.save(dir_salvare + "/Export BOMs " + save_time + ".xlsx")
            log_file("Creat Export Lungimi.xlsx")
        except PermissionError:
            log_file("Eroare scriere. Nu am salvat Export BOMs.xlsx")
            messagebox.showerror('Eroare scriere', "Fisierul este read-only!")
            quit()
    pbar.destroy()
    pbargui.destroy()
    end = time.time()
    messagebox.showinfo('Finalizat!', 'Prelucrate ' + str(counter) + " fisiere in " + str(end - start)[:6] + " secunde.")


def extragere_variatii():
    pbargui = Tk()
    pbargui.title("Extragere variatii lungimi")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    start = time.time()
    counter = 0
    array_print = [["KSK", "Type", "Abbreviation", "Part_Number_2", "Side", "q-VLA", "r-VLA/RQT", "s-Radstand",	"t-NLA",
                    "u-Aoeberhang", "Heck module", "Identic"]]

    dir_selectat = os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/8000/"
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_selectat):
        if file_all.endswith(".xlsx"):
            file_counter = file_counter + 1
    pbar['value'] = 0
    for file_all in os.listdir(dir_selectat):
        if file_all.endswith(".xlsx"):
            counter = counter + 1
            file_progres = file_progres + 1
            statuslabel["text"] = "8000 = " + str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 1
            pbargui.update_idletasks()
            wb = load_workbook(dir_selectat + "/" + file_all)
            ws = wb.worksheets[3]
            for row in ws['A']:
                if row.value != "Type":
                    array_print.append([file_all[:-5], ws.cell(row=row.row, column=1).value,
                                        ws.cell(row=row.row, column=2).value, ws.cell(row=row.row, column=3).value,
                                        ws.cell(row=row.row, column=4).value, ws.cell(row=row.row, column=5).value,
                                        ws.cell(row=row.row, column=6).value, ws.cell(row=row.row, column=7).value,
                                        ws.cell(row=row.row, column=8).value, ws.cell(row=row.row, column=9).value,
                                        ws.cell(row=row.row, column=10).value, ws.cell(row=row.row, column=11).value])
            continue
        else:
            continue

    dir_selectat = os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/8011/"
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_selectat):
        if file_all.endswith(".xlsx"):
            file_counter = file_counter + 1
    pbar['value'] = 0
    for file_all in os.listdir(dir_selectat):
        if file_all.endswith(".xlsx"):
            counter = counter + 1
            file_progres = file_progres + 1
            statuslabel["text"] = "8011 = " + str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 1
            pbargui.update_idletasks()
            wb = load_workbook(dir_selectat + "/" + file_all)
            ws = wb.worksheets[3]
            for row in ws['A']:
                if row.value != "Type":
                    array_print.append([file_all[:-5], ws.cell(row=row.row, column=1).value,
                                        ws.cell(row=row.row, column=2).value, ws.cell(row=row.row, column=3).value,
                                        ws.cell(row=row.row, column=4).value, ws.cell(row=row.row, column=5).value,
                                        ws.cell(row=row.row, column=6).value, ws.cell(row=row.row, column=7).value,
                                        ws.cell(row=row.row, column=8).value, ws.cell(row=row.row, column=9).value,
                                        ws.cell(row=row.row, column=10).value, ws.cell(row=row.row, column=11).value])
            continue
        else:
            continue

    dir_selectat = os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/8023/"
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_selectat):
        if file_all.endswith(".xlsx"):
            file_counter = file_counter + 1
    pbar['value'] = 0
    for file_all in os.listdir(dir_selectat):
        if file_all.endswith(".xlsx"):
            counter = counter + 1
            file_progres = file_progres + 1
            statuslabel["text"] = "8023 = " + str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 1
            pbargui.update_idletasks()
            wb = load_workbook(dir_selectat + "/" + file_all)
            ws = wb.worksheets[3]
            for row in ws['A']:
                if row.value != "Type":
                    array_print.append([file_all[:-5], ws.cell(row=row.row, column=1).value,
                                        ws.cell(row=row.row, column=2).value, ws.cell(row=row.row, column=3).value,
                                        ws.cell(row=row.row, column=4).value, ws.cell(row=row.row, column=5).value,
                                        ws.cell(row=row.row, column=6).value, ws.cell(row=row.row, column=7).value,
                                        ws.cell(row=row.row, column=8).value, ws.cell(row=row.row, column=9).value,
                                        ws.cell(row=row.row, column=10).value, ws.cell(row=row.row, column=11).value])
            continue
        else:
            continue
    for i in range(len(array_print)):
        for x in range(len(array_print[i])):
            if array_print[i][x] is None:
                array_print[i][x] = ""
        else:
            continue
    pbar.destroy()
    pbargui.destroy()
    prn_excel_variatii(array_print)
    end = time.time()

    messagebox.showinfo('Finalizat!', 'Prelucrate ' + str(counter) + " fisiere in " + str(end - start)[:6] + " secunde.")


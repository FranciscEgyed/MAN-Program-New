import csv
import os
import time
from tkinter import messagebox, filedialog, Tk, ttk, HORIZONTAL, Label
import pandas as pd
from openpyxl import load_workbook
from diverse import log_file
from functii_print import prn_excel_separare_ksk, prn_excel_bom_complete, prn_excel_wires_complete_leoni, \
    prn_excel_wirelistsallinone, prn_excel_ksk_neprelucrate
import sqlite3


def sortare_jit():
    pbargui = Tk()
    pbargui.title("Sortare JIT")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    timelabel = Label(pbargui, text="Time . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2)
    timelabel.grid(row=2, column=2)
    c8000 = 0
    c8011 = 0
    c8023 = 0
    cnec = 0
    tip = ""
    # Import data files
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Sortare Module.txt", newline='') as csvfile:
        array_sortare_module = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/KSKLight.txt", newline='') as csvfile:
        array_sortare_light = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Module Active.txt", newline='') as csvfile:
        array_module_active = list(csv.reader(csvfile, delimiter=';'))
    normal = ["04.37161-9100", "81.25484-5259", "81.25484-5263", "81.25484-5260", "81.25484-5264", "81.25484-5273",
              "81.25484-5272", "81.25484-5267", "81.25484-5268"]
    ADR = ["04.37161-9000", "81.25484-5261", "81.25484-5265", "81.25484-5262", "81.25484-5266", "81.25484-5275",
           "81.25484-5274", "81.25484-5271", "81.25484-5270"]
    # Open JIT file
    fisier_calloff = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                                title="Incarcati fisierul care necesita sortare")
    if len(fisier_calloff) == 0:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showinfo("Nu ati selectat nimic")
        return None
    statuslabel["text"] = "Incarcare fisier excel"
    timelabel["text"] = "0.15 secunde / KSK"
    pbar['value'] += 2
    pbargui.update_idletasks()
    data_download = os.path.basename(fisier_calloff)[11:21]
    try:
        wb = load_workbook(fisier_calloff)
    except:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showinfo("Fisier invalid", fisier_calloff + " extensie incompatibila!")
        return None
    try:
        conn = sqlite3.connect("//SVRO8FILE01/Groups/General/EFI/DBMAN/database.db")
    except sqlite3.OperationalError:
        conn = sqlite3.connect(os.path.abspath(os.curdir) + "/MAN/Input/Others/database.db")
    statuslabel["text"] = "Sortare fisier JIT"
    timelabel["text"] = "0.15 secunde / KSK"
    pbar['value'] += 2
    pbargui.update_idletasks()
    ws = wb.worksheets[0]
    lista_calloff_total = []
    for row in ws['A']:
        if row.value is not None and row.value != "ext.Abrufnummer" and row.value != "PRODN" and\
                row.value != "Ext.JIT Call No":
            lista_calloff_total.append(row.value)

    lista_calloff_unice = list(dict.fromkeys(lista_calloff_total))
    file_counter = len(lista_calloff_unice)
    file_progres = 0
    start = time.time()
    # Extract each KSK from JIT file
    for element in lista_calloff_unice:
        is_light = "NO"
        array_temporar = []
        array_temporar_module = []
        array_database = []
        for row in ws['A']:
            if row.value == element:
                if len(str(ws.cell(row=row.row, column=4).value)) > 4:
                    qty = "0"
                elif ws.cell(row=row.row, column=4).value == 1000:
                    qty = "1"
                elif ws.cell(row=row.row, column=4).value == 2000:
                    qty = "2"
                elif ws.cell(row=row.row, column=4).value == 3000:
                    qty = "3"
                else:
                    qty = ws.cell(row=row.row, column=4).value
                array_temporar.append(
                    [ws.cell(row=row.row, column=1).value[1:], ws.cell(row=row.row, column=2).value,
                     ws.cell(row=row.row, column=3).value, qty, ws.cell(row=row.row, column=5).value,
                     ws.cell(row=row.row, column=12).value, ws.cell(row=row.row, column=13).value,
                     ws.cell(row=row.row, column=16).value, ws.cell(row=row.row, column=14).value])
                array_temporar_module.append(ws.cell(row=row.row, column=2).value.replace('PM.', '81.').replace('VM.', '81.'))
        for i in range(len(array_temporar)):
            if "PM." in array_temporar[i][1]:
                array_temporar[i][1] = array_temporar[i][1].replace('PM.', '81.')
        for i in range(len(array_temporar)):
            if "VM." in array_temporar[i][1]:
                array_temporar[i][1] = array_temporar[i][1].replace('VM.', '81.')

        for item in array_temporar:
            if item[1] in array_sortare_module[0]:
                tip = "8000"
                c8000 = c8000 + 1
                break
            elif item[1] in array_sortare_module[1]:
                tip = "8011"
                c8011 = c8011 + 1
                break
            elif item[1] in array_sortare_module[2]:
                tip = "8023"
                c8023 = c8023 + 1
                break
            else:
                tip = "Necunoscut"
        if tip == "Necunoscut":
            cnec = cnec + 1
        #Chech KSK if part of KSK Light project
        if set(array_temporar_module).issubset(array_sortare_light[0]):
            prn_excel_separare_ksk(array_temporar_module, element[1:])
            is_light = "YES"
        # Check harness type
        harnesstype = []
        for m in range(len(array_temporar_module)):
            for n in range(len(array_module_active)):
                if array_temporar_module[m] == array_module_active[n][0] and array_module_active[n][3] != "XXXX":
                    harnesstype.append(array_module_active[n][3].replace(' LHD', '').replace(' RHD', ''))
        harnesstype = list(set(harnesstype))
        # Check ss type
        sstype = "None"
        for x in range(len(array_temporar_module)):
            if array_temporar_module[x] in normal:
                sstype = "NON ADR"
                break
            elif array_temporar_module[x] in ADR:
                sstype = "ADR"
                break
        # Write to database
        primarykey = os.path.basename(fisier_calloff) + element[1:]
        array_database.append([primarykey, os.path.basename(fisier_calloff), ';'.join(harnesstype), is_light,
                               array_temporar[1][8], data_download, element[1:], array_temporar[1][7], sstype,
                               ';'.join(array_temporar_module)])
        conn = sqlite3.connect(os.path.abspath(os.curdir) + "/MAN/Input/Others/database.db")
        cursor = conn.cursor()
        # create a table
        cursor.execute("""CREATE TABLE IF NOT EXISTS KSKDatabase
                          (primarykey text UNIQUE, numejit text, TipHarness text, Light text, DataLivrare text, 
                          DataJIT text, KSKNo text, TrailerNO text, SSType text, Module text) """)
        # insert multiple records using the more secure "?" method
        cursor.executemany("INSERT OR IGNORE INTO KSKDatabase VALUES (?,?,?,?,?,?,?,?,?,?)", array_database)
        conn.commit()
        conn.close()

        # Write to disk
        with open(os.path.abspath(os.curdir) + "/MAN/Input/Module Files/" + tip + "/"
                  + element[1:] + ".csv", 'w', newline='') as myfile:
            wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
            wr.writerow(['Harness', 'Module', 'Side', 'Quantity', tip, "Date", "Time", "Trailer No"])
            wr.writerows(array_temporar)
        del array_temporar
        end = time.time()

        file_progres = file_progres + 1
        statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + "                "
        timelabel["text"] = "Estimated time to complete : " + \
                            str(((file_counter * 0.15) - (end - start)) / 60)[:5] + " minutes."
        pbar['value'] += 2
        pbargui.update_idletasks()
    conn.close()
    pbar.destroy()
    pbargui.destroy()
    log_file("Sortate  8000 = " + str(c8000) + ", 8011 = " + str(c8011) + ", 8023 = " + str(c8023) +
             "Necunoscute = " + str(cnec))
    messagebox.showinfo("Finalizat", "Sortate  8000 = " + str(c8000) + ", 8011 = " + str(c8011) + ", 8023 = " +
                        str(c8023) + ", Necunoscute = " + str(cnec))
    return None

def sortare_jit_dir():
    pbargui = Tk()
    pbargui.title("Sortare JIT din director")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    statuslabel2 = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2)
    statuslabel2.grid(row=2, column=2)
    # Import data files
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Sortare Module.txt", newline='') as csvfile:
        array_sortare_module = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/KSKLight.txt", newline='') as csvfile:
        array_sortare_light = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Module Active.txt", newline='') as csvfile:
        array_module_active = list(csv.reader(csvfile, delimiter=';'))
    normal = ["04.37161-9100", "81.25484-5259", "81.25484-5263", "81.25484-5260", "81.25484-5264", "81.25484-5273",
              "81.25484-5272", "81.25484-5267", "81.25484-5268"]
    ADR = ["04.37161-9000", "81.25484-5261", "81.25484-5265", "81.25484-5262", "81.25484-5266", "81.25484-5275",
           "81.25484-5274", "81.25484-5271", "81.25484-5270"]
    dir_Jit = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir),
                                      title="Selectati directorul cu fisiere JIT:")
    try:
        conn = sqlite3.connect("//SVRO8FILE01/Groups/General/EFI/DBMAN/database.db")
    except sqlite3.OperationalError:
        conn = sqlite3.connect(os.path.abspath(os.curdir) + "/MAN/Input/Others/database.db")
    cursor = conn.cursor()
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_Jit):
        if file_all.endswith(".xlsx"):
            file_counter = file_counter + 1
    if file_counter == 0:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showinfo("Fisier invalid", "Nu am gasit fisiere de prelucrat!")
        return None
    for file_all in os.listdir(dir_Jit):
        if file_all.endswith(".xlsx") and file_all.startswith("JIT"):
            c8000 = 0
            c8011 = 0
            c8023 = 0
            cnec = 0
            tip = ""
            data_download = os.path.basename(file_all)[11:21]
            fisier_calloff = os.path.join(dir_Jit, file_all)
            try:
                wb = load_workbook(fisier_calloff)
            except:
                pbar.destroy()
                pbargui.destroy()
                messagebox.showinfo("Fisier invalid", fisier_calloff + " extensie incompatibila!")
                return None

            ws = wb.worksheets[0]
            lista_calloff_total = []
            for row in ws['A']:
                if row.value is not None and row.value != "ext.Abrufnummer" and row.value != "PRODN" and \
                        row.value != "Ext.JIT Call No":
                    lista_calloff_total.append(row.value)
            lista_calloff_unice = list(dict.fromkeys(lista_calloff_total))
            kskcounter = len(lista_calloff_unice)
            kskprogres = 0
            start = time.time()
            # Extract each KSK from JIT file
            for element in lista_calloff_unice:
                is_light = "NO"
                array_database = []
                kskprogres = kskprogres + 1
                statuslabel2["text"] = "                 " + str(kskprogres) + " / " + str(kskcounter) + " : " + element[1:]
                pbar['value'] += 2
                pbargui.update_idletasks()
                array_temporar = []
                array_temporar_module = []
                for row in ws['A']:
                    if row.value == element:
                        if len(str(ws.cell(row=row.row, column=4).value)) > 4:
                            qty = "0"
                        elif ws.cell(row=row.row, column=4).value == 1000:
                            qty = "1"
                        elif ws.cell(row=row.row, column=4).value == 2000:
                            qty = "2"
                        elif ws.cell(row=row.row, column=4).value == 3000:
                            qty = "3"
                        else:
                            qty = ws.cell(row=row.row, column=4).value
                        array_temporar.append(
                            [ws.cell(row=row.row, column=1).value[1:], ws.cell(row=row.row, column=2).value,
                             ws.cell(row=row.row, column=3).value, qty, ws.cell(row=row.row, column=5).value,
                             ws.cell(row=row.row, column=12).value, ws.cell(row=row.row, column=13).value,
                             ws.cell(row=row.row, column=16).value, ws.cell(row=row.row, column=14).value])
                        array_temporar_module.append(
                            ws.cell(row=row.row, column=2).value.replace('PM.', '81.').replace('VM.', '81.'))
                for i in range(len(array_temporar)):
                    if "PM." in array_temporar[i][1]:
                        array_temporar[i][1] = array_temporar[i][1].replace('PM.', '81.')
                for i in range(len(array_temporar)):
                    if "VM." in array_temporar[i][1]:
                        array_temporar[i][1] = array_temporar[i][1].replace('VM.', '81.')

                for item in array_temporar:
                    if item[1] in array_sortare_module[0]:
                        tip = "8000"
                        c8000 = c8000 + 1
                        break
                    elif item[1] in array_sortare_module[1]:
                        tip = "8011"
                        c8011 = c8011 + 1
                        break
                    elif item[1] in array_sortare_module[2]:
                        tip = "8023"
                        c8023 = c8023 + 1
                        break
                    else:
                        tip = "Necunoscut"
                if tip == "Necunoscut":
                    cnec = cnec + 1
                # Chech KSK if part of KSK Light project
                if set(array_temporar_module).issubset(array_sortare_light[0]):
                    prn_excel_separare_ksk(array_temporar_module, element[1:])
                    is_light = "YES"
                else:
                    prn_excel_ksk_neprelucrate(array_temporar_module, element[1:])
                # Check harness type
                harnesstype = []
                for m in range(len(array_temporar_module)):
                    for n in range(len(array_module_active)):
                        if array_temporar_module[m] == array_module_active[n][0] and array_module_active[n][3] != "XXXX":
                            harnesstype.append(array_module_active[n][3].replace(' LHD', '').replace(' RHD', ''))
                harnesstype = list(set(harnesstype))
                # Check ss type
                sstype = "None"
                for x in range(len(array_temporar_module)):
                    if array_temporar_module[x] in normal:
                        sstype = "NON ADR"
                        break
                    elif array_temporar_module[x] in ADR:
                        sstype = "ADR"
                        break
                # Write to database
                primarykey = os.path.basename(fisier_calloff) + element[1:]
                array_database.append([primarykey, os.path.basename(fisier_calloff), ';'.join(harnesstype), is_light,
                                       array_temporar[1][8], data_download, element[1:], array_temporar[1][7], sstype,
                                       ';'.join(array_temporar_module)])

                # create a table
                cursor.execute("""CREATE TABLE IF NOT EXISTS KSKDatabase
                                  (primarykey text UNIQUE, numejit text, TipHarness text, Light text, DataLivrare text, 
                                  DataJIT text, KSKNo text, TrailerNO text, SSType text, Module text) """)
                # insert multiple records using the more secure "?" method
                cursor.executemany("INSERT OR IGNORE INTO KSKDatabase VALUES (?,?,?,?,?,?,?,?,?,?)", array_database)
                conn.commit()
                with open(os.path.abspath(os.curdir) + "/MAN/Input/Module Files/" + tip + "/"
                          + element[1:] + ".csv", 'w', newline='') as myfile:
                    wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                    wr.writerow(['Harness', 'Module', 'Side', 'Quantity', tip, "Date", "Time", "Trailer No"])
                    wr.writerows(array_temporar)
                del array_temporar
                end = time.time()


            file_progres = file_progres + 1
            statuslabel["text"] = str(file_progres) + " / " + str(file_counter) + " : " + file_all
            pbar['value'] += 2
            pbargui.update_idletasks()
    conn.close()
    pbar.destroy()
    pbargui.destroy()
    messagebox.showinfo('Finalizat!', str(file_progres) + " fisiere din " + str(file_counter))


def boms():
    pbargui = Tk()
    pbargui.title("Prelucrare BOM-uri")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    global nume_fisier
    array_boms = ["8000", "8001", "8011", "8012", "8013", "8014", "8023", "8024", "8025", "8026", "8030", "8031",
                  "8052", "8053"]
    dir_BOM = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir), title="Selectati directorul cu fisiere:")
    start = time.time()
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_BOM):
        if file_all.endswith(".csv"):
            file_counter = file_counter + 1
    for file_all in os.listdir(dir_BOM):
        if file_all.endswith(".csv"):
            file_progres = file_progres + 1
            statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 2
            pbargui.update_idletasks()
            for item in array_boms:
                if item in file_all:
                    try:
                        with open(dir_BOM + "/" + file_all, newline='') as csvfile:
                            array_original = list(csv.reader(csvfile, delimiter=';'))
                        if "AEM" in array_original[0][1]:
                            for i in range(len(array_original)):
                                del array_original[i][1]
                        if array_original[0][0] != "1" and len(array_original[0]) != 2:
                            messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect')
                            return
                        if len(array_original) < 50:
                            for i in range(0, len(array_original)):
                                if array_original[i][0] == "0":
                                    if str(array_original[i][-1])[:1] != "V":
                                        messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect')
                                        return
                        else:
                            for i in range(0, 50):
                                if array_original[i][0] == "0":
                                    if str(array_original[i][-1])[:1] != "V":
                                        messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect')
                                        return
                    except:
                        messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect')
                        return
                    for q in range(len(array_original)):
                        try:
                            if "Workplace" in array_original[q][7]:
                                array_original[q].pop(7)
                            if "Workplace" in array_original[q][6]:
                                array_original[q].pop(6)
                        except:
                            continue
                    nume_fisier = os.path.splitext(os.path.basename(file_all))[0]
                    array_prelucrat = [["Module", "Quantity", "Bezeichnung", "VOBES-ID", "Benennung", "Verwendung",
                                        "Verwendung", "Kurzname", "xy", "Teilenummer", "Vorzugsteil", "TAB-Nummer",
                                        "Referenzteil", "Farbe", "E-Komponente", "E-Komponente Part-Nr.", "Einh."]]
                    array_temp = [["Bezeichnung", "VOBES-ID", "Benennung", "Verwendung", "Verwendung", "Kurzname", "xy",
                                   "Teilenummer", "Vorzugsteil", "TAB-Nummer", "Referenzteil", "Farbe", "E-Komponente",
                                   "E-Komponente Part-Nr.", "Einh."]]
                    ultimul = len(array_original)
                    for i in range(len(array_original)):
                        if array_original[i][0] == "1":
                            for x in range(15, len(array_temp[0])):
                                for y in range(1, len(array_temp)):
                                    try:
                                        if array_temp[y][x] != "-":
                                            array_prelucrat.append(
                                                [array_temp[0][x], array_temp[y][x], array_temp[y][0],
                                                 array_temp[y][1], array_temp[y][2], array_temp[y][3],
                                                 array_temp[y][4], array_temp[y][5], array_temp[y][6],
                                                 array_temp[y][7], array_temp[y][8], array_temp[y][9],
                                                 array_temp[y][10], array_temp[y][11], array_temp[y][12],
                                                 array_temp[y][13], array_temp[y][14]])
                                    except:
                                        continue
                            del array_temp
                            array_temp = [["Bezeichnung", "VOBES-ID", "Benennung", "Verwendung", "Verwendung",
                                           "Kurzname", "xy", "Teilenummer", "Vorzugsteil", "TAB-Nummer", "Referenzteil",
                                           "Farbe", "E-Komponente", "E-Komponente Part-Nr.", "Einh."]]
                        if array_original[i][0] == "2":
                            array_temp[0].append(array_original[i][2])
                        if array_original[i][0] == "3" or array_original[i][0] == "4":
                            array_temporar2 = []
                            for t in range(16, len(array_original[i])):
                                array_temporar2.append(array_original[i][t])
                            array_temp.append(
                                [array_original[i][1], array_original[i][2], array_original[i][3], array_original[i][4],
                                 array_original[i][5], array_original[i][6], array_original[i][7], array_original[i][8],
                                 array_original[i][9], array_original[i][10], array_original[i][11],
                                 array_original[i][12], array_original[i][13], array_original[i][14],
                                 array_original[i][15]])
                            rng = len(array_temp)
                            for q in range(len(array_temporar2)):
                                array_temp[rng - 1].append(array_temporar2[q])
                            del array_temporar2
                        if i == ultimul - 1:
                            for x in range(15, len(array_temp[0])):
                                for y in range(1, len(array_temp)):
                                    try:
                                        if array_temp[y][x] != "-":
                                            array_prelucrat.append(
                                                [array_temp[0][x], array_temp[y][x], array_temp[y][0],
                                                 array_temp[y][1], array_temp[y][2], array_temp[y][3],
                                                 array_temp[y][4], array_temp[y][5], array_temp[y][6],
                                                 array_temp[y][7], array_temp[y][8], array_temp[y][9],
                                                 array_temp[y][10], array_temp[y][11], array_temp[y][12],
                                                 array_temp[y][13], array_temp[y][14]])
                                    except:
                                        continue
                    with open(os.path.abspath(os.curdir) + "/MAN/Input/BOMs/" + nume_fisier + ".csv", 'w',
                              newline='') as myfile:
                        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                        wr.writerows(array_prelucrat)
    log_file("Creat " + nume_fisier)
    end = time.time()
    pbar.destroy()
    pbargui.destroy()
    messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")
    return None

def wires():
    pbargui = Tk()
    pbargui.title("Prelucrare WIRELIST-uri")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    array_wirelisturi = ["8000", "8001", "8011", "8012", "8013", "8014", "8023", "8024", "8025", "8026", "8030", "8031",
                         "8052", "8053"]
    dir_WIRE = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir), title="Selectati directorul cu fisiere:")
    start = time.time()
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_WIRE):
        if file_all.endswith(".csv"):
            file_counter = file_counter + 1
    for file_all in os.listdir(dir_WIRE):
        if file_all.endswith(".csv"):
            file_progres = file_progres + 1
            statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 2
            pbargui.update_idletasks()
            for item in array_wirelisturi:
                if item in file_all:
                    try:
                        array_original = []
                        with open(dir_WIRE + "/" + file_all, newline='') as csvfile:
                            array_incarcat = list(csv.reader(csvfile, delimiter=';'))
                        if "AEM" in array_incarcat[0][1]:
                            for i in range(len(array_incarcat)):
                                del array_incarcat[i][1]
                        for x in range(len(array_incarcat)):
                            array_t = []
                            for i in range(len(array_incarcat[x])):
                                if i < 27 and array_incarcat[x][0] == "3":
                                    if array_incarcat[x][i] == "":
                                        array_t.append("X")
                                    else:
                                        array_t.append(array_incarcat[x][i])
                                elif array_incarcat[x][i] != "":
                                    array_t.append(array_incarcat[x][i])
                            array_original.append(array_t)
                        if array_original[0][0] != 1 and len(array_original[0]) != 2:
                            messagebox.showerror('Eroare fisier'+file_all, 'Nu ai incarcat fisierul corect ...')
                            return
                        if len(array_original) < 50:
                            for i in range(0, len(array_original)):
                                if array_original[i][0] == "0":
                                    if str(array_original[i][1]) != "Ltg-Nr.":
                                        messagebox.showerror('Eroare fisier' + file_all,
                                                             'Nu ai incarcat fisierul corect')
                                        return
                        else:
                            for i in range(0, 50):
                                if array_original[i][0] == "0":
                                    if str(array_original[i][1]) != "Ltg-Nr.":
                                        messagebox.showerror('Eroare fisier'+file_all, 'Nu ai incarcat fisierul corect')
                                        return
                    except:
                        messagebox.showerror('Eroare fisier'+file_all, 'Nu ai incarcat fisierul corect')
                        return
                    nume_fisier = os.path.splitext(os.path.basename(file_all))[0]
                    array_temp = [["Module", "Ltg No", "Leitung", "Farbe", "Quer.", "Kurzname", "Pin", "Kurzname",
                                   "Pin", "Sonderltg.", "Lange"]]
                    dict_v = []

                    def cutare_modul(numarv):
                        w = ""
                        if len(numarv) > 4:
                            numarv = numarv.split("  ")[1]
                        for q in range(len(dict_v)):
                            if numarv in dict_v[q]:
                                w = dict_v[q][1]
                        return w

                    for i in range(len(array_original)):
                        if array_original[i][0] == "1":
                            del dict_v
                            dict_v = []
                        if array_original[i][0] == "2":
                            dict_v.append([array_original[i][1], array_original[i][2]])
                        if array_original[i][0] == "0":
                            zero_index = i
                        if array_original[i][0] == "3":
                            for lungime_arr in range(29, len(array_original[i]) - 1):
                                if array_original[i][lungime_arr] != "-":
                                    try:
                                        array_temp.append([cutare_modul(array_original[zero_index][lungime_arr]),
                                                           array_original[i][1],
                                                           array_original[i][21], array_original[i][26],
                                                           float(array_original[i][27].replace(',', '.')),
                                                           array_original[i][7],
                                                           array_original[i][8], array_original[i][16],
                                                           array_original[i][17], array_original[i][24],
                                                           array_original[i][lungime_arr]])
                                    except ValueError:
                                        array_temp.append([cutare_modul(array_original[zero_index][lungime_arr]),
                                                           array_original[i][1],
                                                           array_original[i][15], array_original[i][19],
                                                           float(array_original[i][20].replace(',', '.')),
                                                           array_original[i][4],
                                                           array_original[i][5], array_original[i][10],
                                                           array_original[i][11], array_original[i][24],
                                                           array_original[i][lungime_arr]])
                    with open(os.path.abspath(os.curdir) + "/MAN/Input/Wire Lists/" + nume_fisier + ".csv", 'w',
                              newline='') as myfile:
                        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                        wr.writerows(array_temp)
                        log_file("Creat " + nume_fisier)

    end = time.time()
    pbar.destroy()
    pbargui.destroy()
    messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")
    return None

def boms_leoni():
    pbargui = Tk()
    pbargui.title("Prelucrare BOM-uri cu PN Leoni")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    global nume_fisier
    array_boms = ["8000", "8001", "8011", "8012", "8013", "8014", "8023", "8024", "8025", "8026", "8030", "8031",
                  "8052", "8053"]
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Component Overview.txt", newline='') as csvfile:
        compover = list(csv.reader(csvfile, delimiter=';'))
    dir_BOM = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir), title="Selectati directorul cu fisiere:")
    start = time.time()
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_BOM):
        if file_all.endswith(".csv"):
            file_counter = file_counter + 1
    for file_all in os.listdir(dir_BOM):
        if file_all.endswith(".csv"):
            file_progres = file_progres + 1
            statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 2
            pbargui.update_idletasks()
            for item in array_boms:
                if item in file_all:
                    try:
                        with open(dir_BOM + "/" + file_all, newline='') as csvfile:
                            array_original = list(csv.reader(csvfile, delimiter=';'))
                        if "AEM" in array_original[0][1]:
                            for i in range(len(array_original)):
                                del array_original[i][1]
                        if array_original[0][0] != "1" and len(array_original[0]) != 2:
                            messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect')
                            return
                        if len(array_original) < 50:
                            for i in range(0, len(array_original)):
                                if array_original[i][0] == "0":
                                    if str(array_original[i][-1])[:1] != "V":
                                        messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect')
                                        return
                        else:
                            for i in range(0, 50):
                                if array_original[i][0] == "0":
                                    if str(array_original[i][-1])[:1] != "V":
                                        messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect')
                                        return
                    except:
                        messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect')
                        return
                    for q in range(len(array_original)):
                        try:
                            if "Workplace" in array_original[q][7]:
                                array_original[q].pop(7)
                            if "Workplace" in array_original[q][6]:
                                array_original[q].pop(6)
                        except:
                            continue
                    nume_fisier = os.path.splitext(os.path.basename(file_all))[0]
                    array_prelucrat = [["Module", "Quantity", "Bezeichnung", "VOBES-ID", "Benennung", "Verwendung",
                                        "Verwendung", "Kurzname", "xy", "Teilenummer", "Vorzugsteil", "TAB-Nummer",
                                        "Referenzteil", "Farbe", "E-Komponente", "E-Komponente Part-Nr.", "Einh."]]
                    array_temp = [["Bezeichnung", "VOBES-ID", "Benennung", "Verwendung", "Verwendung", "Kurzname", "xy",
                                   "Teilenummer", "Vorzugsteil", "TAB-Nummer", "Referenzteil", "Farbe", "E-Komponente",
                                   "E-Komponente Part-Nr.", "Einh."]]
                    ultimul = len(array_original)
                    for i in range(len(array_original)):
                        if array_original[i][0] == "1":
                            for x in range(15, len(array_temp[0])):
                                for y in range(1, len(array_temp)):
                                    try:
                                        if array_temp[y][x] != "-":
                                            array_prelucrat.append(
                                                [array_temp[0][x], array_temp[y][x], array_temp[y][0],
                                                 array_temp[y][1], array_temp[y][2], array_temp[y][3],
                                                 array_temp[y][4], array_temp[y][5], array_temp[y][6],
                                                 array_temp[y][7], array_temp[y][8], array_temp[y][9],
                                                 array_temp[y][10], array_temp[y][11], array_temp[y][12],
                                                 array_temp[y][13], array_temp[y][14]])
                                    except:
                                        continue
                            del array_temp
                            array_temp = [["Bezeichnung", "VOBES-ID", "Benennung", "Verwendung", "Verwendung",
                                           "Kurzname", "xy", "Teilenummer", "Vorzugsteil", "TAB-Nummer", "Referenzteil",
                                           "Farbe", "E-Komponente", "E-Komponente Part-Nr.", "Einh."]]
                        if array_original[i][0] == "2":
                            array_temp[0].append(array_original[i][2])
                        if array_original[i][0] == "3" or array_original[i][0] == "4":
                            array_temporar2 = []
                            for t in range(16, len(array_original[i])):
                                array_temporar2.append(array_original[i][t])
                            array_temp.append(
                                [array_original[i][1], array_original[i][2], array_original[i][3], array_original[i][4],
                                 array_original[i][5], array_original[i][6], array_original[i][7], array_original[i][8],
                                 array_original[i][9], array_original[i][10], array_original[i][11],
                                 array_original[i][12], array_original[i][13], array_original[i][14],
                                 array_original[i][15]])
                            rng = len(array_temp)
                            for q in range(len(array_temporar2)):
                                array_temp[rng - 1].append(array_temporar2[q])
                            del array_temporar2
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                        if i == ultimul - 1:
                            for x in range(15, len(array_temp[0])):
                                for y in range(1, len(array_temp)):
                                    try:
                                        if array_temp[y][x] != "-":
                                            array_prelucrat.append(
                                                [array_temp[0][x], array_temp[y][x], array_temp[y][0],
                                                 array_temp[y][1], array_temp[y][2], array_temp[y][3],
                                                 array_temp[y][4], array_temp[y][5], array_temp[y][6],
                                                 array_temp[y][7], array_temp[y][8], array_temp[y][9],
                                                 array_temp[y][10], array_temp[y][11], array_temp[y][12],
                                                 array_temp[y][13], array_temp[y][14]])
                                    except:
                                        continue
                    array_prelucrat[0].extend(["PN Leoni 1", "PN Leoni 2"])
                    for i in range(1, len(array_prelucrat)):
                        for x in range(len(compover)):
                            if array_prelucrat[i][9] == compover[x][0]:
                                array_prelucrat[i].extend([compover[x][1], compover[x][2]])
                    prn_excel_bom_complete(array_prelucrat, nume_fisier)
                    pbar['value'] += 2
                    pbargui.update_idletasks()
    log_file("Creat " + nume_fisier)
    pbar.destroy()
    pbargui.destroy()
    end = time.time()
    messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")
    return None

def wires_leoni():
    pbargui = Tk()
    pbargui.title("Prelucrare WIRELIST-uri cu PN Leoni")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Component Overview.txt", newline='') as csvfile:
        compover = list(csv.reader(csvfile, delimiter=';'))

    array_wirelisturi = ["8000", "8001", "8011", "8012", "8013", "8014", "8023", "8024", "8025", "8026", "8030", "8031",
                         "8052", "8053"]
    dir_WIRE = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir), title="Selectati directorul cu fisiere:")
    start = time.time()
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_WIRE):
        if file_all.endswith(".csv"):
            file_counter = file_counter + 1
    for file_all in os.listdir(dir_WIRE):
        if file_all.endswith(".csv"):
            file_progres = file_progres + 1
            statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 2
            pbargui.update_idletasks()
            for item in array_wirelisturi:
                if item in file_all:
                    try:
                        array_original = []
                        with open(dir_WIRE + "/" + file_all, newline='') as csvfile:
                            array_incarcat = list(csv.reader(csvfile, delimiter=';'))
                        if "AEM" in array_incarcat[0][1]:
                            for i in range(len(array_incarcat)):
                                del array_incarcat[i][1]
                        for x in range(len(array_incarcat)):
                            array_t = []
                            for i in range(len(array_incarcat[x])):
                                if i < 27 and array_incarcat[x][0] == "3":
                                    if array_incarcat[x][i] == "":
                                        array_t.append("X")
                                    else:
                                        array_t.append(array_incarcat[x][i])
                                elif array_incarcat[x][i] != "":
                                    array_t.append(array_incarcat[x][i])
                            array_original.append(array_t)
                        if array_original[0][0] != 1 and len(array_original[0]) != 2:
                            messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect')
                            return
                        if len(array_original) < 50:
                            for i in range(0, len(array_original)):
                                if array_original[i][0] == "0":
                                    if str(array_original[i][1]) != "Ltg-Nr.":
                                        messagebox.showerror('Eroare fisier' + file_all,
                                                             'Nu ai incarcat fisierul corect')
                                        return
                        else:
                            for i in range(0, 50):
                                if array_original[i][0] == "0":
                                    if str(array_original[i][1]) != "Ltg-Nr.":
                                        messagebox.showerror('Eroare fisier' + file_all,
                                                             'Nu ai incarcat fisierul corect')
                                        return
                    except:
                        messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect')
                        return

                    nume_fisier = os.path.splitext(os.path.basename(file_all))[0]
                    array_temp = [["Module", "Ltg-Nr.", "Verbindung", "Komp. Ben.", "Verwendung", "Verwendung", "von",
                                   "Kurzname", "Pin", "xy", "Kontakt", "Dichtung", "Komp. Ben.", "Verwendung",
                                   "Verwendung", "nach", "Kurzname", "Pin", "xy", "Kontakt", "Dichtung", "Leitung",
                                   "Typ", "Vorzugsteil", "Sonderltg.", "Innenleiter", "Farbe", "Quer.", "Pot.",
                                   "Lange"]]
                    dict_v = []

                    def cutare_modul(numarv):
                        w = ""
                        if len(numarv) > 4:
                            numarv = numarv.split("  ")[1]
                        for e in range(len(dict_v)):
                            if numarv in dict_v[e]:
                                w = dict_v[e][1]
                        return w

                    for i in range(len(array_original)):
                        if array_original[i][0] == "1":
                            del dict_v
                            dict_v = []
                        if array_original[i][0] == "2":
                            dict_v.append([array_original[i][1], array_original[i][2]])
                        if array_original[i][0] == "0":
                            zero_index = i
                        if array_original[i][0] == "3":
                            for lungime_arr in range(29, len(array_original[i]) - 1):
                                if array_original[i][lungime_arr] != "-":
                                    try:
                                        array_temp.append([cutare_modul(array_original[zero_index][lungime_arr]),
                                                           array_original[i][1], array_original[i][2],
                                                           array_original[i][3], array_original[i][4],
                                                           array_original[i][5], array_original[i][6],
                                                           array_original[i][7], array_original[i][8],
                                                           array_original[i][9], array_original[i][10],
                                                           array_original[i][11], array_original[i][12],
                                                           array_original[i][13], array_original[i][14],
                                                           array_original[i][15], array_original[i][16],
                                                           array_original[i][17], array_original[i][18],
                                                           array_original[i][19], array_original[i][20],
                                                           array_original[i][21], array_original[i][22],
                                                           array_original[i][23], array_original[i][24],
                                                           array_original[i][25], array_original[i][26],
                                                           float(array_original[i][27].replace(',', '.')),
                                                           array_original[i][28], array_original[i][lungime_arr]])
                                    except ValueError:
                                        array_temp.append([cutare_modul(array_original[zero_index][lungime_arr]),
                                                           array_original[i][1], array_original[i][2],
                                                           array_original[i][3], array_original[i][4],
                                                           array_original[i][5], array_original[i][6],
                                                           array_original[i][7], array_original[i][8],
                                                           array_original[i][9], array_original[i][10],
                                                           array_original[i][11], array_original[i][12],
                                                           array_original[i][13], array_original[i][14],
                                                           array_original[i][15], array_original[i][16],
                                                           array_original[i][17], array_original[i][18],
                                                           array_original[i][19], array_original[i][20],
                                                           array_original[i][21], array_original[i][22],
                                                           array_original[i][23], array_original[i][24],
                                                           array_original[i][25], array_original[i][26],
                                                           float(array_original[i][27].replace(',', '.')),
                                                           array_original[i][28], array_original[i][lungime_arr]])
                    pbar['value'] += 2
                    pbargui.update_idletasks()
                    array_man = [["Module", "Ltg-Nr.", "Kurzname", "Pin", "Kontakt", "Dichtung", "Kurzname", "Pin",
                                  "Kontakt", "Dichtung", "Leitung", "Typ", "Sonderltg.", "Farbe", "Quer.",
                                  "Pot.", "Lange"]]
                    array_write_leoni = [["Module", "Nr. Fir", "Conector", "Pin", "Contact", "Seal", "Conector", "Pin",
                                          "Contact", "Seal", "Cablu", "Typ", "Sonderltg.", "Culoare", "Sectiune",
                                          "Pot.", "Lungime"]]
                    array_print = [["Module", "Ltg-Nr.", "Kurzname", "Pin", "Kontakt", "Dichtung", "Kurzname", "Pin",
                                    "Kontakt", "Dichtung", "Leitung", "Typ", "Sonderltg.", "Farbe", "Quer.", "Pot.",
                                    "Lange"]]
                    for x in range(1, len(array_temp)):
                        array_man.append([array_temp[x][0], array_temp[x][1], array_temp[x][7], array_temp[x][8],
                                          array_temp[x][10], array_temp[x][11], array_temp[x][16], array_temp[x][17],
                                          array_temp[x][19], array_temp[x][20], array_temp[x][21], array_temp[x][22],
                                          array_temp[x][24], array_temp[x][26], array_temp[x][27], array_temp[x][28],
                                          array_temp[x][29]])
                        array_print.append([array_temp[x][0], array_temp[x][1], array_temp[x][7], array_temp[x][8],
                                            array_temp[x][10], array_temp[x][11], array_temp[x][16], array_temp[x][17],
                                            array_temp[x][19], array_temp[x][20], array_temp[x][21], array_temp[x][22],
                                            array_temp[x][24], array_temp[x][26], array_temp[x][27], array_temp[x][28],
                                            array_temp[x][29]])
                    for x in range(1, len(array_man)):
                        array_write_leoni.append(array_man[x])

                    def compsearch(search_string):
                        for d in range(len(compover)):
                            if compover[d][0] == search_string:
                                return compover[d][1]

                    for y in range(1, len(array_write_leoni)):
                        for q in [4, 5, 8, 9, 10]:
                            if array_write_leoni[y][q] != "-":
                                array_write_leoni[y][q] = compsearch(array_write_leoni[y][q])
                    prn_excel_wires_complete_leoni(array_print, array_write_leoni, nume_fisier)

    pbar.destroy()
    pbargui.destroy()
    end = time.time()
    messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")
    return None


def wirelist_all_simplu():
    pbargui = Tk()
    pbargui.title("Wirelist All In One")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    dir_selectat = os.path.abspath(os.curdir) + "/MAN/Input/Wire Lists/"
    array_print = []
    array_print2 = []
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_selectat):
        if file_all.endswith(".csv") and not file_all.startswith("All"):
            file_counter = file_counter + 1
    for file_all in os.listdir(dir_selectat):
        if file_all.endswith(".csv"):
            file_progres = file_progres + 1
            statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 2
            pbargui.update_idletasks()
            with open(dir_selectat + file_all, newline='') as csvfile:
                array_wirelist = list(csv.reader(csvfile, delimiter=';'))
            for i in range(len(array_wirelist)):
                array_print.append([array_wirelist[i][0], array_wirelist[i][1], array_wirelist[i][2],
                                   array_wirelist[i][3], array_wirelist[i][4], array_wirelist[i][5],
                                   array_wirelist[i][6], array_wirelist[i][9]])
                pbar['value'] += 2
                pbargui.update_idletasks()
            for i in range(len(array_wirelist)):
                array_print.append([array_wirelist[i][0], array_wirelist[i][1], array_wirelist[i][2],
                                   array_wirelist[i][3], array_wirelist[i][4], array_wirelist[i][7],
                                   array_wirelist[i][8], array_wirelist[i][9]])
                pbar['value'] += 2
                pbargui.update_idletasks()
            for i in range(len(array_wirelist)):
                array_print2.append(array_wirelist[i])
            continue
        else:
            continue
    prn_excel_wirelistsallinone(array_print, array_print2)
    pbar.destroy()
    pbargui.destroy()
    return None


def wirelist_all_complet():
    pbargui = Tk()
    pbargui.title("Wirelist All In One Complet")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)

    dir_selectat = os.path.abspath(os.curdir) + "/MAN/Output/Complete BOM and WIRELIST/Wirelist/"
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_selectat):
        if file_all.endswith(".xlsx") and not file_all.startswith("Leoni"):
            file_counter = file_counter + 1
    if file_counter == 0:
        messagebox.showinfo('Finalizat!', "Directorul Complete BOM and WIRELIST/Wirelist/ este gol")
        pbar.destroy()
        pbargui.destroy()
        return None
    all_data = pd.DataFrame()
    for file_all in os.listdir(dir_selectat):
        if file_all.endswith(".xlsx") and not file_all.startswith("Leoni"):
            path = os.path.join(dir_selectat, file_all)
            df = pd.read_excel(path)
            all_data = all_data.append(df, ignore_index=True)
            file_progres = file_progres + 1
            statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 2
    all_data.to_csv(os.path.abspath(os.curdir) + "/MAN/Input/Others/Wirelist Complet.txt", encoding='utf-8',
                    index=False, sep=';')
    pbar.destroy()
    pbargui.destroy()
    messagebox.showinfo('Finalizat!')
    return None



import csv
import os
import time
from tkinter import messagebox, filedialog, Tk, ttk, HORIZONTAL, Label
from openpyxl import load_workbook
from diverse import log_file
from functii_print import prn_excel_separare_ksk
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
    # Open JIT file
    fisier_calloff = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                                title="Incarcati fisierul care necesita sortare")
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
                     ws.cell(row=row.row, column=16).value, ws.cell(row=row.row, column=16).value])
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
        # Create LHD and RHD files for each KSK from JIT
        stanga = []
        dreapta = []
        for x in range(len(array_temporar)):
            if array_temporar[x][2] == "BODYL":
                stanga.append([array_temporar[x][1]])
            elif array_temporar[x][2] == "BODYR":
                dreapta.append([array_temporar[x][1]])
        with open(os.path.abspath(os.curdir) + "/MAN/Input/Module Files/Comparatii/LHD/"
                  + element[1:] + ".csv", 'w', newline='') as myfile:
            wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
            wr.writerows(stanga)
        with open(os.path.abspath(os.curdir) + "/MAN/Input/Module Files/Comparatii/RHD/"
                  + element[1:] + ".csv", 'w', newline='') as myfile:
            wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
            wr.writerows(dreapta)

        # Check harness type
        harnesstype = []
        for m in range(len(array_temporar_module)):
            for n in range(len(array_module_active)):
                if array_temporar_module[m] == array_module_active[n][0] and array_module_active[n][3] != "XXXX":
                    harnesstype.append(array_module_active[n][3].replace(' LHD', '').replace(' RHD', ''))
        harnesstype = list(set(harnesstype))
        # Write to database
        array_database.append([os.path.basename(fisier_calloff), ';'.join(harnesstype), is_light, array_temporar[1][8], data_download,
                               element[1:], array_temporar[1][7], ';'.join(array_temporar_module)])
        conn = sqlite3.connect(os.path.abspath(os.curdir) + "/MAN/Input/Others/database.db")
        cursor = conn.cursor()
        # create a table
        cursor.execute("""CREATE TABLE IF NOT EXISTS KSKDatabase
                          (numejit text, tip text, light text, datalivrare text, datajit text,harness text, 
                           trailerno text, listamodule text) 
                       """)
        # insert multiple records using the more secure "?" method
        albums = []
        cursor.executemany("INSERT INTO KSKDatabase VALUES (?,?,?,?,?,?,?,?)", array_database)
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

    dir_Jit = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir),
                                      title="Selectati directorul cu fisiere JIT:")
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
                             ws.cell(row=row.row, column=16).value])
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
                # Create LHD and RHD files for each KSK from JIT
                stanga = []
                dreapta = []
                for x in range(len(array_temporar)):
                    if array_temporar[x][2] == "BODYL":
                        stanga.append([array_temporar[x][1]])
                    elif array_temporar[x][2] == "BODYR":
                        dreapta.append([array_temporar[x][1]])
                with open(os.path.abspath(os.curdir) + "/MAN/Input/Module Files/Comparatii/LHD/"
                          + element[1:] + ".csv", 'w', newline='') as myfile:
                    wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                    wr.writerows(stanga)
                with open(os.path.abspath(os.curdir) + "/MAN/Input/Module Files/Comparatii/RHD/"
                          + element[1:] + ".csv", 'w', newline='') as myfile:
                    wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                    wr.writerows(dreapta)

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
    pbar.destroy()
    pbargui.destroy()
    messagebox.showinfo('Finalizat!', str(file_progres) + " fisiere din " + str(file_counter))


def golire_directoare_comparati():
    dir_input1 = os.path.abspath(os.curdir) + "/MAN/Input/Module Files/Comparatii/LHD/"
    dir_input2 = os.path.abspath(os.curdir) + "/MAN/Input/Module Files/Comparatii/RHD/"
    for file_all in os.listdir(dir_input1):
        try:
            os.remove(dir_input1 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_input2):
        try:
            os.remove(dir_input2 + file_all)
        except:
            continue
    messagebox.showinfo("Golire", "Directoarele Input si Output au fost golite!!")

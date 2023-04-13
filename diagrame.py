import csv
import os
import time
from tkinter import Tk, ttk, HORIZONTAL, Label, filedialog, messagebox
from openpyxl import load_workbook
from functii_print import prn_excel_diagrame, prn_excel_asocierediagramemodule, prn_excel_diagrameinksk, \
    prn_excel_infoindiagrame, prn_excel_diagrameinksk_total
import pandas as pd
import datetime


def comparatiediagrame():
    pbargui = Tk()
    pbargui.title("Comparatie diagrame")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    dir_old = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir),
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
    array_output = []
    try:
        with open(file_load, newline='') as csvfile:
            array_module_file = list(csv.reader(csvfile, delimiter=';'))
        if array_module_file[0][0] != "Harness" and array_module_file[0][0] != "Module":
            messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect. Eroare cap de tabel!')
            return
        file_name = array_module_file[2][0]
        moduleinksk = [row[1] for row in array_module_file]
        with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/MatrixDiagrame.txt", newline='') as csvfile:
            array_diagrame = list(csv.reader(csvfile, delimiter=';'))
        array_diagrame[0].append("Nota")
        statuslabel = Label(pbargui, text="Working . . .")
        for i in range(1, len(array_diagrame)):
            counter = 0
            for x in range(len(moduleinksk)):
                if moduleinksk[x] in array_diagrame[i]:
                    counter += 1
                    pbar['value'] += 2
                    pbargui.update_idletasks()
            array_diagrame[i].append(counter)
        dfdiagrame = pd.DataFrame(array_diagrame)
        dfdiagrame.columns = dfdiagrame.iloc[0]
        dfdiagrame = dfdiagrame[1:]
        df_max = dfdiagrame.groupby('Diagrama', as_index=False)['Nota'].max()
        array_output_temp = df_max.values.tolist()
        for i in range(len(array_output_temp)):
            if array_output_temp[i][1] != 0:
                array_output.append(array_output_temp[i])
        statuslabel["text"] = "Printing file . . . "
        pbar['value'] += 2
        pbargui.update_idletasks()
        prn_excel_diagrameinksk(array_output, file_name)
        pbar.destroy()
        pbargui.destroy()
        end = time.time()
        messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")
    except FileNotFoundError:
        pbar.destroy()
        pbargui.destroy()
        return None


def extragere_informatii_diagrame():
    pbargui = Tk()
    pbargui.title("Extragere informatii din diagrame")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    dir_diagrame = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir),
                                           title="Selectati directorul cu diagramele:")
    start = time.time()
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_diagrame):
        if file_all.endswith(".xlsx"):
            file_counter = file_counter + 1
    if file_counter == 0:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showwarning('Eroare!', "Directorul selectat este gol.")
    array_informatii = []
    for file_all in os.listdir(dir_diagrame):
        if file_all.endswith(".xlsx"):
            file_progres = file_progres + 1
            statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 2
            pbargui.update_idletasks()
            wb1 = load_workbook(dir_diagrame + "/" + file_all)
            sheet1 = wb1.worksheets[0]
            for row in range(1, sheet1.max_row + 1):
                for col in range(1, sheet1.max_column + 1):
                    cell1 = sheet1.cell(row, col)
                    if cell1.value is not None:
                        array_informatii.append([file_all, cell1.value, row, col])
    array_informatii.insert(0, ["Nume Diagrama", "Text celula", "Rand", "Coloana"])
    prn_excel_infoindiagrame(array_informatii)
    pbar.destroy()
    pbargui.destroy()
    end = time.time()
    messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")


def diagrameinkskfolder():
    pbargui = Tk()
    pbargui.title("Lista diagrame in KSK")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    dir_files = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir) + "/MAN/Input/Module Files",
                                        title="Selectati directorul cu fisiere:")
    start = time.time()
    file_counter = 0
    file_progres = 0
    array_output_total = []
    for file_all in os.listdir(dir_files):
        if file_all.endswith(".csv"):
            file_counter = file_counter + 1
    for file_all in os.listdir(dir_files):
        if file_all.endswith(".csv"):
            array_output = []
            file_progres = file_progres + 1
            statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 2
            pbargui.update_idletasks()
            with open(dir_files + "/" + file_all, newline='') as csvfile:
                array_module_file = list(csv.reader(csvfile, delimiter=';'))
            if array_module_file[0][0] != "Harness" and array_module_file[0][0] != "Module":
                messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect. Eroare cap de tabel!')
                return
            file_name = array_module_file[2][0]
            moduleinksk = [row[1] for row in array_module_file]
            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/MatrixDiagrame.txt", newline='') as csvfile:
                array_diagrame = list(csv.reader(csvfile, delimiter=';'))
            array_diagrame[0].append("Nota")
            for i in range(1, len(array_diagrame)):
                counter = 0
                for x in range(len(moduleinksk)):
                    if moduleinksk[x] in array_diagrame[i]:
                        counter += 1
                array_diagrame[i].append(counter)
            dfdiagrame = pd.DataFrame(array_diagrame)
            dfdiagrame.columns = dfdiagrame.iloc[0]
            dfdiagrame = dfdiagrame[1:]
            df_max = dfdiagrame.groupby('Diagrama', as_index=False)['Nota'].max()
            array_output_temp = df_max.values.tolist()
            for i in range(len(array_output_temp)):
                if array_output_temp[i][1] != 0:
                    array_output.append(array_output_temp[i])
            pbar['value'] += 2
            pbargui.update_idletasks()
            prn_excel_diagrameinksk(array_output, file_name)
            for x in range(len(array_output)):
                array_output_total.append(array_output[x][0])
    save_time = datetime.datetime.now().strftime("%d.%m.%Y_%H.%M.%S")
    prn_excel_diagrameinksk_total(array_output_total, save_time + " - " + str(file_counter))
    pbar.destroy()
    pbargui.destroy()
    end = time.time()
    messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")


def asocierediagramemodule():
    pbargui = Tk()
    pbargui.title("Creare Matrix Diagrame (Matrix_Module)")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    dir_files = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir),
                                        title="Selectati directorul cu fisierele Matrix_Module:")
    start = time.time()
    file_counter = 0
    file_progres = 0
    array_output = [["Drawing", "Nr Crt", "Nota", "CT", "AEM", "Nota pentru AEM", "Loc", "Nr. PB/Art.", "GROUP",
                     "Numele modulului", "Nume_Norm", "Nume_ADR", "Conector", "Fir", "KN", "XCode1", "Cavity1",
                     "XCode2", "Cavity2", "Patrunderea modulului", "adr module",
                     "ALL_Supersleeve", "Module_ADR", "Module_NonADR", "Module 1", "Module 2",
                     "Module 3", "Module 4", "Module 5", "Module 6", "Module 7", "Module 1", "Module 2", "Module 3",
                     "Module 4", "Module 5", "Module 6", "Module 7"]]
    for file_all in os.listdir(dir_files):
        if file_all.endswith(".xlsm"):
            file_counter = file_counter + 1
    for file_all in os.listdir(dir_files):
        if file_all.endswith(".xlsm"):
            file_progres = file_progres + 1
            statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + " : " + file_all
            pbar['value'] += 2
            pbargui.update_idletasks()
            wb = load_workbook(dir_files + "/" + file_all)
            ws = wb.active
            mod1counter = 0
            mod2counter = 0
            mod3counter = 0
            mod4counter = 0
            mod5counter = 0
            mod6counter = 0
            mod7counter = 0
            pbart = False
            pb = False
            index_mod6 = 0
            index_mod7 = 0
            for row in ws['A']:
                if row.value == "№":
                    index_nrcrt = row.row
                    for column_cells in ws.iter_cols(min_row=index_nrcrt, max_row=index_nrcrt):
                        for cell in column_cells:
                            if cell.value == "№ПБ/Ст.":
                                pbart = True
                            if cell.value == "PB":
                                pb = True
                    if pbart is True:
                        for row_cells in ws.iter_rows(min_row=row.row, max_row=row.row):
                            for cell in row_cells:
                                if cell.value == "Примітка":
                                    index_nota = cell.column
                                if cell.value == "СТ.":
                                    index_ct = cell.column
                                else:
                                    index_ct = 1
                                if cell.value == "АЕМ":
                                    index_aem = cell.column
                                if cell.value == "Примітка до АЕМ" or cell.value == "Прим до АЕМ":
                                    index_notaaem = cell.column
                                if cell.value == "Місце":
                                    index_loc = cell.column
                                elif cell.value == "№ПБ/Ст.":
                                    index_pbart = cell.column
                                if cell.value == "GROUP":
                                    index_group = cell.column
                                if cell.value == "Назва модуля":
                                    index_modul = cell.column
                                if cell.value == "Name_Norm":
                                    index_numenormal = cell.column
                                if cell.value == "Name_ADR":
                                    index_numeadr = cell.column
                                if cell.value == "Роз'єм":
                                    index_conector = cell.column
                                if cell.value == "Провід":
                                    index_fir = cell.column
                                if cell.value == "KN":
                                    index_kn = cell.column
                                if cell.value == "XCode1":
                                    index_xc1 = cell.column
                                if cell.value == "Cavity1":
                                    index_cav1 = cell.column
                                if cell.value == "XCode2":
                                    index_xc2 = cell.column
                                if cell.value == "Cavity2":
                                    index_cav2 = cell.column
                                if cell.value == "Пенетрація модуля":
                                    index_patrundere = cell.column
                                if cell.value == "adr module":
                                    index_adrmod = cell.column
                                if cell.value == "ALL_Supersleeve":
                                    index_allss = cell.column
                                if cell.value == "Module_ADR" or cell.value == "Adr":
                                    index_modadr = cell.column
                                if cell.value == "Module_NonADR" or cell.value == "normal":
                                    index_modnonadr = cell.column
                                if (cell.value == "Module 1" or cell.value == "Module1") and mod1counter == 0:
                                    index_mod1 = cell.column
                                    mod1counter = 1
                                if (cell.value == "Module 2" or cell.value == "Module2") and mod2counter == 0:
                                    index_mod2 = cell.column
                                    mod2counter = 1
                                if (cell.value == "Module 3" or cell.value == "Module3") and mod3counter == 0:
                                    index_mod3 = cell.column
                                    mod3counter = 1
                                if (cell.value == "Module 4" or cell.value == "Module4") and mod4counter == 0:
                                    index_mod4 = cell.column
                                    mod4counter = 1
                                if (cell.value == "Module 5" or cell.value == "Module5") and mod5counter == 0:
                                    index_mod5 = cell.column
                                    mod5counter = 1
                                if (cell.value == "Module 6" or cell.value == "Module6") and mod6counter == 0:
                                    index_mod6 = cell.column
                                    mod6counter = 1
                                if (cell.value == "Module 7" or cell.value == "Module7") and mod7counter == 0:
                                    index_mod7 = cell.column
                                    mod7counter = 1
                                if (cell.value == "Module 1" or cell.value == "Module1") and mod1counter == 1:
                                    index_mod11 = cell.column
                                if (cell.value == "Module 2" or cell.value == "Module2") and mod2counter == 1:
                                    index_mod22 = cell.column
                                if (cell.value == "Module 3" or cell.value == "Module3") and mod3counter == 1:
                                    index_mod33 = cell.column
                                if (cell.value == "Module 4" or cell.value == "Module4") and mod4counter == 1:
                                    index_mod44 = cell.column
                                if (cell.value == "Module 5" or cell.value == "Module5") and mod5counter == 1:
                                    index_mod55 = cell.column
                                if (cell.value == "Module 6" or cell.value == "Module6") and mod6counter == 1:
                                    index_mod66 = cell.column
                                if (cell.value == "Module 7" or cell.value == "Module7") and mod7counter == 1:
                                    index_mod77 = cell.column
                    if pb is True:
                        for row_cells in ws.iter_rows(min_row=row.row, max_row=row.row):
                            for cell in row_cells:
                                if cell.value == "Примітка":
                                    index_nota = cell.column
                                if cell.value == "СТ.":
                                    index_ct = cell.column
                                if cell.value == "АЕМ":
                                    index_aem = cell.column
                                if cell.value == "Примітка до АЕМ":
                                    index_notaaem = cell.column
                                if cell.value == "Місце":
                                    index_loc = cell.column
                                elif cell.value == "PB":
                                    index_pbart = cell.column
                                if cell.value == "GROUP":
                                    index_group = cell.column
                                if cell.value == "Назва модуля":
                                    index_modul = cell.column
                                if cell.value == "Name_Norm":
                                    index_numenormal = cell.column
                                if cell.value == "Name_ADR":
                                    index_numeadr = cell.column
                                if cell.value == "Роз'єм":
                                    index_conector = cell.column
                                if cell.value == "Провід":
                                    index_fir = cell.column
                                if cell.value == "KN":
                                    index_kn = cell.column
                                if cell.value == "XCode1":
                                    index_xc1 = cell.column
                                if cell.value == "Cavity1":
                                    index_cav1 = cell.column
                                if cell.value == "XCode2":
                                    index_xc2 = cell.column
                                if cell.value == "Cavity2":
                                    index_cav2 = cell.column
                                if cell.value == "Пенетрація модуля":
                                    index_patrundere = cell.column
                                if cell.value == "adr module":
                                    index_adrmod = cell.column
                                if cell.value == "ALL_Supersleeve":
                                    index_allss = cell.column
                                if cell.value == "Module_ADR" or cell.value == "Adr":
                                    index_modadr = cell.column
                                if cell.value == "Module_NonADR" or cell.value == "normal":
                                    index_modnonadr = cell.column
                                if (cell.value == "Module 1" or cell.value == "Module1") and mod1counter == 0:
                                    index_mod1 = cell.column
                                    mod1counter = 1
                                if (cell.value == "Module 2" or cell.value == "Module2") and mod2counter == 0:
                                    index_mod2 = cell.column
                                    mod2counter = 1
                                if (cell.value == "Module 3" or cell.value == "Module3") and mod3counter == 0:
                                    index_mod3 = cell.column
                                    mod3counter = 1
                                if (cell.value == "Module 4" or cell.value == "Module4") and mod4counter == 0:
                                    index_mod4 = cell.column
                                    mod4counter = 1
                                if (cell.value == "Module 5" or cell.value == "Module5") and mod5counter == 0:
                                    index_mod5 = cell.column
                                    mod5counter = 1
                                if (cell.value == "Module 6" or cell.value == "Module6") and mod6counter == 0:
                                    index_mod6 = cell.column
                                    mod6counter = 1
                                if (cell.value == "Module 7" or cell.value == "Module7") and mod7counter == 0:
                                    index_mod7 = cell.column
                                    mod7counter = 1
                                if (cell.value == "Module 1" or cell.value == "Module1") and mod1counter == 1:
                                    index_mod11 = cell.column
                                if (cell.value == "Module 2" or cell.value == "Module2") and mod2counter == 1:
                                    index_mod22 = cell.column
                                if (cell.value == "Module 3" or cell.value == "Module3") and mod3counter == 1:
                                    index_mod33 = cell.column
                                if (cell.value == "Module 4" or cell.value == "Module4") and mod4counter == 1:
                                    index_mod44 = cell.column
                                if (cell.value == "Module 5" or cell.value == "Module5") and mod5counter == 1:
                                    index_mod55 = cell.column
                                if (cell.value == "Module 6" or cell.value == "Module6") and mod6counter == 1:
                                    index_mod66 = cell.column
                                if (cell.value == "Module 7" or cell.value == "Module7") and mod7counter == 1:
                                    index_mod77 = cell.column
                    if pbart is False and pb is False:
                        for row_cells in ws.iter_rows(min_row=row.row, max_row=row.row):
                            for cell in row_cells:
                                if cell.value == "Примітка":
                                    index_nota = cell.column
                                if cell.value == "СТ.":
                                    index_ct = cell.column
                                if cell.value == "АЕМ":
                                    index_aem = cell.column
                                if cell.value == "Примітка до АЕМ":
                                    index_notaaem = cell.column
                                if cell.value == "Place":
                                    index_loc = cell.column
                                elif cell.value == "Місце":
                                    index_pbart = cell.column
                                if cell.value == "GROUP":
                                    index_group = cell.column
                                if cell.value == "Назва модуля":
                                    index_modul = cell.column
                                if cell.value == "Name_Norm":
                                    index_numenormal = cell.column
                                if cell.value == "Name_ADR":
                                    index_numeadr = cell.column
                                if cell.value == "Роз'єм":
                                    index_conector = cell.column
                                if cell.value == "Провід":
                                    index_fir = cell.column
                                if cell.value == "KN":
                                    index_kn = cell.column
                                if cell.value == "XCode1":
                                    index_xc1 = cell.column
                                if cell.value == "Cavity1":
                                    index_cav1 = cell.column
                                if cell.value == "XCode2":
                                    index_xc2 = cell.column
                                if cell.value == "Cavity2":
                                    index_cav2 = cell.column
                                if cell.value == "Пенетрація модуля":
                                    index_patrundere = cell.column
                                if cell.value == "adr module":
                                    index_adrmod = cell.column
                                if cell.value == "ALL_Supersleeve":
                                    index_allss = cell.column
                                if cell.value == "Module_ADR" or cell.value == "Adr":
                                    index_modadr = cell.column
                                if cell.value == "Module_NonADR" or cell.value == "normal":
                                    index_modnonadr = cell.column
                                if (cell.value == "Module 1" or cell.value == "Module1") and mod1counter == 0:
                                    index_mod1 = cell.column
                                    mod1counter = 1
                                if (cell.value == "Module 2" or cell.value == "Module2") and mod2counter == 0:
                                    index_mod2 = cell.column
                                    mod2counter = 1
                                if (cell.value == "Module 3" or cell.value == "Module3") and mod3counter == 0:
                                    index_mod3 = cell.column
                                    mod3counter = 1
                                if (cell.value == "Module 4" or cell.value == "Module4") and mod4counter == 0:
                                    index_mod4 = cell.column
                                    mod4counter = 1
                                if (cell.value == "Module 5" or cell.value == "Module5") and mod5counter == 0:
                                    index_mod5 = cell.column
                                    mod5counter = 1
                                if (cell.value == "Module 6" or cell.value == "Module6") and mod6counter == 0:
                                    index_mod6 = cell.column
                                    mod6counter = 1
                                if (cell.value == "Module 7" or cell.value == "Module7") and mod7counter == 0:
                                    index_mod7 = cell.column
                                    mod7counter = 1
                                if (cell.value == "Module 1" or cell.value == "Module1") and mod1counter == 1:
                                    index_mod11 = cell.column
                                if (cell.value == "Module 2" or cell.value == "Module2") and mod2counter == 1:
                                    index_mod22 = cell.column
                                if (cell.value == "Module 3" or cell.value == "Module3") and mod3counter == 1:
                                    index_mod33 = cell.column
                                if (cell.value == "Module 4" or cell.value == "Module4") and mod4counter == 1:
                                    index_mod44 = cell.column
                                if (cell.value == "Module 5" or cell.value == "Module5") and mod5counter == 1:
                                    index_mod55 = cell.column
                                if (cell.value == "Module 6" or cell.value == "Module6") and mod6counter == 1:
                                    index_mod66 = cell.column
                                if (cell.value == "Module 7" or cell.value == "Module7") and mod7counter == 1:
                                    index_mod77 = cell.column

            for row in range(index_nrcrt + 1, ws.max_row):
                if index_mod6 != 0 and index_mod7 != 0 and ws.cell(row, index_loc).value != "":
                    array_output.append([file_all,
                                         ws.cell(row, index_nrcrt).value, ws.cell(row, index_nota).value,
                                         ws.cell(row, index_ct).value, ws.cell(row, index_aem).value,
                                         ws.cell(row, index_notaaem).value, ws.cell(row, index_loc).value,
                                         ws.cell(row, index_pbart).value, ws.cell(row, index_group).value,
                                         ws.cell(row, index_modul).value,
                                         ws.cell(row, index_numenormal).value,
                                         ws.cell(row, index_numeadr).value,
                                         ws.cell(row, index_conector).value,
                                         ws.cell(row, index_fir).value, ws.cell(row, index_kn).value,
                                         ws.cell(row, index_xc1).value, ws.cell(row, index_cav1).value,
                                         ws.cell(row, index_xc2).value, ws.cell(row, index_cav2).value,
                                         ws.cell(row, index_patrundere).value,
                                         ws.cell(row, index_adrmod).value,
                                         ws.cell(row, index_allss).value, ws.cell(row, index_modadr).value,
                                         ws.cell(row, index_modnonadr).value,
                                         ws.cell(row, index_mod1).value,
                                         ws.cell(row, index_mod2).value, ws.cell(row, index_mod3).value,
                                         ws.cell(row, index_mod4).value, ws.cell(row, index_mod5).value,
                                         ws.cell(row, index_mod6).value, ws.cell(row, index_mod7).value,
                                         ws.cell(row, index_mod11).value, ws.cell(row, index_mod22).value,
                                         ws.cell(row, index_mod33).value, ws.cell(row, index_mod44).value,
                                         ws.cell(row, index_mod55).value, ws.cell(row, index_mod66).value,
                                         ws.cell(row, index_mod77).value])
                if index_mod6 != 0 and index_mod7 == 0 and ws.cell(row, index_loc).value != "":
                    array_output.append([file_all,
                                         ws.cell(row, index_nrcrt).value, ws.cell(row, index_nota).value,
                                         ws.cell(row, index_ct).value, ws.cell(row, index_aem).value,
                                         ws.cell(row, index_notaaem).value, ws.cell(row, index_loc).value,
                                         ws.cell(row, index_pbart).value, ws.cell(row, index_group).value,
                                         ws.cell(row, index_modul).value,
                                         ws.cell(row, index_numenormal).value,
                                         ws.cell(row, index_numeadr).value,
                                         ws.cell(row, index_conector).value,
                                         ws.cell(row, index_fir).value, ws.cell(row, index_kn).value,
                                         ws.cell(row, index_xc1).value, ws.cell(row, index_cav1).value,
                                         ws.cell(row, index_xc2).value, ws.cell(row, index_cav2).value,
                                         ws.cell(row, index_patrundere).value,
                                         ws.cell(row, index_adrmod).value,
                                         ws.cell(row, index_allss).value, ws.cell(row, index_modadr).value,
                                         ws.cell(row, index_modnonadr).value,
                                         ws.cell(row, index_mod1).value,
                                         ws.cell(row, index_mod2).value, ws.cell(row, index_mod3).value,
                                         ws.cell(row, index_mod4).value, ws.cell(row, index_mod5).value,
                                         ws.cell(row, index_mod6).value,
                                         ws.cell(row, index_mod11).value, ws.cell(row, index_mod22).value,
                                         ws.cell(row, index_mod33).value, ws.cell(row, index_mod44).value,
                                         ws.cell(row, index_mod55).value, ws.cell(row, index_mod66).value])
                elif ws.cell(row, index_loc).value != "":
                    array_output.append([file_all,
                                         ws.cell(row, index_nrcrt).value, ws.cell(row, index_nota).value,
                                         ws.cell(row, index_ct).value, ws.cell(row, index_aem).value,
                                         ws.cell(row, index_notaaem).value, ws.cell(row, index_loc).value,
                                         ws.cell(row, index_pbart).value, ws.cell(row, index_group).value,
                                         ws.cell(row, index_modul).value,
                                         ws.cell(row, index_numenormal).value,
                                         ws.cell(row, index_numeadr).value,
                                         ws.cell(row, index_conector).value,
                                         ws.cell(row, index_fir).value, ws.cell(row, index_kn).value,
                                         ws.cell(row, index_xc1).value, ws.cell(row, index_cav1).value,
                                         ws.cell(row, index_xc2).value, ws.cell(row, index_cav2).value,
                                         ws.cell(row, index_patrundere).value,
                                         ws.cell(row, index_adrmod).value,
                                         ws.cell(row, index_allss).value, ws.cell(row, index_modadr).value,
                                         ws.cell(row, index_modnonadr).value,
                                         ws.cell(row, index_mod1).value,
                                         ws.cell(row, index_mod2).value, ws.cell(row, index_mod3).value,
                                         ws.cell(row, index_mod4).value, ws.cell(row, index_mod5).value,
                                         ws.cell(row, index_mod11).value, ws.cell(row, index_mod22).value,
                                         ws.cell(row, index_mod33).value, ws.cell(row, index_mod44).value,
                                         ws.cell(row, index_mod55).value])
    for i in range(0, 100):
        print(array_output[i])
    statuslabel["text"] = "Printing file . . . "
    pbar['value'] += 2
    pbargui.update_idletasks()
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/MatrixDiagrame.txt", 'w', newline='', encoding='utf8') as myfile:
        wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
        wr.writerows(array_output)
    pbar.destroy()
    pbargui.destroy()
    end = time.time()
    messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")


def indexarediagrame():
    pbargui = Tk()
    pbargui.title("Indexare Matrix_Module")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    start = time.time()
    file_progres = 0
    array_output = [["Drawing", "Nr Crt", "Nota", "CT", "AEM", "Nota pentru AEM", "Loc", "Nr. PB/Art.", "GROUP",
                     "Numele modulului", "Nume_Norm", "Nume_ADR", "Conector", "Fir", "KN", "XCode1", "Cavity1",
                     "XCode2", "Cavity2", "Patrunderea modulului", "adr module",
                     "ALL_Supersleeve", "Module_ADR", "Module_NonADR", "Module", "Index"]]
    try:
        with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/MatrixDiagrame.txt", newline='', encoding='utf8') as csvfile:
            matrix = list(csv.reader(csvfile, delimiter=';'))
    except FileNotFoundError:
        messagebox.showerror("Missing file", "Fisierul MatrixDiagrame nu exista!")
        pbar.destroy()
        pbargui.destroy()
    statuslabel["text"] = "Loading CSV file . . . "
    pbar['value'] += 2
    pbargui.update_idletasks()
    for i in range(len(matrix)):
        if len(matrix[i]) < 38:
            for x in range(0, 38 - len(matrix[i])):
                matrix[i].append("")
                pbar['value'] += 2
                pbargui.update_idletasks()
    statuslabel["text"] = "Indexing diagrams . . . "
    pbar['value'] += 2
    pbargui.update_idletasks()
    file_counter = len(matrix)
    for i in range(1, len(matrix)):  # len(matrix)
        file_progres = file_progres + 1
        statuslabel["text"] = str(file_progres) + "/" + str(file_counter) + "                       "
        pbar['value'] += 2
        pbargui.update_idletasks()
        for x in range(24, 38):
            modulelist = matrix[i][x].split("/")
            for module in modulelist:
                if (x == 22 and matrix[i][x] != "") or (x == 23 and matrix[i][x] != ""):
                    temp_array = matrix[i][0:24]
                    temp_array.append(module)
                    temp_array.append(1)
                    array_output.append(temp_array)
                    del temp_array
                elif (x == 24 and matrix[i][x] != "") or (x == 32 and matrix[i][x] != ""):
                    temp_array = matrix[i][0:24]
                    temp_array.append(module)
                    temp_array.append(2)
                    array_output.append(temp_array)
                    del temp_array
                elif (x == 25 and matrix[i][x] != "") or (x == 33 and matrix[i][x] != ""):
                    temp_array = matrix[i][0:24]
                    temp_array.append(module)
                    temp_array.append(3)
                    array_output.append(temp_array)
                    del temp_array
                elif (x == 26 and matrix[i][x] != "") or (x == 34 and matrix[i][x] != ""):
                    temp_array = matrix[i][0:24]
                    temp_array.append(module)
                    temp_array.append(4)
                    array_output.append(temp_array)
                    del temp_array
                elif (x == 27 and matrix[i][x] != "") or (x == 35 and matrix[i][x] != ""):
                    temp_array = matrix[i][0:24]
                    temp_array.append(module)
                    temp_array.append(5)
                    array_output.append(temp_array)
                    del temp_array
                elif (x == 28 and matrix[i][x] != "") or (x == 36 and matrix[i][x] != ""):
                    temp_array = matrix[i][0:24]
                    temp_array.append(module)
                    temp_array.append(6)
                    array_output.append(temp_array)
                    del temp_array
                elif (x == 29 and matrix[i][x] != "") or (x == 37 and matrix[i][x] != ""):
                    temp_array = matrix[i][0:24]
                    temp_array.append(module)
                    temp_array.append(7)
                    array_output.append(temp_array)
                    del temp_array
                elif (x == 30 and matrix[i][x] != "") or (x == 38 and matrix[i][x] != ""):
                    temp_array = matrix[i][0:24]
                    temp_array.append(module)
                    temp_array.append(8)
                    array_output.append(temp_array)
                    del temp_array
    statuslabel["text"] = "Printing file . . . "
    pbar['value'] += 2
    pbargui.update_idletasks()
    prn_excel_asocierediagramemodule(array_output)
    pbar.destroy()
    pbargui.destroy()
    end = time.time()
    messagebox.showinfo('Finalizat!', "Prelucrate in " + str(end - start)[:6] + " secunde.")


def crearematrixmodule():
    pbargui = Tk()
    pbargui.title("Matrix module complet")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    array_output = ["Drawing", "Place", "PB", "Group", "Module ID", "Name Normal", "Name ADR", "Conector", "Wire Print",
                    "Wire Number", "XCode1", "Cavity1", "XCode2", "Cavity2", "All Supersleeve", "Module ID ADR",
                    "Module ID NonADR", "Module 1",	"Module 2",	"Module 3",	"Module 4",	"Module 5",	"Module 6",
                    "Module 7", "Module 8", "MY23 Module 1", "MY23 Module 2", "MY23 Module 3", "MY23 Module 4",
                    "MY23 Module 5", "MY23 Module 6", "MY23 Module 7", "MY23 Module 8"]

    place_index = 0
    pb_index = 0
    group_index = 0
    modid_index = 0
    normal_index = 0
    adr_index = 0
    con_index = 0
    wirep_index = 0
    wiren_index = 0
    Xcode1_index = 0
    cavity1_index = 0
    xcode2_index = 0
    cavity2_index = 0
    allss_index = 0
    modidadr_index = 0
    modidnonadr_index = 0
    module1 = 0
    module2 = 0
    module3 = 0
    module4 = 0
    module5 = 0
    module6 = 0
    module7 = 0
    module8 = 0
    my23module1 = 0
    my23module2 = 0
    my23module3 = 0
    my23module4 = 0
    my23module5 = 0
    my23module6 = 0
    my23module7 = 0
    my23module8 = 0
    dir_matrix = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir) + '/MAN',
                                         title="Selectati directorul cu fisiere Matrix Module:")
    file_counter = 0
    file_progres = 0
    for file_all in os.listdir(dir_matrix):
        if file_all.endswith(".xlsm"):
            file_counter = file_counter + 1
    if file_counter == 0:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showinfo("Fisier invalid", "Nu am gasit fisiere de prelucrat!")
        return None
    for file_all in os.listdir(dir_matrix):
        try:
            if file_all.endswith(".xlsm"):
                array_output_temp = []
                wb = load_workbook(dir_matrix + "/" + file_all)
                ws = wb.active
                for row in ws.iter_rows():
                    array_temp = []
                    for cell in row:
                        array_temp.append(cell.value)
                    array_output_temp.append(array_temp)
                for i in range(len(array_output_temp)):
                    if array_output_temp[i][0] == "№":
                        for x, y in enumerate(array_output_temp[i]):
                            if y == 'Місце' or y == 'Place':
                                place_index = x
                            elif y == '№ПБ/Ст.' or y == 'PB':
                                pb_index = x
                            elif y == 'GROUP' or y == 'Group':
                                group_index = x
                            elif y == 'Назва модуля':
                                modid_index = x
                            elif y == 'Name_Norm' or y == 'NAME_norm' or y == 'Name':
                                normal_index = x
                            elif y == 'Name_ADR' or y == 'NAME_AD':
                                adr_index = x
                            elif y == "Роз"єм":
                                con_index = x

















                        print(place_index, pb_index, group_index, adr_index, con_index)
                        break








        except PermissionError:
            messagebox.showerror('Eroare scriere', "Fisierul " + file_all + "este read-only!")
            quit()

    pbargui.destroy()
    end = time.time()
    messagebox.showinfo('Finalizat!', 'Prelucrate  fisiere.')
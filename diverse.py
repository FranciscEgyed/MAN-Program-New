import csv
import os
import datetime
import shutil
import sqlite3
from tkinter import messagebox


def error_file(log):
    with open(os.path.abspath(os.curdir) + "/MAN/Error file.txt", 'a') as myfile:
        myfile.write(datetime.datetime.now().strftime("%d.%m.%Y_%H.%M.%S") + " - " + str(log) + "\n")
    return None


def log_file(log):
    with open(os.path.abspath(os.curdir) + "/MAN/Log file.txt", 'a') as myfile:
        myfile.write(datetime.datetime.now().strftime("%d.%m.%Y_%H.%M.%S") + " - " + str(log) + "\n")
    return None


def structura_directoare():
    os.chdir("..")
    directoareinput = [os.path.abspath(os.curdir) + "/MAN/Input/Module Files/8011",
                       os.path.abspath(os.curdir) + "/MAN/Input/Module Files/8023",
                       os.path.abspath(os.curdir) + "/MAN/Input/Module Files/8000",
                       os.path.abspath(os.curdir) + "/MAN/Input/Module Files/Necunoscut",
                       os.path.abspath(os.curdir) + "/MAN/Input/Wire Lists",
                       os.path.abspath(os.curdir) + "/MAN/Input/BOMs",
                       os.path.abspath(os.curdir) + "/MAN/Input/Others",
                       os.path.abspath(os.curdir) + "/MAN/Input"]
    for d in directoareinput:
        if not os.path.exists(d):
            os.makedirs(d)
    directoareoutput = [os.path.abspath(os.curdir) + "/MAN/Output/Excel Files",
                        os.path.abspath(os.curdir) + "/MAN/Output/Report Files",
                        os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/8000",
                        os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/8011",
                        os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/8023",
                        os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/Beius/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Complete BOM and WIRELIST/BOM/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Complete BOM and WIRELIST/Wirelist/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/Beius/Neprelucrate/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/Beius/Neprelucrate/Light/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/Beius/Neprelucrate/Light+/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/Beius/Prelucrate/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/Beius/Taiere/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Database/KSK Export/",
                        os.path.abspath(os.curdir) + "/MAN/Output/LDorado/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/",
                        os.path.abspath(os.curdir) + "/MAN/Output/QR Images/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Validare/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Validare/Wirelist/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Validare/BOM/",
                        os.path.abspath(os.curdir) + "/MAN/Output/Clustering/"]
    for d in directoareoutput:
        if not os.path.exists(d):
            os.makedirs(d)

    return None


def file_checker():
    arr_others = ["Bracket Side.txt", "Combinatii sectiuni.txt", "Heck Modules.txt", "Langenmodule.txt",
                  "Module Active.txt", "Module Excluse.txt", "Module Implementate.txt", "Module MY2023.txt",
                  "Prufung.txt", "Sortare Module.txt", "Tabel BKK.txt", "Tabel klappschale.txt", "CKD.txt"]
    arr_wires = ["8011.Wirelist.csv", "8012.Wirelist.csv", "8013.Wirelist.csv", "8014.Wirelist.csv",
                 "8023.Wirelist.csv", "8024.Wirelist.csv", "8025.Wirelist.csv", "8026.Wirelist.csv",
                 "8000.Wirelist.csv", "8001.Wirelist.csv"]

    arr_boms = ["8011.BOM.csv", "8012.BOM.csv", "8013.BOM.csv", "8014.BOM.csv", "8023.BOM.csv", "8024.BOM.csv",
                "8025.BOM.csv", "8026.BOM.csv", "8000.BOM.csv", "8001.BOM.csv"]

    arr_wire_9999 = [["Module", "Ltg No", "Leitung", "Farbe", "Quer.", "Kurzname", "Pin", "Kurzname", "Pin", "Lange"],
                     ["9999", "9999", "9999", "9999", "9999", "9999", "9999", "9999", "9999", "9999"],
                     ["9999", "9999", "9999", "9999", "9999", "9999", "9999", "9999", "9999", "9999"]]
    arr_bom_9999 = [["Module", "Quantity", "Bezeichnung", "VOBES-ID", "Benennung", "Verwendung", "Verwendung",
                     "Kurzname", "xy", "Teilenummer", "Vorzugsteil", "TAB-Nummer", "Referenzteil", "Farbe",
                     "E-Komponente", "E-Komponente Part-Nr.", "Einh."],
                    ["85.25480-6000", "80", "Wellrohr NW22 sw", "-", "-", "-", "-", "-", "-", "04.37135-9942",
                     "-", "-", "-", "-", "-", "-", "mm"],
                    ["85.25480-6000", "80", "Wellrohr NW22 sw", "-", "-", "-", "-", "-", "-", "04.37135-9942", "-",
                     "-", "-", "-", "-", "-", "mm"],
                    ["85.25480-6000", "80", "Wellrohr NW22 sw", "-", "-", "-", "-",
                     "-", "-", "04.37135-9942", "-", "-", "-", "-", "-", "-", "mm"],
                    ["85.25480-6000", "80", "Wellrohr NW22 sw", "-", "-", "-", "-", "-", "-", "04.37135-9942", "-", "-",
                     "-", "-", "-", "-", "mm"]]
    arr_temp_bom = []
    arr_temp_wire = []
    arr_temp_others = []
    missing_files = []
    for file_all in os.listdir(os.path.abspath(os.curdir) + "/MAN/Input/BOMs"):
        arr_temp_bom.append(file_all)
    for item in arr_boms:
        if item not in arr_temp_bom:
            missing_files.append(item)
    if "9999.BOM.csv" not in arr_temp_bom:
        with open(os.path.abspath(os.curdir) + "/MAN/Input/BOMs/9999.BOM.csv", 'w', newline='') as myfile:
            wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
            wr.writerows(arr_bom_9999)
    for file_all in os.listdir(os.path.abspath(os.curdir) + "/MAN/Input/Wire Lists"):
        arr_temp_wire.append(file_all)
    if "9999.Wirelist.csv" not in arr_temp_others:
        with open(os.path.abspath(os.curdir) + "/MAN/Input/Wire Lists/9999.Wirelist.csv", 'w', newline='') as myfile:
            wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
            wr.writerows(arr_wire_9999)

    for item in arr_wires:
        if item not in arr_temp_wire:
            missing_files.append(item)
    for file_all in os.listdir(os.path.abspath(os.curdir) + "/MAN/Input/Others"):
        arr_temp_others.append(file_all)
    for item in arr_others:
        if item not in arr_temp_others:
            missing_files.append(item)
    if len(missing_files) > 0:
        messagebox.showwarning("Files missing!", missing_files)
    return None


def skip_file(filename):
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Erori.txt", 'a') as myfile:
        myfile.write(datetime.datetime.now().strftime("%d.%m.%Y_%H.%M.%S") + " - " + str(filename) + "\n")


def pivotare(sheet, item):
    """Load required data files"""
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Combinatii sectiuni.txt", newline='') as csvfile:
        combinatii = list(csv.reader(csvfile, delimiter=';'))
    sec_combi = []
    sec_combi_sort = []
    for i in range(1, len(sheet)):
        if sheet[i][9] == item:
            if sheet[i][5].split(".")[1] == "0":
                sec_combi.append(int(sheet[i][5].split(".")[0]))
            else:
                sec_combi.append(float(sheet[i][5]))
    sec_combi_sort = sorted(sec_combi)
    combi = str(sec_combi_sort[0]) + "x" + str(sec_combi_sort[-1])
    for i in range(len(combinatii)):
        if combinatii[i][0] == combi:
            result = combinatii[i][1]
            break
        else:
            result = "Not in list"
    return result


def istsoll(sheet, sheet2):
    """Verificare tip harness"""
    tipharness = ""
    mixed = ""
    dateheck = []
    rowsinerror = []
    langenmodule = []
    if sheet.cell(row=1, column=17).value == "OK":
        if sheet.cell(row=2, column=17).value > 0:
            tipharness = sheet.cell(row=2, column=16).value
        elif sheet.cell(row=3, column=17).value > 0:
            tipharness = sheet.cell(row=3, column=16).value
        elif sheet.cell(row=4, column=17).value > 0:
            tipharness = sheet.cell(row=4, column=16).value
        elif sheet.cell(row=5, column=17).value > 0:
            tipharness = sheet.cell(row=5, column=16).value
        elif sheet.cell(row=6, column=17).value > 0:
            tipharness = sheet.cell(row=6, column=16).value
    else:
        mixed = " - Mixed Platforms"
    dateheck = [0, 0, 0, 0, 0]
    for row in sheet2['L']:
        if row.value != "Heck module" and row.value is not None:
            dateheck[0] = sheet2.cell(row=row.row, column=7).value
            dateheck[1] = sheet2.cell(row=row.row, column=8).value
            dateheck[2] = sheet2.cell(row=row.row, column=9).value
            dateheck[3] = sheet2.cell(row=row.row, column=10).value
            dateheck[4] = sheet2.cell(row=row.row, column=11).value
            break
    diffsd = 0
    dateheck2 = [0, 0, 0, 0, 0]
    if tipharness == "SATTEL": dateheck2[0] = dateheck[0] - 100
    if tipharness == "CHASSIS": dateheck2[0] = dateheck[0] + 250
    if tipharness == "TGLM": dateheck2[0] = dateheck[0]
    if tipharness == "4AXEL": dateheck2[0] = dateheck[0] + 600
    if tipharness == "Military": dateheck2[0] = dateheck[0] + 600
    dateheck2[1] = dateheck[1]
    dateheck2[2] = dateheck[2]
    dateheck2[3] = dateheck[3]
    dateheck2[4] = dateheck[4]
    for row in sheet2['F']:
        unu = True
        doi = True
        trei = True
        patru = True
        cinci = True
        if row.value == "RIGHT":
            if sheet2.cell(column=7, row=row.row).value != "": unu = dateheck[0] == sheet2.cell(column=7,
                                                                                                row=row.row).value
            if sheet2.cell(column=8, row=row.row).value != "": doi = dateheck[1] == sheet2.cell(column=8,
                                                                                                row=row.row).value
            if sheet2.cell(column=9, row=row.row).value != "": trei = dateheck[2] == sheet2.cell(column=9,
                                                                                                 row=row.row).value
            if sheet2.cell(column=10, row=row.row).value != "": patru = dateheck[3] == sheet2.cell(column=10,
                                                                                                   row=row.row).value
            if sheet2.cell(column=11, row=row.row).value != "": cinci = dateheck[4] == sheet2.cell(column=11,
                                                                                                   row=row.row).value
            if not unu == doi == trei == patru == cinci: rowsinerror.append(sheet2.cell(column=4, row=row.row).value)
        elif row.value == "LEFT":
            if sheet2.cell(column=7, row=row.row).value != "": unu = dateheck2[0] == sheet2.cell(column=7,
                                                                                                 row=row.row).value
            if sheet2.cell(column=8, row=row.row).value != "": doi = dateheck2[1] == sheet2.cell(column=8,
                                                                                                 row=row.row).value
            if sheet2.cell(column=9, row=row.row).value != "": trei = dateheck2[2] == sheet2.cell(column=9,
                                                                                                  row=row.row).value
            if sheet2.cell(column=10, row=row.row).value != "": patru = dateheck2[3] == sheet2.cell(column=10,
                                                                                                    row=row.row).value
            if sheet2.cell(column=11, row=row.row).value != "": cinci = dateheck2[4] == sheet2.cell(column=11,
                                                                                                    row=row.row).value
            if not unu == doi == trei == patru == cinci: rowsinerror.append(sheet2.cell(column=4, row=row.row).value)
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Langenmodule.txt", newline='') as csvfile:
        langenmodule = list(csv.reader(csvfile, delimiter=';'))

    for i in range(len(rowsinerror)):
        kabcounter = ""
        listainlocuitori = []
        inlocuitori = []
        for row in sheet2['D']:
            if sheet2.cell(column=4, row=row.row).value == rowsinerror[i]:
                newrow = row.row
                for x in range(len(langenmodule)):
                    if langenmodule[x][1] == rowsinerror[i]:
                        kabcounter = langenmodule[x][2]
                for x in range(len(langenmodule)):
                    if langenmodule[x][2] == kabcounter:
                        listainlocuitori.append(langenmodule[x])
                for q in range(len(listainlocuitori)):
                    if listainlocuitori[q][15] != "":
                        segmentei = 5
                    elif listainlocuitori[q][14] != "":
                        segmentei = 4
                    elif listainlocuitori[q][13] != "":
                        segmentei = 3
                    elif listainlocuitori[q][12] != "":
                        segmentei = 2
                    elif listainlocuitori[q][11] != "":
                        segmentei = 1
                    if sheet2.cell(column=11, row=row.row).value != "":
                        segmente = 5
                    elif sheet2.cell(column=10, row=row.row).value != "":
                        segmente = 4
                    elif sheet2.cell(column=9, row=row.row).value != "":
                        segmente = 3
                    elif sheet2.cell(column=8, row=row.row).value != "":
                        segmente = 2
                    elif sheet2.cell(column=7, row=row.row).value != "":
                        segmente = 1
                    if sheet2.cell(column=6, row=newrow).value == "RIGHT":
                        if segmente == 5 and segmentei == 5:
                            if dateheck[0] == float(listainlocuitori[q][11]) and dateheck[1] == float(
                                    listainlocuitori[q][12]) and dateheck[2] == float(listainlocuitori[q][13]) and \
                                    dateheck[3] == float(listainlocuitori[q][14]) and dateheck[4] == float(
                                    listainlocuitori[q][15]):
                                inlocuitori.append(
                                    ["SOLL", listainlocuitori[q][0], listainlocuitori[q][10], listainlocuitori[q][1],
                                     listainlocuitori[q][9], "RIGHT", listainlocuitori[q][11], listainlocuitori[q][12],
                                     listainlocuitori[q][13], listainlocuitori[q][14], listainlocuitori[q][15]])
                        elif segmente == 4 and segmentei == 4:
                            if dateheck[0] == float(listainlocuitori[q][11]) and dateheck[1] == float(
                                    listainlocuitori[q][12]) and dateheck[2] == float(listainlocuitori[q][13]) and \
                                    dateheck[3] == float(listainlocuitori[q][14]):
                                inlocuitori.append(
                                    ["SOLL", listainlocuitori[q][0], listainlocuitori[q][10], listainlocuitori[q][1],
                                     listainlocuitori[q][9], "RIGHT", listainlocuitori[q][11], listainlocuitori[q][12],
                                     listainlocuitori[q][13], listainlocuitori[q][14], ""])
                        elif segmente == 3 and segmentei == 3:
                            if dateheck[0] == float(listainlocuitori[q][11]) and dateheck[1] == float(
                                    listainlocuitori[q][12]) and dateheck[2] == float(listainlocuitori[q][13]):
                                inlocuitori.append(
                                    ["SOLL", listainlocuitori[q][0], listainlocuitori[q][10], listainlocuitori[q][1],
                                     listainlocuitori[q][9], "RIGHT", listainlocuitori[q][11], listainlocuitori[q][12],
                                     listainlocuitori[q][13], "", ""])
                        elif segmente == 2 and segmentei == 2:
                            if dateheck[0] == float(listainlocuitori[q][11]) and dateheck[1] == float(
                                    listainlocuitori[q][12]):
                                inlocuitori.append(
                                    ["SOLL", listainlocuitori[q][0], listainlocuitori[q][10], listainlocuitori[q][1],
                                     listainlocuitori[q][9], "RIGHT", listainlocuitori[q][11], listainlocuitori[q][12],
                                     "", "", ""])
                        elif segmente == 1 and segmentei == 1:
                            if dateheck[0] == float(listainlocuitori[q][11]):
                                inlocuitori.append(
                                    ["SOLL", listainlocuitori[q][0], listainlocuitori[q][10], listainlocuitori[q][1],
                                     listainlocuitori[q][9], "RIGHT", listainlocuitori[q][11], "", "", "", ""])
                    if sheet2.cell(column=6, row=newrow).value == "LEFT":
                        if segmente == 5 and segmentei == 5:
                            if dateheck2[0] == float(listainlocuitori[q][11]) and dateheck2[1] == float(
                                    listainlocuitori[q][12]) and dateheck2[2] == float(listainlocuitori[q][13]) and \
                                    dateheck2[3] == float(listainlocuitori[q][14]) and dateheck2[4] == float(
                                    listainlocuitori[q][15]):
                                inlocuitori.append(
                                    ["SOLL", listainlocuitori[q][0], listainlocuitori[q][10], listainlocuitori[q][1],
                                     listainlocuitori[q][9], "LEFT", listainlocuitori[q][11], listainlocuitori[q][12],
                                     listainlocuitori[q][13], listainlocuitori[q][14], listainlocuitori[q][15]])
                        elif segmente == 4 and segmentei == 4:
                            if dateheck2[0] == float(listainlocuitori[q][11]) and dateheck2[1] == float(
                                    listainlocuitori[q][12]) and dateheck2[2] == float(listainlocuitori[q][13]) and \
                                    dateheck2[3] == float(listainlocuitori[q][14]):
                                inlocuitori.append(
                                    ["SOLL", listainlocuitori[q][0], listainlocuitori[q][10], listainlocuitori[q][1],
                                     listainlocuitori[q][9], "LEFT", listainlocuitori[q][11], listainlocuitori[q][12],
                                     listainlocuitori[q][13], listainlocuitori[q][14], ""])
                        elif segmente == 3 and segmentei == 3:
                            if dateheck2[0] == float(listainlocuitori[q][11]) and dateheck2[1] == float(
                                    listainlocuitori[q][12]) and dateheck2[2] == float(listainlocuitori[q][13]):
                                inlocuitori.append(
                                    ["SOLL", listainlocuitori[q][0], listainlocuitori[q][10], listainlocuitori[q][1],
                                     listainlocuitori[q][9], "LEFT", listainlocuitori[q][11], listainlocuitori[q][12],
                                     listainlocuitori[q][13], "", ""])
                        elif segmente == 2 and segmentei == 2:
                            if dateheck2[0] == float(listainlocuitori[q][11]) and dateheck2[1] == float(
                                    listainlocuitori[q][12]):
                                inlocuitori.append(
                                    ["SOLL", listainlocuitori[q][0], listainlocuitori[q][10], listainlocuitori[q][1],
                                     listainlocuitori[q][9], "LEFT", listainlocuitori[q][11], listainlocuitori[q][12],
                                     "", "", ""])
                        elif segmente == 1 and segmentei == 1:
                            if dateheck2[0] == float(listainlocuitori[q][11]):
                                inlocuitori.append(
                                    ["SOLL", listainlocuitori[q][0], listainlocuitori[q][10], listainlocuitori[q][1],
                                     listainlocuitori[q][9], "LEFT", listainlocuitori[q][11], "", "", "", ""])
        for c in range(len(inlocuitori)):
            sheet2.insert_rows(newrow + 1)
            for s in range(len(inlocuitori[c])):
                try:
                    sheet2.cell(column=s + 1, row=newrow + 1, value=float(inlocuitori[c][s]))
                except:
                    sheet2.cell(column=s + 1, row=newrow + 1, value=inlocuitori[c][s])

    for row in sheet2['E']:
        if row.value is None:
            for i in range(len(langenmodule)):
                if sheet2.cell(column=4, row=row.row).value in langenmodule[i]:
                    sheet2.cell(column=5, row=row.row).value = langenmodule[i][9]


def golire_directoare():
    dir_input1 = os.path.abspath(os.curdir) + "/MAN/Input/Module Files/8000/"
    dir_input2 = os.path.abspath(os.curdir) + "/MAN/Input/Module Files/8011/"
    dir_input3 = os.path.abspath(os.curdir) + "/MAN/Input/Module Files/8023/"
    dir_output1 = os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/8000/"
    dir_output2 = os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/8011/"
    dir_output3 = os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/8023/"
    dir_output4 = os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/8000/"
    dir_output5 = os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/8011/"
    dir_output6 = os.path.abspath(os.curdir) + "/MAN/Output/Excel Files/8023/"
    dir_outputr = os.path.abspath(os.curdir) + "/MAN/Output/Report Files/"
    dir_output_separare = os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/"
    dir_output_separare2 = os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/Beius/"
    dir_output_separare3 = os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/Beius/Prelucrate/"
    dir_output_separare4 = os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/Beius/Neprelucrate/Light/"
    dir_output_separare41 = os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/Beius/Neprelucrate/Light+/"
    dir_output_separare5 = os.path.abspath(os.curdir) + "/MAN/Output/Separare KSK/Beius/Taiere/"
    dir_output_BOM_complet = os.path.abspath(os.curdir) + "/MAN/Output/Complete BOM and WIRELIST/BOM/"
    dir_output_wire_complet = os.path.abspath(os.curdir) + "/MAN/Output/Complete BOM and WIRELIST/Wirelist/"
    dir_ldorado = os.path.abspath(os.curdir) + "/MAN/Output/LDorado/"
    dir_output_database1 = os.path.abspath(os.curdir) + "/MAN/Output/Database/KSK Export/"
    dir_output_database2 = os.path.abspath(os.curdir) + "/MAN/Output/Database/"

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
    for file_all in os.listdir(dir_input3):
        try:
            os.remove(dir_input3 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output1):
        try:
            os.remove(dir_output1 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output2):
        try:
            os.remove(dir_output2 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output3):
        try:
            os.remove(dir_output3 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output4):
        try:
            os.remove(dir_output4 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output5):
        try:
            os.remove(dir_output5 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output6):
        try:
            os.remove(dir_output6 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_outputr):
        try:
            os.remove(dir_outputr + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output_separare):
        try:
            os.remove(dir_output_separare + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output_BOM_complet):
        try:
            os.remove(dir_output_BOM_complet + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output_wire_complet):
        try:
            os.remove(dir_output_wire_complet + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output_database1):
        try:
            os.remove(dir_output_database1 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output_database2):
        try:
            os.remove(dir_output_database2 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output_separare2):
        try:
            os.remove(dir_output_separare2 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output_separare3):
        try:
            os.remove(dir_output_separare3 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output_separare4):
        try:
            os.remove(dir_output_separare4 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output_separare41):
        try:
            os.remove(dir_output_separare41 + file_all)
        except:
            continue
    for file_all in os.listdir(dir_output_separare5):
        try:
            os.remove(dir_output_separare5 + file_all)
        except:
            continue
    messagebox.showinfo("Golire", "Directoarele Input si Output au fost golite!!")


def databesemerge():
    try:
        conn = sqlite3.connect("//SVRO8FILE01/Groups/General/EFI/DBMAN/database.db")
        conn2 = sqlite3.connect(os.path.abspath(os.curdir) + "/MAN/Input/Others/database.db")
        cursor2 = conn2.cursor()
        query = "SELECT * FROM `KSKDatabase`"
        cursor2.execute(query)
        final_result = list(cursor2.fetchall())
        cursor = conn.cursor()
        # create a table
        cursor.execute("""CREATE TABLE IF NOT EXISTS KSKDatabase
                          (primarykey text UNIQUE, numejit text, TipHarness text, Light text, DataLivrare text, 
                          DataJIT text, KSKNo text, TrailerNO text, SSType text, Module text) """)
        # insert multiple records using the more secure "?" method
        cursor.executemany("INSERT OR IGNORE INTO KSKDatabase VALUES (?,?,?,?,?,?,?,?,?,?)", final_result)
        conn.commit()
        conn.close()
        cursor2.execute('DELETE FROM KSKDatabase;', )
        conn2.commit()
        conn2.close()

    except sqlite3.OperationalError:
        # messagebox.showerror("Database Error", "Online database unreachable, using local database")
        return None


def databasebackup():
    try:
        shutil.copy2("//SVRO8FILE01/Groups/General/EFI/DBMAN/database.db",
                     "//SVRO8FILE01/Groups/General/EFI/DBMAN/BACKUP/database.db")
    except FileNotFoundError:
        return None

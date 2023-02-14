import csv
import os
from tkinter import filedialog

fisier_calloff = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                            title="Incarcati fisierul care necesita sortare")
#dir_wirelist = filedialog.askdirectory(initialdir=os.path.abspath(os.curdir), title="Selectati directorul cu fisiere:")
#for file_all in os.listdir(dir_wirelist):
    #if file_all.endswith(".csv"):
with open(fisier_calloff, newline='') as csvfile:
    array_incarcat = list(csv.reader(csvfile, delimiter=';'))
array_output = [["Module", "Quantity", "Bezeichnung", "VOBES-ID", "Benennung", "Verwendung", "Verwendung",
                 "Kurzname", "xy", "Teilenummer", "Vorzugsteil", "TAB-Nummer", "Referenzteil", "Farbe",
                 "E-Komponente", "E-Komponente Part-Nr.", "Einh."]]
array_module = []
array_comp = []
for i in range(1, len(array_incarcat)):
    if i == len(array_incarcat) - 1:
        array_comp.append(array_incarcat[i])
        array_out_temp = []
        for vercomp in range(len(array_comp[0])):
            for vermodule in range(len(array_module)):
                if array_module[vermodule][1] in array_comp[0][vercomp]:
                    array_comp[0][vercomp] = array_module[vermodule][2]
        last_position = len(array_comp[0])
        index_beze = array_comp[0].index('Bezeichnung')
        index_VID = array_comp[0].index('VOBES-ID')
        index_bene = array_comp[0].index('Benennung')
        index_verew1 = array_comp[0].index('Verwendung')
        index_verew2 = array_comp[0][index_verew1:].index('Verwendung') + index_verew1
        index_kurz = array_comp[0].index('Kurzname')
        index_xy = array_comp[0].index('xy')
        index_teile = array_comp[0].index('Teilenummer')
        index_tab = array_comp[0].index('TAB-Nummer')
        index_refe = array_comp[0].index('Referenzteil')
        index_farbe = array_comp[0].index('Farbe')
        index_ekomp = array_comp[0].index('E-Komponente')
        index_ekomppn = array_comp[0].index('E-Komponente Part-Nr.')
        index_einh = array_comp[0].index('Einh.')
        for comp in range(1, len(array_comp)):
            for index in range(0, last_position):
                if array_comp[comp][index] != "0":
                    array_out_temp.append([array_comp[0][index], array_comp[comp][index],
                                           array_comp[comp][index_beze], array_comp[comp][index_VID],
                                           array_comp[comp][index_bene], array_comp[comp][index_verew1],
                                           array_comp[comp][index_verew2], array_comp[comp][index_kurz],
                                           array_comp[comp][index_xy], array_comp[comp][index_teile],
                                           array_comp[comp][index_tab], array_comp[comp][index_refe],
                                           array_comp[comp][index_farbe], array_comp[comp][index_ekomp],
                                           array_comp[comp][index_ekomppn], array_comp[comp][index_einh]])
        array_module = []
        array_comp = []
        array_output.extend(array_out_temp)
    else:
        if array_incarcat[i][0] == "2":
            array_module.append(array_incarcat[i])
        elif array_incarcat[i][0] == "3" or array_incarcat[i][0] == "0":
            array_comp.append(array_incarcat[i])
        elif array_incarcat[i][0] == "1":
            for vercomp in range(len(array_comp[0])):
                for vermodule in range(len(array_module)):
                    if array_module[vermodule][1] in array_comp[0][vercomp]:
                        array_comp[0][vercomp] = array_module[vermodule][2]
            array_out_temp = []
            last_position = len(array_comp[0])
            index_beze = array_comp[0].index('Bezeichnung')
            index_VID = array_comp[0].index('VOBES-ID')
            index_bene = array_comp[0].index('Benennung')
            index_verew1 = array_comp[0].index('Verwendung')
            index_verew2 = array_comp[0][index_verew1:].index('Verwendung') + index_verew1
            index_kurz = array_comp[0].index('Kurzname')
            index_xy = array_comp[0].index('xy')
            index_teile = array_comp[0].index('Teilenummer')
            index_tab = array_comp[0].index('TAB-Nummer')
            index_refe = array_comp[0].index('Referenzteil')
            index_farbe = array_comp[0].index('Farbe')
            index_ekomp = array_comp[0].index('E-Komponente')
            index_ekomppn = array_comp[0].index('E-Komponente Part-Nr.')
            index_einh = array_comp[0].index('Einh.')
            for comp in range(1, len(array_comp)):
                for index in range(0, last_position):
                    if array_comp[comp][index] != "0":
                        array_out_temp.append([array_comp[0][index], array_comp[comp][index],
                                               array_comp[comp][index_beze], array_comp[comp][index_VID],
                                               array_comp[comp][index_bene], array_comp[comp][index_verew1],
                                               array_comp[comp][index_verew2], array_comp[comp][index_kurz],
                                               array_comp[comp][index_xy], array_comp[comp][index_teile],
                                               array_comp[comp][index_tab], array_comp[comp][index_refe],
                                               array_comp[comp][index_farbe], array_comp[comp][index_ekomp],
                                               array_comp[comp][index_ekomppn], array_comp[comp][index_einh]])
            array_module = []
            array_comp = []
            array_output.extend(array_out_temp)
for i in range(len(array_output)):
    print(array_output)

import csv
import fnmatch
import itertools
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Entry
from functii_print import prn_excel_clustering


def wirelistprep(array_incarcat):
    array_original = []
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
        messagebox.showerror('Eroare fisier', 'Nu ai incarcat fisierul corect ...')
        return
    if len(array_original) < 50:
        for i in range(0, len(array_original)):
            if array_original[i][0] == "0":
                if str(array_original[i][1]) != "Ltg-Nr.":
                    messagebox.showerror('Eroare fisier',
                                         'Nu ai incarcat fisierul corect')
                    return
    else:
        for i in range(0, 50):
            if array_original[i][0] == "0":
                if str(array_original[i][1]) != "Ltg-Nr.":
                    messagebox.showerror('Eroare fisier',
                                         'Nu ai incarcat fisierul corect')
                    return
    array_output = [
        ["Module", "Ltg No", "Leitung", "Farbe", "Quer.", "Kurzname", "Pin", "Kurzname", "Pin",
         "Sonderltg.", "Lange", "Kontakt1", "Dichtung1", "Typ", "Kontakt2", "Dichtung2"]]
    array_module = []
    array_wires = []
    for i in range(1, len(array_incarcat)):
        if i == len(array_incarcat) - 1:
            array_wires.append(array_incarcat[i])
            array_out_temp = []
            for verwires in range(len(array_wires[0])):
                for vermodule in range(len(array_module)):
                    if array_wires[0][verwires] == array_module[vermodule][1] or \
                            array_wires[0][verwires] == "Length  " + array_module[vermodule][1] or \
                            array_wires[0][verwires] == "LÃ¤nge  " + array_module[vermodule][1] or \
                            fnmatch.fnmatch(array_wires[0][verwires],
                                            "LÃ¤nge  " + array_module[vermodule][1] + " (KSW*"):
                        array_wires[0][verwires] = array_module[vermodule][2]
            pot_position = array_wires[0].index('Pot.') + 1
            ltgno_position = array_wires[0][pot_position:].index('Ltg-Nr.') + pot_position
            index_ltgno = array_wires[0].index('Ltg-Nr.')
            index_leitung = array_wires[0].index('Leitung')
            index_farbe = array_wires[0].index('Farbe')
            index_quer = array_wires[0].index('Quer.')
            index_kurz1 = array_wires[0].index('Kurzname')
            index_pin1 = array_wires[0].index('Pin')
            index_kurz2 = array_wires[0][index_kurz1 + 1:].index('Kurzname') + index_kurz1 + 1
            index_pin2 = array_wires[0][index_pin1 + 1:].index('Pin') + index_pin1 + 1
            index_sonder = array_wires[0].index('Sonderltg.')
            index_contact1 = array_wires[0].index('Kontakt')
            index_contact2 = array_wires[0].index("Kontakt", index_contact1 + 1)
            index_type = array_wires[0].index('Typ')
            index_dichtung1 = array_wires[0].index('Dichtung')
            index_dichtung2 = array_wires[0].index("Dichtung", index_dichtung1 + 1)
            innenleiter = array_wires[0].index('Innenleiter')
            for wire in range(1, len(array_wires)):
                for index in range(pot_position, ltgno_position):
                    if array_wires[wire][index] != "-":
                        array_out_temp.append([array_wires[0][index], array_wires[wire][index_ltgno],
                                               array_wires[wire][index_leitung],
                                               array_wires[wire][index_farbe],
                                               float(array_wires[wire][index_quer].replace(',', '.')),
                                               array_wires[wire][index_kurz1],
                                               array_wires[wire][index_pin1],
                                               array_wires[wire][index_kurz2],
                                               array_wires[wire][index_pin2],
                                               array_wires[wire][index_sonder],
                                               array_wires[wire][index],
                                               array_wires[wire][index_contact1],
                                               array_wires[wire][index_dichtung1],
                                               array_wires[wire][index_type],
                                               array_wires[wire][index_contact2],
                                               array_wires[wire][index_dichtung2],
                                               array_wires[wire][innenleiter]
                                               ])
            array_module = []
            array_wires = []
            array_output.extend(array_out_temp)
        else:
            if array_incarcat[i][0] == "2":
                array_module.append(array_incarcat[i])
            elif array_incarcat[i][0] == "3" or array_incarcat[i][0] == "0":
                array_wires.append(array_incarcat[i])
            elif array_incarcat[i][0] == "1":
                for verwires in range(len(array_wires[0])):
                    for vermodule in range(len(array_module)):
                        if array_wires[0][verwires] == array_module[vermodule][1] or \
                                array_wires[0][verwires] == "Length  " + array_module[vermodule][1] or \
                                array_wires[0][verwires] == "LÃ¤nge  " + array_module[vermodule][1] or \
                                fnmatch.fnmatch(array_wires[0][verwires],
                                                "LÃ¤nge  " + array_module[vermodule][1] + " (KSW*"):
                            array_wires[0][verwires] = array_module[vermodule][2]
                array_out_temp = []
                index_ltgno = array_wires[0].index('Ltg-Nr.')
                index_leitung = array_wires[0].index('Leitung')
                index_farbe = array_wires[0].index('Farbe')
                index_quer = array_wires[0].index('Quer.')
                index_kurz1 = array_wires[0].index('Kurzname')
                index_pin1 = array_wires[0].index('Pin')
                index_kurz2 = array_wires[0][index_kurz1 + 1:].index('Kurzname') + index_kurz1 + 1
                index_pin2 = array_wires[0][index_pin1 + 1:].index('Pin') + index_pin1 + 1
                index_sonder = array_wires[0].index('Sonderltg.')
                pot_position = array_wires[0].index('Pot.') + 1
                ltgno_position = array_wires[0][pot_position:].index('Ltg-Nr.') + pot_position
                index_contact1 = array_wires[0].index('Kontakt')
                index_contact2 = array_wires[0].index("Kontakt", index_contact1 + 1)
                index_type = array_wires[0].index('Typ')
                index_dichtung1 = array_wires[0].index('Dichtung')
                index_dichtung2 = array_wires[0].index("Dichtung", index_dichtung1 + 1)
                innenleiter = array_wires[0].index('Innenleiter')
                for wire in range(1, len(array_wires)):
                    for index in range(pot_position, ltgno_position):
                        if array_wires[wire][index] != "-":
                            array_out_temp.append([array_wires[0][index], array_wires[wire][index_ltgno],
                                                   array_wires[wire][index_leitung],
                                                   array_wires[wire][index_farbe],
                                                   float(array_wires[wire][index_quer].replace(',', '.')),
                                                   array_wires[wire][index_kurz1],
                                                   array_wires[wire][index_pin1],
                                                   array_wires[wire][index_kurz2],
                                                   array_wires[wire][index_pin2],
                                                   array_wires[wire][index_sonder],
                                                   array_wires[wire][index],
                                                   array_wires[wire][index_contact1],
                                                   array_wires[wire][index_dichtung1],
                                                   array_wires[wire][index_type],
                                                   array_wires[wire][index_contact2],
                                                   array_wires[wire][index_dichtung2],
                                                   array_wires[wire][innenleiter]
                                                   ])
                array_module = []
                array_wires = []
                array_output.extend(array_out_temp)
    for line in array_output[1:]:
        if line[13] == "FLK" or line[13] == "Schirm" or line[13] == "FHLR":
            array_output.remove(line)
    array_output = sorted(array_output[1:], key=lambda w: (w[1], float(w[10])))
    return array_output


def wirelistpreph(array_incarcat):
    array_output = [
        ["Design PartNumber", "Wire Name", "INT part number", "Color", "CSA", "From Conn1", "From Pin1", "To Conn2",
         "To Pin2", "Multicore name", "Length", "INT Term1", "Strip length1", "Type", "INT Term2", "Strip length2"]]
    dsgpn = array_incarcat[0].index('Design PartNumber')
    wireno = array_incarcat[0].index('Wire Name')
    cablu = array_incarcat[0].index('INT part number')
    color = array_incarcat[0].index('Color')
    crosssec = array_incarcat[0].index('CSA')
    con1 = array_incarcat[0].index('From Conn1')
    pin1 = array_incarcat[0].index('From Pin1')
    con2 = array_incarcat[0].index('To Conn2')
    pin2 = array_incarcat[0].index('To Pin2')
    multicore = array_incarcat[0].index('Multicore name')
    lenght = array_incarcat[0].index('Length')
    term1 = array_incarcat[0].index('INT Term1')
    strip1 = array_incarcat[0].index('Strip length1')
    typec = array_incarcat[0].index('Type')
    term2 = array_incarcat[0].index('INT Term2')
    strip2 = array_incarcat[0].index('Strip length2')

    for line in array_incarcat[1:]:
        array_output.append(
            [line[dsgpn], line[wireno], line[cablu], line[color], line[crosssec], line[con1], line[pin1],
             line[con2], line[pin2], line[multicore], line[lenght], line[term1], line[strip1],
             line[typec], line[term2], line[strip2]])
    array_output = sorted(array_output[1:], key=lambda w: (w[1], float(w[10])))

    return array_output


def group_data(data, prefix, clusteringlength):
    grouped_data = {}
    key_incrementer = 1
    for item in data:
        ltg_no = item[2]
        length = int(float(item[11]))
        matched = False
        for key, value in grouped_data.items():
            lower_limit = value[1]
            upper_limit = value[2]
            if value[0] == ltg_no and lower_limit <= length <= upper_limit:
                grouped_data[key].append(item)
                matched = True
                break
        if not matched:
            key = prefix + "_" + str(key_incrementer)
            grouped_data[key] = [ltg_no, length, length + clusteringlength]
            grouped_data[key].append(item)
            key_incrementer += 1
    # Create a new dictionary with modified keys
    return grouped_data


def cm(input, nume_fisier):
    output = [["Dwg", "LPN", "What changed", "(CW)", "Change", "Leadset", "REAL NAME", "Kanban name", "von",
               "Pin", "nach", "Pin", "Splice", "Leitung", "Farbe", "Typ", "Kontakt", "Dichtung", "Kontakt2",
               "Dichtung2", "Labeling", "FORS PN", "Crosssec", "Strip_1", "Kontakt PN", "Seal 1 PN", "Strip_2",
               "Kontakt2 PN", "Seal 2 PN", "Length", "OLD Length", "AG", "OpText", "ResGroup", "Sonderltg.",
               "Twist", "End Product", "max_v_puchok", "AGNr", "Alt+F2 1", "Alt+F2 2", "Alt+F2 3", "Alt+F2 4",
               "Alt+F2 5", "Alt+F2 6", "Alt+F2 7", "Alt+F2 8", "Alt+F2 9", "Alt+F2 10", "Suppl.text 1",
               "Suppl.text 2", "APAB_1", "APAB_2"]]
    lista_multicores = ["07.08304-0111", "07.08304-0131", "07.08304-0182", "07.08134-4348", "07.08304-0165"]
    lista_module = {}
    for line in input:
        key = line[0]
        if line[0] not in lista_module:
            lista_module[key] = [line[2]]
        else:
            lista_module[key].append(line[2])
    lista_module_list = sorted(list(set(line[2] for line in input)))
    lista_twisturi = sorted(
        list(set((line[11], line[12]) for line in input if line[11] != "-" and line[4] not in lista_multicores)))
    lista_twisturi_4fire = sorted(
        list(set((line[11], line[12]) for line in input if line[11] != "-" and line[4] in lista_multicores)))

    def merge_tuples(input_list):
        merged_dict = {}
        # Merge tuples based on the first element
        for code, value in input_list:
            if code in merged_dict:
                merged_dict[code].append(value)
            else:
                merged_dict[code] = [value]
        # Create merged tuples
        merged_tuples = [(code, *values) for code, values in merged_dict.items()]
        return merged_tuples

    lista_twisturi = merge_tuples(lista_twisturi)
    output[0].extend(lista_module_list)
    used_wireno = []
    for line in input:
        if line[0] not in used_wireno and line[12] != "0":
            output.append(["", "", "", "", "", line[1], line[3], line[0], line[7], line[8], line[9], line[10],
                           "", line[4], line[5], line[15], line[13], line[14], line[16], line[17], "", "",
                           line[6], "", "", "", "", "", "", line[12], "", "", "", "", line[11],
                           ])
            used_wireno.append(line[0])
    for line in output:
        if len(line) < len(output[0]):
            mulptiplicator = len(output[0]) - len(line)
            line.extend([""] * mulptiplicator)

    for key, values in lista_module.items():
        for line in output[1:]:
            if line[7] == key:
                for module in values:
                    index = output[0].index(module)
                    line[index] = "X"

    keep_list = output[0]
    output = sorted(output[1:], key=lambda x: (x[34], x[29]))
    output.insert(0, keep_list)
    for line in output[1:]:
        cuttingnokont = ["", "", "", "", "", "", line[6], line[7], "", "", "", "", "Cutting",
                         line[13], line[14], line[15], "", "", "", "", "", "", line[22], "", "", "", "", "", "",
                         line[29], "", "", "", "", line[34]]
        cuttingkont1 = ["", "", "", "", "", "", line[6], line[7], "", "", "", "", "Cutting",
                        line[13], line[14], line[15], line[16], line[17], "", "", "", "", line[22], "", "", "", "", "",
                        "",
                        line[29], "", "", "", "", line[34]]
        cuttingkont2 = ["", "", "", "", "", "", line[6], line[7], "", "", "", "", "Cutting",
                        line[13], line[14], line[15], "", "", line[18], line[19], "", "", line[22], "", "", "", "", "",
                        "",
                        line[29], "", "", "", "", line[34]]
        cuttingkont12 = ["", "", "", "", "", "", line[6], line[7], "", "", "", "", "Cutting",
                         line[13], line[14], line[15], line[16], line[17], line[18], line[19], "", "", line[22], "", "",
                         "", "", "", "",
                         line[29], "", "", "", "", line[34]]
        crimpwpa1 = ["", "", "", "", "", "", line[6], line[7], "", "", "", "", "Crimp WPA",
                     line[13], line[14], line[15], line[16], line[17], "", "", "", "", line[22], "", "", "", "", "", "",
                     line[29], "", "", "", "", line[34]]
        crimpwpa2 = ["", "", "", "", "", "", line[6], line[7], "", "", "", "", "Crimp WPA",
                     line[13], line[14], line[15], "", "", line[18], line[19], "", "", line[22], "", "", "", "", "", "",
                     line[29], "", "", "", "", line[34]]
        if float(line[22]) > 4:
            index = output.index(line) + 1
            if line[16] != "-" and line[18] != "-" and line[13] not in lista_multicores:
                output.insert(index, cuttingnokont)
                output.insert(index + 1, crimpwpa1)
                output.insert(index + 2, crimpwpa2)
            elif line[16] != "-" and line[18] == "-":
                output.insert(index, cuttingnokont)
                output.insert(index + 1, crimpwpa1)
            elif line[16] == "-" and line[18] != "-":
                output.insert(index, cuttingnokont)
                output.insert(index + 1, crimpwpa2)
            else:
                output.insert(index, cuttingnokont)
        elif line[13] not in lista_multicores:
            index = output.index(line) + 1
            output.insert(index, cuttingkont12)

    for twist in lista_twisturi:
        for line in output:
            if twist[0] == line[34] and twist[1] == line[29]:
                index = output.index(line)
                indexfir1 = index
                indexfir2 = index + 2
                twistwpa = ["", "", "", "", "", "", output[indexfir1][6] + "/" + output[indexfir2][6],
                            output[indexfir1][7] + "/" + output[indexfir2][7], "", "", "", "", "Twist WPA",
                            line[13], line[14], line[15], "", "", "", "", "", "", line[22], "", "", "", "",
                            "", "", line[29], "", "", "", "", line[34]]
                tapetwistwpa = ["", "", "", "", "", "", output[indexfir1][6] + "/" + output[indexfir2][6],
                            output[indexfir1][7] + "/" + output[indexfir2][7], "", "", "", "", "Tape Twist WPA",
                            line[13], line[14], line[15], "", "", "", "", "", "", line[22], "", "", "", "",
                            "", "", line[29], "", "", "", "", line[34]]
                output.insert(index + 4, twistwpa)
                output.insert(index + 5, tapetwistwpa)
                break
    for twist in lista_twisturi_4fire:
        for line in output:
            if twist[0] == line[34] and twist[1] == line[29]:
                index = output.index(line)
                output[index + 1][7] = line[7]
                output[index + 2][7] = line[7]
                output[index + 3][7] = line[7]

                cuttingnokont = ["", "", "", "", "", "", line[6], line[7], "", "", "", "",
                                 "Cutting MC", line[13], line[14], line[15], "", "", "", "", "", "", line[22],
                                 "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa11 = ["", "", "", "", "", "", line[6], line[7], "", "", "", "",
                              "Crimp WPA MC", line[13], output[index][14], line[15], output[index][16],
                              output[index][17], "", "",
                              "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa12 = ["", "", "", "", "", "", output[index + 1][6], line[7], "", "", "", "",
                              "Crimp WPA MC", line[13], line[14], line[15], output[index + 1][16],
                              output[index + 1][17],
                              "", "", "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa13 = ["", "", "", "", "", "", output[index + 2][6], line[7], "", "", "", "",
                              "Crimp WPA MC", line[13], line[14], line[15], output[index + 2][16],
                              output[index + 2][17],
                              "", "", "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa14 = ["", "", "", "", "", "", output[index + 3][6], line[7], "", "", "", "",
                              "Crimp WPA MC", line[13], line[14], line[15], output[index + 3][16],
                              output[index + 3][17],
                              "", "", "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa21 = ["", "", "", "", "", "", output[index][6], line[7], "", "", "", "",
                              "Crimp WPA MC", line[13], line[14], line[15], "", "", output[index][18],
                              output[index][19],
                              "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa22 = ["", "", "", "", "", "", output[index + 1][6], line[7], "", "", "", "",
                              "Crimp WPA MC", line[13], line[14], line[15], "", "", output[index + 1][18],
                              output[index + 1][19],
                              "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa23 = ["", "", "", "", "", "", output[index + 2][6], line[7], "", "", "", "",
                              "Crimp WPA MC", line[13], line[14], line[15], "", "", output[index + 2][18],
                              output[index + 2][19],
                              "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa24 = ["", "", "", "", "", "", output[index + 3][6], line[7], "", "", "", "",
                              "Crimp WPA MC", line[13], line[14], line[15], "", "", output[index + 3][18],
                              output[index + 3][19],
                              "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                output.insert(index + 4, cuttingnokont)
                output.insert(index + 5, crimpwpa11)
                output.insert(index + 6, crimpwpa12)
                output.insert(index + 7, crimpwpa13)
                output.insert(index + 8, crimpwpa14)
                output.insert(index + 9, crimpwpa21)
                output.insert(index + 10, crimpwpa22)
                output.insert(index + 11, crimpwpa23)
                output.insert(index + 12, crimpwpa24)
                break

    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Component Overview.txt",
              newline='') as csvfile:
        array_componente = list(csv.reader(csvfile, delimiter=';'))
    for line in output:
        for lista in array_componente:
            if line[13] == lista[0]:
                line[21] = lista[1]
            if line[16] == lista[0]:
                line[24] = lista[1]
            if line[17] == lista[0]:
                line[25] = lista[1]
            if line[18] == lista[0]:
                line[27] = lista[1]
            if line[19] == lista[0]:
                line[28] = lista[1]

    extragere_welding(input, nume_fisier)
    extragere_crimp(input, nume_fisier)
    prn_excel_clustering(output, "CM " + nume_fisier)


def extragere_welding(input, nume_fisier):
    # create dictionary
    lista_welding = []
    for line in input:
        if line[7].startswith(('X9', 'X10', 'X11')) and line[9].startswith(('X9', 'X10', 'X11')):
            lista_welding.append([line[7], line[3], line[4], line[6], line[2]])
            lista_welding.append([line[9], line[3], line[4], line[6], line[2]])
        elif line[7].startswith(('X9', 'X10', 'X11')):
            lista_welding.append([line[7], line[3], line[4], line[6], line[2]])
        elif line[9].startswith(('X9', 'X10', 'X11')):
            lista_welding.append([line[9], line[3], line[4], line[6], line[2]])
    lista_puncte_weld = list(set([line[0] for line in lista_welding]))
    output_combinatii = []
    for punct in lista_puncte_weld:
        used_prints = []
        punct_weld = []
        for weld in lista_welding:
            if punct == weld[0] and weld[1] not in used_prints:
                used_prints.append(weld[1])
                punct_weld.append(weld)

        cross_sections = [sublist[3] for sublist in punct_weld]
        # Generate all unique combinations of cross sections with a minimum length of 2
        unique_combinations = set()
        for r in range(2, len(cross_sections) + 1):
            combinations = itertools.combinations(cross_sections, r)
            unique_combinations.update(tuple(sorted(combination)) for combination in combinations)

        # Convert the unique combinations back to tuples for the final list (if needed)
        all_combinations = list(unique_combinations)
        for line in all_combinations:
            temp_list = []
            for crosssec in line:
                temp_list.append(crosssec)
            temp_list.insert(0, punct)
            output_combinatii.append(temp_list)
    for line in output_combinatii:
        for linie in input:
            if line[0] == linie[7] or line[0] == linie[9]:
                line.insert(1, linie[4])
                break
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Component Overview.txt",
              newline='') as csvfile:
        array_componente = list(csv.reader(csvfile, delimiter=';'))
    for line in output_combinatii:
        for linie in array_componente:
            if line[1] == linie[0]:
                line[1] = linie[1]
    prn_excel_clustering(output_combinatii, 'Welding ' + nume_fisier)


def extragere_crimp(input, nume_fisier):
    # create dictionary
    lista_crimp = [["PN Cablu", "Sectiune", "Terminal 1", "Seal 1", "Terminal 2", "Seal 2"]]
    for line in input:
        if 13 >= len(line[13]) > 1 and 13 >= len(line[16]) > 1:
            new_line = [line[4], line[6], line[13], line[14], line[16], line[17]]
            if new_line not in lista_crimp:
                lista_crimp.append([line[4], line[6], line[13], line[14], line[16], line[17]])
        elif 13 >= len(line[13]) <= 1 and 13 >= len(line[16]) > 1:
            new_line = [line[4], line[6], "", "", line[16], line[17]]
            if new_line not in lista_crimp:
                lista_crimp.append([line[4], line[6], "", "", line[16], line[17]])
        elif 13 >= len(line[13]) > 1 and 13 >= len(line[16]) <= 1:
            new_line = [line[4], line[6], line[13], line[14], "", ""]
            if new_line not in lista_crimp:
                lista_crimp.append([line[4], line[6], line[13], line[14], "", ""])
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Component Overview.txt",
              newline='') as csvfile:
        array_componente = list(csv.reader(csvfile, delimiter=';'))
    lista_crimp[0].extend(["PN Cablu Leoni", "Sectiune", "Terminal 1 Leoni", "Seal 1 Leoni", "Terminal 2 Leoni",
                           "Seal 2 Leoni"])
    for line in lista_crimp[1:]:
        line.extend(["", line[1], "", "", "", ""])
    for line in lista_crimp[1:]:
        for linie in array_componente:
            if line[0] == linie[0]:
                line[6] = linie[1]
            if line[2] == linie[0]:
                line[8] = linie[1]
            if line[3] == linie[0]:
                line[9] = linie[1]
            if line[4] == linie[0]:
                line[10] = linie[1]
            if line[5] == linie[0]:
                line[11] = linie[1]
    prn_excel_clustering(lista_crimp, 'Crimping ' + nume_fisier)


def cpah(input, nume_fisier):
    pass


def clustering():
    fisier_wirelist = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                                 title="Incarcati fisierul care necesita prelucrare")

    root = tk.Tk()
    root.title("2022 MAN file processor")
    # Set the geometry of Tkinter frame
    root.geometry("750x250")

    def display_text():
        string = entrywno.get()
        string2 = int(entrycl.get())
        # pentru LDorado
        try:
            with open(fisier_wirelist, 'r') as file:
                first_line = file.readline()
        except:
            messagebox.showerror('Fisier gresit', 'Nu ai incarcat fisier wirelist!')
            root.destroy()

        if 'Antrieb' in first_line and '1'in first_line:
            print("LDorado")
            with open(fisier_wirelist, newline='') as csvfile:
                table = list(csv.reader(csvfile, delimiter=';'))
            nume_fisier = os.path.splitext(os.path.basename(fisier_wirelist))[0]
            tablenew = wirelistprep(table)
            nume_fisier = os.path.splitext(os.path.basename(fisier_wirelist))[0]
            tablenew = wirelistprep(table)
            for i in range(len(tablenew)):
                tablenew[i].insert(0, tablenew[i][0] + tablenew[i][1])
            grouped_data = group_data(tablenew, string, string2)
            output = []
            for key, values in grouped_data.items():
                # output.append([key])
                for item in values:
                    temp = [key]
                    if type(item) is list:
                        temp.extend(item)
                        if len(temp) > 1:
                            output.append(temp)
                # output.append([])
            prn_excel_clustering(output, nume_fisier)
            cm(output, nume_fisier)
            root.destroy()
            messagebox.showinfo("Finalizat", "Finalizat clustering!")
        # pentru CapH
        elif "Wire Name" in first_line and "CSA" in first_line:
            print("CAPH")
            with open(fisier_wirelist, newline='') as csvfile:
                table = list(csv.reader(csvfile, delimiter=','))
            nume_fisier = os.path.splitext(os.path.basename(fisier_wirelist))[0]
            tablenew = wirelistpreph(table)
            for i in range(len(tablenew)):
                tablenew[i].insert(0, tablenew[i][0] + tablenew[i][1])
            print(tablenew)
            grouped_data = group_data(tablenew, string, string2)
            output = []
            for key, values in grouped_data.items():
                # output.append([key])
                for item in values:
                    temp = [key]
                    if type(item) is list:
                        temp.extend(item)
                        if len(temp) > 1:
                            output.append(temp)
                # output.append([])
            prn_excel_clustering(output, nume_fisier)
            cpah(output, nume_fisier)



            messagebox.showinfo("Finalizat", "Finalizat clustering!")





    # Initialize a Label to display the User Input
    label = tk.Label(root, text="Introduceti prefixul dorit pentru fire:", font="Courier 11 bold")

    label2 = tk.Label(root, text="Introduceti lungimea maxima pentru clustering:", font="Courier 11 bold")
    # Create an Entry widget to accept User Input

    entrywno = Entry(root, width=40)
    entrywno.insert(0, 'WIRE')
    entrywno.focus_set()
    entrycl = Entry(root, width=40)
    entrycl.insert(0, '200')
    label.pack()
    entrywno.pack()
    label2.pack()
    entrycl.pack()

    # Create a Button to validate Entry Widget
    tk.Button(root, text="Okay", width=20, command=display_text).pack(pady=20)

    root.mainloop()

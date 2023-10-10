import csv
import fnmatch
import itertools
import os
import tkinter as tk
from tkinter import filedialog, messagebox, Entry
from functii_print import prn_excel_clustering


def wirelistprep(array_incarcat):
    def find_second_instance(input_list, target_string):
        count = 0
        for item in input_list:
            if item == target_string:
                count += 1
                if count == 2:
                    return input_list.index(item)
        return None

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
            index_contact2 = find_second_instance(array_wires[0], 'Kontakt')
            index_type = array_wires[0].index('Typ')
            index_dichtung1 = array_wires[0].index('Dichtung')
            index_dichtung2 = find_second_instance(array_wires[0], 'Dichtung')
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
                                               array_wires[wire][index_contact2],
                                               array_wires[wire][index_type],
                                               array_wires[wire][index_dichtung1],
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
                index_contact2 = find_second_instance(array_wires[0], 'Kontakt')
                index_type = array_wires[0].index('Typ')
                index_dichtung1 = array_wires[0].index('Dichtung')
                index_dichtung2 = find_second_instance(array_wires[0], 'Dichtung')
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
                                                   array_wires[wire][index_contact2],
                                                   array_wires[wire][index_type],
                                                   array_wires[wire][index_dichtung1],
                                                   array_wires[wire][index_dichtung2],
                                                   array_wires[wire][innenleiter]
                                                   ])
                array_module = []
                array_wires = []
                array_output.extend(array_out_temp)
    array_output = sorted(array_output[1:], key=lambda w: (w[1], float(w[10])))
    for i in range(0, 10):
        print(array_output[i])
    return array_output


def group_data(data, prefix, clusteringlength):
    grouped_data = {}
    key_incrementer = 1
    for item in data:
        ltg_no = item[2]
        length = int(item[11])
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

    filtered_data = {}
    for code, value in lista_twisturi:
        value = int(value)
        if code in filtered_data:
            # Check if the absolute difference is greater than or equal to 100
            if abs(value - filtered_data[code]) >= 100:
                filtered_data[code] = value
        else:
            filtered_data[code] = value

    lista_twisturi = [(code, str(value)) for code, value in filtered_data.items()]

    output[0].extend(lista_module_list)
    used_wireno = []
    for line in input:
        if line[0] not in used_wireno:
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
            if line[17] != "-" and line[19] != "-" and line[13] not in lista_multicores:
                output.insert(index, cuttingnokont)
                output.insert(index + 1, crimpwpa1)
                output.insert(index + 2, crimpwpa2)
            elif line[17] != "-" and line[19] == "-":
                output.insert(index, cuttingnokont)
                output.insert(index + 1, crimpwpa1)
            elif line[17] == "-" and line[19] != "-":
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
                output.insert(index + 4, twistwpa)
                break
    for twist in lista_twisturi_4fire:
        for line in output:
            if twist[0] == line[34] and twist[1] == line[29]:
                index = output.index(line)
                cuttingnokont = ["", "", "", "", "", "", line[6], line[7], "", "", "", "",
                                 "Cutting", line[13], line[14], line[15], "", "", "", "", "", "", line[22],
                                 "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa11 = ["", "", "", "", "", "", line[6], line[7], "", "", "", "",
                              "Crimp WPA", line[13], output[index][14], line[15], output[index][16], output[index][17], "", "",
                              "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa12 = ["", "", "", "", "", "", output[index + 1][6], line[7], "", "", "", "",
                              "Crimp WPA", line[13], line[14], line[15], output[index + 1][16], output[index + 1][17],
                              "", "", "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa13 = ["", "", "", "", "", "", output[index + 2][6], line[7], "", "", "", "",
                              "Crimp WPA", line[13], line[14], line[15], output[index + 2][16], output[index + 2][17],
                              "", "", "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa14 = ["", "", "", "", "", "", output[index + 3][6], line[7], "", "", "", "",
                              "Crimp WPA", line[13], line[14], line[15], output[index + 3][16], output[index + 3][17],
                              "", "", "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa21 = ["", "", "", "", "", "", output[index][6], line[7], "", "", "", "",
                              "Crimp WPA", line[13], line[14], line[15], "", "", output[index][18], output[index][19],
                              "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa22 = ["", "", "", "", "", "", output[index + 1][6], line[7], "", "", "", "",
                              "Crimp WPA", line[13], line[14], line[15], "", "", output[index + 1][18], output[index + 1][19],
                              "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa23 = ["", "", "", "", "", "", output[index + 2][6], line[7], "", "", "", "",
                              "Crimp WPA", line[13], line[14], line[15], "", "", output[index + 2][18], output[index + 2][19],
                              "", "", line[22], "", "", "", "", "", "", line[29], "", "", "", "", line[34]]
                crimpwpa24 = ["", "", "", "", "", "", output[index + 3][6], line[7], "", "", "", "",
                              "Crimp WPA", line[13], line[14], line[15], "", "", output[index + 3][18], output[index + 3][19],
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
    extragere_welding(input, nume_fisier)
    prn_excel_clustering(output, "CM " + nume_fisier)

def extragere_welding(input, nume_fisier):
    # create dictionary
    print(input[0])
    lista_welding = []
    for line in input:
        if line[7].startswith(('X9', 'X10', 'X11')):
            lista_welding.append([line[7], line[3], line[4], line[6], line[2]])
        elif line[9].startswith(('X9', 'X10', 'X11')):
            lista_welding.append([line[9], line[3], line[4], line[6], line[2]])
    lista_puncte_weld = list(set([line[0] for line in lista_welding]))
    for punct in lista_puncte_weld:
        punct_weld = []
        for weld in lista_welding:
            if punct == weld[0]:
                punct_weld.append(weld)

    values_at_index_1 = list(set([sublist[1] for sublist in punct_weld]))
    # Generate combinations of different lengths from 2 to the length of the list
    all_combinations = []
    for r in range(2, len(values_at_index_1) + 1):
        combinations = list(itertools.combinations(values_at_index_1, r))
        all_combinations.extend(combinations)
    output_combinatii = []
    for line in all_combinations:
        temp_line = []
        for element in line:
            for wire in punct_weld:
                if element == wire[1]:
                    temp_line.append([wire[1], wire[3], wire[2]])
        if temp_line not in output_combinatii:
            output_combinatii.append(temp_line)

    prn_excel_clustering(output_combinatii, 'Welding ' + nume_fisier)




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
        try:
            with open(fisier_wirelist, newline='') as csvfile:
                table = list(csv.reader(csvfile, delimiter=';'))
        except:
            messagebox.showerror('Fisier gresit', 'Nu ai incarcat fisier wirelist!')
            root.destroy()
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

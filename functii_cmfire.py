import csv
import fnmatch
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



    array_wirelisturi = ["8000", "8001", "8011", "8012", "8013", "8014", "8023", "8024", "8025", "8026",
                         "8030", "8031", "8032", "8052", "8053", "8041", "8042", "8010", "8034", "8035"]
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
         "Sonderltg.",
         "Lange"]]
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
                                               array_wires[wire][index_dichtung2]
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
                                                   array_wires[wire][index_dichtung2]
                                                   ])
                array_module = []
                array_wires = []
                array_output.extend(array_out_temp)
    array_output = sorted(array_output[1:], key=lambda w: (w[1], float(w[10])))
    return array_output


def group_data(data, ltg_no, clusteringlength):
    grouped_data = {}
    key_incrementer = 1
    for item in data:
        length = int(item[11])
        key = ltg_no + "_" + str(key_incrementer)
        if key not in grouped_data:
            grouped_data[key] = [length, length + clusteringlength]
            grouped_data[key].append(item)
        else:
            min_length = grouped_data[key][0]
            max_length = grouped_data[key][1]
            if min_length <= length <= max_length:
                grouped_data[key].append(item)
            else:
                key_incrementer += 1
                key = ltg_no + "_" + str(key_incrementer)
                grouped_data[key] = [length, length + clusteringlength]
                grouped_data[key].append(item)
    return grouped_data



def cm(input, nume_fisier):
    output = [["Dwg", "LPN", "What changed", "(CW)", "Change", "Leadset", "REAL NAME", "Kanban name", "von",
              "Pin", "nach", "Pin", "Splice", "Leitung", "Farbe", "Typ", "Kontakt", "Dichtung", "Kontakt2",
              "Dichtung2", "Labeling", "FORS PN", "Crosssec", "Strip_1", "Kontakt PN", "Seal 1 PN", "Strip_2",
              "Kontakt2 PN", "Seal 2 PN", "Length", "OLD Length", "AG", "OpText", "ResGroup", "Sonderltg.",
              "Twist", "End Product", "max_v_puchok", "AGNr", "Alt+F2 1", "Alt+F2 2", "Alt+F2 3", "Alt+F2 4",
              "Alt+F2 5", "Alt+F2 6", "Alt+F2 7", "Alt+F2 8", "Alt+F2 9", "Alt+F2 10", "Suppl.text 1",
              "Suppl.text 2", "APAB_1", "APAB_2"]]
    lista_module = sorted(list(set([line[2] for line in input])))

    output[0].extend(lista_module)
    for line in input:
        output.append(["X", "X", "X", "X", "X", "X", line[3], line[0], line[7], line[8], line[9], line[10],
                       line[11], line[4], line[5], line[15], line[13], line[14], line[16], line[17], "X", "X",
                       line[6]
                       ])
    for line in output:
        if len(line) < len(output[0]):
            mulptiplicator = len(output[0]) - len(line)
            line.extend([""] * mulptiplicator)

    for line in output[1:]:
        for linie in input:
            if line[7] == linie[0]:
                module = linie[2]
                index_x = output[0].index(module)
                line[index_x] = "X"
                break
    for i in range(0, 5):
        print(output[i])
        print(len(output[i]))
    prn_excel_clustering(output, "CM cu " + nume_fisier)






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

        with open(fisier_wirelist, newline='') as csvfile:
            table = list(csv.reader(csvfile, delimiter=';'))
        nume_fisier = os.path.splitext(os.path.basename(fisier_wirelist))[0]
        tablenew = wirelistprep(table)
        for i in range(len(tablenew)):
            tablenew[i].insert(0, tablenew[i][0] + tablenew[i][1])
        grouped_data = group_data(tablenew, string, string2)
        # Print each line separately
        output = []
        for key, values in grouped_data.items():
            # output.append([key])
            for item in values:
                temp = [key]
                if type(item) is list:
                    temp.extend(item)
                    if len(temp) > 1:
                        output.append(temp)
            #output.append([])
        prn_excel_clustering(output, nume_fisier)
        for i in range(0, 5):
            print(output[i])
        cm(output, nume_fisier)







        root.destroy()
        messagebox.showinfo("Finalizat", "Finalizat clustering!")

    # Initialize a Label to display the User Input
    label = tk.Label(root, text="Introduceti prefixul dorit pentru fire:", font="Courier 11 bold")

    label2 = tk.Label(root, text="Introduceti lungimea maxima pentru clustering:", font="Courier 11 bold")
    # Create an Entry widget to accept User Input

    entrywno = Entry(root, width=40)
    entrywno.focus_set()
    entrycl = Entry(root, width=40)
    label.pack()
    entrywno.pack()
    label2.pack()
    entrycl.pack()

    # Create a Button to validate Entry Widget
    tk.Button(root, text="Okay", width=20, command=display_text).pack(pady=20)

    root.mainloop()

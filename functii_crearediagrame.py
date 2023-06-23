import csv
import fnmatch
import os
from collections import defaultdict
from itertools import combinations, product, permutations
from tkinter import messagebox, filedialog, Tk, ttk, HORIZONTAL

from functii_print import prn_excel_diagramenew


def wirelistprepd(array_incarcat):
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
                                               array_wires[wire][index]])
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
                for wire in range(1, len(array_wires)):
                    for index in range(pot_position, ltgno_position):
                        if array_wires[wire][index] != "-":
                            array_out_temp.append(
                                [array_wires[0][index], array_wires[wire][index_ltgno],
                                 array_wires[wire][index_leitung], array_wires[wire][index_farbe],
                                 float(array_wires[wire][index_quer].replace(',', '.')),
                                 array_wires[wire][index_kurz1], array_wires[wire][index_pin1],
                                 array_wires[wire][index_kurz2], array_wires[wire][index_pin2],
                                 array_wires[wire][index_sonder],
                                 array_wires[wire][index]])
                array_module = []
                array_wires = []
                array_output.extend(array_out_temp)
    return array_output


def count_occurrences(data):
    kurzname_counts = {}
    for row in data:
        kurzname = row[5]  # Assuming 'Kurzname' is always at index 5
        if kurzname in kurzname_counts:
            kurzname_counts[kurzname] += 1
        else:
            kurzname_counts[kurzname] = 1
    return kurzname_counts


def create_groups(data):
    kurzname_counts = count_occurrences(data)
    kurzname_groups = defaultdict(dict)
    used_ltg_module = set()

    for row in data[1:]:  # Skip the header row
        module = row[0]  # Assuming 'Module' is always at index 0
        ltg_no = row[1]  # Assuming 'Ltg No' is always at index 1
        kurzname = row[5]  # Assuming 'Kurzname' is always at index 5
        pin_no = row[6]
        if ltg_no in used_ltg_module:
            continue

        used_ltg_module.add(ltg_no)

        if kurzname not in kurzname_groups:
            kurzname_groups[kurzname] = {}

        if module not in kurzname_groups[kurzname]:
            kurzname_groups[kurzname][module] = []

        kurzname_groups[kurzname][module].append({'Ltg No': ltg_no, 'Pin No': pin_no})

    # Sort the groups by occurrence count of Kurzname in ascending order
    sorted_groups = {k: kurzname_groups[k] for k in sorted(kurzname_groups, key=lambda x: kurzname_counts.get(x, 0))}

    return sorted_groups



def print_combinations(dictionary):
    pbargui = Tk()
    pbargui.title("Creare combinatii")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    pbar.grid(row=1, column=1, padx=5, pady=5)
    output = {}
    combined_groups = {}
    for kurzname, kurzname_group in dictionary.items():
        module_groups = kurzname_group.keys()
        module_combinations = []
        for i in range(2, len(module_groups)):
            module_combinations.extend(list(permutations(module_groups, i)))

        with open("filename", 'w') as file:
            for tpl in module_combinations:
                file.write(str(tpl) + '\n')
        for combination in module_combinations:
            pbar['value'] += 2
            pbargui.update_idletasks()
            combined_key = kurzname + "===" + ''.join(str(x) for x in combination)
            outlst = []
            duplicate_pin_no = False  # Flag for duplicate 'Pin No' values
            pin_no_values = set()
            for module in combination:
                for key, inner_dict in dictionary.items():
                    for inner_key, lst in inner_dict.items():
                        pbar['value'] += 2
                        pbargui.update_idletasks()

                        if module == inner_key:
                            for item in lst:
                                pin_no = item.get('Pin No')
                                if pin_no in pin_no_values:
                                    duplicate_pin_no = True
                                    break  # Break if duplicate 'Pin No' found
                                pin_no_values.add(pin_no)
                                for sub_key, value in item.items():
                                    outlst.append([sub_key, value])
                if duplicate_pin_no:
                    break  # Break if duplicate 'Pin No' found
            if not duplicate_pin_no:
                combined_groups[combined_key] = outlst
    output.update(combined_groups)
    pbar.destroy()
    pbargui.destroy()

    return output


def crearediagrame():
    fisier_wirelist = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                                title="Incarcati fisierul care necesita prelucrare")
    with open(fisier_wirelist, newline='') as csvfile:
        table = list(csv.reader(csvfile, delimiter=';'))
    nume_fisier = os.path.splitext(os.path.basename(fisier_wirelist))[0]
    tablenew = wirelistprepd(table)
    groups = create_groups(tablenew)
    output1 = []
    for key, inner_dict in groups.items():
        #print(f"Kurzname: {key}")
        output1.append(["Kurzname", key])
        for inner_key, lst in inner_dict.items():
            #print(f"Module: {inner_key}")
            output1.append(["Module", inner_key])
            for item in lst:
                #print(f"Ltg No: {item['Ltg No']}, {item['Pin No']}")
                output1.append(["Ltg No", item['Ltg No'], item['Pin No']])
        #print()
        output1.append([])
    #output2 = print_combinations(groups)
    output2 = []
    for key, lst in print_combinations(groups).items():
        output2.append(["Diagrama", key])
        for q in range(0, len(lst), 2):
            output2.append([lst[q][0] + " " + lst[q][1], lst[q + 1][0] + " " + lst[q + 1][1]])
        output2.append([])

    prn_excel_diagramenew(output1, output2, "Test")
    messagebox.showinfo("Finalizat", "Finalizat diagrame!")

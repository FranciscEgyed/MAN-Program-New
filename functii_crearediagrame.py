import csv
import fnmatch
import math
import os
from collections import defaultdict
from itertools import combinations, product, permutations
from tkinter import messagebox, filedialog, Tk, ttk, HORIZONTAL
from functii_print import prn_excel_diagramenew


def wirelistpreparation(array_incarcat):
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


def find_max_pin_by_kurzname(data):
    kurzname_max_pin = {}
    for item in data[1:]:
        kurzname = item[5]
        try:
            pin = int(item[6])
        except ValueError:
            pin = int(item[6][:-1])
        if kurzname.startswith(("X9", "X10", "X11", "SP")):
            continue
        if kurzname in kurzname_max_pin:
            if pin > kurzname_max_pin[kurzname]:
                kurzname_max_pin[kurzname] = math.ceil(pin / 2) * 2
        else:
            kurzname_max_pin[kurzname] = math.ceil(pin / 2) * 2
        kurzname2 = item[7]
        try:
            pin2 = int(item[8])
        except ValueError:
            pin2 = int(item[8][:-1])
        if kurzname2.startswith(("X9", "X10", "X11", "SP")):
            continue
        if kurzname2 in kurzname_max_pin:
            if pin2 > kurzname_max_pin[kurzname2]:
                kurzname_max_pin[kurzname2] = math.ceil(pin2 / 2) * 2
        else:
            kurzname_max_pin[kurzname2] = math.ceil(pin2 / 2) * 2
    kurzname_max_pin = dict(sorted(kurzname_max_pin.items(), key=lambda x: x[1]))
    return kurzname_max_pin


def create_nested_dictionary(data):
    result = {}
    for item in data[1:]:
        module = item[0]
        kurzname = item[5]
        ltg_no = item[1]
        pin_no = item[6]
        if kurzname not in result:
            result[kurzname] = {}

        if module not in result[kurzname]:
            result[kurzname][module] = []
        result[kurzname][module].append({'Ltg No': ltg_no, 'Pin No': pin_no})

    return result


def print_combinations(dictionary):
    output = {}
    combined_groups = {}
    for kurzname, kurzname_group in dictionary.items():
        module_groups = kurzname_group.keys()
        module_combinations = []
        for i in range(1, len(module_groups)+1):
            module_combinations.extend(list(combinations(module_groups, i)))

        with open(os.path.abspath(os.curdir) + "/filename", 'w') as file:
            for tpl in module_combinations:
                file.write(str(tpl) + '\n')
        for combination in module_combinations:
            combined_key = kurzname + "===" + ''.join(str(x) for x in combination)
            outlst = []
            duplicate_pin_no = False  # Flag for duplicate 'Pin No' values
            pin_no_values = set()
            for module in combination:
                for key, inner_dict in dictionary.items():
                    for inner_key, lst in inner_dict.items():
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

    return output

def get_ltg_pin_pairs(data):
    ltg_pin_pairs = []
    for key, values in data.items():
        pairs = []
        for i in range(0, len(values), 2):
            ltg_no = values[i][1]
            pin_no = values[i+1][1]
            pairs.append((ltg_no, pin_no))
        ltg_pin_pairs.append((key, pairs))
    return ltg_pin_pairs
def print_components_dict(components_dict):
    ltg_pin_pairs = get_ltg_pin_pairs(components_dict)
    for key, pairs in ltg_pin_pairs:
        print(key)
        for ltg_no, pin_no in pairs:
            print(f'Ltg No: {ltg_no}, Pin No: {pin_no}')
        print()


def crearediagrame():
    pbargui = Tk()
    pbargui.title("Creare combinatii")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    pbar.grid(row=1, column=1, padx=5, pady=5)
    fisier_wirelist = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                                title="Incarcati fisierul care necesita prelucrare")
    with open(fisier_wirelist, newline='') as csvfile:
        table = list(csv.reader(csvfile, delimiter=';'))
    nume_fisier = os.path.splitext(os.path.basename(fisier_wirelist))[0]

    data = wirelistpreparation(table)
    nesteddic = create_nested_dictionary(data)
    find_max_pin_by_kurzname(data)
    permutations_dict = print_combinations(nesteddic)
    print_components_dict(permutations_dict)

    pbar.destroy()
    pbargui.destroy()
    #prn_excel_diagramenew(output1, output2, "Test")
    messagebox.showinfo("Finalizat", "Finalizat diagrame!")

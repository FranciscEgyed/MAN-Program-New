import csv
import fnmatch
import os
from tkinter import filedialog, messagebox
from functii_print import prn_excel_clustering


def wirelistprep(array_incarcat):
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
    array_output = sorted(array_output[1:], key=lambda w: float(w[10]))
    return array_output


def group_data_by_ltg_no(data, range_difference):
    grouped_data = {}
    # Iterate through each row of data
    for row in data[1:]:
        ltg_no = row[2]  # Get the "Ltg No" value from the row
        lange = int(row[-1])  # Get the "Lange" value from the row and convert it to an integer
        # Check if the "Ltg No" already exists in the grouped data dictionary
        if ltg_no in grouped_data:
            max_lange = min(grouped_data[ltg_no], key=lambda x: x[-1][-1])[-1][-1]
            # Check if the current "Lange" falls within the range difference
            if abs(lange - int(max_lange)) <= range_difference and lange != int(max_lange):
                # Append the current row to the existing group
                grouped_data[ltg_no][-1].append(row)
            else:
                # Create a new group for the current "Ltg No" and "Lange"
                grouped_data[ltg_no].append([row])
        else:
            # Create a new group for the current "Ltg No" and "Lange"
            grouped_data[ltg_no] = [[row]]

    return grouped_data


def clustering():
    fisier_wirelist = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                                title="Incarcati fisierul care necesita prelucrare")
    with open(fisier_wirelist, newline='') as csvfile:
        table = list(csv.reader(csvfile, delimiter=';'))
    nume_fisier = os.path.splitext(os.path.basename(fisier_wirelist))[0]
    tablenew = wirelistprep(table)
    for i in range(len(tablenew)):
        tablenew[i].insert(0, tablenew[i][0]+tablenew[i][1] )
    grouped_data = group_data_by_ltg_no(tablenew, 300)
    # Print each line separately
    output = []
    for group in grouped_data.values():
        for lines in group:
            if len(lines) > 1:
                for line in lines:
                    output.append(line)
                output.append([])
    prn_excel_clustering(output, nume_fisier)
    messagebox.showinfo("Finalizat", "Finalizat clustering!")



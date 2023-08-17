import csv
import math
import os
import tkinter as tk
from tkinter import messagebox, filedialog, Tk, ttk, HORIZONTAL, Label
import json
import xml.etree.ElementTree as ET

from openpyxl.reader.excel import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.workbook import Workbook
from itertools import combinations
from collections import Counter


def crearediagrame():
    fisier_xml = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                                title="Incarcati fisierul care necesita prelucrare")

    nume_fisier = os.path.splitext(os.path.basename(fisier_xml))[0]
    print(nume_fisier)
    parse_xml(fisier_xml)

    messagebox.showinfo("Finalizat", "Finalizat diagrame!")


def parse_xml(file):
    def process_element(element):
        # Create a dictionary for the current tag
        tag_dict = {}

        # Store the attributes of the element in the dictionary
        tag_dict.update(element.attrib)

        # Check if the element has text content
        if element.text and element.text.strip():
            # Store the text content under the "__text__" key
            tag_dict["__text__"] = element.text.strip()

        # Process the child elements recursively
        for child in element:
            child_dict = process_element(child)
            tag_dict.setdefault(child.tag, []).append(child_dict)

        return tag_dict

    # Root elements to process
    root_element_names = ["Modules", "LengthVariants", "Wires", "CavitySeals", "Connectors",
                          "Tapes", "Terminals", "CavityPlugs", "Accessories", "FusePMDs", "RelayPMDs", "AccessoryPMDs",
                          "AssemblyPartPMDs", "HarnessConfigurationPMDs", "WiringGroupPMDs", "ConnectorPMDs",
                          "FixingPMDs", "GeneralSpecialWirePMDs", "GeneralWirePMDs", "PartPMDs", "SealPMDs",
                          "PlugPMDs", "TerminalPMDs", "TapePMDs", "SymbolPMDs","ComponentBoxPMDs", "Configurations"]

    # Parse the XML file
    tree = ET.parse(file)
    for element in root_element_names:
        # Get the root element based on the given root_element_name
        root = tree.find(".//{}".format(element))

        if root is not None:
            # Process the root element
            root_dict = process_element(root)
            # Save the dictionary to a JSON file
            json_file_path = element + ".json"
            with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/JSON/"+ element + ".json", "w") as json_file:
                json.dump(root_dict, json_file)
        else:
            print(f"Root element '{element}' not found in the XML file.")


def prelucrare_json():
    files = ["Modules", "LengthVariants", "Wires", "CavitySeals", "Connectors",
             "Tapes", "Terminals", "CavityPlugs", "Accessories", "Configurations", "AccessoryPMDs"]

    def extract_wire_ids(wire_list):
        with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/JSON/GeneralSpecialWirePMDs.json", "r") as json_file:
            gsw = json.load(json_file)
        with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/JSON/GeneralWirePMDs.json", "r") as json_file:
            gw = json.load(json_file)
        wires = []
        for wire in wire_list['Wire']:
            cablu = "Empty"
            sectiune = "Empty"
            culoare = "Empty"
            pmd = wire['PMD']
            for fir in gsw['GeneralSpecialWirePMD']:
                for core in fir['CorePMD']:
                    if core["CustomerPartNo"] == pmd:
                        cablu = core["CustomerPartNo"]
                        sectiune = core["CSA"]
                        try:
                            culoare = core["Colour"].split("_")[1]
                        except IndexError:
                            culoare = core["Colour"]
                    else:
                        for firut in gw['GeneralWirePMD']:
                            if firut["ID"] == pmd:
                                cablu = firut["ID"]
                                sectiune = firut["CSA"]
                                try:
                                    culoare = core["Colour"].split("_")[1]
                                except IndexError:
                                    culoare = core["Colour"]
            for moduleid in wire["WireModuleRefs"][0]["WireModuleRef"]:
                wires.append([moduleid["ModuleID"], wire["ID"], wire["WireNo"], wire["RouteLength"], cablu, sectiune,
                              culoare])
        wires.insert(0, ["Module ID", "Wire ID", "Wire No", "Length", "Cablu", "Sectiune", "Culoare"])
        printfile(wires, "Wires list")

    def extract_conector_ids(wire_list):
        conectors = []
        for conector in wire_list['Connector']:
            conid = conector["ID"]
            eleID = conector["ElementID"]
            pmdID = conector["PMD"]
            juncID = conector["JunctionID"]
            txID = conector["TableX"]
            tyID = conector["TableY"]
            for ele in conector["Slots"][0]["Slot"][0]["Cavities"][0]["Cavity"]:
                pinid = ele["PinNo"]
                try:
                    plugid = ele["PlugID"]
                except KeyError:
                    plugid = "Empty"
                try:
                    wireid = ele["ConnectorWire"][0]["WireID"]
                except KeyError:
                    wireid = "Empty"
                try:
                    terminalid = ele["ConnectorWire"][0]["Terminals"][0]["Terminal"][0]["TerminalID"]
                except KeyError:
                    terminalid = "Empty"
                try:
                    sealid = ele["ConnectorWire"][0]["Seals"][0]["Seal"][0]["SealID"]
                except KeyError:
                    sealid = "Empty"
                conectors.append([conid, pmdID, eleID, juncID, txID, tyID, pinid, wireid, terminalid, sealid, plugid])
        conectors.insert(0, ["Conector ID", "PMD", "Name", "Junction", "X coord", "Y coord", "Pin", "Wire ID", "Terminal ID",
                             "Seal ID", "Plug ID"])
        printfile(conectors, "Conectors list")

    def extract_module_ids(module_list):
        modules = []
        for module in module_list["Module"]:
            moduleid = module["ID"]
            eleID = module["TitleBlock"][0]["CustomerPartNo"]
            familyref = module["FamilyRef"]
            logisticsID = module["TitleBlock"][0]["Description"]
            modules.append([moduleid, eleID, familyref, logisticsID])
        modules.insert(0, ["Module ID", "MAN ID", "Family", "Description"])
        printfile(modules, "Modules list")

    def extract_variants_ids(variant_list):
        nume_variatii = [["Nume"]]
        for variant in variant_list["VariantNames"][0]["VariantName"]:
            varname = variant["__text__"]
            nume_variatii[0].append(varname)
        for variant in variant_list["VariantParameters"][0]["VariantParameter"]:
            valoare = []
            for value in variant["VariantParameterValue"]:
                valoare.append(value["__text__"])
            valoare.insert(0, variant["Name"])
            nume_variatii.append(valoare)
        printfile(nume_variatii, "Length Variations list")

    def extract_tape_ids(tape_list):
        tapes = []
        for tape in tape_list['Tape']:
            tapeid = tape["ID"]
            tapepmd = tape["PMD"]
            length = tape["TapeLength"]
            try:
                txID = tape["TableX"]
                tyID = tape["TableY"]
            except KeyError:
                txID = "Empty"
                tyID = "Empty"
            for ele in tape["TapeModuleRefs"][0]["TapeModuleRef"]:
                tapes.append([tapeid, tapepmd, length, txID, tyID, ele["ModuleID"]])
        tapes.insert(0, ["Tape ID", "PMD", "Length", "X coord", "Y coord", "Module ID"])
        printfile(tapes, "Tapes list")

    def extract_terminal_ids(terminal_list):
        terminals = []
        for terminal in terminal_list["Terminal"]:
            terminalid = terminal["ID"]
            terminalpmd = terminal["PMD"]
            for ele in terminal["TerminalModuleRefs"][0]["TerminalModuleRef"]:
                terminals.append([terminalid, terminalpmd, ele["ModuleID"]])
        terminals.insert(0, ["Terminal ID", "PMD", "Module ID"])
        printfile(terminals, "Terminals list")

    def extract_seal_ids(seal_list):
        seals = []
        for seal in seal_list["CavitySeal"]:
            sealid = seal["ID"]
            sealpmd = seal["PMD"]
            for ele in seal["CavitySealModuleRefs"][0]["CavitySealModuleRef"]:
                seals.append([sealid, sealpmd, ele["ModuleID"]])
        seals.insert(0, ["Seal ID", "PMD", "Module ID"])
        printfile(seals, "Seals list")

    def extract_plugs_ids(plug_list):
        plugs = []
        for plug in plug_list["CavityPlug"]:
            plugid = plug["ID"]
            plugpmd = plug["PMD"]
            for ele in plug["CavityPlugModuleRefs"][0]["CavityPlugModuleRef"]:
                plugs.append([plugid, plugpmd, ele["ModuleID"]])
        plugs.insert(0, ["Plug ID", "PMD", "Module ID"])
        printfile(plugs, "Plugs list")

    def extract_accessory_ids(accessory_list):
        accessorys = []
        for accessory in accessory_list["Accessory"]:
            accessoryid = accessory["ID"]
            accessorypmd = accessory["PMD"]
            accessoryconectorID = accessory["ReferencedConnectors"][0]["ReferencedConnector"][0]["ConnectorID"]
            for ele in accessory["AccessoryModuleRefs"][0]["AccessoryModuleRef"]:
                accessorys.append([accessoryid, accessorypmd, accessoryconectorID, ele["ModuleID"]])
        accessorys.insert(0, ["Accessory ID", "PMD", "ConnectorID", "Module ID"])
        printfile(accessorys, "Accessory list")


    def extract_configurations(configurations_list):
        configurati = []
        for configuration in configurations_list["Configuration"]:
            configurationid = configuration["ID"]
            configurationnickname = configuration["NickName"]
            for moduleid in configuration["ConfigurationModule"]:
                configurati.append([configurationid, configurationnickname, moduleid["ModuleID"]])
        configurati.insert(0, ["Configuration ID", "NickName", "Module ID"])
        printfile(configurati, "Configurations")


    def extract_accessorypmds(accesory_list):
        accesorypmd = []
        for accesory in accesory_list["AccessoryPMD"]:
            accesorypmdid = accesory["ID"]
            abbr = accesory["Abbreviation"]
            desc = accesory["Description"]
            accesorypmd.append([accesorypmdid, abbr, desc])
        accesorypmd.insert(0, ["AccessoryPMD ID", "Abbreviation", "Description"])
        printfile(accesorypmd, "AccessoryPMDs")

    def printfile(list_to_print, file_name):
        wb = Workbook()
        ws1 = wb.active
        ws1.title = file_name
        for i in range(len(list_to_print)):
            for x in range(len(list_to_print[i])):
                try:
                    ws1.cell(column=x + 1, row=i + 1, value=float(list_to_print[i][x]))
                except:
                    ws1.cell(column=x + 1, row=i + 1, value=str(list_to_print[i][x]))
        try:
            wb.save(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/" + file_name + ".xlsx")
        except PermissionError:
            messagebox.showerror('Eroare scriere', "Fisierul este read-only!")
            return None

        with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/" + file_name + ".txt", 'w', newline='',
                  encoding='utf-8') as myfile:
            wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
            wr.writerows(list_to_print)




    for file in files:
        with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/JSON/" + file + ".json", "r") as json_file:
            loaded_dictionary = json.load(json_file)

            if file == "Modules":
                extract_module_ids(loaded_dictionary)
            if file == "LengthVariants":
                extract_variants_ids(loaded_dictionary)
            if file == "Wires":
                extract_wire_ids(loaded_dictionary)
            if file == "CavitySeals":
                extract_seal_ids(loaded_dictionary)
            if file == "Connectors":
                extract_conector_ids(loaded_dictionary)
            if file == "Tapes":
                extract_tape_ids(loaded_dictionary)
            if file == "Terminals":
                extract_terminal_ids(loaded_dictionary)
            if file == "CavityPlugs":
                extract_plugs_ids(loaded_dictionary)
            if file == "Accessories":
                extract_accessory_ids(loaded_dictionary)
            if file == "Configurations":
                extract_configurations(loaded_dictionary)
            if file == "AccessoryPMDs":
                extract_accessorypmds(loaded_dictionary)
    messagebox.showinfo('Finalizat', "Fisierele JSON finalizate!")

def display_directory_contents(directory, wires):
    def list_directory_contents(directory):
        try:
            # Get the list of files and directories in the specified directory
            contents = os.listdir(directory)
            return contents
        except OSError:
            return []
    def sort_by_first_two_characters(name):
        # Extract the first two characters of the filename
        first_two_characters = name[:2]

        # Convert the first two characters to an integer (if possible)
        try:
            return int(first_two_characters)
        except ValueError:
            # If conversion to an integer fails, return a large value
            return float('inf')

    # Create the main window
    root = tk.Tk()
    root.title("Directory Contents")
    root.geometry("500x400")

    # Create a canvas with a vertical scrollbar
    canvas = tk.Canvas(root)
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(root, command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    canvas.configure(yscrollcommand=scrollbar.set)

    # Create a frame inside the canvas to hold the radio buttons
    frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=frame, anchor=tk.NW)

    # Get the directory contents
    contents = list_directory_contents(directory)
    # Sort the contents based on the first two characters as numbers
    contents.sort(key=sort_by_first_two_characters)
    # Create a list to store the selected items
    selected_items = []



    def print_selection(selected_items):
        on_select()
        print("Selected Items:")
        for itemul in selected_items:
            print("-", itemul)

    def on_select():
        selected_items.clear()
        for item, var in selected_var_dict.items():
            if var.get() == 1:
                selected_items.append(item)

    def clear_selection():
        for var in selected_var_dict.values():
            var.set(0)
        selected_items.clear()

    # Create a dictionary to store the radio buttons' variables
    selected_var_dict = {}

    # Insert radio buttons for each item in the directory
    for item in contents:
        selected_var_dict[item] = tk.IntVar()
        radio_button = tk.Radiobutton(frame, text=item, variable=selected_var_dict[item], value=1)
        radio_button.pack(anchor="w")

    # Create the "Print Selection" button
    print_button = tk.Button(root, text="Print Selection", command=lambda: print_selection(selected_items))
    print_button.pack()

    # Create the "Clear Selection" button
    clear_button = tk.Button(root, text="Clear Selection", command=clear_selection)
    clear_button.pack()

    def on_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))

    frame.bind("<Configure>", on_frame_configure)

    # Start the main event loop
    root.mainloop()



def creare_wirelist():
    #creare GUI
    pbargui = Tk()
    pbargui.title("Creare fisiere pentru diagrame")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    timelabel = Label(pbargui, text="Time . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2)
    timelabel.grid(row=2, column=2)
    # clear folder from old files
    test_excel = []
    for file_all in os.listdir(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Conectori/Compatibili/"):
        try:
            os.remove(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Conectori/Compatibili/" + file_all)
        except:
            continue
    for file_all in os.listdir(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Conectori/Incompatibili/"):
        try:
            os.remove(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Conectori/Incompatibili/" + file_all)
        except:
            continue
    for file_all in os.listdir(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Diagrame create/"):
        try:
            os.remove(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Diagrame create/" + file_all)
        except:
            continue
    #Load required data files
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Wires list.txt", newline='') as csvfile:
        array_wires = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Conectors list.txt", newline='') as csvfile:
        array_conectors  = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Length Variations list.txt", newline='') as csvfile:
        array_variatii = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Modules list.txt", newline='') as csvfile:
        array_modules = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Plugs list.txt", newline='') as csvfile:
        array_plugs = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Seals list.txt", newline='') as csvfile:
        array_seals = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Tapes list.txt", newline='') as csvfile:
        array_tapes = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Terminals list.txt", newline='') as csvfile:
        array_terminals = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Accessory list.txt", newline='') as csvfile:
        array_accessory = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/AccessoryPMDs.txt", newline='') as csvfile:
        array_accessorypmd = list(csv.reader(csvfile, delimiter=';'))

# function needed
    statuslabel["text"] = "Incarcare fisiere . . . "
    pbar['value'] += 2
    pbargui.update_idletasks()
    output = []
    # Extragere conectori si fire
    for i in range(len(array_conectors)):
        if array_conectors[i][7] == "Empty":
            temp_list = [item for item in array_conectors[i]]
            temp_list.append("Empty")
            temp_list.append("Empty")
            temp_list.append("Empty")
            output.append(temp_list)
        else:
            for x in range(len(array_wires)):
                if array_conectors[i][7] == array_wires[x][1]:
                    temp_list = [item for item in array_conectors[i]]
                    temp_list.append(array_wires[x][0])
                    temp_list.append(array_wires[x][2])
                    temp_list.append(array_wires[x][3])
                    output.append(temp_list)
    # atasare supersleeve la conector
    output[0].append("Super S")
    output[0].append("SS Length")
    for i in range(1, len(output)):
        if output[i][10] == "Empty":
            output[i].append("Empty")
            output[i].append("Empty")
        else:
            for x in range(len(array_tapes)):
                if output[i][10] == array_tapes[x][5]:
                    output[i].append(array_tapes[x][1])
                    output[i].append(array_tapes[x][2])
    # atasare accesorii la conector
    for i in range(1, len(output)):
        temp_list_acc = []
        for x in range(len(array_accessory)):
            if output[i][0] == array_accessory[x][2] and output[i][11] == array_accessory[x][3]:
                temp_list_acc.append(array_accessory[x][1])
        output[i].extend(temp_list_acc)
    # inlocuire ID Modul cu part number MAN
    for i in range(1, len(output)):
        for x in range(len(array_modules)):
            if output[i][11] == array_modules[x][0]:
                output[i][11] = array_modules[x][1].replace("PM", "81").replace("VM", "81")
                break
    # inlocuire ID terminal cu part number MAN
    for i in range(1, len(output)):
        for x in range(len(array_terminals)):
            if output[i][8] == array_terminals[x][0]:
                output[i][8] = array_terminals[x][1]
                break
    # inlocuire ID seal cu part number MAN
    for i in range(1, len(output)):
        for x in range(len(array_seals)):
            if output[i][9] == array_seals[x][0]:
                output[i][9] = array_seals[x][1]
                break
    # inlocuire ID plug cu part number MAN
    for i in range(1, len(output)):
        for x in range(len(array_plugs)):
            if output[i][10] == array_plugs[x][0]:
                output[i][10] = array_plugs[x][1]
                break

    # list unica conectori
    lista_conectori = []
    lista_conectori_cu_pin = []
    for i in range(1, len(output)):
        if output[i][0] not in lista_conectori:
            lista_conectori.append(output[i][0])
    # creare lista conectori pentru diagrame
    for i in range(len(lista_conectori)):
        max_pin = 0
        for x in range(len(output)):
            if lista_conectori[i] == output[x][0]:
                current_pin = int(output[x][6].replace("S", ""))
                if current_pin > max_pin:
                    max_pin = current_pin
        if 1 < max_pin < 50:
            lista_conectori_cu_pin.append((lista_conectori[i], max_pin))
    statuslabel["text"] = "Salvare lista fire . . . "
    pbar['value'] += 2
    pbargui.update_idletasks()
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "output"
    for i in range(len(output)):
        for x in range(len(output[i])):
            try:
                ws1.cell(column=x + 1, row=i + 1, value=float(output[i][x]))
            except:
                ws1.cell(column=x + 1, row=i + 1, value=str(output[i][x]))
    wb.save(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Lista output.xlsx")

    lista_conectori_neprelucrati = []
    lista_diagrame_create = []
    for pereche in lista_conectori_cu_pin:
        statuslabel["text"] = "Prelucrare combinatii pentru : " + str(pereche[0]) + " factor de:"
        pbar['value'] += 2
        pbargui.update_idletasks()
        array_temp_wirelist = []
        lista_combinatii_incompatibile = []
        lista_combinatii_compatibile = []
        if pereche[1] != 1 and pereche[1] < 60:
            for x in range(len(output)):
                if pereche[0] == output[x][0]:
                    array_temp_wirelist.append(output[x])
        lista_module_conector = set([x[11] for x in array_temp_wirelist if x[11] != "Empty"])

        module_pins = {}
        for row in output:
            if row[11] != "Empty":
                module = row[11]
                pin = row[6]
                if module in module_pins:
                    module_pins[module].add(pin)
                else:
                    module_pins[module] = {pin}

        if len(lista_module_conector) < 19:
            for r in range(1, len(lista_module_conector) + 1):
                counter_combinatii = 0
                combinatii_posibile = math.factorial(len(lista_module_conector) + 1) // \
                                      (math.factorial(r) * math.factorial(len(lista_module_conector) + 1 - r))
                statuslabel["text"] = str(pereche[0]) + "-Prelucrare "+ str(combinatii_posibile) + " combinatii                                  "
                for combination in combinations(lista_module_conector, r):
                    counter_combinatii = counter_combinatii + 1
                    timelabel["text"] = "Verificate " + str(counter_combinatii) + "                                              "
                    pbar['value'] += 2
                    pbargui.update_idletasks()
                    if len(combination) == 1:
                        num_common_pins = 0
                    else:
                        pins_lists = [module_pins.get(module_id, set()) for module_id in combination]
                        common_pins = set.intersection(*pins_lists)
                        num_common_pins = len(common_pins)
                    if num_common_pins > 0:
                        lista_combinatii_incompatibile.append(combination)
                    else:
                        lista_combinatii_compatibile.append(combination)

                wb = Workbook()
                ws1 = wb.active
                ws1.title = "Module compatibile"
                for i in range(len(lista_combinatii_compatibile)):
                    for x in range(len(lista_combinatii_compatibile[i])):
                        try:
                            ws1.cell(column=x + 1, row=i + 1, value=float(lista_combinatii_compatibile[i][x]))
                        except:
                            ws1.cell(column=x + 1, row=i + 1, value=str(lista_combinatii_compatibile[i][x]))
                wb.save(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Conectori/Compatibili/" + str(pereche[1]) + " " +
                        pereche[0] + ".xlsx")

                wb2 = Workbook()
                ws21 = wb2.active
                ws21.title = "Module incompatibile"
                for i in range(len(lista_combinatii_incompatibile)):
                    for x in range(len(lista_combinatii_incompatibile[i])):
                        try:
                            ws21.cell(column=x + 1, row=i + 1, value=float(lista_combinatii_incompatibile[i][x]))
                        except:
                            ws21.cell(column=x + 1, row=i + 1, value=str(lista_combinatii_incompatibile[i][x]))
                wb2.save(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Conectori/Incompatibili/ " +
                        pereche[0] + ".xlsx")
        else:
            lista_conectori_neprelucrati.append(pereche[0])

    wb3 = Workbook()
    ws31 = wb3.active
    ws31.title = "Conectori neprelucrati"
    for i in range(len(lista_conectori_neprelucrati)):
        try:
            ws31.cell(column=1, row=i + 1, value=float(lista_conectori_neprelucrati[i]))
        except:
            ws31.cell(column=1, row=i + 1, value=str(lista_conectori_neprelucrati[i]))
    wb3.save(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Conectori/Incompatibili/Conectori neperlucrati.xlsx")

    def sort_by_first_two_characters(name):
        # Extract the first two characters of the filename
        first_two_characters = name[:2]

        # Convert the first two characters to an integer (if possible)
        try:
            return int(first_two_characters)
        except ValueError:
            # If conversion to an integer fails, return a large value
            return float('inf')

    def colorpicker(codculoare):
        if codculoare == "bk":
            cell_fill = PatternFill(start_color="000000", end_color="000000", fill_type='solid')
        elif codculoare == "bl":
            cell_fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type='solid')
        elif codculoare == "ws":
            cell_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type='solid')
        elif codculoare == "rs":
            cell_fill = PatternFill(start_color="ff99ff", end_color="ff99ff", fill_type='solid')
        elif codculoare == "bn":
            cell_fill = PatternFill(start_color="996600", end_color="996600", fill_type='solid')
        elif codculoare == "rd":
            cell_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type='solid')
        else:
            cell_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type='solid')
        return cell_fill


    contents = os.listdir(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Conectori/Compatibili/")
    # Sort the contents based on the first two characters as numbers
    contents.sort(key=sort_by_first_two_characters)
    lista_fire = output[1:]

    for i in range(len(contents)):
        counter_diagrama = 1
        wb = load_workbook(os.path.abspath(os.curdir) +
                           "/MAN/Output/Diagrame/EXCELS/Conectori/Compatibili/" + contents[i])
        ws1 = wb.worksheets[0]
        row_count = 1
        conectoropus = "WIP"
        for row in ws1.iter_rows():
            lista_module_diagrama = []
            output_diagrama = {key: None for key in range(1, int(contents[i][0:contents[i].find(' ')]) + 1)}
            # Iterate through worksheet and print cell contents
            for cell in row:
                if cell.value is not None:
                    lista_module_diagrama.append(cell.value)
            for key in output_diagrama.keys():
                for conector in array_conectors:
                    if conector[0] == contents[i][contents[i].find(' ') + 1:-5] and conector[7] != "Empty":
                        fir = conector[7]
                        for wire in array_wires:
                            if wire[1] == fir:
                                modul = wire[0]
                                for plug in array_plugs:
                                    if plug[2] == modul:
                                        numarfir = plug[1]
                                        break
                modulul = ""
                culoare = "bk"
                codcul1 = "bk"
                codcul2 = "bk"
                sectiune = ""
                for fir in lista_fire:
                    if int(fir[6].replace("S", "")) == key and fir[11] in lista_module_diagrama and\
                            fir[0] == contents[i][contents[i].find(' ') + 1:-5]:
                        modulul = fir[11]
                        numarfir = fir[12]
                        for wire in array_wires:
                            if wire[1] == fir[7]:
                                culoare = wire[6]
                                codcul1 = wire[6][:2]
                                if wire[6][2:] != "":
                                    codcul2 = wire[6][2:]
                                else:
                                    codcul2 = wire[6][:2]
                                sectiune = wire[5]

                output_diagrama[key] = modulul, numarfir, culoare, codcul1, codcul2, sectiune, conectoropus
            output_diagrama[0] = "Modul PN", "Imprimare Fir", "Culoare", "", "", "Sectiune", "Conector opus"

            print()
            wb = Workbook()
            ws1 = wb.active
            ws1.title = "V" + str(counter_diagrama)
            # Write conector information
            nume_conector = ""
            partman_conector = ""
            descrierecon = ""
            for con in array_conectors:
                if con[0] == contents[i][contents[i].find(' ') + 1:-5]:
                    partman_conector = con[1]
                    nume_conector = con[2]
                    break
            for pmd in array_accessorypmd:
                if pmd[0] == partman_conector:
                    descrierecon = pmd[2]
                    break
            print(contents[i][contents[i].find(' ') + 1:-5])
            print(nume_conector)
            print(partman_conector)
            print(descrierecon)
            print()

            ws1.cell(row=1, column=2, value=nume_conector)
            ws1.cell(row=2, column=2, value=partman_conector)
            ws1.cell(row=3, column=2, value=descrierecon)

            # Write the data from the dictionary to the Excel worksheet
            for row_num, row_data in output_diagrama.items():
                ws1.cell(row=row_num + 4, column=1, value=row_data[0])  # Column 1: '81.25481-7038'
                ws1.cell(row=row_num + 4, column=2, value=row_num)  # Column 2: Key
                ws1.cell(row=row_num + 4, column=3, value=row_data[1])  # Column 3: 'blau_001'
                ws1.cell(row=row_num + 4, column=4, value=row_data[2])  # Column 4: 'bl'
                ws1.cell(row=row_num + 4, column=5, value="").fill = colorpicker(row_data[3])  # Column 5: 'bl'
                ws1.cell(row=row_num + 4, column=6, value="").fill = colorpicker(row_data[4])  # Column 6: 'bl'
                ws1.cell(row=row_num + 4, column=7, value=row_data[5])  # Column 7: '1'
                ws1.cell(row=row_num + 4, column=8, value=row_data[6])  # Column 8: 'WIP'

            wb.save(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Diagrame create/" +
                    contents[i][contents[i].find(' ') + 1:-5] + " V" + str(counter_diagrama)  + ".xlsx")
            row_count = row_count + 1
            lista_diagrame_create.append([contents[i][contents[i].find(' ') + 1:-5] + " V" + str(counter_diagrama),
                                      lista_module_diagrama])
            counter_diagrama = counter_diagrama + 1



    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Module compatibile"
    for i in range(len(lista_diagrame_create)):
        for x in range(len(lista_diagrame_create[i])):
            try:
                ws1.cell(column=x + 1, row=i + 1, value=float(lista_diagrame_create[i][x]))
            except:
                ws1.cell(column=x + 1, row=i + 1, value=str(lista_diagrame_create[i][x]))
    wb.save(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Diagrame create/Lista diagrame create.xlsx")
    print(lista_diagrame_create)
    print()
    print("FINISH")
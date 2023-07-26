import csv
import os
from tkinter import filedialog, messagebox
import json
import xml.etree.ElementTree as ET
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
                          "Tapes", "Terminals", "CavityPlugs", "Accessories"]

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
             "Tapes", "Terminals", "CavityPlugs", "Accessories"]

    def extract_wire_ids(wire_list):
        wires = []
        for wire in wire_list['Wire']:
            for moduleid in wire["WireModuleRefs"][0]["WireModuleRef"]:
                wires.append([moduleid["ModuleID"], wire["ID"], wire["WireNo"], wire["RouteLength"]])
        wires.insert(0, ["Module ID", "Wire ID", "Wire No", "Length"])
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

    messagebox.showinfo('Finalizat', "Fisierele JSON finalizate!")


def creare_wirelist():
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

# function needed

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

    for pereche in lista_conectori_cu_pin:
        array_temp_wirelist = []
        lista_combinatii_incompatibile = []
        lista_combinatii_compatibile = []
        if pereche[1] != 1 and pereche[1] < 60:
            for x in range(len(output)):
                if pereche[0] == output[x][0]:
                    array_temp_wirelist.append(output[x])
        lista_module_conector = set([x[11] for x in array_temp_wirelist if x[11] != "Empty"])
        print(lista_module_conector)
        for r in range(1, len(lista_module_conector) + 1):
            print(r)
            for combination in combinations(lista_module_conector, r):
                print(combination)
                incompatibil = False
                if len(combination) == 1:
                    lista_combinatii_compatibile.append(combination)
                else:
                    common_pins = []
                    for modul in combination:
                        for row in output:
                            if modul in row:
                                common_pins.append(row[6])
                    for count in Counter(common_pins).values():
                        if count > 1:
                            incompatibil = True
                if incompatibil:
                    lista_combinatii_incompatibile.append(combination)
                else:
                    lista_combinatii_compatibile.append(combination)

        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Module compatibile"
        ws2 = wb.create_sheet("Module incompatibile")
        for i in range(len(lista_combinatii_compatibile)):
            for x in range(len(lista_combinatii_compatibile[i])):
                try:
                    ws1.cell(column=x + 1, row=i + 1, value=float(lista_combinatii_compatibile[i][x]))
                except:
                    ws1.cell(column=x + 1, row=i + 1, value=str(lista_combinatii_compatibile[i][x]))
        for i in range(len(lista_combinatii_incompatibile)):
            for x in range(len(lista_combinatii_incompatibile[i])):
                try:
                    ws2.cell(column=x + 1, row=i + 1, value=float(lista_combinatii_incompatibile[i][x]))
                except:
                    ws2.cell(column=x + 1, row=i + 1, value=str(lista_combinatii_incompatibile[i][x]))
        wb.save(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Conectori/Combinatii module " +
                pereche[0] + ".xlsx")




    print("FINISH")
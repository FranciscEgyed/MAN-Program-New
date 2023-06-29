import csv
import os
from tkinter import filedialog, messagebox
import json
import xml.etree.ElementTree as ET
from openpyxl.workbook import Workbook


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
             "Tapes", "Terminals", "CavityPlugs"]

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
                    plugid = "None"
                try:
                    wireid = ele["ConnectorWire"][0]["WireID"]
                    terminalid = ele["ConnectorWire"][0]["Terminals"][0]["Terminal"][0]["TerminalID"]
                except KeyError:
                    wireid = "Empty"
                    terminalid = "Empty"
                conectors.append([conid, pmdID, eleID, juncID, txID, tyID, pinid, wireid, terminalid, plugid])
        conectors.insert(0, ["Conector ID", "PMD", "Name", "Junction", "X coord", "Y coord", "Pin", "Wire ID", "Terminal ID",
                             "Seal ID"])
        printfile(conectors, "Conectors list")

    def extract_module_ids(module_list):
        modules = []
        for module in module_list["Module"]:
            moduleid = module["ID"]
            eleID = module["ElementID"]
            familyref = module["FamilyRef"]
            logisticsID = module["LogisticsID"]
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
            if file == "Accessory":
                extract_accessory_ids(loaded_dictionary)

    messagebox.showinfo('Finalizat', "Fisierele JSON finalizate!")


def creare_wirelist():
    #Load required data files
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Wires list.txt", newline='') as csvfile:
        array_wires = list(csv.reader(csvfile, delimiter=';'))
    with open(os.path.abspath(os.curdir) + "/MAN/Output/Diagrame/EXCELS/Conectors  list.txt", newline='') as csvfile:
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

    for i in range(len(array_wires)):
        print(array_wires[i])
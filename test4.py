import json
import os
import time
from tkinter import messagebox
from openpyxl.workbook import Workbook


files = ["Modules", "LengthVariants", "Wires", "CavitySeals", "Connectors",
                          "Tapes", "Terminals", "CavityPlugs"]
def print_list_of_dictionaries(lst, indent=0):
    for item in lst:
        if isinstance(item, dict):
            print_dictionary(item, indent)
        elif isinstance(item, list):
            print_list_of_dictionaries(item, indent)
        else:
            print(f"{' ' * indent}{item}")

def print_dictionary(dictionary, indent=0):
    output = []
    for key, value in dictionary.items():
        if isinstance(value, dict):
            print(f"{' ' * indent}{key}:")
            print_dictionary(value, indent + 4)
        elif isinstance(value, list):
            print(f"{' ' * indent}{key}:")
            print_list_of_dictionaries(value, indent + 4)
        else:
            print(f"{' ' * indent}{key}: {value}")
        time.sleep(0.1)


def extract_wire_ids(wire_list):
    wires = []
    for wire in wire_list['Wire']:
        for moduleid in wire["WireModuleRefs"][0]["WireModuleRef"]:
            wires.append([moduleid["ModuleID"], wire["ID"], wire["WireNo"],wire["RouteLength"]])
    wires.insert(0,["Module ID", "Wire ID", "Wire No", "Length"])
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
            plugid = ele["PlugID"]
            try:
                wireid = ele["ConnectorWire"][0]["WireID"]
                terminalid = ele["ConnectorWire"][0]["Terminals"][0]["Terminal"][0]["TerminalID"]
            except KeyError:
                wireid = "Empty"
                terminalid = "Empty"
            conectors.append([conid, pmdID, eleID, juncID, txID, tyID, pinid, wireid, terminalid, plugid])
    conectors.insert(0,["Conector ID", "PMD", "Junction", "X coord", "Y coord","Pin", "Wire ID", "Terminal ID",
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
    modules.insert(0,["Module ID", "MAN ID", "Family", "Description"])
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
        valoare.insert(0,variant["Name"])
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
    tapes.insert(0,["Tape ID", "PMD", "Length", "X coord", "Y coord", "Module ID"])
    printfile(tapes, "Tapes list")


def extract_terminal_ids(terminal_list):
    terminals = []
    for terminal in terminal_list["Terminal"]:
        terminalid = terminal["ID"]
        terminalpmd = terminal["PMD"]
        for ele in terminal["TerminalModuleRefs"][0]["TerminalModuleRef"]:
            terminals.append([terminalid, terminalpmd, ele["ModuleID"]])
    terminals.insert(0,["Terminal ID", "PMD", "Module ID"])
    printfile(terminals, "Terminals list")


def extract_seal_ids(seal_list):
    seals = []
    for seal in seal_list["CavitySeal"]:
        sealid = seal["ID"]
        sealpmd = seal["PMD"]
        for ele in seal["CavitySealModuleRefs"][0]["CavitySealModuleRef"]:
            seals.append([sealid, sealpmd, ele["ModuleID"]])
    seals.insert(0,["Seal ID", "PMD", "Module ID"])
    printfile(seals, "Seals list")


def extract_plugs_ids(plug_list):
    plugs = []
    for plug in plug_list["CavityPlug"]:
        plugid = plug["ID"]
        plugpmd = plug["PMD"]
        for ele in plug["CavityPlugModuleRefs"][0]["CavityPlugModuleRef"]:
            plugs.append([plugid, plugpmd, ele["ModuleID"]])
    plugs.insert(0,["Plug ID", "PMD", "Module ID"])
    printfile(plugs, "Plugs list")

def extract_accessory_ids(accessory_list):
    accessorys = []
    for accessory in accessory_list["Accessory"]:
        accessoryid = accessory["ID"]
        accessorypmd = accessory["PMD"]
        accessoryconectorID = accessory["ReferencedConnectors"][0]["ReferencedConnector"][0]["ConnectorID"]
        for ele in accessory["AccessoryModuleRefs"][0]["AccessoryModuleRef"]:
            accessorys.append([accessoryid, accessorypmd, accessoryconectorID, ele["ModuleID"]])
            print(accessoryid, accessorypmd, accessoryconectorID, ele["ModuleID"])

def printfile(list_to_print, file_name):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "file_name"
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


with open("F:\Python Projects\MAN 2022\MAN\Output\Diagrame\JSON/Accessories.json", "r") as json_file:
    loaded_dictionary = json.load(json_file)
extract_accessory_ids(loaded_dictionary)




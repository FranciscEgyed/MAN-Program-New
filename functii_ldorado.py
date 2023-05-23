import os
from tkinter import filedialog
from xml.etree.ElementTree import parse
from functii_print import prn_excel_module_ldorado


def segment_test():
    Ldorado_file = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                              title="Incarcati fisierul LDorado")
    tree = parse(Ldorado_file)
    root = tree.getroot()

    array_conector = [['Connector ID', 'Connector Name', 'PMD', 'Cavity PinNo', 'Connector WireID',
                       'ConnectorModuleRefs']]
    array_module = [['Module ID', 'CustomerPartNo', 'AscertainedPPSPartNo', 'Description', 'Signature']]
    array_conectorpmd = [['ConnectorPMD ID', 'Abbreviation', 'Description', 'HousingColour', 'HousingType',
                          'NumberOfCavities', 'Cavity ID', 'Terminal Type', 'CustomerPartNo']]
    array_wires = [["Wire ID", "ElementID", "PMD", "WireNo", "MultiCoreID", "WireModuleRef ModuleID"]]
    array_wirepmd = [["WirePMD ID", "Abbreviation", "Description", "WireType", "CSA", "Colour"]]
    module_conector = [['Connector ID', 'Module ID']]
    # drawing number
    drawing_number = root.find("Harness/TitleBlock").attrib['CustomerPartNo']

    for connectors in root.find("Harness/Connectors"):
        for connector in connectors:
            if connector.tag == "Slots":
                for slots in connector:
                    for slot in slots:
                        for cavities in slot:
                            for cavity in cavities:
                                array_conector.append(
                                    [connectors.attrib['ID'].strip(), connectors.attrib['ElementID'].strip(),
                                     connectors.attrib['PMD'].strip(), cavities.attrib['PinNo'].strip(),
                                     cavity.attrib['WireID'].strip(), module_conector])

    for connectors in root.find("Harness/Connectors"):
        for connector in connectors:
            if connector.tag == "ConnectorModuleRefs":
                for modref in connector:
                    module_conector.append([connectors.attrib['ID'].strip(), modref.attrib['ModuleID'].strip()])

    for module in root.find("Harness/Modules"):
        for titleblock in module:
            if titleblock.tag == "TitleBlock":
                array_module.append([module.attrib['ID'].strip(), titleblock.attrib['CustomerPartNo'].strip(),
                                     module.attrib['AscertainedPPSPartNo'].strip(),
                                     titleblock.attrib['Description'].strip()])
    cpn = []
    for connectorpmd in root.find("PMDs/ConnectorPMDs"):
        for node in connectorpmd:
            if node.tag == "Accessory":
                cpn.append([connectorpmd.attrib['ID'].strip(), node.attrib['CustomerPartNo'].strip()])
            else:
                cpn.append([connectorpmd.attrib['ID'].strip(), "None"])
            if node.tag == "SlotPMD":
                for cavitypmd in node:
                    for validtypes in cavitypmd:
                        for terminal in validtypes:
                            array_conectorpmd.append([connectorpmd.attrib['ID'].strip(),
                                                      connectorpmd.attrib['Abbreviation'].strip(),
                                                      connectorpmd.attrib['Description'].strip(),
                                                      connectorpmd.attrib['HousingColour'].strip(),
                                                      connectorpmd.attrib['HousingType'].strip(),
                                                      node.attrib['NumberOfCavities'].strip(),
                                                      cavitypmd.attrib['ID'].strip(),
                                                      terminal.attrib['Type'].strip()])
    for i in range(1, len(array_conectorpmd)):
        for x in range(len(cpn)):
            if cpn[x][0] == array_conectorpmd[i][0]:
                array_conectorpmd[i].append(cpn[x][1])
                break

    for wire in root.find("Harness/Wires"):
        for node in wire:
            if node.tag == "WireModuleRefs":
                for wiremoduleref in node:
                    array_wires.append([wire.attrib['ID'].strip(), wire.attrib['ElementID'].strip(),
                                        wire.attrib['PMD'].strip(), wire.attrib['WireNo'].strip(),
                                        wire.attrib['MultiCoreID'].strip(), wiremoduleref.attrib['ModuleID'].strip()])

    for generalspecialwirepmd in root.find("PMDs/GeneralSpecialWirePMDs"):
        array_wirepmd.append([generalspecialwirepmd.attrib['ID'].strip(),
                              generalspecialwirepmd.attrib['Abbreviation'].strip(),
                              generalspecialwirepmd.attrib['Description'].strip(),
                              generalspecialwirepmd.attrib['WireType'].strip(),
                              generalspecialwirepmd.attrib['CSA'].strip(),
                              generalspecialwirepmd.attrib['Colour'].strip()])
    for generalwirepmd in root.find("PMDs/GeneralWirePMDs"):
        array_wirepmd.append([generalwirepmd.attrib['ID'].strip(),
                              generalwirepmd.attrib['Abbreviation'].strip(),
                              generalwirepmd.attrib['Description'].strip(),
                              generalwirepmd.attrib['WireType'].strip(),
                              generalwirepmd.attrib['CSA'].strip(),
                              generalwirepmd.attrib['Colour'].strip()])


    prn_excel_module_ldorado(array_conector, array_module, array_conectorpmd, array_wires, array_wirepmd, module_conector)

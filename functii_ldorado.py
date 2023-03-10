import os
from tkinter import filedialog
from xml.etree.ElementTree import parse
from functii_print import prn_excel_module_LDorado


def segment_test():
    Ldorado_file = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                              title="Incarcati fisierul LDorado")
    tree = parse(Ldorado_file)
    root = tree.getroot()

    array_conector = [['Connector ID', 'Connector Name', 'PMD', 'Cavity PinNo', 'Connector WireID']]
    array_module = [['Module ID', 'CustomerPartNo', 'AscertainedPPSPartNo', 'Description', 'Signature']]
    array_conectorpmd = [['ConnectorPMD ID', 'Abbreviation', 'Description', 'HousingColour', 'HousingType',
                          'NumberOfCavities', 'Cavity ID', 'Terminal Type', 'CustomerPartNo']]
    array_slots = []

    # drawing number
    drawing_number = root.find("Harness/TitleBlock").attrib['CustomerPartNo']

    for connectors in root.find("Harness/Connectors"):
        for connector in connectors:
            if connector.tag == "Slots":
                for slots in connector:
                    for slot in slots:
                        for cavities in slot:
                            for cavity in cavities:
                                array_conector.append([connectors.attrib['ID'].strip(), connectors.attrib['ElementID'].strip(),
                                      connectors.attrib['PMD'].strip(), cavities.attrib['PinNo'].strip(),
                                      cavity.attrib['WireID'].strip()])

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
    for i in range(len(array_conectorpmd)):
        print(array_conectorpmd[i])





    #prn_excel_module_LDorado(array_print, "SegAttTubes " + drawing_number)








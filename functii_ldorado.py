import os
from tkinter import filedialog
from xml.etree.ElementTree import parse
from functii_print import prn_excel_module_LDorado


def segment_test():
    Ldorado_file = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                              title="Incarcati fisierul LDorado")
    tree = parse(Ldorado_file)
    root = tree.getroot()

    array_conector = ['Connector ID', 'Connector Name', 'PMD', 'Cavity PinNo', 'Connector WireID']
    array_module = ['Module ID', 'CustomerPartNo', 'AscertainedPPSPartNo', 'Description', 'Signature']
    # drawing number
    drawing_number = root.find("Harness/TitleBlock").attrib['CustomerPartNo']
    for child in root.find("Harness/Connectors"):
        for node in child:
            if node.tag == "Slots":
                for node2 in node:
                    for node3 in node2:
                        for node4 in node3:
                            for node5 in node4:
                                array_conector.append([child.attrib['ID'].strip(), child.attrib['ElementID'].strip(),
                                      child.attrib['PMD'].strip(), node4.attrib['PinNo'].strip(),
                                      node5.attrib['WireID'].strip()])
    for child in root.find("Harness/Modules"):
        for node in child:
            if node.tag == "Modules":
                for node2 in node:  # TitleBlock
                    array_module.append([child.attrib['ID'].strip(), node2.attrib['CustomerPartNo'].strip(),
                                         node2.attrib['AscertainedPPSPartNo'].strip(),
                                         node2.attrib['Description'].strip(), child.attrib['Signature'].strip()])
    for i in range(len(array_conector)):
        print(array_conector[i])
    for i in range(len(array_module)):
        print(array_module[i])
    #prn_excel_module_LDorado(array_print, "SegAttTubes " + drawing_number)








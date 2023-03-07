import os
from tkinter import filedialog
import xml.etree.ElementTree as ET
from functii_print import prn_excel_module_LDorado


def segment_att_tubes():
    Ldorado_file = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                              title="Incarcati fisierul LDorado")
    tree = ET.parse(Ldorado_file)
    root = tree.getroot()

    array_write = []

    # drawing number
    drawing_number = root.find("Harness/TitleBlock").attrib['CustomerPartNo']

    # lista cu modulele (ID si ElementID)
    lista_module = []
    for module in root.findall("Harness/Modules/Module"):
        lista_module.append([module.attrib['ElementID'], module.attrib['ID']])

    # lista tape IDs(TapeModuleRefs===ModuleID, PMD, ElementID
    lista_tape = [
        ["ID", "PMD", "TapeLength", "SegmentID", "ModuleID"]]  # ElementID, PMD, TapeLength, SegmentID,ModuleID
    for tape in root.findall("Harness/Tapes/Tape"):
        lista_tape.append([tape.attrib['ID'], tape.attrib['PMD'], tape.attrib['TapeLength']])
    i = 1
    for tape in root.findall("Harness/Tapes/Tape/TapeSegments/TapeSegment"):
        lista_tape[i].append(tape.attrib['SegmentID'])
        i = i + 1
    i = 1
    for tape in root.findall("Harness/Tapes/Tape/TapeModuleRefs"):
        tape_module_id_tmp = []
        for child in tape:
            tape_module_id_tmp.append(child.attrib['ModuleID'])
        lista_tape[i].extend(tape_module_id_tmp)
        i = i + 1

    # segmentIDs ID ElementID Value Length
    lista_segmenete = []
    for segment in root.findall("Harness/Segments/Segment"):
        for child in segment:
            for infant in child:
                try:
                    lista_segmenete.append([segment.attrib['ID'], segment.attrib['ElementID'], infant.attrib["Value"],
                                            segment.attrib['Length']])
                except:
                    continue

    # tape PMD
    lista_tapePMD = []
    for PMD in root.findall("PMDs/TapePMDs/TapePMD"):
        lista_tapePMD.append([PMD.attrib['ID'], PMD.attrib['Description']])

    array_print = [["Drawing No.", "SegmentID", "Segment Element ID", "CA_Segment_ID", "Segment Lenght", "ID_Tap",
                    "Part No MAN", "Tape Lenght", "Description"]]

    for i in range(1, len(lista_tape)):
        array_print.append([drawing_number, lista_tape[i][3], "", "", "", lista_tape[i][0], lista_tape[i][1],
                            lista_tape[i][2]])
    for i in range(1, len(array_print)):
        for x in range(len(lista_tapePMD)):
            if lista_tapePMD[x][0] == array_print[i][6]:
                array_print[i].extend([lista_tapePMD[x][1]])
    for i in range(1, len(array_print)):
        for x in range(len(lista_segmenete)):
            if lista_segmenete[x][0] == array_print[i][1]:
                array_print[i][3] = lista_segmenete[x][2]
                array_print[i][2] = lista_segmenete[x][1]
                array_print[i][4] = lista_segmenete[x][3]

    for i in range(len(lista_tape)):
        for x in range(len(lista_module)):
            if lista_module[x][1] in lista_tape[i]:
                lista_tape[i].append(lista_module[x][0])

    for i in range(len(lista_module)):
        array_print[0].append(lista_module[i][0])

    for i in range(9, len(array_print[0])):
        for x in range(1, len(array_print)):
            for y in range(1, len(lista_tape)):
                if array_print[0][i] in lista_tape[y] and array_print[x][5] in lista_tape[y]:
                    array_print[x].extend([''] * (len(array_print[0]) - 9))
                    array_print[x][i] = lista_tape[y][2]

    prn_excel_module_LDorado(array_print, "SegAttTubes " + drawing_number)



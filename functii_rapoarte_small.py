import csv
import os


def file_namer(sheet):
    file_name = sheet["A2"].value
    return file_name


def heckmoduler(sheet):
    hecklist = []
    for row in sheet["J"]:
        if row.value != "Heck module" and row.value is not None:
            hecklist.append(row.value)
    if len(hecklist) < 1:
        return "No Heck module"
    else:
        return ', '.join(hecklist)


def klappschaller(sheet, sheet1):
    steering_side = sheet["H2"].value
    module_mising = []
    module_prezente = [[], []]
    verificare_side_module = []
    verificare_inversa = []
    klappschalle1 = ""
    lista_selectie = (["BODYL", "8011/8012"], ["BODYR", "8013/8014"], ["BODYL", "8023"], ["BODYR", "8024"],
                      ["BODYL", "8025"], ["BODYR", "8026"])
    # "Verificare steering side"
    if steering_side == "LHD":
        check = "Check LL "
    elif steering_side == "RHD":
        check = "Check LR "
    else:
        check = "Error "
    #"Verificare kalppschalle"
    for row in sheet["J"]:
        if row.value != "Module" and row.value is not None:
            module_prezente[0].append(row.value)
    for row in sheet["K"]:
        if row.value != "Drawing" and row.value is not None:
            module_prezente[1].append(row.value)
    for row in sheet["C"]:
        if row.value == steering_side:
            rand = row.row
            if sheet.cell(row=rand, column=2).value in module_prezente[0]:
                for i in range(len(lista_selectie)):
                    if lista_selectie[i][1] == sheet.cell(row=rand, column=4).value:
                        verificare_side_module.append([sheet.cell(row=rand, column=2).value, lista_selectie[i][0]])
    for i in range(len(module_prezente[0])):
        for x in range(len(verificare_side_module)):
            if module_prezente[0][i] == verificare_side_module[x][0]:
                if module_prezente[1][i] == verificare_side_module[x][1]:
                    klappschalle1 = ""
                else:
                    klappschalle1 = module_prezente[0][i] + " wrong side"
    #"Module dublate"
    duplicat0315 = ""
    duplicat0316 = ""
    counter0315 = 0
    counter0316 = 0
    for row in sheet["J"]:
        if row.value != "Module" and row.value is not None:
            if row.value == "81.25433-0315":
                counter0315 = counter0315 + 1
            elif row.value == "81.25433-6095":
                counter0315 = counter0315 + 1
            elif row.value == "81.25433-0316":
                counter0316 = counter0316 + 1
            elif row.value == "81.25433-6096":
                counter0316 = counter0316 + 1
    if counter0316 > 1:
        duplicat0316 = "81.25433-0316/81.25433-6096 both present "
    if counter0315 > 1:
        duplicat0315 = "81.25433-0315/81.25433-6095 both present "
    duplicataklappschalle = duplicat0316 + duplicat0315
    #"Module missing"
    #"to be rewieed"
    for row in sheet["N"]:
        if row.value != "Module absente" and row.value is not None:
            if row.value == "81.25433-0315" and "81.25433-6095" in module_prezente[0]:
                'module_mising.append("")'
                continue
            elif row.value == "81.25433-0316" and "81.25433-6096" in module_prezente[0]:
                'module_mising.append("")'
                continue
            else:
                module_mising.append(row.value)
    if not len(module_mising) == 0:
        if len(module_mising[0]) > 0:
            klappschalle2 = "Missing " + str(', '.join(module_mising))
        elif len(module_mising[1]) > 0:
            klappschalle2 = "Missing " + str(', '.join(module_mising))
        else:
            klappschalle2 = ""
    else:
        klappschalle2 = ""
    count2450 = 0
    count2451 = 0
    for row in sheet1["B"]:
        if row.value == "81.99918-2450":
            count2450 = count2450 + 1
        elif row.value == "81.99918-2451":
            count2451 = count2451 + 1
    if count2451 == 2:
        indication = ""
    elif count2450 == 2:
        indication = ""
    else:
        indication = " Missing LHD-RHD indication"
    if klappschalle1 == "" and klappschalle2 == "":
        klappschalle = "OK" + indication
    else:
        klappschalle = klappschalle1 + klappschalle2 + indication
    #"verificare inversa"
    for row in sheet["C"]:
        if row.value == steering_side:
            if not sheet.cell(row=row.row, column=2).value in verificare_inversa and \
                    (sheet.cell(row=row.row, column=5).value == "X" or sheet.cell(row=row.row, column=6).value == "X"):
                verificare_inversa.append(sheet.cell(row=row.row, column=2).value)
    # verificare 6095
    text6095 = ""
    for row in sheet["J"]:
        if (row.value == "81.25433-6095" and
                sheet.cell(row=row.row, column=12).value != sheet.cell(row=row.row, column=15).value):
            text6095 = (" 81.25433-6095 found " + str(sheet.cell(row=row.row, column=12).value) +
                        " need " + str(sheet.cell(row=row.row, column=15).value))
    # verificare 3 instante klappschale
    texttreiklap = ""
    for row in sheet["L"]:
        if (sheet.cell(row=row.row, column=12).value != "Quantity" and
                sheet.cell(row=row.row, column=12).value is not None):
            if int(sheet.cell(row=row.row, column=12).value) > 2:
                texttreiklap = " " + (sheet.cell(row=row.row, column=10).value  + " found " +
                                str(sheet.cell(row=row.row, column=12).value) + " need " +
                                str(sheet.cell(row=row.row, column=15).value))

    if check == "Error ":
        return check + duplicataklappschalle + text6095 + texttreiklap
    else:
        return check + duplicataklappschalle + klappschalle + text6095 + texttreiklap


def oldnewcheckr(sheet):
    oldnew = []
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Module Excluse.txt", newline='') as csvfile:
        excluse = list(csv.reader(csvfile, delimiter=';'))
    for row in sheet["F"]:
        if row.value == "XXXX" and sheet.cell(row=row.row, column=3).value == "XXXX" and \
                not sheet.cell(row=row.row, column=2).value in excluse[0]:
            oldnew.append(sheet.cell(row=row.row, column=2).value)
    if len(oldnew) > 0:
        return str(', '.join(oldnew))
    else:
        return "None"


def extraautarker(sheet):
    extra = []
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Module Excluse.txt", newline='') as csvfile:
        excluse = list(csv.reader(csvfile, delimiter=';'))
    for row in sheet["F"]:
        if row.value == "XXXX" and sheet.cell(row=row.row, column=7).value == "XXXX" and \
                not sheet.cell(row=row.row, column=2).value in excluse[0] and \
                not sheet.cell(row=row.row, column=3).value == "XXXX":
            extra.append(sheet.cell(row=row.row, column=2).value)
    if len(extra) > 0:
        return str(', '.join(extra))
    else:
        return "None"


def bkkr(sheet):
    return sheet.cell(row=2, column=19).value


def x1555r(sheet1, sheet2):
    x1555 = []
    for row in sheet1["F"]:
        if row.value == "81.25481-7059" and "X1555" in sheet1.cell(row=row.row, column=2):
            x1555.append("Not OK")
        else:
            x1555.append("OK")
    for row in sheet2["F"]:
        if row.value == "81.25481-7059" and "X1555" in sheet1.cell(row=row.row, column=2):
            x1555.append("Not OK")
        else:
            x1555.append("OK")
    if "Not OK" in x1555:
        return "Not OK"
    else:
        return "OK"


def splicewirer(sheet):
    splicewire = []
    resistor = []
    for row in sheet['C']:
        if "RESISTOR" in row.value:
            resistor.append(sheet.cell(row=row.row, column=10).value)
    for row in sheet['L']:
        if row.row == 1:
            continue
        else:
            if row.value != "Duplicat" and sheet.cell(row=row.row, column=11).value <= 4 and \
                    not sheet.cell(row=row.row, column=10).value in resistor and \
                    sheet.cell(row=row.row, column=3).value != "591003_1":
                splicewire.append(sheet.cell(row=row.row, column=1).value)
    splicewire = list(dict.fromkeys(splicewire))
    if len(splicewire) > 0:
        return "Not OK - " + str(splicewire)[1:-1]
    else:
        return "OK"


def x2799r(sheet, sheet2):
    x2799 = []
    count1a1 = 0
    count1a1_1 = 0
    for row in sheet['G']:
        if "X2799.1A1" == row.value:
            count1a1 = count1a1 + 1
        elif "X2799.1A1_1" == row.value:
            count1a1_1 = count1a1_1 + 1
    if count1a1 > 0 and count1a1_1 > 0:
        for row in sheet['G']:
            if "X2799.1A1_1" == row.value:
                x2799.append(sheet.cell(row=row.row, column=2).value)
    for row in sheet2['G']:
        if "X2799.1A1" == row.value:
            count1a1 = count1a1 + 1
        elif "X2799.1A1_1" == row.value:
            count1a1_1 = count1a1_1 + 1
    if count1a1 > 0 and count1a1_1 > 0:
        for row in sheet2['G']:
            if "X2799.1A1_1" == row.value:
                x2799.append(sheet2.cell(row=row.row, column=2).value)
    x2799 = list(dict.fromkeys(x2799))
    if len(x2799) > 0:
        return "Not OK - " + str(x2799)
    else:
        return "OK"


def samewirer(sheet):
    samewire = []
    for row in sheet['L']:
        if row.value != 1 and row.value != "Verificare":
            samewire.append(sheet.cell(row=row.row, column=3).value)
    samewire = list(dict.fromkeys(samewire))
    if len(samewire) > 0:
        return "Wire No. - " + str(', '.join(samewire))
    else:
        return "OK"


def supersleever(sheet):
    supers = []
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Supersleeve.txt", newline='') as csvfile:
        array_supersleeve = list(csv.reader(csvfile, delimiter=';'))
    for row in sheet['B']:
        for x in range(len(array_supersleeve[0])):
            if row.value in array_supersleeve[0][x]:
                supers.append(array_supersleeve[0][x])
        if row.value == "04.37161-9100":
            supers.append("Error 04.37161-9100")
        elif row.value == "04.37161-9000":
            supers.append("Error 04.37161-9000")
    supers = list(dict.fromkeys(supers))
    if len(supers) > 0:
        return str(', '.join(supers))
    else:
        return "No Module"


def xab6101r(sheet, sheet2):
    xab = 0
    for row in sheet['G']:
        if row.value == "XA.B129.1" or row.value == "XA.B610.1":
            xab = xab + 1
    for row in sheet2['G']:
        if row.value == "XA.B129.1" or row.value == "XA.B610.1":
            xab = xab + 1
    if xab > 0:
        return "Present"
    else:
        return "Not present"


def module_implementater(sheet):
    implementate = []
    for row in sheet['G']:
        if row.value == "Not OK":
            implementate.append(sheet.cell(row=row.row, column=2).value)
    if len(implementate) > 0:
        return str(', '.join(implementate))
    else:
        return "OK"


def verificarelungimir(sheet, sheet2):
    """Verificare tip harness"""
    tipharness = ""
    mixed = ""
    comment = ""
    comment1 = ""
    comment2 = ""
    comment3 = ""
    comment4 = ""
    comment5 = ""
    wrcomment = ""
    if sheet.cell(row=1, column=17).value == "OK":
        if sheet.cell(row=2, column=17).value > 0:
            tipharness = sheet.cell(row=2, column=16).value
        elif sheet.cell(row=3, column=17).value > 0:
            tipharness = sheet.cell(row=3, column=16).value
        elif sheet.cell(row=4, column=17).value > 0:
            tipharness = sheet.cell(row=4, column=16).value
        elif sheet.cell(row=5, column=17).value > 0:
            tipharness = sheet.cell(row=5, column=16).value
        elif sheet.cell(row=6, column=17).value > 0:
            tipharness = sheet.cell(row=6, column=16).value
    else:
        mixed = " - Mixed Platforms"
    lungimi_stanga = []
    lungimi_dreapta = []
    lungimi_1 = []
    lungimi_2 = []
    lungimi_3 = []
    lungimi_4 = []

    if tipharness == "SATTEL":
        for row in sheet2['D']:
            if row.value == "LEFT" and sheet2.cell(row=row.row, column=5).value != 0:
                lungimi_stanga.append(sheet2.cell(row=row.row, column=5).value)
            elif row.value == "RIGHT" and sheet2.cell(row=row.row, column=5).value != 0:
                lungimi_dreapta.append(sheet2.cell(row=row.row, column=5).value)
        if len(set(lungimi_stanga)) != 1 or len(set(lungimi_dreapta)) != 1:
            comment1 = "Error q-VLA"
        if len(set(lungimi_stanga)) == 0 or len(set(lungimi_dreapta)) == 0:
            comment = "One side only"
        else:
            if float(lungimi_stanga[0]) == float(lungimi_dreapta[0]) - 100:
                comment = "Left-Right OK"
            else:
                comment = "LEFT RIGHT values error"
        for row in sheet2['F']:
            if row.value is not None and row.value != "r-VLA/RQT":
                lungimi_1.append(row.value)
        for row in sheet2['G']:
            if row.value is not None and row.value != "s-Radstand":
                lungimi_2.append(row.value)
        for row in sheet2['H']:
            if row.value is not None and row.value != "t-NLA":
                lungimi_3.append(row.value)
        for row in sheet2['I']:
            if row.value is not None and row.value != "u-Aoeberhang":
                lungimi_4.append(row.value)
        if len(set(lungimi_1)) > 1:
            comment2 = "Error r-VLA/RQT"
        if len(set(lungimi_2)) > 1:
            comment3 = "Error s-Radstand"
        if len(set(lungimi_3)) > 1:
            comment4 = "Error t-NLA"
        if len(set(lungimi_4)) > 1:
            comment5 = "Error u-Aoeberhang"

    if tipharness == "CHASSIS":
        for row in sheet2['D']:
            if row.value == "LEFT" and sheet2.cell(row=row.row, column=5).value != 0:
                lungimi_stanga.append(sheet2.cell(row=row.row, column=5).value)
            elif row.value == "RIGHT" and sheet2.cell(row=row.row, column=5).value != 0:
                lungimi_dreapta.append(sheet2.cell(row=row.row, column=5).value)
        if len(set(lungimi_stanga)) != 1 or len(set(lungimi_dreapta)) != 1:
            comment1 = "Error q-VLA"
        if len(set(lungimi_stanga)) == 0 or len(set(lungimi_dreapta)) == 0:
            comment = "One side only"
        else:
            if float(lungimi_stanga[0]) == float(lungimi_dreapta[0]) + 250:
                comment = "Left-Right OK"
            else:
                comment = "LEFT RIGHT values error"
        for row in sheet2['F']:
            if row.value is not None and row.value != "r-VLA/RQT":
                lungimi_1.append(row.value)
        for row in sheet2['G']:
            if row.value is not None and row.value != "s-Radstand":
                lungimi_2.append(row.value)
        for row in sheet2['H']:
            if row.value is not None and row.value != "t-NLA":
                lungimi_3.append(row.value)
        for row in sheet2['I']:
            if row.value is not None and row.value != "u-Aoeberhang":
                lungimi_4.append(row.value)
        if len(set(lungimi_1)) > 1:
            comment2 = "Error r-VLA/RQT"
        if len(set(lungimi_2)) > 1:
            comment3 = "Error s-Radstand"
        if len(set(lungimi_3)) > 1:
            comment4 = "Error t-NLA"
        if len(set(lungimi_4)) > 1:
            comment5 = "Error u-Aoeberhang"

    if tipharness == "TGLM":
        for row in sheet2['D']:
            if row.value == "LEFT" and sheet2.cell(row=row.row, column=5).value != 0:
                lungimi_stanga.append(sheet2.cell(row=row.row, column=5).value)
            elif row.value == "RIGHT" and sheet2.cell(row=row.row, column=5).value != 0:
                lungimi_dreapta.append(sheet2.cell(row=row.row, column=5).value)
        if len(set(lungimi_stanga)) != 1 or len(set(lungimi_dreapta)) != 1:
            comment1 = "Error q-VLA"
        if len(set(lungimi_stanga)) == 0 or len(set(lungimi_dreapta)) == 0:
            comment = "One side only"
        else:
            if float(lungimi_stanga[0]) == float(lungimi_dreapta[0]):
                comment = "Left-Right OK"
            else:
                comment = "LEFT RIGHT values error"
        for row in sheet2['F']:
            if row.value is not None and row.value != "r-VLA/RQT":
                lungimi_1.append(row.value)
        for row in sheet2['G']:
            if row.value is not None and row.value != "s-Radstand":
                lungimi_2.append(row.value)
        for row in sheet2['H']:
            if row.value is not None and row.value != "t-NLA":
                lungimi_3.append(row.value)
        for row in sheet2['I']:
            if row.value is not None and row.value != "u-Aoeberhang":
                lungimi_4.append(row.value)
        if len(set(lungimi_1)) > 1:
            comment2 = "Error r-VLA/RQT"
        if len(set(lungimi_2)) > 1:
            comment3 = "Error s-Radstand"
        if len(set(lungimi_3)) > 1:
            comment4 = "Error t-NLA"
        if len(set(lungimi_4)) > 1:
            comment5 = "Error u-Aoeberhang"

    if tipharness == "4AXEL":
        for row in sheet2['D']:
            if row.value == "LEFT" and sheet2.cell(row=row.row, column=5).value != 0:
                lungimi_stanga.append(sheet2.cell(row=row.row, column=5).value)
            elif row.value == "RIGHT" and sheet2.cell(row=row.row, column=5).value != 0:
                lungimi_dreapta.append(sheet2.cell(row=row.row, column=5).value)
        if len(set(lungimi_stanga)) != 1 or len(set(lungimi_dreapta)) != 1:
            comment1 = "Error q-VLA"
        if len(set(lungimi_stanga)) == 0 or len(set(lungimi_dreapta)) == 0:
            comment = "One side only"
        else:
            if float(lungimi_stanga[0]) == float(lungimi_dreapta[0]) + 600:
                comment = "Left-Right OK"
            else:
                comment = "LEFT RIGHT values error"
        for row in sheet2['F']:
            if row.value is not None and row.value != "r-VLA/RQT":
                lungimi_1.append(row.value)
        for row in sheet2['G']:
            if row.value is not None and row.value != "s-Radstand":
                lungimi_2.append(row.value)
        for row in sheet2['H']:
            if row.value is not None and row.value != "t-NLA":
                lungimi_3.append(row.value)
        for row in sheet2['I']:
            if row.value is not None and row.value != "u-Aoeberhang":
                lungimi_4.append(row.value)
        if len(set(lungimi_1)) > 1:
            comment2 = "Error r-VLA/RQT"
        if len(set(lungimi_2)) > 1:
            comment3 = "Error s-Radstand"
        if len(set(lungimi_3)) > 1:
            comment4 = "Error t-NLA"
        if len(set(lungimi_4)) > 1:
            comment5 = "Error u-Aoeberhang"

    if tipharness == "Military":
        for row in sheet2['D']:
            if row.value == "LEFT" and sheet2.cell(row=row.row, column=5).value != 0:
                lungimi_stanga.append(sheet2.cell(row=row.row, column=5).value)
            elif row.value == "RIGHT" and sheet2.cell(row=row.row, column=5).value != 0:
                lungimi_dreapta.append(sheet2.cell(row=row.row, column=5).value)
        if len(set(lungimi_stanga)) != 1 or len(set(lungimi_dreapta)) != 1:
            comment1 = "Error q-VLA"
        if len(set(lungimi_stanga)) == 0 or len(set(lungimi_dreapta)) == 0:
            comment = "One side only"
        else:
            if float(lungimi_stanga[0]) == float(lungimi_dreapta[0]) + 600:
                comment = "Left-Right OK"
            else:
                comment = "LEFT RIGHT values error"
        for row in sheet2['F']:
            if row.value is not None and row.value != "r-VLA/RQT":
                lungimi_1.append(row.value)
        for row in sheet2['G']:
            if row.value is not None and row.value != "s-Radstand":
                lungimi_2.append(row.value)
        for row in sheet2['H']:
            if row.value is not None and row.value != "t-NLA":
                lungimi_3.append(row.value)
        for row in sheet2['I']:
            if row.value is not None and row.value != "u-Aoeberhang":
                lungimi_4.append(row.value)
        if len(set(lungimi_1)) != 1:
            comment2 = "Error r-VLA/RQT"
        if len(set(lungimi_2)) != 1:
            comment3 = "Error s-Radstand"
        if len(set(lungimi_3)) != 1:
            comment4 = "Error t-NLA"
        if len(set(lungimi_4)) != 1:
            comment5 = "Error u-Aoeberhang"
    if len(comment) > 0:
        wrcomment = comment
    if len(comment1) > 0:
        wrcomment = comment + " " + comment1
    if len(comment2) > 0:
        wrcomment = comment + " " + comment1 + " " + comment2
    if len(comment3) > 0:
        wrcomment = comment + " " + comment1 + " " + comment2 + " " + comment3
    if len(comment4) > 0:
        wrcomment = comment + " " + comment1 + " " + comment2 + " " + comment3 + " " + comment4
    if len(comment5) > 0:
        wrcomment = comment + " " + comment1 + " " + comment2 + " " + comment3 + " " + comment4 + " " + comment5
    return wrcomment + " " + mixed


def copylenghtvaluesleftr(sheet):
    side = "LEFT"
    abbrl = ""
    ql = 0
    rl = 0
    sl = 0
    tl = 0
    ul = 0
    arr_celmailung = []
    for row in sheet['D']:
        if row.value == "LEFT":
            arr_celmailung.append(
                [row.row, sheet.cell(row=row.row, column=5).value, sheet.cell(row=row.row, column=6).value,
                 sheet.cell(row=row.row, column=7).value, sheet.cell(row=row.row, column=8).value,
                 sheet.cell(row=row.row, column=9).value])
    for i in range(len(arr_celmailung)):
        for item in reversed(arr_celmailung[i]):
            if item is None:
                arr_celmailung[i].remove(item)
    try:
        col = max(enumerate(arr_celmailung), key=lambda tup: len(tup[1]))[1][0]
        abbrl = sheet.cell(row=col, column=2).value
        ql = sheet.cell(row=col, column=5).value
        rl = sheet.cell(row=col, column=6).value
        sl = sheet.cell(row=col, column=7).value
        tl = sheet.cell(row=col, column=8).value
        ul = sheet.cell(row=col, column=9).value
    except:
        if ql is None:
            ql = 0
        if rl is None:
            rl = 0
        if sl is None:
            sl = 0
        if tl is None:
            tl = 0
        if ul is None:
            ul = 0
        return [side, abbrl, ql, rl, sl, tl, ul]
    if ql is None:
        ql = 0
    if rl is None:
        rl = 0
    if sl is None:
        sl = 0
    if tl is None:
        tl = 0
    if ul is None:
        ul = 0
    return [side, abbrl, ql, rl, sl, tl, ul]


def copylenghtvaluesrightr(sheet):
    side = "RIGHT"
    abbrr = ""
    qr = 0
    rr = 0
    sr = 0
    tr = 0
    ur = 0
    for row in sheet['D']:
        if row.value == "RIGHT" and sheet.cell(row=row.row, column=10).value is not None:
            abbrr = sheet.cell(row=row.row, column=2).value
            qr = sheet.cell(row=row.row, column=5).value
            rr = sheet.cell(row=row.row, column=6).value
            sr = sheet.cell(row=row.row, column=7).value
            tr = sheet.cell(row=row.row, column=8).value
            ur = sheet.cell(row=row.row, column=9).value
            break
    if qr is None:
        qr = 0
    if rr is None:
        rr = 0
    if sr is None:
        sr = 0
    if tr is None:
        tr = 0
    if ur is None:
        ur = 0
    return [side, abbrr, qr, rr, sr, tr, ur]


def dokar(sheet):
    doka6110 = 0
    doka6111 = 0
    doka7339 = 0
    doka7340 = 0
    doka7341 = 0
    doka7342 = 0
    for row in sheet['B']:
        if row.value == "85.AA962-6110":
            doka6110 = doka6110 + 1
        elif row.value == "85.AA962-6111":
            doka6110 = doka6111 + 1
        elif row.value == "81.25482-7340":
            doka7340 = doka7340 + 1
        elif row.value == "81.25482-7341":
            doka7341 = doka7341 + 1
        elif row.value == "81.25482-7342":
            doka7342 = doka7342 + 1
        elif row.value == "81.25482-7339":
            doka7339 = doka7339 + 1

    if doka6110 > 0 and doka6111 > 0:
        doka = "OK DOKA Modules"
        return doka
    elif doka6110 > 0 and doka6111 == 0:
        doka = "Missing  85.AA962-6111"
        return doka
    elif doka6110 == 0 and doka6111 > 0:
        doka = "Missing  85.AA962-6110"
        return doka
    elif doka7339 > 0 and doka7340 > 0:
        doka = "OK DOKA Modules"
        return doka
    elif doka7339 > 0 and doka7340 == 0:
        doka = "Missing  81.25482-7340"
        return doka
    elif doka7339 == 0 and doka7340 > 0:
        doka = "Missing  81.25482-7339"
        return doka
    elif doka7341 > 0 and doka7342 > 0:
        doka = "Non DOKA Modules"
        return doka
    elif doka7341 > 0 and doka7342 == 0:
        doka = "Missing  81.25482-7342"
        return doka
    elif doka7341 == 0 and doka7342 > 0:
        doka = "Missing  81.25482-7341"
        return doka
    elif doka7341 == 0 and doka7342 == 0:
        doka = "No Modules"
        return doka
    elif doka6110 == 0 and doka6111 == 0:
        doka = "No Modules"
        return doka
    elif doka7339 == 0 and doka7340 == 0:
        doka = "No Modules"
        return doka


def x6616r(sheet):
    return "X6616 - " + sheet.cell(row=1, column=22).value + "/" + "X6490 - " + sheet.cell(row=3, column=22).value


def kswr(sheet):
    array_ksw = []
    for row in sheet['F']:
        if "KSW" in row.value:
            array_ksw.append(sheet.cell(column=2, row=row.row).value)
    array_ksw_unice = list(set(array_ksw))
    if len(array_ksw_unice) == 0:
        return "No KSW"
    else:
        output_ksw = ""
        for i in range(len(array_ksw_unice)):
            output_ksw = output_ksw + "/" + array_ksw_unice[i]
        return output_ksw


def militaryr(sheet):
    mil = 0
    for row in sheet['B']:
        if "85.25480-5" in row.value or "85.25480-6" in row.value or "85.25480-7" in row.value:
            mil = mil + 1
    if mil > 0:
        return "Military"
    else:
        return "Not Military"


def my2023r(sheet, sheet2):
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Module MY2023.txt", newline='') as csvfile:
        my2023 = list(csv.reader(csvfile, delimiter=';'))
    tipharness = ""
    if sheet.cell(row=1, column=17).value == "OK":
        if sheet.cell(row=2, column=17).value > 0:
            tipharness = sheet.cell(row=2, column=16).value
        elif sheet.cell(row=3, column=17).value > 0:
            tipharness = sheet.cell(row=3, column=16).value
        elif sheet.cell(row=4, column=17).value > 0:
            tipharness = sheet.cell(row=4, column=16).value
        elif sheet.cell(row=5, column=17).value > 0:
            tipharness = sheet.cell(row=5, column=16).value
        elif sheet.cell(row=6, column=17).value > 0:
            tipharness = sheet.cell(row=6, column=16).value
    else:
        tipharness = "Mixed"
    module_my2023 = []
    mycount = 0
    if tipharness != "Mixed":
        for i in range(len(my2023)):
            if my2023[i][1] == tipharness:
                module_my2023.append(my2023[i][0])
        for row in sheet2['B']:
            if row.value in module_my2023:
                mycount = mycount + 1
        if mycount > 0:
            return "True"
        else:
            return "False"
    else:
        return "Mixed Platforms"


def x6616stvbr(sheet, sheet2, sheet3):
    c_x64901_a1 = 0
    c_x64902_a1 = 0
    for row in sheet2['G']:
        if row.value == "X6490.1A1":
            c_x64901_a1 = c_x64901_a1 + 1
    for row in sheet2['J']:
        if row.value == "X6490.2A1":
            c_x64902_a1 = c_x64902_a1 + 1
    c_x6490 = c_x64901_a1 + c_x64902_a1
    tipharness = ""
    if sheet3.cell(row=1, column=17).value == "OK":
        if sheet3.cell(row=2, column=17).value > 0:
            tipharness = sheet3.cell(row=2, column=16).value
        elif sheet3.cell(row=3, column=17).value > 0:
            tipharness = sheet3.cell(row=3, column=16).value
        elif sheet3.cell(row=4, column=17).value > 0:
            tipharness = sheet3.cell(row=4, column=16).value
        elif sheet3.cell(row=5, column=17).value > 0:
            tipharness = sheet3.cell(row=5, column=16).value
        elif sheet3.cell(row=6, column=17).value > 0:
            tipharness = sheet3.cell(row=6, column=16).value
    else:
        tipharness = "Mixed"

    counter1 = 0
    counter2 = 0
    stvb = "OK"

    if tipharness == "SATTEL" and c_x6490 > 0:
        for row in sheet['B']:
            if row.value == "81.25482-7988":
                counter1 = counter1 + 1
            elif row.value == "81.25482-7987":
                counter2 = counter2 + 1
        if counter1 == 0 and counter2 == 0:
            stvb = "Missing 81.25482-7987 and 81.25482-7988"
        elif counter1 == 0:
            stvb = "Missing 81.25482-7988"
        elif counter2 == 0:
            stvb = "Missing 81.25482-7987"
        else:
            stvb = "OK"
    if tipharness == "CHASSIS" and c_x6490 > 0:
        for row in sheet['B']:
            if row.value == "81.25482-7989":
                counter1 = counter1 + 1
            elif row.value == "81.25482-7990":
                counter2 = counter2 + 1
        if counter1 == 0 and counter2 == 0:
            stvb = "Missing 81.25482-7989 and 81.25482-7990"
        elif counter1 == 0:
            stvb = "Missing 81.25482-7989"
        elif counter2 == 0:
            stvb = "Missing 81.25482-7990"
        else:
            stvb = "OK"
    if tipharness == "4AXEL" and c_x6490 > 0:
        for row in sheet['B']:
            if row.value == "81.25483-5024":
                counter1 = counter1 + 1
            elif row.value == "81.25483-5038":
                counter2 = counter2 + 1
        if counter1 == 0 and counter2 == 0:
            stvb = "Missing 81.25483-5024 and 81.25483-5038"
        elif counter1 == 0:
            stvb = "Missing 81.25483-5024"
        elif counter2 == 0:
            stvb = "Missing 81.25483-5038"
        else:
            stvb = "OK"
    return stvb


def prufungr(sheet, sheetm, sheet1, sheet2):
    tipharness = ""
    available = []
    module_prezente = []
    prufung_final = []
    "Load required data files"
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Prufung.txt", newline='') as csvfile:
        prufung = list(csv.reader(csvfile, delimiter=';'))
    for i in range(len(prufung)):
        if "Availability" in prufung[i][2] and not prufung[i][2] in available:
            available.append(prufung[i][2])
    if sheet.cell(row=1, column=17).value == "OK":
        if sheet.cell(row=2, column=17).value > 0:
            tipharness = sheet.cell(row=2, column=16).value
        elif sheet.cell(row=3, column=17).value > 0:
            tipharness = sheet.cell(row=3, column=16).value
        elif sheet.cell(row=4, column=17).value > 0:
            tipharness = sheet.cell(row=4, column=16).value
        elif sheet.cell(row=5, column=17).value > 0:
            tipharness = sheet.cell(row=5, column=16).value
        elif sheet.cell(row=6, column=17).value > 0:
            tipharness = sheet.cell(row=6, column=16).value
    else:
        tipharness = "Mixed"
    "Avalability  "
    for row in sheetm['B']:
        if row.value != "Module":
            module_prezente.append(row.value)
    sh1tip = sheet1.cell(row=2, column=1).value
    sh2tip = sheet2.cell(row=2, column=1).value
    conectori1 = []
    conectori2 = []
    for row in sheet1['G']:
        if row.value != "Kurzname":
            conectori1.append(row.value)
    for row in sheet2['G']:
        if row.value != "Kurzname":
            conectori2.append(row.value)

    for i in range(len(available)):
        for x in range(len(prufung)):
            arr_module_ava = [[], [""]]
            counter = 0
            if available[i] == prufung[x][2] and str(sh1tip) == prufung[x][1]:
                if prufung[x][0] not in conectori1:
                    arr_module_ava[0].append(prufung[x][0])
                    arr_module_ava[1][0] = prufung[x][3]
                    counter = counter + 1
            if available[i] == prufung[x][2] and str(sh2tip) == prufung[x][1]:
                if prufung[x][0] not in conectori2:
                    arr_module_ava[0].append(prufung[x][0])
                    arr_module_ava[1][0] = prufung[x][3]
                    counter = counter + 1
            if counter != 0 and arr_module_ava[1][0] != "":
                prufung_final.append(arr_module_ava[1][0])
    "Exclude"
    lista_selectie = (["SATTEL", "8011-8013"], ["CHASSIS", "8012-8014"], ["TGLM", "8023-8024"], ["4AXEL", "8025-8026"])
    for i in range(1, len(prufung)):
        if prufung[i][2] == "Exclude":
            for x in range(len(lista_selectie)):
                if lista_selectie[x][1] == prufung[i][4]:
                    tipharness2 = lista_selectie[x][0]
                    if tipharness == tipharness2 and prufung[i][0] in module_prezente and \
                            prufung[i][1] in module_prezente:
                        if not prufung[i][3] in prufung_final:
                            prufung_final.append(prufung[i][3])
    "If M1 and not M2 add M2"
    for i in range(1, len(prufung)):
        if prufung[i][2] == "If (Modul1) then (Modul2)":
            for x in range(len(lista_selectie)):
                if lista_selectie[x][1] == prufung[i][4]:
                    tipharness2 = lista_selectie[x][0]
                    if tipharness == tipharness2 and prufung[i][0] in module_prezente and not prufung[i][
                                                                                                  1] in module_prezente:
                        if not prufung[i][3] in prufung_final:
                            prufung_final.append(prufung[i][3])
    if len(prufung_final) == 0:
        return "OK"
    else:
        return str(', '.join(prufung_final))


def module_check(sheet):
    module_lines = []
    for row in sheet['H']:
        if row.value != "Desen" and sheet.cell(row=row.row, column=6).value != "XXXX":
            if row.value == "BODYL" and "LHD" not in sheet.cell(row=row.row, column=6).value:
                module_lines.append(row.row)
            elif row.value == "BODYR" and "RHD" not in sheet.cell(row=row.row, column=6).value:
                module_lines.append(row.row)
    if len(module_lines) > 0:
        return "Error on line " + str(module_lines)
    else:
        return "OK"


def ckd(sheet):
    result = ""
    with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/CKD.txt", newline='') as csvfile:
        ckd_list = list(csv.reader(csvfile, delimiter=';'))
    for i in range(len(ckd_list[0])):
        if ckd_list[0][i] in sheet.cell(row=2, column=1).value:
            result = "CKD"
        else:
            result = "None"
    return result


def delivery(sheet):
    if sheet.cell(row=2, column=10).value is not None:
        return sheet.cell(row=2, column=10).value
    else:
        return "No date"

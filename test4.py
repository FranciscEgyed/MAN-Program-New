import tkinter as tk
data = [('ID_CON129', 'X4482.1C1_LI', 21, 0, 0, 0), ('ID_CON130', 'X4481.1C1_LI', 21, 0, 0, 0), ('id_316_4', 'X3.Y263.1PT', 4, 1, 1, 0), ('id_316_7', 'X23.S149.1', 1, 1, 1, 0), ('id_316_19', 'X6511.1A1', 8, 1, 1, 0), ('id_316_22', 'XA.A1397.1', 1, 1, 1, 0), ('id_316_24', 'XA.B103.1.LI', 4, 1, 1, 0), ('id_316_25', 'XA.B110.1', 4, 1, 1, 0), ('id_316_33', 'XA.B128.1', 4, 1, 1, 0), ('id_316_34', 'XA.B540.1', 4, 1, 1, 0), ('id_316_39', 'XA.B634.1.LI', 2, 1, 1, 0), ('id_316_42', 'XA.E366.1.LI', 2, 1, 1, 0), ('id_316_44', 'XA.M137.1.LI', 2, 1, 1, 0), ('id_316_45', 'XA.M175.1.LI', 2, 1, 1, 0), ('id_316_52', 'XA.Y160.1', 3, 1, 1, 0), ('id_316_53', 'XA.Y175.1', 2, 1, 1, 0), ('id_316_56', 'XA.Y288.1.LI', 7, 1, 1, 0), ('id_316_61', 'XA.Y507.1', 4, 1, 1, 0), ('id_316_75', 'XF1.A1356.1', 1, 1, 1, 0), ('Connector_occurrence_47', 'X3.Y263.1PT_1', 4, 1, 1, 0), ('Connector_occurrence_61', 'X23.S149.1_20', 1, 1, 1, 0), ('Connector_occurrence_97', 'XA.B110.1_1', 4, 1, 1, 0), ('ID_CON5', 'XB.A1397.1', 1, 1, 1, 0), ('id_316_41', 'X3.X1170.1', 2, 1, 1, 0), ('ID_CON13', 'XM2.G101.3.LI_1', 1, 1, 1, 0), ('ID_CON18', 'XA.Y308.1PT.LI', 2, 1, 1, 0), ('Connector_occurrence_118', 'XA.B284.1_1', 4, 1, 1, 0), ('id_316_30', 'XA.B371.1.LI', 2, 1, 1, 0), ('id_316_31', 'XA.B432.1', 2, 1, 1, 0), ('id_316_26', 'XA.B119.2', 2, 1, 1, 0), ('ID_CON31', 'XBROTL.1A1', 8, 1, 1, 0), ('ID_CON32', 'X4481.2A1', 21, 1, 1, 0), ('ID_CON33', 'X4482.2A1', 18, 1, 1, 0), ('id_316_16', 'X6407.1A1', 1, 1, 1, 0), ('ID_CON36', 'XA.B597.1_1', 3, 1, 1, 0), ('ID_CON37', 'XA.B597.1', 3, 1, 1, 0), ('ID_CON38', 'XM2.G101.3.LI_3', 1, 1, 1, 0), ('ID_CON19', 'X6816.1A1', 12, 1, 1, 0), ('ID_CON42', 'XA.F1590.1_LIBK', 2, 1, 1, 0), ('ID_CON43', 'XA.F1590.1_LIBKK', 2, 1, 1, 0), ('ID_CON50', 'X23.S149.2', 1, 1, 1, 0), ('id_316_67', 'XA.B333.4', 4, 1, 1, 0), ('ID_CON70', 'XA.B597.1_2', 3, 1, 1, 0), ('ID_CON72', 'XA.B1357.1', 5, 1, 1, 0), ('ID_CON74', 'XA.B1358.1', 5, 1, 1, 0), ('ID_CON71', 'X7042.1A1', 5, 1, 1, 0), ('ID_CON73', 'X7043.1A1', 5, 1, 1, 0), ('ID_CON76', 'XM2.G101.3.LI_6', 1, 1, 1, 0), ('ID_CON75', 'XM2.G101.3.LI_7', 1, 1, 1, 0), ('id_316_81', 'XF7.A1356.1', 1, 1, 1, 0), ('id_316_78', 'XF4.A1356.1', 1, 1, 1, 0), ('id_316_76', 'XF9.A1356.1', 1, 1, 1, 0), ('Connector_occurrence_63', 'XC.Q105.3_20', 1, 1, 1, 0), ('Connector_occurrence_62', 'XC.Q105.2_20', 1, 1, 1, 0), ('ID_CON44', 'XP1.G100.3.LI', 1, 1, 1, 0), ('id_316_64', 'XC.Q105.5', 1, 1, 1, 0), ('id_316_69', 'XC.Q105.2', 1, 1, 1, 0), ('ID_CON45', 'XP1.G100.3.LI_20', 1, 1, 1, 0), ('ID_CON52', 'XC.Q105.4', 1, 1, 1, 0), ('ID_CON51', 'XC.Q105.3', 1, 1, 1, 0), ('ID_CON79', 'X4482.1B1', 18, 1, 1, 0), ('ID_CON78', 'X4481.1B1', 21, 1, 1, 0), ('ID_CON25', 'XB.XV802.1', 1, 1, 1, 0), ('ID_CON27', 'XC.XV802.1', 1, 1, 1, 0), ('ID_CON15', 'XM2.G101.3.LI_2', 1, 1, 1, 0), ('ID_CON69', 'XM2.G101.3.LI_5', 1, 1, 1, 0), ('ID_CON80', 'XA.Y278.1PT_3', 7, 1, 1, 0), ('ID_CON92', 'XA.B432.1_1', 2, 1, 1, 0), ('ID_CON97', 'XA.Y212.1_LI', 4, 1, 1, 0), ('ID_CON98', 'XM2.G101.3.LI_26', 1, 1, 1, 0), ('ID_CON99', 'XM2.G101.3.LI_27', 1, 1, 1, 0), ('ID_CON100', 'XM2.G101.3.LI_28', 1, 1, 1, 0), ('ID_CON101', 'XB.XV802.1_1', 1, 1, 1, 0), ('ID_CON103', 'X1.Y505.1PT_1', 2, 1, 1, 0), ('ID_CON104', 'X2.Y505.1PT_1', 6, 1, 1, 0), ('ID_CON106', 'X3.Y505.1PT_1', 4, 1, 1, 0), ('ID_CON108', 'X4482.2B1', 18, 1, 1, 0), ('ID_CON109', 'XA.B597.1_3', 3, 1, 1, 0), ('id_316_12', 'X5050.1B1', 1, 1, 1, 0), ('ID_CON107', 'X4481.2B1', 21, 1, 1, 0), ('id_316_43', 'XA.M103.1.LI', 4, 1, 1, 0), ('id_316_46', 'XA.M196.1', 4, 1, 1, 0), ('ID_CON119', 'X4481.2B2', 21, 1, 1, 0), ('ID_CON120', 'X4482.2B2', 18, 1, 1, 0), ('ID_CON123', 'X4481.1A2', 21, 1, 1, 0), ('ID_CON124', 'X4481.2A2', 21, 1, 1, 0), ('ID_CON121', 'X4482.1A2', 18, 1, 1, 0), ('ID_CON122', 'X4482.2A2', 18, 1, 1, 0), ('ID_CON125', 'XA.B128.3', 3, 1, 1, 0), ('ID_CON126', 'X4481.2C1_LI', 21, 1, 1, 0), ('ID_CON127', 'X4482.2C1_LI', 21, 1, 1, 0), ('id_316_1', 'X1.Y505.1PT', 2, 2, 3, 0), ('id_316_3', 'X2.Y505.1PT', 6, 2, 3, 0), ('id_316_6', 'X3.Y505.1PT', 4, 2, 3, 0), ('id_316_32', 'XA.B476.1', 4, 2, 3, 0), ('id_316_35', 'XA.B546.1', 4, 2, 3, 0), ('id_316_58', 'XA.Y373.1', 4, 2, 3, 0), ('Connector_occurrence_41', 'XA.Y507.1_1', 4, 2, 3, 0), ('id_316_63', 'XX1.1.A1690.1', 8, 2, 3, 0), ('id_316_23', 'XX1.2.A1690.1', 8, 2, 3, 0), ('Connector_occurrence_42', 'XA.B119.2_1', 2, 2, 3, 0), ('ID_CON34', 'X6491.1A1', 21, 2, 3, 0), ('ID_CON35', 'X6492.1A1', 18, 2, 3, 0), ('Connector_occurrence_38', 'XA.B333.4_1', 4, 2, 3, 0), ('ID_CON40', 'X1535.2A1_LI', 2, 2, 3, 0), ('ID_CON41', 'X1535.2A1_LI1', 2, 2, 3, 0), ('ID_CON49', 'XA.A1192.1_2', 8, 2, 3, 0), ('ID_CON48', 'XA.Y437.2.LI_2', 2, 2, 3, 0), ('id_316_40', 'XA.B994.2', 5, 2, 3, 0), ('ID_CON102', 'XA.Y278.1PT_4', 7, 2, 3, 0), ('id_316_94', 'X9172', 30, 2, 3, 0), ('ID_CON128', 'XA.B994.2_1', 5, 2, 3, 0), ('id_316_83', 'XF2.A1356.1', 1, 3, 7, 0), ('id_316_80', 'XF6.A1356.1', 1, 3, 7, 0), ('id_316_79', 'XF5.A1356.1', 1, 3, 7, 0), ('ID_CON83', 'XA.B1319.1', 2, 3, 7, 0), ('ID_CON84', 'XX1.Y507.2', 3, 3, 7, 0), ('ID_CON85', 'XX1.Y507.3', 4, 3, 7, 0), ('ID_CON93', 'XA.B1314.1_1', 4, 3, 7, 0), ('ID_CON94', 'XA.A1192.2', 6, 3, 7, 0), ('ID_CON110', 'X3.Y264.1_1', 4, 3, 7, 0), ('ID_CON112', 'XA.B1314.1', 4, 3, 7, 0), ('ID_CON113', 'XA.B1319.1_1', 2, 3, 7, 0), ('ID_CON114', 'XX1.Y507.3_1', 4, 3, 7, 0), ('ID_CON115', 'XX1.Y507.2_1', 3, 3, 7, 0), ('ID_CON116', 'X7793.1A1_LI', 21, 3, 7, 0), ('ID_CON118', 'X7794.1A1_LI', 21, 3, 7, 0), ('Connector_occurrence_113', 'XA.Y278.1PT_2', 7, 4, 15, 0), ('ID_CON21', 'X4556.1A1', 2, 4, 15, 0), ('id_316_18', 'X6499.1A1', 1, 4, 15, 0), ('ID_CON46', 'X5508.1B1_2', 8, 4, 15, 0), ('ID_CON20', 'XF8.A1356.1', 1, 4, 15, 0), ('id_316_77', 'XF3.A1356.1', 1, 4, 15, 0), ('ID_CON89', 'X9693', 30, 4, 15, 0), ('ID_CON91', 'X9694', 30, 4, 15, 0), ('id_316_14', 'X5508.1B1', 8, 5, 31, 0), ('id_316_68', 'XC.A1397.1', 6, 5, 31, 0), ('ID_CON95', 'XA.B129.1_1', 4, 5, 31, 0), ('ID_CON96', 'XA.B610.1_1', 4, 5, 31, 0), ('Connector_occurrence_112', 'XA.Y278.1PT_1', 7, 6, 63, 0), ('id_316_38', 'XA.B610.1', 4, 7, 127, 0), ('id_316_27', 'XA.B129.1', 4, 10, 1023, 0), ('id_316_87', 'XX1.A1356.1', 1, 10, 1023, 0), ('id_316_88', 'X9088', 30, 10, 1023, 0), ('id_316_5', 'X3.Y264.1', 4, 11, 2047, 0), ('id_316_84', 'X6478.1A1', 43, 12, 4095, 0), ('Connector_occurrence_72', 'X2799.1A1_1', 51, 19, 524287, 2), ('id_316_93', 'X9178', 30, 20, 1048575, 5), ('ID_CON81', 'XA.B129.2', 3, 27, 134217727, 745), ('ID_CON87', 'XX1.2.A1690.1_2', 8, 39, 549755813887, 3054198), ('ID_CON86', 'XX1.1.A1690.1_2', 8, 40, 1099511627775, 6108397), ('id_316_92', 'X9180', 30, 45, 35184372088831, 195468733), ('id_316_11', 'X2799.1A1', 51, 48, 281474976710655, 1563749870), ('id_316_17', 'X6490.2A1', 51, 66, 73786976294838198272, 409927646082434), ('id_316_91', 'X9179', 30, 73, 9444732965739294621696, 52470738698551640), ('id_316_20', 'X6616.1A1', 51, 121, 2658455991569831745807614120560689152, 14769199953165732628919113220096), ('id_316_15', 'X6302.1A1', 72, 125, 42535295865117307932921825928971026432, 236307199250651722062705811521536)]
hederlist = ['Connector Name','Connector ID','Pin count','Module count','Number of combinations','Durata minute']

root = tk.Tk()
root.title("Checkbox Example")

checkboxes_frame = tk.Frame(root)
checkboxes_frame.pack(side="left", padx=10, pady=10)

buttons_frame = tk.Frame(root)
buttons_frame.pack(side="right", padx=10, pady=10)

checkbox_var = []

for header in hederlist:
    label_text = header
    label = tk.Label(checkboxes_frame, text=label_text)
    label.grid(row=0, column=hederlist.index(header), sticky="w")

for idx, (connector_id, name, pin, module_count, combinations, divisions) in enumerate(data):
    var = tk.BooleanVar(root)
    var.set(False)
    checkbox = tk.Checkbutton(checkboxes_frame, text=name, variable=var)
    checkbox.grid(row=idx + 1, column=0, sticky="w")
    checkbox_var.append(var)

    label = tk.Label(checkboxes_frame, text=connector_id)
    label.grid(row=idx +1, column=1, sticky="w")
    label = tk.Label(checkboxes_frame, text=pin)
    label.grid(row=idx +1, column=2, sticky="w")
    label = tk.Label(checkboxes_frame, text=module_count)
    label.grid(row=idx +1, column=3, sticky="w")
    label = tk.Label(checkboxes_frame, text=combinations)
    label.grid(row=idx +1, column=4, sticky="w")
    label = tk.Label(checkboxes_frame, text=divisions)
    label.grid(row=idx +1, column=5, sticky="w")

def check_all():
    for var in checkbox_var:
        var.set(True)


def clear_all():
    for var in checkbox_var:
        var.set(False)


def print_selected():
    selected_items = []
    for idx, var in enumerate(checkbox_var):
        if var.get():
            selected_items.append(data[idx][0])
    print("Selected Checkboxes:", selected_items)


check_all_button = tk.Button(buttons_frame, text="Check All", command=check_all)
check_all_button.pack(fill="x", pady=5)

clear_all_button = tk.Button(buttons_frame, text="Clear All", command=clear_all)
clear_all_button.pack(fill="x", pady=5)

print_selected_button = tk.Button(buttons_frame, text="Print Selected", command=print_selected)
print_selected_button.pack(fill="x", pady=5)

root.mainloop()



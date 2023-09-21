import csv
import os
from tkinter import messagebox


def extragere_welding():
    try:
        with open(os.path.abspath(os.curdir) + "/MAN/Input/Wire Lists/8000.Wirelist.csv",
                  newline='') as csvfile:
            array_wires_1 = list(csv.reader(csvfile, delimiter=';'))
    except FileNotFoundError:
        messagebox.showerror('Eroare fisier ' , 'Lipsa fisierul 8000')
        quit()

    def create_kurzname_pin_dictionary(wirelist):
        # Initialize an empty dictionary to store the result
        kurzname_pin_dict = {}

        # Loop through the rows starting from the second row (index 1)
        for row in wirelist[1:]:
            if row[5] not in kurzname_pin_dict:
                kurzname_pin_dict[row[5]] = [row[6]]
            elif row[6] not in kurzname_pin_dict.get(row[5], []):
                print(kurzname_pin_dict.get(row[5], []))
                kurzname_pin_dict[row[5]].append(row[6])
        return kurzname_pin_dict
    def combinatii_valabile(lista_module):
        for t in range(1, len(lista_module) + 1):
            for combination in combinations(lista_module_, t):
                if len(combination) == 1:
                    common_pins = False
                else:
                    pins_lists = [module_pins.get(module_id, set()) for module_id in combination]
                    common_pins = has_common_elements(pins_lists)
                if common_pins:
                    lista_combinatii_incompatibile.append(combination)
                else:
                    lista_combinatii_compatibile.append(combination)



    print(create_kurzname_pin_dictionary(array_wires_1))
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


def extract_module_ids(module_list):
    modules = []
    for module in module_list["Module"]:
        print(module["TitleBlock"][0]["Description"])

with open("F:\Python Projects\MAN 2022\MAN\Output\Diagrame\JSON/Modules.json", "r") as json_file:
    loaded_dictionary = json.load(json_file)
#print_dictionary(loaded_dictionary)
extract_module_ids(loaded_dictionary)



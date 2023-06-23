from itertools import permutations, combinations


def print_combinations(dictionary):
    output = {}
    combined_groups = {}
    for kurzname, kurzname_group in dictionary.items():
        module_groups = kurzname_group.keys()
        module_combinations = []
        for i in range(1, len(module_groups)+1):
            module_combinations.extend(list(combinations(module_groups, i)))

        with open("filename", 'w') as file:
            for tpl in module_combinations:
                file.write(str(tpl) + '\n')
        for combination in module_combinations:
            combined_key = kurzname + "===" + ''.join(str(x) for x in combination)
            outlst = []
            duplicate_pin_no = False  # Flag for duplicate 'Pin No' values
            pin_no_values = set()
            for module in combination:
                for key, inner_dict in dictionary.items():
                    for inner_key, lst in inner_dict.items():
                        if module == inner_key:
                            for item in lst:
                                pin_no = item.get('Pin No')
                                if pin_no in pin_no_values:
                                    duplicate_pin_no = True
                                    break  # Break if duplicate 'Pin No' found
                                pin_no_values.add(pin_no)
                                for sub_key, value in item.items():
                                    outlst.append([sub_key, value])
                if duplicate_pin_no:
                    break  # Break if duplicate 'Pin No' found
            if not duplicate_pin_no:
                combined_groups[combined_key] = outlst
    output.update(combined_groups)

    return output

def get_ltg_pin_pairs(data):
    ltg_pin_pairs = []
    for key, values in data.items():
        pairs = []
        for i in range(0, len(values), 2):
            ltg_no = values[i][1]
            pin_no = values[i+1][1]
            pairs.append((ltg_no, pin_no))
        ltg_pin_pairs.append((key, pairs))
    return ltg_pin_pairs
def print_components_dict(components_dict):
    ltg_pin_pairs = get_ltg_pin_pairs(components_dict)
    for key, pairs in ltg_pin_pairs:
        print(key)
        for ltg_no, pin_no in pairs:
            print(f'Ltg No: {ltg_no}, Pin No: {pin_no}')
        print()


components_dict = {'X6490.2A1':
                       {'81.25484-5824':
                            [{'Ltg No': '723003_3', 'Pin No': '13'}],
                        '85.25480-6734':
                            [{'Ltg No': '431012_7', 'Pin No': '46'}, {'Ltg No': '431013_7', 'Pin No': '45'}],
                        '85.25480-6012':
                            [{'Ltg No': '600021_8', 'Pin No': '13'}],
                        '85.25480-6003':
                            [{'Ltg No': '601098_001', 'Pin No': '9'}, {'Ltg No': '601228_001', 'Pin No': '10'}]}}

permutations_dict = print_combinations(components_dict)
print(permutations_dict)
print_components_dict(permutations_dict)

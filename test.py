from itertools import product, combinations, permutations


def print_combinations(dictionary):
    combined_groups = {}
    for kurzname, kurzname_group in dictionary.items():
        module_groups = kurzname_group.keys()
        print(len(module_groups))
        module_combinations = []
        for i in range(2, len(module_groups)):
            module_combinations.extend(list(combinations(module_groups, i)))
        print(module_combinations)
        for combination in module_combinations:

            combined_key = kurzname + ''.join(str(x) for x in combination)
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
    for k, v in combined_groups.items():
        print(k, v)



# Example usage:
dictionary = {
    'XK.A1358.1': {
        '81.25481-6621': [{'Ltg No': '401011_2', 'Pin No': '1'}, {'Ltg No': '403034_2', 'Pin No': '2'}],
        '81.25480-5661': [{'Ltg No': '271057_1', 'Pin No': '3'}],
        '81.25480-5660': [{'Ltg No': '310000_83', 'Pin No': '4'}, {'Ltg No': '713001_9', 'Pin No': '5'}],
        '81.25480-5669': [{'Ltg No': '310000_83', 'Pin No': '4'}, {'Ltg No': '713501_9', 'Pin No': '1'}]
    },
    # Add more keys and subkeys as needed
}
dictionary2 ={
    'X2799.1A1': {
        '85.25480-6000': [{'Ltg No': '605007_001', 'Pin No': '30'}],
        '85.25480-6001': [{'Ltg No': '600003_001', 'Pin No': '28'}],
        '85.25480-6003': [{'Ltg No': '601097_001', 'Pin No': '19'} ,{'Ltg No': '601227_001', 'Pin No': '21'}],
        '85.25480-6006': [{'Ltg No': '150002_001', 'Pin No': '4'} ,{'Ltg No': '170000_001', 'Pin No': '25'}],
        '85.25480-6007': [{'Ltg No': '642009_001', 'Pin No': '6'}],
        '85.25480-6009': [{'Ltg No': '903002_003', 'Pin No': '37'} ,{'Ltg No': '903003_003', 'Pin No': '38'}],
        '85.25480-6013': [{'Ltg No': '901033_001', 'Pin No': '49'}],
        '85.25480-6015': [{'Ltg No': '191_024', 'Pin No': '33'} ,{'Ltg No': '192_024', 'Pin No': '32'}],
        '85.25480-6017': [{'Ltg No': '431004_002', 'Pin No': '41'} ,{'Ltg No': '431005_002', 'Pin No': '42'}],
        '85.25480-6019': [{'Ltg No': '903002_003', 'Pin No': '37'} ,{'Ltg No': '903003_003', 'Pin No': '38'}],
    },

}

print_combinations(dictionary2)
my_dict = {
    '032014_001_1': [200, 350],
    '032075_001_2': [200, 350],
    '032075_001_3': [200, 350],
    '032075_001_4': [200, 350],
    '032075_001_5': [200, 350]
}

search_key = "032014_001"

for key in my_dict:
    if search_key in key:
        print("Substring found in key:", key)
        break
else:
    print("Substring not found in any key")
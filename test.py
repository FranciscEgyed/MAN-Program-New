from fuzzywuzzy import fuzz
array_fuzz = []
array_fuzz_index = []
array = ["weissblau_201"]
for i in range(len(array)):
    array_fuzz.append(fuzz.ratio(array[i], "weiss_blau_201"))
    array_fuzz_index.append(i)

max_index = array_fuzz.index(max(array_fuzz))
print(max_index)
print(array[max_index])
print(fuzz.ratio("weiss-blau_101", "weissblau_101"))




import csv

with open("F:\Python Projects\MAN 2022\MAN\Input\Others/Tabel Klappschale.txt", newline='') as csvfile:
    arr_tabel_klappschale = list(csv.reader(csvfile, delimiter=';'))

print(arr_tabel_klappschale)
lista_platforme =list(set([x[0] for x in arr_tabel_klappschale]))
lista_klapps =list(set([x[2] for x in arr_tabel_klappschale]))
lista_side =list(set([x[3] for x in arr_tabel_klappschale]))

lista_calcul = []
print(lista_platforme)

print(lista_klapps)

print(lista_side)
array_combinatii_klapp = []
for side in lista_side:
    for platforma in lista_platforme:
        for klapp in lista_klapps:
            if side+platforma+klapp not in array_combinatii_klapp:
                array_combinatii_klapp.append(side+platforma+klapp)

print(array_combinatii_klapp)
for combinatie in array_combinatii_klapp:
    array_temp= []
    for linie in arr_tabel_klappschale:
        if linie[3] == combinatie[:3] and linie[4] == combinatie[3:7] and linie[2] == combinatie[7:]:
            array_temp.append(linie[1])
    array_temp.insert(0, combinatie)
    if len(array_temp) > 1:
        lista_calcul.append(array_temp)
print(lista_calcul)






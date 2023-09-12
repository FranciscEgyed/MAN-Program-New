import csv
import os
from tkinter import filedialog

file_load = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                       title="Incarcati fisierul cu informatiile sursa:")
with open(file_load, newline='') as csvfile:
    array_apfr = list(csv.reader(csvfile, delimiter='\t'))

array_delete = [
['\x0c+===================================================================================================================================+'],
['+===================================================================================================================================+'],
['|===================================================================================================================================|'],
['| product number    :    0000000000000     - 9999999999999                                                                          |'],
['|===================================================================================================================================|'],
['|external number               product number                      designation                                                      |'],
['|ressource group          designation ressource group                    total   unit                                               |'],
['|===================================================================================================================================|'],
[]
]
array_output = [["Client PN", "Leoni PN", "Designation", "Resource group", "Resignation resource group", "Total", "Unit"]]
print(len(array_apfr))
array_curatat = [line for line in array_apfr if line not in array_delete]
print(len(array_curatat))
array_curatat2 = [line for line in array_curatat if "" not in line]
array_curatat3 = [line for line in array_curatat2 if 'FAVLS' not in line[0]]

array_breakers = []
for i in range(len(array_curatat3)):
    if array_curatat3[i] == ['+-----------------------------------------------------------------------------------------------------------------------------------+']:
        array_breakers.append(i)

for i in range(0, 100):#len(array_breakers) - 1):
    array_output_temp = []
    for x in range(array_breakers[i] + 1, array_breakers[i + 1]-1):
        array_output_temp.append(array_curatat3[x])
    clientpn = [s.strip() for s in array_output_temp[0][0].split("   ") if s][0]
    leonipn = [s.strip() for s in array_output_temp[0][0].split("   ") if s][1]
    designation = [s.strip() for s in array_output_temp[0][0].split("   ") if s][2]
    first_three = [clientpn, leonipn, designation]
    for lista in array_output_temp[1:]:
        item_temp_list = [s.strip() for s in lista[0].split("   ") if s]
        merged_list = [x for x in first_three] + [x for x in item_temp_list]
        array_output.append(merged_list)
lista_rg = []
for p in array_output:
    if p[3] not in lista_rg:
        lista_rg.append(p[3])
print(lista_rg)




#for i in range(0, 50):
#    print(array_output[i])





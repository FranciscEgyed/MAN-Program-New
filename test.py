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
print(fuzz.ratio("weissblau201", "weiss_blau_201"))



def cmcsr():
    pbargui = Tk()
    pbargui.title("Control Matrix CSR")
    pbargui.geometry("500x50+50+550")
    pbar = ttk.Progressbar(pbargui, orient=HORIZONTAL, length=200, mode='indeterminate')
    statuslabel = Label(pbargui, text="Waiting . . .")
    pbar.grid(row=1, column=1, padx=5, pady=5)
    statuslabel.grid(row=1, column=2, padx=5, pady=5)
    listatwist = ["131_002", "131_102", "131_102", "grau_047"]
    fisier_cm = filedialog.askopenfilename(initialdir=os.path.abspath(os.curdir),
                                           title="Incarcati fisierul control matrix")
    if fisier_cm[-3:] == "csv":
        with open(fisier_cm, newline='') as csvfile:
            array_sortare = list(csv.reader(csvfile, delimiter=','))
        if array_sortare[0][5] == "8014":
            array_print = [["Module ID+REAL NAME", "KANBAN-AG", "REAL NAME", "Kanban name", "Module ID", "Ledset",
                            "Type", "Material PN", "Conector 1", "Pin 1", "Conector 2", "Pin 2"]]
            statuslabel["text"] = "Working on it . . . "
            pbar['value'] += 2
            pbargui.update_idletasks()
            for i in range(3, len(array_sortare)):
                for x in range(111, 2682):
                    if array_sortare[i][x] == "X" or array_sortare[i][x] == "x":
                        array_print.append([array_sortare[3][x] + array_sortare[i][11].lower(),
                                            array_sortare[i][47].replace("U", "W"),
                                            array_sortare[i][11].lower(), array_sortare[i][9], array_sortare[3][x],
                                            array_sortare[i][46], "FIR", array_sortare[i][13], array_sortare[i][91],
                                            array_sortare[i][92], array_sortare[i][93], array_sortare[i][94]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][x] == "Y" or array_sortare[i][x] == "y":
                        array_print.append([array_sortare[3][x] + array_sortare[i][11].lower(),
                                            array_sortare[i][47].replace("U", "W"),
                                            array_sortare[i][11].lower(), array_sortare[i][9], array_sortare[3][x],
                                            array_sortare[i][46], "OPERATIE", array_sortare[i][13],
                                            array_sortare[i][91],
                                            array_sortare[i][92], array_sortare[i][93], array_sortare[i][94]])
                    elif array_sortare[i][x] == "S" or array_sortare[i][x] == "s" \
                            and array_sortare[i][11] not in listatwist:
                        array_print.append([array_sortare[3][x] + array_sortare[i][11].lower(),
                                            array_sortare[i][47].replace("U", "W"),
                                            array_sortare[i][11].lower(), array_sortare[i][9], array_sortare[3][x],
                                            array_sortare[i][46], "COMPONENT", array_sortare[i][13],
                                            array_sortare[i][91], array_sortare[i][92], array_sortare[i][93],
                                            array_sortare[i][94]])
                        pbar['value'] += 2
                        pbargui.update_idletasks()
                    elif array_sortare[i][x] == "S" or array_sortare[i][x] == "s" \
                            and array_sortare[i][11] in listatwist:
                        array_print.append([array_sortare[3][x] + array_sortare[i][11].lower(),
                                            array_sortare[i + 4][47].replace("U", "W"),
                                            array_sortare[i][11].lower(), array_sortare[i + 4][9], array_sortare[3][x],
                                            array_sortare[i][46], "FIR", array_sortare[i][13],
                                            array_sortare[i][91], array_sortare[i][92], array_sortare[i][93],
                                            array_sortare[i][94]])

            with open(os.path.abspath(os.curdir) + "/MAN/Input/Others/Control_Matrix_CSR.txt", 'w', newline='',
                      encoding='utf-8') as myfile:
                wr = csv.writer(myfile, quoting=csv.QUOTE_ALL, delimiter=';')
                wr.writerows(array_print)
            pbar.destroy()
            pbargui.destroy()
            messagebox.showinfo('Finalizat!', "Finalizat.")
        else:
            pbar.destroy()
            pbargui.destroy()
            messagebox.showerror('Fisier gresit!', "Nu ati incarcat fisierul CSR")
    else:
        pbar.destroy()
        pbargui.destroy()
        messagebox.showerror('Extensie gresita!', "Incarcati fisierul CSV")

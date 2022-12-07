import tkinter as tk
from PIL import ImageTk, Image
from functii_input import *
from diverse import structura_directoare, file_checker
from functii_prelucrare import *
from functii_prelucrare_ksk import *


def statusbusy():
    statuslabel["text"] = "Working on it . . . "


def statusidle():
    statuslabel["text"] = "Finished."


root = tk.Tk()
root.title("2022 MAN file processor")
root.geometry("450x450+50+50")
root.iconbitmap("Img/ICON.ico")
img = ImageTk.PhotoImage(Image.open("Img/MAN.jpg"))
container = tk.Frame(root, bg="black")
container.grid_rowconfigure(0, weight=0)
container.grid_columnconfigure(0, weight=1)
container = tk.Frame(root)
container.grid_rowconfigure(0, weight=0)
container.grid_columnconfigure(0, weight=1)
label = tk.Label(container, image=img)
menu_frame = tk.Frame(container)
statuslabel = tk.Label(root, text="Waiting . . .")

menu1 = tk.Menubutton(menu_frame, text="Fisiere Sursa", background="grey", font=("Arial", 10))
menu1.grid(row=0, column=0)
submenu1 = tk.Menu(menu1, tearoff=0, background="grey", font=("Arial", 10))
submenu1.add_command(label="Incarcare Input File", command=lambda: [statusbusy(), load_source(), statusidle()])
submenu1.add_separator()
submenu1.add_command(label="Control Matrix CSR", command=lambda: [statusbusy(), cmcsr(), statusidle()])
submenu1.add_command(label="Control Matrix CSL", command=lambda: [statusbusy(), cmcsl(), statusidle()])
submenu1.add_command(label="Control Matrix TGLM L", command=lambda: [statusbusy(), cmtglml(), statusidle()])
submenu1.add_command(label="Control Matrix TGLM R", command=lambda: [statusbusy(), cmtglmr(), statusidle()])
submenu1.add_command(label="Control Matrix 4AXEL L(8025)", command=lambda: [statusbusy(), cm4axell(), statusidle()])
submenu1.add_command(label="Control Matrix 4AXEL R(8026)", command=lambda: [statusbusy(), cm4axelr(), statusidle()])
submenu1.add_separator()
submenu1.add_command(label="Control Matrix Super Sleeve - in development",
                     command=lambda: [statusbusy(), cmss(), statusidle()])
menu1.configure(menu=submenu1)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu2 = tk.Menubutton(menu_frame, text="Fisiere Input", background="gray", font=("Arial", 10))
menu2.grid(row=0, column=1)
submenu2 = tk.Menu(menu2, tearoff=0, background="grey", font=("Arial", 10))
submenu2.add_command(label="Sortare JIT(SAP)", command=lambda: [statusbusy(), sortare_jit(), statusidle()])
submenu2.add_command(label="Sortare JIT(SAP) din director",
                     command=lambda: [statusbusy(), sortare_jit_dir(), statusidle()])
submenu2.add_command(label="Golire directoare LHD si RHD",
                     command=lambda: [statusbusy(), golire_directoare_comparati(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Prelucrare BOM-uri", command=lambda: [statusbusy(), boms(), statusidle()])
submenu2.add_command(label="Prelucrare WIRELIST-uri", command=lambda: [statusbusy(), wires(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Prelucrare BOM-uri cu PN Leoni",
                     command=lambda: [statusbusy(), boms_leoni(), statusidle()])
submenu2.add_command(label="Prelucrare WIRELIST-uri cu PN Leoni",
                     command=lambda: [statusbusy(), wires_leoni(), statusidle()])
menu2.configure(menu=submenu2)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu3 = tk.Menubutton(menu_frame, text="Prelucrare KSK", background="grey", font=("Arial", 10))
menu3.grid(row=0, column=2)
submenu3 = tk.Menu(menu3, tearoff=0, background="grey", font=("Arial", 10))
submenu3.add_command(label="Wirelist individual",
                     command=lambda: [statusbusy(), wirelist_individual(), statusidle()])
submenu3.add_command(label="Wirelist toate", command=lambda: [statusbusy(), wirelist_director(), statusidle()])
submenu3.add_separator()
submenu3.add_command(label="BOM individual", command=lambda: [statusbusy(), prelucrare_individuala_bom(), statusidle()])
submenu3.add_command(label="BOM toate", command=lambda: [statusbusy(), bom_director(), statusidle()])
menu3.configure(menu=submenu3)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu4 = tk.Menubutton(menu_frame, text="Rapoarte", background="grey", font=("Arial", 10))
menu4.grid(row=0, column=3)
submenu4 = tk.Menu(menu4, tearoff=0, background="grey", font=("Arial", 10))
submenu4.add_command(label="Raport", command=lambda: [statusbusy(), creare_raport(), statusidle()])
submenu4.add_command(label="Rapoarte din director",
                     command=lambda: [statusbusy(), creare_raport_director(), statusidle()])
submenu4.add_command(label="Rapoarte toate", command=lambda: [statusbusy(), creare_raport_all(), statusidle()])
submenu4.add_separator()
submenu4.add_command(label="Extragere lungimi KSK",
                     command=lambda: [statusbusy(), extragere_lungimi_ksk(), statusidle()])
submenu4.add_separator()
submenu4.add_command(label="Extragere BOM KSK", command=lambda: [statusbusy(), extragere_bom_ksk(), statusidle()])
submenu4.add_separator()
submenu4.add_command(label="Extragere Variatii de lungimi",
                     command=lambda: [statusbusy(), extragere_variatii(), statusidle()])
menu4.configure(menu=submenu4)








menu_frame.grid(row=0, column=0)
container.grid(row=1, column=0)
label.grid(row=2, column=0)
statuslabel.grid(row=12, column=0)
structura_directoare()
file_checker()
root.mainloop()

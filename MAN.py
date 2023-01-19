import tkinter as tk
from PIL import ImageTk, Image
from functii_database import databasecontent
from functii_diverse import *
from functii_input import *
from diverse import *
from functii_ksklight import *
from functii_prelucrare import *
from functii_prelucrare_ksk import *
from functii_rapoarte import *
from masterdata import *


def statusbusy():
    statuslabel["text"] = "Working on it . . . "


def statusidle():
    statuslabel["text"] = "Finished."


root = tk.Tk()
root.title("2022 MAN file processor")
root.geometry("500x450+50+50")
root.iconbitmap("Img/ICON.ico")
img = ImageTk.PhotoImage(Image.open("Img/MAN.jpg"))
container = tk.Frame(root, bg="gray")
container.grid_rowconfigure(0, weight=0)
container.grid_columnconfigure(0, weight=1)
container = tk.Frame(root)
container.grid_rowconfigure(0, weight=0)
container.grid_columnconfigure(0, weight=1)
label = tk.Label(container, image=img)
menu_frame = tk.Frame(container)
statuslabel = tk.Label(root, text="Waiting . . .")

menu1 = tk.Menubutton(menu_frame, text="Fisiere Sursa", background="DarkSeaGreen1", font="Arial 10 bold")
menu1.grid(row=0, column=0)
submenu1 = tk.Menu(menu1, tearoff=0, background="DarkSeaGreen1", font="Arial 10 bold")
submenu1.add_command(label="Incarcare Fisier Sursa", command=lambda: [statusbusy(), load_source(), statusidle()])
submenu1.add_separator()
submenu1.add_command(label="Control Matrix CSR", command=lambda: [statusbusy(), cmcsrnew(), statusidle()])
submenu1.add_command(label="Control Matrix CSL", command=lambda: [statusbusy(), cmcslnew(), statusidle()])
submenu1.add_command(label="Control Matrix TGLM L", command=lambda: [statusbusy(), cmtglmlnew(), statusidle()])
submenu1.add_command(label="Control Matrix TGLM R", command=lambda: [statusbusy(), cmtglmrnew(), statusidle()])
submenu1.add_command(label="Control Matrix 4AXEL L(8025)", command=lambda: [statusbusy(), cm4axellnew(), statusidle()])
submenu1.add_command(label="Control Matrix 4AXEL R(8026)", command=lambda: [statusbusy(), cm4axelrnew(), statusidle()])
submenu1.add_separator()
submenu1.add_command(label="Control Matrix Super Sleeve - in development",
                     command=lambda: [statusbusy(), cmss(), statusidle()])
submenu1.add_separator()
menu1.configure(menu=submenu1)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu2 = tk.Menubutton(menu_frame, text="Fisiere Input", background="DarkSeaGreen1", font="Arial 10 bold")
menu2.grid(row=0, column=1)
submenu2 = tk.Menu(menu2, tearoff=0, background="DarkSeaGreen1", font="Arial 10 bold")
submenu2.add_command(label="Sortare JIT(SAP)", command=lambda: [statusbusy(), sortare_jit(), statusidle()])
submenu2.add_command(label="Sortare JIT(SAP) din director",
                     command=lambda: [statusbusy(), sortare_jit_dir(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Incarcare BOM-uri", command=lambda: [statusbusy(), boms(), statusidle()])
submenu2.add_command(label="Incarcare WIRELIST-uri", command=lambda: [statusbusy(), wires(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Prelucrare BOM-uri cu PN Leoni",
                     command=lambda: [statusbusy(), boms_leoni(), statusidle()])
submenu2.add_command(label="Prelucrare WIRELIST-uri cu PN Leoni",
                     command=lambda: [statusbusy(), wires_leoni(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Creare wirelist toate platformele simplu",
                     command=lambda: [statusbusy(), wirelist_all_simplu(), statusidle()])
submenu2.add_command(label="Creare wirelist toate platformele complet",
                     command=lambda: [statusbusy(), wires_leoni(), statusidle()])
menu2.configure(menu=submenu2)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu3 = tk.Menubutton(menu_frame, text="Prelucrare KSK", background="DarkSeaGreen1", font="Arial 10 bold")
menu3.grid(row=0, column=2)
submenu3 = tk.Menu(menu3, tearoff=0, background="DarkSeaGreen1", font="Arial 10 bold")
submenu3.add_command(label="Wirelist individual",
                     command=lambda: [statusbusy(), wirelist_individual(), statusidle()])
submenu3.add_command(label="Wirelist toate", command=lambda: [statusbusy(), wirelist_director(), statusidle()])
submenu3.add_separator()
submenu3.add_command(label="BOM individual", command=lambda: [statusbusy(), prelucrare_individuala_bom(), statusidle()])
submenu3.add_command(label="BOM toate", command=lambda: [statusbusy(), bom_director(), statusidle()])
menu3.configure(menu=submenu3)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu4 = tk.Menubutton(menu_frame, text="Rapoarte", background="DarkSeaGreen1", font="Arial 10 bold")
menu4.grid(row=0, column=3)
submenu4 = tk.Menu(menu4, tearoff=0, background="DarkSeaGreen1", font="Arial 10 bold")
submenu4.add_command(label="Raport", command=lambda: [statusbusy(), creare_raport(), statusidle()])
submenu4.add_command(label="Rapoarte din director",
                     command=lambda: [statusbusy(), creare_raport_director(), statusidle()])
submenu4.add_command(label="Rapoarte toate", command=lambda: [statusbusy(), creare_raport_all(), statusidle()])
menu4.configure(menu=submenu4)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu5 = tk.Menubutton(menu_frame, text="KSK Light", background="DarkSeaGreen1", font="Arial 10 bold")
menu5.grid(row=0, column=4)
submenu5 = tk.Menu(menu5, tearoff=0, background="DarkSeaGreen1", font="Arial 10 bold")
submenu5.add_command(label="Raport KSK Light", command=lambda: [statusbusy(), raport_light(), statusidle()])
#submenu5.add_command(label="Comparatie KSK Light", command=lambda: [statusbusy(), compare_ksk_light(), statusidle()])
submenu5.add_separator()
submenu5.add_command(label="Lista taiere KSK Light", command=lambda: [statusbusy(), cutting_ksklight(), statusidle()])
submenu5.add_separator()
submenu5.add_command(label="Lista taiere forecast module",
                     command=lambda: [statusbusy(), cutting_ksklight_module(), statusidle()])
submenu5.add_separator()
# submenu5.add_command(label="Lista SuperSleeve KSK Light",
#                     command=lambda: [statusbusy(), ss_ksklight(), statusidle()])
menu5.configure(menu=submenu5)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu7 = tk.Menubutton(menu_frame, text="Diverse", background="DarkSeaGreen1", font="Arial 10 bold")
menu7.grid(row=0, column=5)
submenu7 = tk.Menu(menu7, tearoff=0, background="DarkSeaGreen1", font="Arial 10 bold")
submenu7.add_command(label="Extragere lungimi KSK",
                     command=lambda: [statusbusy(), extragere_lungimi_ksk(), statusidle()])
submenu7.add_command(label="Extragere BOM KSK", command=lambda: [statusbusy(), extragere_bom_ksk(), statusidle()])
submenu7.add_command(label="Extragere Variatii de lungimi",
                     command=lambda: [statusbusy(), extragere_variatii(), statusidle()])
submenu7.add_separator()
submenu7.add_command(label="Prelucrare masterdata",
                     command=lambda: [statusbusy(), inlocuire(), statusidle()])
submenu7.add_command(label="Database", command=lambda: [statusbusy(), databasecontent(), statusidle()])
submenu7.add_separator()
submenu7.add_command(label="Stergere fisiere", command=golire_directoare)
submenu7.add_separator()
menu7.configure(menu=submenu7)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


menu_frame.grid(row=0, column=0)
container.grid(row=1, column=0)
label.grid(row=2, column=0)
statuslabel.grid(row=12, column=0)
structura_directoare()
file_checker()
databesemerge()
databasecopy()
root.mainloop()

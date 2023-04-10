import tkinter as tk
from PIL import ImageTk, Image
from diagrame import *
from functii_database import databasecontent, exportdatabase, database_delete_record
from functii_diverse import *
from functii_input import *
from diverse import *
from functii_ksklight import *
from functii_ldorado import *
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
root.geometry("600x450+50+50")
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
submenu1 = tk.Menu(menu1, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu1.add_command(label="Incarcare Fisier Sursa", command=lambda: [statusbusy(), load_source(), statusidle()])
submenu1.add_separator()
submenu1.add_command(label="Control Matrix CSR(8012)", command=lambda: [statusbusy(), cmcsrnew(), statusidle()])
submenu1.add_command(label="Control Matrix CSL(8011)", command=lambda: [statusbusy(), cmcslnew(), statusidle()])
submenu1.add_command(label="Control Matrix TGLM L(8023)", command=lambda: [statusbusy(), cmtglml(), statusidle()])
submenu1.add_command(label="Control Matrix TGLM R(8024)", command=lambda: [statusbusy(), cmtglmr(), statusidle()])
submenu1.add_command(label="Control Matrix 4AXEL L(8025)", command=lambda: [statusbusy(), cm4axell(), statusidle()])
submenu1.add_command(label="Control Matrix 4AXEL R(8026)", command=lambda: [statusbusy(), cm4axelr(), statusidle()])
submenu1.add_separator()
submenu1.add_command(label="Control Matrix Super Sleeve - in development",
                     command=lambda: [statusbusy(), cmss(), statusidle()])
submenu1.add_separator()
menu1.configure(menu=submenu1)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu2 = tk.Menubutton(menu_frame, text="Fisiere Input", background="DarkSeaGreen1", font="Arial 10 bold")
menu2.grid(row=0, column=1)
submenu2 = tk.Menu(menu2, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu2.add_command(label="Sortare JIT(SAP)", command=lambda: [statusbusy(), sortare_jit(), statusidle()])
submenu2.add_command(label="Sortare JIT(SAP) din director",
                     command=lambda: [statusbusy(), sortare_jit_dir(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Incarcare BOM-uri", command=lambda: [statusbusy(), boms(), statusidle()])
submenu2.add_command(label="Incarcare WIRELIST-uri", command=lambda: [statusbusy(), wires(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Creare Wirelist complet", command=lambda: [statusbusy(), wires_complet(), statusidle()])
submenu2.add_command(label="Creare Wirelist complet cu PN Leoni",
                     command=lambda: [statusbusy(), wires_pnleoni(), statusidle()])
submenu2.add_command(label="Creare BOM cu PN Leoni",
                     command=lambda: [statusbusy(), boms_pnleoni(), statusidle()])
submenu2.add_separator()
submenu2.add_command(label="Creare BOM toate platformele",
                     command=lambda: [statusbusy(), boms_cumulat(), statusidle()])
submenu2.add_command(label="Creare Wirelist toate platformele",
                     command=lambda: [statusbusy(), wires_cumulat(), statusidle()])
menu2.configure(menu=submenu2)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu3 = tk.Menubutton(menu_frame, text="Prelucrare KSK", background="DarkSeaGreen1", font="Arial 10 bold")
menu3.grid(row=0, column=2)
submenu3 = tk.Menu(menu3, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
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
submenu4 = tk.Menu(menu4, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu4.add_command(label="Raport", command=lambda: [statusbusy(), creare_raport(), statusidle()])
submenu4.add_command(label="Rapoarte din director",
                     command=lambda: [statusbusy(), creare_raport_director(), statusidle()])
submenu4.add_command(label="Rapoarte toate", command=lambda: [statusbusy(), creare_raport_all(), statusidle()])
menu4.configure(menu=submenu4)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu5 = tk.Menubutton(menu_frame, text="KSK Light", background="DarkSeaGreen1", font="Arial 10 bold")
menu5.grid(row=0, column=4)
submenu5 = tk.Menu(menu5, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu5.add_command(label="Raport KSK Light", command=lambda: [statusbusy(), raport_light(), statusidle()])
submenu5.add_separator()
submenu5.add_command(label="Lista taiere KSK Light", command=lambda: [statusbusy(), cutting_ksklight(), statusidle()])
submenu5.add_command(label="Lista SuperSleeve KSK Light", command=lambda: [statusbusy(), ss_ksklight(), statusidle()])
submenu5.add_separator()
menu5.configure(menu=submenu5)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++a
menu7 = tk.Menubutton(menu_frame, text="Diverse", background="DarkSeaGreen1", font="Arial 10 bold")
menu7.grid(row=0, column=5)
submenu7 = tk.Menu(menu7, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenucascade = tk.Menu(submenu7, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenucascade2 = tk.Menu(submenu7, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu7.add_command(label="Extragere lungimi KSK",
                     command=lambda: [statusbusy(), extragere_lungimi_ksk(), statusidle()])
submenu7.add_command(label="Extragere BOM KSK", command=lambda: [statusbusy(), extragere_bom_ksk(), statusidle()])
submenu7.add_command(label="Extragere Variatii de lungimi",
                     command=lambda: [statusbusy(), extragere_variatii(), statusidle()])
submenu7.add_separator()
submenu7.add_command(label="Prelucrare masterdata",
                     command=lambda: [statusbusy(), inlocuire_masterdata(), statusidle()])
submenu7.add_separator()
submenu7.add_command(label="Stergere fisiere", command=golire_directoare, background="red")
submenu7.add_separator()

submenu7.add_cascade(label="Diagrame . . . ", menu=submenucascade)
submenucascade.add_command(label="Comparatie diagrame", command=lambda: [statusbusy(),
                                                                         comparatiediagrame(), statusidle()])
submenucascade.add_command(label="Asociere diagrame cu module din matrix",
                           command=lambda: [statusbusy(), asocierediagramemodule(), statusidle()])
submenucascade.add_command(label="Indexare diagrame dupa matrix",
                           command=lambda: [statusbusy(), indexarediagrame(), statusidle()])
submenucascade.add_command(label="Extragere diagrame pentru KSK",
                           command=lambda: [statusbusy(), diagrameinksk(), statusidle()])
submenucascade.add_command(label="Extragere diagrame pentru KSK din folder",
                           command=lambda: [statusbusy(), diagrameinkskfolder(), statusidle()])
submenucascade.add_command(label="Extragere informatii din diagrame",
                           command=lambda: [statusbusy(), extragere_informatii_diagrame(), statusidle()])

submenu7.add_cascade(label="LDorado . . . ", menu=submenucascade2)
submenucascade2.add_command(label="segment_test", command=lambda: [statusbusy(),
                                                                         segment_test(), statusidle()])
menu7.configure(menu=submenu7)
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
menu8 = tk.Menubutton(menu_frame, text="Database", background="DarkSeaGreen1", font="Arial 10 bold")
menu8.grid(row=0, column=6)
submenu8 = tk.Menu(menu8, tearoff=0, background="DarkSeaGreen1", font="Arial 15 bold")
submenu8.add_command(label="Extragere KSK din database", command=lambda: [statusbusy(), databasecontent(), statusidle()])
submenu8.add_separator()
submenu8.add_command(label="Incarcare stock pentru comparatie", command=stockcompa)
submenu8.add_separator()
submenu8.add_command(label="Extragere database completa", command=lambda: [statusbusy(), exportdatabase(), statusidle()])
submenu8.add_separator()
submenu8.add_command(label="Stergere inregistrari din database",
                     command=lambda: [statusbusy(), database_delete_record(), statusidle()], background="red")
submenu8.add_separator()
menu8.configure(menu=submenu8)

menu_frame.grid(row=0, column=0)

container.grid(row=1, column=0)
label.grid(row=2, column=0)
statuslabel.grid(row=12, column=0)
databasebackup()
structura_directoare()
file_checker()
databesemerge()
root.mainloop()

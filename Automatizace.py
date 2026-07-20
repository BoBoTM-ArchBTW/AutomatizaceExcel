import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd

# Globální proměnné pro cesty a listy
cesta_zakaznik = ""
cesta_cil = ""

# Seznam pro uložení ovládacích prvků ze strany 2
seznam_operaci = []

# Seznamy sloupců, které naplní Pandas
sloupce_zakaznik = []
sloupce_nase = []


# --- AUTOMATICKÝ DETEKTOR ENGINŮ S FALLBACKEM ---
def nacist_listy(cesta_souboru):
    enginy = ['openpyxl', 'xlrd', 'pyxlsb', 'calamine']
    for eng in enginy:
        try:
            xl = pd.ExcelFile(cesta_souboru, engine=eng)
            return xl.sheet_names
        except Exception:
            continue
    try:
        xl = pd.ExcelFile(cesta_souboru)
        return xl.sheet_names
    except Exception as e:
        messagebox.showerror("Chyba formátu", f"Nepodařilo se otevřít Excel:\n{str(e)}")
        return []


def nacist_sloupce_bezpecne(cesta, list_name):
    enginy = ['openpyxl', 'xlrd', 'pyxlsb', 'calamine']
    for eng in enginy:
        try:
            df = pd.read_excel(cesta, sheet_name=list_name, nrows=0, engine=eng)
            return list(df.columns)
        except Exception:
            continue
    try:
        df = pd.read_excel(cesta, sheet_name=list_name, nrows=0)
        return list(df.columns)
    except Exception:
        return []


# --- AKCE NA PRVNÍ STRÁNCE ---
def vybrat_zakaznika():
    global cesta_zakaznik
    soubor = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
    if soubor:
        cesta_zakaznik = soubor
        lbl_zakaznik.config(text=os.path.basename(soubor))
        
        listy = nacist_listy(soubor)
        cb_list_zakaznik['values'] = listy
        if listy:
            cb_list_zakaznik.set(listy[0])
            frame_list_zakaznik.pack(pady=(0, 15))
            aktualizovat_parovani_zakaznika()

def aktualizovat_parovani_zakaznika(event=None):
    global sloupce_zakaznik
    sloupce_zakaznik = nacist_sloupce_bezpecne(cesta_zakaznik, cb_list_zakaznik.get())
    cb_klic_zakaznik['values'] = sloupce_zakaznik
    if sloupce_zakaznik:
        cb_klic_zakaznik.set(sloupce_zakaznik[0])
        zobrazit_frame_parovani()


def vybrat_cil():
    global cesta_cil
    soubor = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
    if soubor:
        cesta_cil = soubor
        lbl_cil.config(text=os.path.basename(soubor))
        
        listy = nacist_listy(soubor)
        cb_list_cil['values'] = listy
        if listy:
            cb_list_cil.set(listy[0])
            frame_list_cil.pack(pady=(0, 15))
            aktualizovat_parovani_cile()

def aktualizovat_parovani_cile(event=None):
    global sloupce_nase
    sloupce_nase = nacist_sloupce_bezpecne(cesta_cil, cb_list_cil.get())
    cb_klic_cil['values'] = sloupce_nase
    if sloupce_nase:
        cb_klic_cil.set(sloupce_nase[0])
        zobrazit_frame_parovani()


def zobrazit_frame_parovani():
    # Zobrazí panel pro výběr párovacích klíčů, pokud jsou načtené oba soubory
    if cesta_zakaznik and cesta_cil:
        frame_parovani.pack(pady=(15, 15))


def prejit_na_operace():
    if not cesta_zakaznik or not cesta_cil:
        messagebox.showwarning("Chyba", "Musíš vybrat soubor od zákazníka i váš cílový soubor!")
        return

    frame_strana1.pack_forget()
    frame_strana2.pack(fill=tk.BOTH, expand=True)
    pridat_radek_operace()


# --- STRÁNKA 2: DYNAMICKÝ SEZNAM ---
def pridat_radek_operace():
    radek_frame = ttk.Frame(scrollable_frame)
    radek_frame.pack(fill=tk.X, pady=5)
    
    cb_poradi = ttk.Combobox(radek_frame, values=[str(i) for i in range(1, 31)], width=3)
    cb_poradi.set(str(len(seznam_operaci) + 1))
    cb_poradi.pack(side=tk.LEFT, padx=5)
    
    cb_operace = ttk.Combobox(radek_frame, values=["Kopírovat", "Sečíst", "Vydělit"], width=12, state="readonly")
    cb_operace.set("Kopírovat")
    cb_operace.pack(side=tk.LEFT, padx=5)
    
    ttk.Label(radek_frame, text="Z:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(10, 2))
    cb_odkud = ttk.Combobox(radek_frame, values=sloupce_zakaznik, width=25, state="readonly")
    if sloupce_zakaznik: cb_odkud.set(sloupce_zakaznik[0])
    cb_odkud.pack(side=tk.LEFT, padx=5)
    
    ttk.Label(radek_frame, text="Do:", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=(20, 2))
    cb_kam = ttk.Combobox(radek_frame, values=sloupce_nase, width=25, state="readonly")
    if sloupce_nase: cb_kam.set(sloupce_nase[0])
    cb_kam.pack(side=tk.LEFT, padx=5)
    
    btn_smazat = ttk.Button(radek_frame, text="✕", width=3, command=lambda: smazat_radek_operace(radek_frame, data_radku))
    btn_smazat.pack(side=tk.RIGHT, padx=5)
    
    data_radku = {
        "frame": radek_frame, "poradi": cb_poradi, "operace": cb_operace,
        "odkud": cb_odkud, "kam": cb_kam
    }
    seznam_operaci.append(data_radku)
    
    canvas.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))


def smazat_radek_operace(frame, data_radku):
    frame.destroy()
    seznam_operaci.remove(data_radku)
    for index, radek in enumerate(seznam_operaci):
        radek["poradi"].set(str(index + 1))


# --- FINÁLNÍ MATICE A VÝPOČET ---
def spustit_konverzi():
    try:
        df_nase = None
        df_zakaznik = None
        
        for eng in ['openpyxl', 'xlrd', 'pyxlsb', 'calamine', None]:
            try:
                if df_nase is None: df_nase = pd.read_excel(cesta_cil, sheet_name=cb_list_cil.get(), engine=eng)
            except Exception: pass
            try:
                if df_zakaznik is None: df_zakaznik = pd.read_excel(cesta_zakaznik, sheet_name=cb_list_zakaznik.get(), engine=eng)
            except Exception: pass

        if df_nase is None or df_zakaznik is None:
            messagebox.showerror("Chyba", "Nepodařilo se načíst data z tabulek.")
            return

        # Získání navolených klíčů pro párování z 1. stránky
        klic_zak = cb_klic_zakaznik.get()
        klic_nas = cb_klic_cil.get()

        serazene_operace = sorted(seznam_operaci, key=lambda x: int(x["poradi"].get() if x["poradi"].get().isdigit() else 99))
        
        # Klíčový moment: Spojení (merge) přesně podle vybraných sloupců s čísly součástek!
        vysledek = pd.merge(df_nase, df_zakaznik, left_on=klic_nas, right_on=klic_zak, how="left")

        for radek in serazene_operace:
            op = radek["operace"].get()
            odkud = radek["odkud"].get()
            kam = radek["kam"].get()
            
            if not odkud or not kam: continue
            
            if op == "Kopírovat":
                vysledek[kam] = vysledek[odkud]
            elif op == "Sečíst":
                vysledek[kam] = pd.to_numeric(vysledek[kam], errors='coerce').fillna(0) + pd.to_numeric(vysledek[odkud], errors='coerce').fillna(0)
            elif op == "Vydělit":
                vysledek[kam] = pd.to_numeric(vysledek[kam], errors='coerce') / pd.to_numeric(vysledek[odkud], errors='coerce').replace(0, 1)

        # Ořízneme přebytečné sloupce od zákazníka, co se nabalily při merge, ať zůstane jen čistá struktura cíle
        vysledek = vysledek[df_nase.columns]

        is_xlsm = cesta_cil.lower().endswith('.xlsm')
        with pd.ExcelWriter(cesta_cil, engine='openpyxl', engine_kwargs={'keep_vba': True} if is_xlsm else {}) as writer:
            vysledek.to_excel(writer, index=False, sheet_name=cb_list_cil.get())

        messagebox.showinfo("Hotovo", "Všechny kroky úspěšně proběhly a data byla uložena!")
        root.destroy()

    except Exception as e:
        messagebox.showerror("Chyba při zápisu", f"Něco kleklo:\n{str(e)}")


def navrat_zpet():
    for radek in seznam_operaci: radek["frame"].destroy()
    seznam_operaci.clear()
    frame_strana2.pack_forget()
    frame_strana1.pack(fill=tk.BOTH, expand=True)


# ================= HLAVNÍ OKNO =================
root = tk.Tk()
root.title("Automatizace tabulek")
root.geometry("850x600")

style = ttk.Style()
style.theme_use('vista')

# ================= STRÁNKA 1 =================
frame_strana1 = ttk.Frame(root, padding="20")
frame_strana1.pack(fill=tk.BOTH, expand=True)

# 1. Zákazník
ttk.Label(frame_strana1, text="1. Vyber soubor od zákazníka", font=("Segoe UI", 10, "bold")).pack(pady=(5, 5))
ttk.Button(frame_strana1, text="Procházet zákazníka...", command=vybrat_zakaznika).pack()
lbl_zakaznik = ttk.Label(frame_strana1, text="Není vybráno", font=("Segoe UI", 9, "italic"))
lbl_zakaznik.pack(pady=(0, 5))

frame_list_zakaznik = ttk.Frame(frame_strana1)
ttk.Label(frame_list_zakaznik, text="Vyber list zákazníka:").pack(side=tk.LEFT, padx=5)
cb_list_zakaznik = ttk.Combobox(frame_list_zakaznik, width=25, state="readonly")
cb_list_zakaznik.pack(side=tk.LEFT)
cb_list_zakaznik.bind("<<ComboboxSelected>>", aktualizovat_parovani_zakaznika)

# 2. Váš cílový soubor
ttk.Label(frame_strana1, text="2. Vyber váš soubor (kam uložit výsledek)", font=("Segoe UI", 10, "bold")).pack(pady=(15, 5))
ttk.Button(frame_strana1, text="Vybrat cíl...", command=vybrat_cil).pack()
lbl_cil = ttk.Label(frame_strana1, text="Není vybráno", font=("Segoe UI", 9, "italic"))
lbl_cil.pack(pady=(0, 5))

frame_list_cil = ttk.Frame(frame_strana1)
ttk.Label(frame_list_cil, text="Vyber cílový list:").pack(side=tk.LEFT, padx=5)
cb_list_cil = ttk.Combobox(frame_list_cil, width=25, state="readonly")
cb_list_cil.pack(side=tk.LEFT)
cb_list_cil.bind("<<ComboboxSelected>>", aktualizovat_parovani_cile)

# --- DYNAMICKÝ PANEL PRO PÁROVÁNÍ ---
frame_parovani = ttk.LabelFrame(frame_strana1, text=" Nastavení párování řádků (Čísla součástek) ", padding="10")
# Balí se automaticky až po načtení obou souborů

ttk.Label(frame_parovani, text="Sloupec s klíčem u zákazníka:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
cb_klic_zakaznik = ttk.Combobox(frame_parovani, width=30, state="readonly")
cb_klic_zakaznik.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(frame_parovani, text="Sloupec s klíčem u nás:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
cb_klic_cil = ttk.Combobox(frame_parovani, width=30, state="readonly")
cb_klic_cil.grid(row=1, column=1, padx=5, pady=5)

# Tlačítko Pokračovat
ttk.Button(frame_strana1, text="Pokračovat ➔", command=prejit_na_operace).pack(fill=tk.X, side=tk.BOTTOM, ipady=7)

# ================= STRÁNKA 2 =================
frame_strana2 = ttk.Frame(root, padding="15")

ttk.Label(frame_strana2, text="Nastavení kroků automatizace", font=("Segoe UI", 11, "bold")).pack(anchor=tk.W, pady=(0, 10))

canvas = tk.Canvas(frame_strana2, borderwidth=0, highlightthickness=0)
scrollbar = ttk.Scrollbar(frame_strana2, orient="vertical", command=canvas.yview)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

spodek_frame = ttk.Frame(frame_strana2)
spodek_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))

ttk.Button(spodek_frame, text="+ Přidat krok", command=pridat_radek_operace).pack(anchor=tk.W, pady=(0, 15))

nav_frame = ttk.Frame(spodek_frame)
nav_frame.pack(fill=tk.X)

ttk.Button(nav_frame, text="⮌ Zpět", command=navrat_zpet).pack(side=tk.LEFT, ipady=5)
ttk.Button(nav_frame, text="Spustit konverzi", command=spustit_konverzi).pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0), ipady=5)

root.mainloop()
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import numpy as np

# Globální proměnné pro cesty a listy
cesta_zakaznik = ""
cesta_cil = ""

# Seznam pro uložení ovládacích prvků ze strany 2
seznam_operaci = []

# Seznam sloupců zákazníka, který naplní Pandas
sloupce_zakaznik = []


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


def vybrat_cil():
    global cesta_cil
    # Vynutíme ukládání pouze do standardního formátu .xlsx, aby to Excel neshazoval
    soubor = filedialog.asksaveasfilename(
        filetypes=[("Excel (.xlsx)", "*.xlsx")],
        defaultextension=".xlsx",
        initialdir=os.path.dirname(cesta_zakaznik) if cesta_zakaznik else None
    )
    if soubor:
        cesta_cil = soubor
        lbl_cil.config(text=os.path.basename(soubor))


def prejit_na_operace():
    global sloupce_zakaznik
    if not cesta_zakaznik or not cesta_cil:
        messagebox.showwarning("Chyba", "Musíš vybrat soubor od zákazníka i kam uložit výsledek!")
        return
    
    sloupce_zakaznik = nacist_sloupce_bezpecne(cesta_zakaznik, cb_list_zakaznik.get())
    
    if not sloupce_zakaznik:
        messagebox.showerror("Chyba", "Nepodařilo se načíst sloupce ze souboru.")
        return

    frame_strana1.pack_forget()
    frame_strana2.pack(fill=tk.BOTH, expand=True)
    pridat_radek_operace()


# --- STRÁNKA 2 ---
def dynamic_ui_zmeny(event, cb_operace, frame_z, frame_do, cb_odkud, cb_kam):
    op = cb_operace.get()
    if op == "Sečíst duplicitní řádky":
        frame_z.config(text=" Podle kterého sloupce sloučit? ")
        frame_do.pack_forget()
        cb_odkud['values'] = sloupce_zakaznik
    elif op == "Přičíst sloupec k jinému":
        frame_z.config(text=" Zdrojový sloupec (S): ")
        frame_do.pack(side=tk.LEFT, padx=5)
        cb_odkud['values'] = sloupce_zakaznik
        cb_kam['values'] = sloupce_zakaznik


def pridat_radek_operace():
    radek_frame = ttk.Frame(scrollable_frame)
    radek_frame.pack(fill=tk.X, pady=5)
    
    cb_poradi = ttk.Combobox(radek_frame, values=[str(i) for i in range(1, 31)], width=3)
    cb_poradi.set(str(len(seznam_operaci) + 1))
    cb_poradi.pack(side=tk.LEFT, padx=5)
    
    dostupne_operace = ["Sečíst duplicitní řádky", "Přičíst sloupec k jinému"]
    cb_operace = ttk.Combobox(radek_frame, values=dostupne_operace, width=22, state="readonly")
    cb_operace.set("Sečíst duplicitní řádky")
    cb_operace.pack(side=tk.LEFT, padx=5)
    
    frame_z = ttk.LabelFrame(radek_frame, text=" Podle kterého sloupce sloučit? ")
    frame_z.pack(side=tk.LEFT, padx=5)
    
    cb_odkud = ttk.Combobox(frame_z, values=sloupce_zakaznik, width=25, state="readonly")
    if sloupce_zakaznik: cb_odkud.set(sloupce_zakaznik[0])
    cb_odkud.pack(padx=5, pady=2)
    
    frame_do = ttk.LabelFrame(radek_frame, text=" Do kterého sloupce přičíst (R)? ")
    
    cb_kam = ttk.Combobox(frame_do, values=sloupce_zakaznik, width=25, state="readonly")
    if sloupce_zakaznik: cb_kam.set(sloupce_zakaznik[0])
    cb_kam.pack(padx=5, pady=2)
    
    cb_operace.bind("<<ComboboxSelected>>", lambda e: dynamic_ui_zmeny(e, cb_operace, frame_z, frame_do, cb_odkud, cb_kam))
    
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


# --- MOTOR AUTOMATIZACE ---
def spustit_konverzi():
    try:
        df = None
        for eng in ['openpyxl', 'xlrd', 'pyxlsb', 'calamine', None]:
            try:
                if df is None: 
                    df = pd.read_excel(cesta_zakaznik, sheet_name=cb_list_zakaznik.get(), engine=eng)
            except Exception: pass

        if df is None:
            messagebox.showerror("Chyba", "Nepodařilo se načíst soubor zákazníka.")
            return

        serazene_operace = sorted(seznam_operaci, key=lambda x: int(x["poradi"].get() if x["poradi"].get().isdigit() else 99))

        for radek in serazene_operace:
            op = radek["operace"].get()
            odkud = radek["odkud"].get()
            kam = radek["kam"].get()
            
            if not odkud: continue

            if op == "Sečíst duplicitní řádky":
                agg_dict = {}
                for col in df.columns:
                    if col == odkud:
                        continue
                    if pd.api.types.is_numeric_dtype(df[col]):
                        agg_dict[col] = 'sum'
                    else:
                        agg_dict[col] = 'first'
                
                df = df.groupby(odkud, as_index=False).agg(agg_dict)

            elif op == "Přičíst sloupec k jinému":
                if not kam: continue
                Zdroj_S = pd.to_numeric(df[odkud], errors='coerce').fillna(0)
                Cil_R = pd.to_numeric(df[kam], errors='coerce').fillna(0)
                df[kam] = Cil_R + Zdroj_S

        # Bezpečné ošetření přípony – vynutíme uložení do čistého .xlsx
        vystupni_cesta = cesta_cil
        if vystupni_cesta.lower().endswith('.xlsm'):
            vystupni_cesta = wystupni_cesta[:-5] + ".xlsx"
        elif vystupni_cesta.lower().endswith('.xls'):
            vystupni_cesta = vystupni_cesta[:-4] + ".xlsx"

        # Uložíme jako standardní moderní xlsx (vždy čistý zip/xml, který Excel schválí)
        with pd.ExcelWriter(vystupni_cesta, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=cb_list_zakaznik.get())

        messagebox.showinfo("Hotovo", f"Uloženo jako čistý sešit:\n{os.path.basename(vystupni_cesta)}")
        root.destroy()

    except Exception as e:
        messagebox.showerror("Chyba při zpracování", f"Něco kleklo v Pandas:\n{str(e)}")


def navrat_zpet():
    for radek in seznam_operaci: radek["frame"].destroy()
    seznam_operaci.clear()
    frame_strana2.pack_forget()
    frame_strana1.pack(fill=tk.BOTH, expand=True)


# ================= HLAVNÍ OKNO =================
root = tk.Tk()
root.title("Automatizace tabulek")
root.geometry("850x550")

style = ttk.Style()
style.theme_use('vista')

# ================= STRÁNKA 1 =================
frame_strana1 = ttk.Frame(root, padding="20")
frame_strana1.pack(fill=tk.BOTH, expand=True)

ttk.Label(frame_strana1, text="1. Vyber soubor od zákazníka, který chceš opravit", font=("Segoe UI", 10, "bold")).pack(pady=(15, 5))
ttk.Button(frame_strana1, text="Procházet soubor...", command=vybrat_zakaznika).pack()
lbl_zakaznik = ttk.Label(frame_strana1, text="Není vybráno", font=("Segoe UI", 9, "italic"))
lbl_zakaznik.pack(pady=(0, 5))

frame_list_zakaznik = ttk.Frame(frame_strana1)
ttk.Label(frame_list_zakaznik, text="Vyber list tabulky:").pack(side=tk.LEFT, padx=5)
cb_list_zakaznik = ttk.Combobox(frame_list_zakaznik, width=25, state="readonly")
cb_list_zakaznik.pack(side=tk.LEFT)

ttk.Label(frame_strana1, text="2. Kam uložit opravený výsledek?", font=("Segoe UI", 10, "bold")).pack(pady=(25, 5))
ttk.Button(frame_strana1, text="Určit název nového souboru...", command=vybrat_cil).pack()
lbl_cil = ttk.Label(frame_strana1, text="Není vybráno", font=("Segoe UI", 9, "italic"))
lbl_cil.pack(pady=(0, 30))

ttk.Button(frame_strana1, text="Pokračovat do nastavení kroků ➔", command=prejit_na_operace).pack(fill=tk.X, ipady=7)

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
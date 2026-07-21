import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd

# Globální proměnné pro cesty a listy
cesta_zakaznik = ""
cesta_nas = ""
cesta_cil = ""

# Seznam pro uložení ovládacích prvků ze strany 2
seznam_operaci = []

# Seznamy sloupců
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
            frame_list_zakaznik.pack(pady=(0, 10))


def vybrat_nas_soubor():
    global cesta_nas
    soubor = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
    if soubor:
        cesta_nas = soubor
        lbl_nas.config(text=os.path.basename(soubor))
        
        listy = nacist_listy(soubor)
        cb_list_nas['values'] = listy
        if listy:
            cb_list_nas.set(listy[0])
            frame_list_nas.pack(pady=(0, 10))


def vybrat_cil():
    global cesta_cil
    soubor = filedialog.asksaveasfilename(
        filetypes=[("Excel (.xlsx)", "*.xlsx")],
        defaultextension=".xlsx",
        initialdir=os.path.dirname(cesta_zakaznik) if cesta_zakaznik else None
    )
    if soubor:
        cesta_cil = soubor
        lbl_cil.config(text=os.path.basename(soubor))


def prejit_na_operace():
    global sloupce_zakaznik, sloupce_nase
    if not cesta_zakaznik or not cesta_nas or not cesta_cil:
        messagebox.showwarning("Chyba", "Musíš vybrat soubor zákazníka, náš soubor i cíl pro uložení!")
        return
    
    sloupce_zakaznik = nacist_sloupce_bezpecne(cesta_zakaznik, cb_list_zakaznik.get())
    sloupce_nase = nacist_sloupce_bezpecne(cesta_nas, cb_list_nas.get())
    
    if not sloupce_zakaznik or not sloupce_nase:
        messagebox.showerror("Chyba", "Nepodařilo se načíst sloupce ze souborů.")
        return

    frame_strana1.pack_forget()
    frame_strana2.pack(fill=tk.BOTH, expand=True)
    pridat_radek_operace()


# --- POP-UP OKNO PRO MAPOVÁNÍ SLOUPCŮ ---
def otevrit_popup_mapovani(data_radku):
    popup = tk.Toplevel(root)
    popup.title("Nastavení mapování sloupců")
    popup.geometry("550x500")
    popup.grab_set()  # Uzamkne okno, dokud se nezavře

    ttk.Label(popup, text="Kopírovat vybrané sloupce ze zákaznického souboru do našeho", font=("Segoe UI", 10, "bold")).pack(pady=10)

    # Scrollable oblast v popu
    pop_canvas = tk.Canvas(popup, borderwidth=0, highlightthickness=0)
    pop_scrollbar = ttk.Scrollbar(popup, orient="vertical", command=pop_canvas.yview)
    pop_scroll_frame = ttk.Frame(pop_canvas)

    pop_scroll_frame.bind("<Configure>", lambda e: pop_canvas.configure(scrollregion=pop_canvas.bbox("all")))
    pop_canvas.create_window((0, 0), window=pop_scroll_frame, anchor="nw")
    pop_canvas.configure(yscrollcommand=pop_scrollbar.set)

    pop_canvas.pack(side="top", fill="both", expand=True, padx=10, pady=5)
    pop_scrollbar.pack(side="right", fill="y")

    mapovani_seznam = data_radku.get("mapovani_rules", [])
    ui_mapovani_prvky = []

    def pridat_dvojici_sloupců(zdroj_val=None, cil_val=None):
        f = ttk.LabelFrame(pop_scroll_frame, text=f" Dvojice {len(ui_mapovani_prvky) + 1} ", padding=5)
        f.pack(fill=tk.X, pady=5, padx=5)

        ttk.Label(f, text="Zdroj (Zákazník):").grid(row=0, column=0, sticky=tk.W, padx=5)
        cb_src = ttk.Combobox(f, values=sloupce_zakaznik, width=30, state="readonly")
        if sloupce_zakaznik: cb_src.set(zdroj_val if zdroj_val else sloupce_zakaznik[0])
        cb_src.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(f, text="Cíl (Náš soubor):").grid(row=1, column=0, sticky=tk.W, padx=5)
        cb_dst = ttk.Combobox(f, values=sloupce_nase, width=30, state="readonly")
        if sloupce_nase: cb_dst.set(cil_val if cil_val else sloupce_nase[0])
        cb_dst.grid(row=1, column=1, padx=5, pady=2)

        btn_del = ttk.Button(f, text="✕", width=3, command=lambda: odebrat_dvojici(f, dvojice_data))
        btn_del.grid(row=0, column=2, rowspan=2, padx=5)

        dvojice_data = {"frame": f, "src": cb_src, "dst": cb_dst}
        ui_mapovani_prvky.append(dvojice_data)
        
        pop_canvas.update_idletasks()
        pop_canvas.configure(scrollregion=pop_canvas.bbox("all"))

    def odebrat_dvojici(frame, dvojice_data):
        frame.destroy()
        ui_mapovani_prvky.remove(dvojice_data)

    # Načteme existující pravidla, pokud už byla dříve vytvořena
    if mapovani_seznam:
        for r in mapovani_seznam:
            pridat_dvojici_sloupců(r["src"], r["dst"])
    else:
        pridat_dvojici_sloupců()

    # Spodní část pop-upu pro výběr konce dat a tlačítko uložit
    bot_frame = ttk.Frame(popup, padding=10)
    bot_frame.pack(fill=tk.X, side=tk.BOTTOM)

    ttk.Button(bot_frame, text="+ Přidat další sloupec", command=lambda: pridat_dvojici_sloupců()).pack(anchor=tk.W, pady=(0, 10))

    frame_konec = ttk.LabelFrame(bot_frame, text=" Určení konce dat ", padding=5)
    frame_konec.pack(fill=tk.X, pady=5)

    ttk.Label(frame_konec, text="Kopírovat řádky 2 až po konec sloupce:").pack(side=tk.LEFT, padx=5)
    cb_konec_col = ttk.Combobox(frame_konec, values=sloupce_zakaznik, width=25, state="readonly")
    if sloupce_zakaznik:
        cb_konec_col.set(data_radku.get("konec_podle_sloupce", sloupce_zakaznik[0]))
    cb_konec_col.pack(side=tk.LEFT, padx=5)

    def ulozit_mapovani():
        ulozene_pravidla = []
        for item in ui_mapovani_prvky:
            ulozene_pravidla.append({
                "src": item["src"].get(),
                "dst": item["dst"].get()
            })
        data_radku["mapovani_rules"] = ulozene_pravidla
        data_radku["konec_podle_sloupce"] = cb_konec_col.get()
        data_radku["btn_upravit"].config(text=f"Upravit mapování ({len(ulozene_pravidla)} sloupců)")
        popup.destroy()

    ttk.Button(bot_frame, text="Uložit a zavřít", command=ulozit_mapovani).pack(fill=tk.X, pady=(10, 0), ipady=5)


# --- STRÁNKA 2: DYNAMICKÝ SEZNAM KROKŮ ---
def dynamic_ui_zmeny(event, cb_operace, frame_z, frame_do, btn_upravit, cb_odkud, cb_kam):
    op = cb_operace.get()
    if op == "Sečíst duplicitní řádky":
        frame_z.pack(side=tk.LEFT, padx=5)
        frame_z.config(text=" Podle kterého sloupce sloučit? ")
        frame_do.pack_forget()
        btn_upravit.pack_forget()
        cb_odkud['values'] = sloupce_zakaznik
    elif op == "Přičíst sloupec k jinému":
        frame_z.pack(side=tk.LEFT, padx=5)
        frame_z.config(text=" Zdrojový sloupec (S): ")
        frame_do.pack(side=tk.LEFT, padx=5)
        btn_upravit.pack_forget()
        cb_odkud['values'] = sloupce_zakaznik
        cb_kam['values'] = sloupce_zakaznik
    elif op == "Přesunout/kopírovat data do naší tabulky":
        frame_z.pack_forget()
        frame_do.pack_forget()
        btn_upravit.pack(side=tk.LEFT, padx=10)


def pridat_radek_operace():
    radek_frame = ttk.Frame(scrollable_frame)
    radek_frame.pack(fill=tk.X, pady=5)
    
    cb_poradi = ttk.Combobox(radek_frame, values=[str(i) for i in range(1, 31)], width=3)
    cb_poradi.set(str(len(seznam_operaci) + 1))
    cb_poradi.pack(side=tk.LEFT, padx=5)
    
    dostupne_operace = [
        "Sečíst duplicitní řádky",
        "Přičíst sloupec k jinému",
        "Přesunout/kopírovat data do naší tabulky"
    ]
    cb_operace = ttk.Combobox(radek_frame, values=dostupne_operace, width=32, state="readonly")
    cb_operace.set("Sečíst duplicitní řádky")
    cb_operace.pack(side=tk.LEFT, padx=5)
    
    frame_z = ttk.LabelFrame(radek_frame, text=" Podle kterého sloupce sloučit? ")
    frame_z.pack(side=tk.LEFT, padx=5)
    cb_odkud = ttk.Combobox(frame_z, values=sloupce_zakaznik, width=22, state="readonly")
    if sloupce_zakaznik: cb_odkud.set(sloupce_zakaznik[0])
    cb_odkud.pack(padx=5, pady=2)
    
    frame_do = ttk.LabelFrame(radek_frame, text=" Do kterého sloupce přičíst (R)? ")
    cb_kam = ttk.Combobox(frame_do, values=sloupce_zakaznik, width=22, state="readonly")
    if sloupce_zakaznik: cb_kam.set(sloupce_zakaznik[0])
    cb_kam.pack(padx=5, pady=2)
    
    btn_upravit = ttk.Button(radek_frame, text="Upravit mapování", command=lambda: otevrit_popup_mapovani(data_radku))
    
    cb_operace.bind("<<ComboboxSelected>>", lambda e: dynamic_ui_zmeny(e, cb_operace, frame_z, frame_do, btn_upravit, cb_odkud, cb_kam))
    
    btn_smazat = ttk.Button(radek_frame, text="✕", width=3, command=lambda: smazat_radek_operace(radek_frame, data_radku))
    btn_smazat.pack(side=tk.RIGHT, padx=5)
    
    data_radku = {
        "frame": radek_frame, "poradi": cb_poradi, "operace": cb_operace,
        "odkud": cb_odkud, "kam": cb_kam, "btn_upravit": btn_upravit,
        "mapovani_rules": [], "konec_podle_sloupce": ""
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
        df_zak = None
        df_nas = None
        
        for eng in ['openpyxl', 'xlrd', 'pyxlsb', 'calamine', None]:
            try:
                if df_zak is None:
                    df_zak = pd.read_excel(cesta_zakaznik, sheet_name=cb_list_zakaznik.get(), engine=eng)
            except Exception: pass
            try:
                if df_nas is None:
                    df_nas = pd.read_excel(cesta_nas, sheet_name=cb_list_nas.get(), engine=eng)
            except Exception: pass

        if df_zak is None or df_nas is None:
            messagebox.showerror("Chyba", "Nepodařilo se načíst vstupní Excel soubory.")
            return

        serazene_operace = sorted(seznam_operaci, key=lambda x: int(x["poradi"].get() if x["poradi"].get().isdigit() else 99))

        for radek in serazene_operace:
            op = radek["operace"].get()
            odkud = radek["odkud"].get()
            kam = radek["kam"].get()

            if op == "Sečíst duplicitní řádky":
                if not odkud: continue
                agg_dict = {}
                for col in df_zak.columns:
                    if col == odkud: continue
                    if pd.api.types.is_numeric_dtype(df_zak[col]):
                        agg_dict[col] = 'sum'
                    else:
                        agg_dict[col] = 'first'
                df_zak = df_zak.groupby(odkud, as_index=False).agg(agg_dict)

            elif op == "Přičíst sloupec k jinému":
                if not odkud or not kam: continue
                Zdroj_S = pd.to_numeric(df_zak[odkud], errors='coerce').fillna(0)
                Cil_R = pd.to_numeric(df_zak[kam], errors='coerce').fillna(0)
                df_zak[kam] = Cil_R + Zdroj_S

            elif op == "Přesunout/kopírovat data do naší tabulky":
                pravidla = radek.get("mapovani_rules", [])
                konec_col = radek.get("konec_podle_sloupce", "")
                
                if not pravidla or not konec_col:
                    continue

                # Zjistíme poslední platný řádek ve vybraném určujícím sloupci (ignorueme NaN)
                posledni_index = df_zak[konec_col].dropna().index.max()
                
                if pd.isna(posledni_index):
                    posledni_index = len(df_zak) - 1

                # Ořízneme zákaznická data od řádku 0 (v Excelu řádek 2, protože hlavička je řádek 1) po konec
                df_zak_oriznuto = df_zak.loc[0:posledni_index]

                # Nakopírujeme hodnoty z vybraných zákaznických sloupců do našich cílových sloupců
                for rule in pravidla:
                    src_s = rule["src"]
                    dst_s = rule["dst"]
                    df_nas[dst_s] = df_zak_oriznuto[src_s]

        # Vynutíme uložení do čistého .xlsx
        vystupni_cesta = cesta_cil
        if vystupni_cesta.lower().endswith('.xlsm'):
            vystupni_cesta = vystupni_cesta[:-5] + ".xlsx"
        elif vystupni_cesta.lower().endswith('.xls'):
            vystupni_cesta = vystupni_cesta[:-4] + ".xlsx"

        with pd.ExcelWriter(vystupni_cesta, engine='openpyxl') as writer:
            df_nas.to_excel(writer, index=False, sheet_name=cb_list_nas.get())

        messagebox.showinfo("Hotovo", f"Zpracování proběhlo úspěšně!\nUloženo do:\n{os.path.basename(vystupni_cesta)}")
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
root.geometry("850x600")

style = ttk.Style()
style.theme_use('vista')

# ================= STRÁNKA 1 =================
frame_strana1 = ttk.Frame(root, padding="20")
frame_strana1.pack(fill=tk.BOTH, expand=True)

# 1. Zákazník
ttk.Label(frame_strana1, text="1. Vyber soubor od zákazníka", font=("Segoe UI", 10, "bold")).pack(pady=(5, 2))
ttk.Button(frame_strana1, text="Procházet zákazníka...", command=vybrat_zakaznika).pack()
lbl_zakaznik = ttk.Label(frame_strana1, text="Není vybráno", font=("Segoe UI", 9, "italic"))
lbl_zakaznik.pack(pady=(0, 2))

frame_list_zakaznik = ttk.Frame(frame_strana1)
ttk.Label(frame_list_zakaznik, text="Vyber list zákazníka:").pack(side=tk.LEFT, padx=5)
cb_list_zakaznik = ttk.Combobox(frame_list_zakaznik, width=25, state="readonly")
cb_list_zakaznik.pack(side=tk.LEFT)

# 2. Náš soubor
ttk.Label(frame_strana1, text="2. Vyber náš firemní soubor (šablonu)", font=("Segoe UI", 10, "bold")).pack(pady=(10, 2))
ttk.Button(frame_strana1, text="Procházet náš soubor...", command=vybrat_nas_soubor).pack()
lbl_nas = ttk.Label(frame_strana1, text="Není vybráno", font=("Segoe UI", 9, "italic"))
lbl_nas.pack(pady=(0, 2))

frame_list_nas = ttk.Frame(frame_strana1)
ttk.Label(frame_list_nas, text="Vyber náš list:").pack(side=tk.LEFT, padx=5)
cb_list_nas = ttk.Combobox(frame_list_nas, width=25, state="readonly")
cb_list_nas.pack(side=tk.LEFT)

# 3. Cíl
ttk.Label(frame_strana1, text="3. Kam uložit opravený výsledek?", font=("Segoe UI", 10, "bold")).pack(pady=(10, 2))
ttk.Button(frame_strana1, text="Určit název nového souboru...", command=vybrat_cil).pack()
lbl_cil = ttk.Label(frame_strana1, text="Není vybráno", font=("Segoe UI", 9, "italic"))
lbl_cil.pack(pady=(0, 15))

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
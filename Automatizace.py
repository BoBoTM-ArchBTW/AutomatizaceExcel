import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

# Globální cesty a seznamy sloupců
cesta_zakaznik = ""
cesta_nas = ""
cesta_cil = ""

seznam_operaci = []
sloupce_zakaznik = []
sloupce_nase = []


# --- HELPERY PRO NAČÍTÁNÍ EXCELŮ ---
def nacist_excel_bezpecne(cesta, **kwargs):
    """Pokusí se načíst Excel pomocí různých enginů pro maximální spolehlivost."""
    for eng in ["openpyxl", "xlrd", "pyxlsb", "calamine", None]:
        try:
            return pd.read_excel(cesta, engine=eng, **kwargs)
        except Exception:
            continue
    return None


def nacist_listy(cesta):
    for eng in ["openpyxl", "xlrd", "pyxlsb", "calamine", None]:
        try:
            return pd.ExcelFile(cesta, engine=eng).sheet_names
        except Exception:
            continue
    messagebox.showerror("Chyba", f"Nepodařilo se načíst listy ze souboru:\n{os.path.basename(cesta)}")
    return []


def nacist_sloupce(cesta, sheet_name, header_row):
    df = nacist_excel_bezpecne(cesta, sheet_name=sheet_name, header=header_row, nrows=0)
    return list(df.columns) if df is not None else []


# --- STRÁNKA 1: VÝBĚR SOUBORŮ ---
def vybrat_soubor(typ):
    global cesta_zakaznik, cesta_nas, cesta_cil
    if typ == "cil":
        path = filedialog.asksaveasfilename(
            filetypes=[("Excel (.xlsx)", "*.xlsx")],
            defaultextension=".xlsx",
            initialdir=os.path.dirname(cesta_zakaznik) if cesta_zakaznik else None
        )
        if path:
            cesta_cil = path
            lbl_cil.config(text=os.path.basename(path))
        return

    path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
    if not path:
        return

    listy = nacist_listy(path)
    if typ == "zakaznik":
        cesta_zakaznik = path
        lbl_zakaznik.config(text=os.path.basename(path))
        cb_list_zakaznik["values"] = listy
        if listy:
            cb_list_zakaznik.set(listy[0])
            frame_list_zakaznik.pack(pady=(0, 5))
    elif typ == "nas":
        cesta_nas = path
        lbl_nas.config(text=os.path.basename(path))
        cb_list_nas["values"] = listy
        if listy:
            cb_list_nas.set(listy[0])
            frame_list_nas.pack(pady=(0, 5))


def prejit_na_stranku_1_5():
    if not cesta_zakaznik or not cesta_nas or not cesta_cil:
        messagebox.showwarning("Chyba", "Musíš vybrat soubor zákazníka, náš soubor i cíl pro uložení!")
        return

    frame_strana1.pack_forget()
    frame_strana1_5.pack(fill=tk.BOTH, expand=True)
    obnovit_sloupce_strana1_5()


# --- STRÁNKA 1.5: PÁROVÁNÍ A ZÁHLAVÍ ---
def obnovit_sloupce_strana1_5(event=None):
    global sloupce_zakaznik, sloupce_nase
    try:
        hdr_zak = int(sp_hdr_zakaznik.get()) - 1
        hdr_nas = int(sp_hdr_nas.get()) - 1
    except ValueError:
        hdr_zak, hdr_nas = 0, 0

    sloupce_zakaznik = nacist_sloupce(cesta_zakaznik, cb_list_zakaznik.get(), hdr_zak)
    sloupce_nase = nacist_sloupce(cesta_nas, cb_list_nas.get(), hdr_nas)

    cb_klic_zakaznik["values"] = sloupce_zakaznik
    if sloupce_zakaznik:
        cb_klic_zakaznik.set(sloupce_zakaznik[0])

    cb_klic_nas["values"] = sloupce_nase
    if sloupce_nase:
        cb_klic_nas.set(sloupce_nase[0])


def prejit_na_operace():
    if not sloupce_zakaznik or not sloupce_nase:
        messagebox.showerror("Chyba", "Nepodařilo se načíst sloupce ze souborů.")
        return

    frame_strana1_5.pack_forget()
    frame_strana2.pack(fill=tk.BOTH, expand=True)
    if not seznam_operaci:
        pridat_radek_operace()


# --- POP-UP OKNO MAPOVÁNÍ ---
def otevrit_popup_mapovani(data_radku):
    popup = tk.Toplevel(root)
    popup.title("Nastavení mapování sloupců")
    popup.geometry("550x500")
    popup.grab_set()

    ttk.Label(popup, text="Kopírovat vybrané sloupce ze zákaznického souboru do našeho", font=("Segoe UI", 10, "bold")).pack(pady=10)

    canvas = tk.Canvas(popup, borderwidth=0, highlightthickness=0)
    scrollbar = ttk.Scrollbar(popup, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas)

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="top", fill="both", expand=True, padx=10, pady=5)
    scrollbar.pack(side="right", fill="y")

    ui_prvky = []

    def pridat_dvojici(src_val=None, dst_val=None):
        f = ttk.LabelFrame(scroll_frame, text=f" Dvojice {len(ui_prvky) + 1} ", padding=5)
        f.pack(fill=tk.X, pady=5, padx=5)

        ttk.Label(f, text="Zdroj (Zákazník):").grid(row=0, column=0, sticky=tk.W, padx=5)
        cb_src = ttk.Combobox(f, values=sloupce_zakaznik, width=30, state="readonly")
        cb_src.set(src_val if src_val else (sloupce_zakaznik[0] if sloupce_zakaznik else ""))
        cb_src.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(f, text="Cíl (Náš soubor):").grid(row=1, column=0, sticky=tk.W, padx=5)
        cb_dst = ttk.Combobox(f, values=sloupce_nase, width=30, state="readonly")
        cb_dst.set(dst_val if dst_val else (sloupce_nase[0] if sloupce_nase else ""))
        cb_dst.grid(row=1, column=1, padx=5, pady=2)

        polozka = {"frame": f, "src": cb_src, "dst": cb_dst}

        btn_del = ttk.Button(f, text="✕", width=3, command=lambda: odebrat_dvojici(polozka))
        btn_del.grid(row=0, column=2, rowspan=2, padx=5)

        ui_prvky.append(polozka)
        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    def odebrat_dvojici(polozka):
        polozka["frame"].destroy()
        ui_prvky.remove(polozka)

    mapovani = data_radku.get("mapovani_rules", [])
    if mapovani:
        for r in mapovani:
            pridat_dvojici(r["src"], r["dst"])
    else:
        pridat_dvojici()

    bot_frame = ttk.Frame(popup, padding=10)
    bot_frame.pack(fill=tk.X, side=tk.BOTTOM)

    ttk.Button(bot_frame, text="+ Přidat další sloupec", command=pridat_dvojici).pack(anchor=tk.W, pady=(0, 10))

    frame_konec = ttk.LabelFrame(bot_frame, text=" Určení konce dat u zákazníka ", padding=5)
    frame_konec.pack(fill=tk.X, pady=5)

    ttk.Label(frame_konec, text="Kopírovat data po konec sloupce:").pack(side=tk.LEFT, padx=5)
    cb_konec_col = ttk.Combobox(frame_konec, values=sloupce_zakaznik, width=25, state="readonly")
    cb_konec_col.set(data_radku.get("konec_podle_sloupce", sloupce_zakaznik[0] if sloupce_zakaznik else ""))
    cb_konec_col.pack(side=tk.LEFT, padx=5)

    def ulozit():
        data_radku["mapovani_rules"] = [{"src": x["src"].get(), "dst": x["dst"].get()} for x in ui_prvky]
        data_radku["konec_podle_sloupce"] = cb_konec_col.get()
        data_radku["btn_upravit"].config(text=f"Upravit mapování ({len(ui_prvky)} sloupců)")
        popup.destroy()

    ttk.Button(bot_frame, text="Uložit a zavřít", command=ulozit).pack(fill=tk.X, pady=(10, 0), ipady=5)


# --- STRÁNKA 2: DYNAMICKÝ SEZNAM KROKŮ ---
def dynamic_ui_zmeny(cb_operace, frame_z, frame_do, btn_upravit, cb_odkud, cb_kam):
    op = cb_operace.get()
    if op == "Sečíst duplicitní řádky":
        frame_z.pack(side=tk.LEFT, padx=5)
        frame_z.config(text=" Podle kterého sloupce sloučit? ")
        frame_do.pack_forget()
        btn_upravit.pack_forget()
        cb_odkud["values"] = sloupce_zakaznik
    elif op == "Přičíst sloupec k jinému":
        frame_z.pack(side=tk.LEFT, padx=5)
        frame_z.config(text=" Zdrojový sloupec (S): ")
        frame_do.pack(side=tk.LEFT, padx=5)
        btn_upravit.pack_forget()
        cb_odkud["values"] = sloupce_zakaznik
        cb_kam["values"] = sloupce_zakaznik
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
    if sloupce_zakaznik:
        cb_odkud.set(sloupce_zakaznik[0])
    cb_odkud.pack(padx=5, pady=2)

    frame_do = ttk.LabelFrame(radek_frame, text=" Do kterého sloupce přičíst (R)? ")
    cb_kam = ttk.Combobox(frame_do, values=sloupce_zakaznik, width=22, state="readonly")
    if sloupce_zakaznik:
        cb_kam.set(sloupce_zakaznik[0])
    cb_kam.pack(padx=5, pady=2)

    btn_upravit = ttk.Button(radek_frame, text="Upravit mapování", command=lambda: otevrit_popup_mapovani(data_radku))

    cb_operace.bind("<<ComboboxSelected>>", lambda e: dynamic_ui_zmeny(cb_operace, frame_z, frame_do, btn_upravit, cb_odkud, cb_kam))

    data_radku = {
        "frame": radek_frame, "poradi": cb_poradi, "operace": cb_operace,
        "odkud": cb_odkud, "kam": cb_kam, "btn_upravit": btn_upravit,
        "mapovani_rules": [], "konec_podle_sloupce": ""
    }

    btn_smazat = ttk.Button(radek_frame, text="✕", width=3, command=lambda: smazat_radek_operace(radek_frame, data_radku))
    btn_smazat.pack(side=tk.RIGHT, padx=5)

    seznam_operaci.append(data_radku)
    canvas.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))


def smazat_radek_operace(frame, data_radku):
    frame.destroy()
    seznam_operaci.remove(data_radku)
    for index, radek in enumerate(seznam_operaci):
        radek["poradi"].set(str(index + 1))


# --- ZPRACOVÁNÍ DAT PŘES PANDAS ---
def spustit_konverzi():
    try:
        hdr_zak_idx = int(sp_hdr_zakaznik.get()) - 1
        hdr_nas_idx = int(sp_hdr_nas.get()) - 1

        df_zak = nacist_excel_bezpecne(cesta_zakaznik, sheet_name=cb_list_zakaznik.get(), header=hdr_zak_idx)
        df_nas = nacist_excel_bezpecne(cesta_nas, sheet_name=cb_list_nas.get(), header=hdr_nas_idx)

        if df_zak is None or df_nas is None:
            messagebox.showerror("Chyba", "Nepodařilo se načíst vstupní Excel soubory.")
            return

        serazene_operace = sorted(seznam_operaci, key=lambda x: int(x["poradi"].get() if x["poradi"].get().isdigit() else 99))

        for radek in serazene_operace:
            op = radek["operace"].get()
            odkud = radek["odkud"].get()
            kam = radek["kam"].get()

            if op == "Sečíst duplicitní řádky":
                if not odkud:
                    continue
                agg_dict = {
                    col: ('sum' if pd.api.types.is_numeric_dtype(df_zak[col]) else 'first')
                    for col in df_zak.columns if col != odkud
                }
                df_zak = df_zak.groupby(odkud, as_index=False).agg(agg_dict)

            elif op == "Přičíst sloupec k jinému":
                if not odkud or not kam:
                    continue
                zdroj = pd.to_numeric(df_zak[odkud], errors='coerce').fillna(0)
                cil = pd.to_numeric(df_zak[kam], errors='coerce').fillna(0)
                df_zak[kam] = cil + zdroj

            elif op == "Přesunout/kopírovat data do naší tabulky":
                pravidla = radek.get("mapovani_rules", [])
                konec_col = radek.get("konec_podle_sloupce", "")
                if not pravidla or not konec_col:
                    continue

                posledni_index = df_zak[konec_col].dropna().index.max()
                if pd.isna(posledni_index):
                    posledni_index = len(df_zak) - 1

                df_zak_oriznuto = df_zak.loc[0:posledni_index]
                klic_zak, klic_nas = cb_klic_zakaznik.get(), cb_klic_nas.get()

                for rule in pravidla:
                    src_col, dst_col = rule["src"], rule["dst"]
                    mapa_hodnot = dict(zip(df_zak_oriznuto[klic_zak], df_zak_oriznuto[src_col]))
                    df_nas[dst_col] = df_nas[klic_nas].map(mapa_hodnot).fillna(df_nas[dst_col])

        # Uložení do čisto-čistého .xlsx
        vystupni_cesta = os.path.splitext(cesta_cil)[0] + ".xlsx"
        with pd.ExcelWriter(vystupni_cesta, engine='openpyxl') as writer:
            df_nas.to_excel(writer, index=False, sheet_name=cb_list_nas.get())

        messagebox.showinfo("Hotovo", f"Zpracování proběhlo úspěšně!\nUloženo do:\n{os.path.basename(vystupni_cesta)}")
        root.destroy()

    except PermissionError:
        nazev_souboru = os.path.basename(cesta_cil)
        messagebox.showerror(
            "Soubor je otevřený",
            f"Nepodařilo se uložit výsledek!\n\nSoubor '{nazev_souboru}' je právě otevřený v jiném programu (např. v Excelu).\n\nNejdříve ho prosím zavři a zkus to znovu."
        )
    except Exception as e:
        err_str = str(e)
        if "Permission denied" in err_str or "PermissionError" in err_str:
            nazev_souboru = os.path.basename(cesta_cil)
            messagebox.showerror(
                "Soubor je otevřený",
                f"Nepodařilo se uložit výsledek!\n\nSoubor '{nazev_souboru}' je právě otevřený v jiném programu (např. v Excelu).\n\nNejdříve ho prosím zavři a zkus to znovu."
            )
        else:
            messagebox.showerror("Chyba při zpracování", f"Něco kleklo v Pandas:\n{err_str}")


def navrat_z_1_5():
    frame_strana1_5.pack_forget()
    frame_strana1.pack(fill=tk.BOTH, expand=True)


def navrat_zpet():
    for radek in seznam_operaci:
        radek["frame"].destroy()
    seznam_operaci.clear()
    frame_strana2.pack_forget()
    frame_strana1_5.pack(fill=tk.BOTH, expand=True)


# ================= HLAVNÍ OKNO =================
root = tk.Tk()
root.title("Automatizace tabulek")
root.geometry("850x600")

style = ttk.Style()
style.theme_use('vista')

# ================= STRÁNKA 1 =================
frame_strana1 = ttk.Frame(root, padding="20")
frame_strana1.pack(fill=tk.BOTH, expand=True)

ttk.Label(frame_strana1, text="1. Vyber soubor od zákazníka", font=("Segoe UI", 10, "bold")).pack(pady=(5, 2))
ttk.Button(frame_strana1, text="Procházet zákazníka...", command=lambda: vybrat_soubor("zakaznik")).pack()
lbl_zakaznik = ttk.Label(frame_strana1, text="Není vybráno", font=("Segoe UI", 9, "italic"))
lbl_zakaznik.pack(pady=(0, 2))

frame_list_zakaznik = ttk.Frame(frame_strana1)
ttk.Label(frame_list_zakaznik, text="Vyber list zákazníka:").pack(side=tk.LEFT, padx=5)
cb_list_zakaznik = ttk.Combobox(frame_list_zakaznik, width=25, state="readonly")
cb_list_zakaznik.pack(side=tk.LEFT)

ttk.Label(frame_strana1, text="2. Vyber náš firemní soubor (šablonu)", font=("Segoe UI", 10, "bold")).pack(pady=(10, 2))
ttk.Button(frame_strana1, text="Procházet náš soubor...", command=lambda: vybrat_soubor("nas")).pack()
lbl_nas = ttk.Label(frame_strana1, text="Není vybráno", font=("Segoe UI", 9, "italic"))
lbl_nas.pack(pady=(0, 2))

frame_list_nas = ttk.Frame(frame_strana1)
ttk.Label(frame_list_nas, text="Vyber náš list:").pack(side=tk.LEFT, padx=5)
cb_list_nas = ttk.Combobox(frame_list_nas, width=25, state="readonly")
cb_list_nas.pack(side=tk.LEFT)

ttk.Label(frame_strana1, text="3. Kam uložit opravený výsledek?", font=("Segoe UI", 10, "bold")).pack(pady=(10, 2))
ttk.Button(frame_strana1, text="Určit název nového souboru...", command=lambda: vybrat_soubor("cil")).pack()
lbl_cil = ttk.Label(frame_strana1, text="Není vybráno", font=("Segoe UI", 9, "italic"))
lbl_cil.pack(pady=(0, 15))

ttk.Button(frame_strana1, text="Pokračovat k nastavení tabulek ➔", command=prejit_na_stranku_1_5).pack(fill=tk.X, ipady=7)

# ================= STRÁNKA 1.5 =================
frame_strana1_5 = ttk.Frame(root, padding="20")

ttk.Label(frame_strana1_5, text="Nastavení řádků záhlaví a párování", font=("Segoe UI", 11, "bold")).pack(anchor=tk.W, pady=(0, 15))

frame_hdr = ttk.LabelFrame(frame_strana1_5, text=" Na kterém řádku je v tabulce záhlaví (název sloupců)? ", padding="10")
frame_hdr.pack(fill=tk.X, pady=(0, 15))

ttk.Label(frame_hdr, text="Řádek záhlaví u zákazníka:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
sp_hdr_zakaznik = ttk.Spinbox(frame_hdr, from_=1, to=100, width=5)
sp_hdr_zakaznik.set(1)
sp_hdr_zakaznik.grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)

ttk.Label(frame_hdr, text="Řádek záhlaví v naší šabloně:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
sp_hdr_nas = ttk.Spinbox(frame_hdr, from_=1, to=100, width=5)
sp_hdr_nas.set(1)
sp_hdr_nas.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

ttk.Button(frame_hdr, text="Obnovit sloupce podle zadaných řádků", command=obnovit_sloupce_strana1_5).grid(row=2, column=0, columnspan=2, pady=10)

frame_keys = ttk.LabelFrame(frame_strana1_5, text=" Párování řádků (Čísla součástek / Kódy) ", padding="10")
frame_keys.pack(fill=tk.X, pady=(0, 20))

ttk.Label(frame_keys, text="Sloupec s kódem u zákazníka:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
cb_klic_zakaznik = ttk.Combobox(frame_keys, width=30, state="readonly")
cb_klic_zakaznik.grid(row=0, column=1, padx=5, pady=5)

ttk.Label(frame_keys, text="Sloupec s kódem u nás:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
cb_klic_nas = ttk.Combobox(frame_keys, width=30, state="readonly")
cb_klic_nas.grid(row=1, column=1, padx=5, pady=5)

nav_frame_1_5 = ttk.Frame(frame_strana1_5)
nav_frame_1_5.pack(fill=tk.X, side=tk.BOTTOM)

ttk.Button(nav_frame_1_5, text="⮌ Zpět", command=navrat_z_1_5).pack(side=tk.LEFT, ipady=5)
ttk.Button(nav_frame_1_5, text="Pokračovat k automatizaci ➔", command=prejit_na_operace).pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0), ipady=5)

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
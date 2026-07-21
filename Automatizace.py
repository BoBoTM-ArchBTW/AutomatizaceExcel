import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

# Soubor pro trvalé ukládání profilů
PROFILY_FILE = "konfigurace_profilu.json"

# Globální cesty a seznamy sloupců
cesta_zakaznik = ""
cesta_nas = ""
cesta_cil = ""

seznam_operaci = []
sloupce_zakaznik = []
sloupce_nase = []


# --- HELPERY PRO PRÁCI S JSON PROFILY ---
def nacist_profily_ze_souboru():
    if not os.path.exists(PROFILY_FILE):
        return {}
    try:
        with open(PROFILY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def ulozit_profily_do_souboru(data):
    try:
        with open(PROFILY_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception as e:
        messagebox.showerror("Chyba", f"Nepodařilo se uložit profily:\n{str(e)}")


# --- HELPER PRO SCROLLOVÁNÍ KOLEČKEM MYŠI ---
def pripojit_scrollovani_koleckem(widget_canvas):
    """Připojí plynulé scrollování kolečkem myši k danému canvasu."""

    def _on_mousewheel(event):
        widget_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    widget_canvas.bind(
        "<Enter>",
        lambda e: widget_canvas.bind_all("<MouseWheel>", _on_mousewheel),
    )
    widget_canvas.bind(
        "<Leave>", lambda e: widget_canvas.unbind_all("<MouseWheel>")
    )


# --- HELPERY PRO NAČÍTÁNÍ EXCELŮ ---
def nacist_excel_bezpecne(cesta, **kwargs):
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
    messagebox.showerror(
        "Chyba",
        f"Nepodařilo se načíst listy ze souboru:\n{os.path.basename(cesta)}",
    )
    return []


def nacist_sloupce(cesta, sheet_name, header_row):
    df = nacist_excel_bezpecne(
        cesta, sheet_name=sheet_name, header=header_row, nrows=0
    )
    return list(df.columns) if df is not None else []


# --- STRÁNKA 1: NAČTENÍ SOUBORŮ ---
def vybrat_soubor(typ):
    global cesta_zakaznik, cesta_nas, cesta_cil
    if typ == "cil":
        path = filedialog.asksaveasfilename(
            filetypes=[("Excel (.xlsx)", "*.xlsx")],
            defaultextension=".xlsx",
            initialdir=(
                os.path.dirname(cesta_zakaznik) if cesta_zakaznik else None
            ),
        )
        if path:
            cesta_cil = path
            lbl_cil.config(text=os.path.basename(path), foreground="#2e7d32")
        return

    path = filedialog.askopenfilename(
        filetypes=[("Excel", "*.xlsx *.xls *.xlsm")]
    )
    if not path:
        return

    listy = nacist_listy(path)
    if typ == "zakaznik":
        cesta_zakaznik = path
        lbl_zakaznik.config(text=os.path.basename(path), foreground="#2e7d32")
        cb_list_zakaznik["values"] = listy
        if listy:
            cb_list_zakaznik.set(listy[0])
    elif typ == "nas":
        cesta_nas = path
        lbl_nas.config(text=os.path.basename(path), foreground="#2e7d32")
        cb_list_nas["values"] = listy
        if listy:
            cb_list_nas.set(listy[0])


def prejit_na_pracovni_plochu():
    global sloupce_zakaznik, sloupce_nase
    if not cesta_zakaznik or not cesta_nas or not cesta_cil:
        messagebox.showwarning(
            "Chyba",
            "Musíš vybrat soubor zákazníka, náš soubor i cíl pro uložení!",
        )
        return

    try:
        hdr_zak = int(sp_hdr_zakaznik.get()) - 1
        hdr_nas = int(sp_hdr_nas.get()) - 1
    except ValueError:
        hdr_zak, hdr_nas = 0, 0

    sloupce_zakaznik = nacist_sloupce(
        cesta_zakaznik, cb_list_zakaznik.get(), hdr_zak
    )
    sloupce_nase = nacist_sloupce(cesta_nas, cb_list_nas.get(), hdr_nas)

    if not sloupce_zakaznik or not sloupce_nase:
        messagebox.showerror(
            "Chyba", "Nepodařilo se načíst sloupce ze souborů."
        )
        return

    cb_klic_zakaznik["values"] = sloupce_zakaznik
    cb_klic_zakaznik.set(sloupce_zakaznik[0])

    cb_klic_nas["values"] = sloupce_nase
    cb_klic_nas.set(sloupce_nase[0])

    cb_konec_col["values"] = sloupce_zakaznik
    cb_konec_col.set(sloupce_zakaznik[0])

    frame_strana1.pack_forget()
    frame_strana2.pack(fill=tk.BOTH, expand=True)

    if not seznam_operaci:
        pridat_radek_operace()


# --- POP-UP OKNO PRO KOPÍROVÁNÍ MAPOVÁNÍ ---
def otevrit_popup_mapovani(data_radku):
    popup = tk.Toplevel(root)
    popup.title("Mapování sloupců pro přesun")
    popup.geometry("620x520")
    popup.grab_set()

    ttk.Label(
        popup,
        text="Kopírování sloupců od Zákazníka ➔ do Naší šablony",
        font=("Segoe UI", 11, "bold"),
    ).pack(pady=10)

    canvas = tk.Canvas(popup, borderwidth=0, highlightthickness=0)
    scrollbar = ttk.Scrollbar(popup, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
    )
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    pripojit_scrollovani_koleckem(canvas)

    canvas.pack(side="top", fill="both", expand=True, padx=10, pady=5)
    scrollbar.pack(side="right", fill="y")

    ui_prvky = []

    def pridat_dvojici(src_val=None, dst_val=None):
        f = ttk.LabelFrame(
            scroll_frame, text=f" Pravidlo {len(ui_prvky) + 1} ", padding=8
        )
        f.pack(fill=tk.X, pady=4, padx=5)

        ttk.Label(f, text="Zdroj:").grid(
            row=0, column=0, sticky=tk.W, padx=5
        )
        cb_src = ttk.Combobox(
            f, values=sloupce_zakaznik, width=32, state="readonly"
        )
        cb_src.set(
            src_val
            if src_val
            else (sloupce_zakaznik[0] if sloupce_zakaznik else "")
        )
        cb_src.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(f, text="Cíl:").grid(
            row=1, column=0, sticky=tk.W, padx=5
        )
        cb_dst = ttk.Combobox(
            f, values=sloupce_nase, width=32, state="readonly"
        )
        cb_dst.set(
            dst_val if dst_val else (sloupce_nase[0] if sloupce_nase else "")
        )
        cb_dst.grid(row=1, column=1, padx=5, pady=2)

        polozka = {"frame": f, "src": cb_src, "dst": cb_dst}

        btn_del = ttk.Button(
            f, text="✕", width=3, command=lambda: odebrat_dvojici(polozka)
        )
        btn_del.grid(row=0, column=2, rowspan=2, padx=10)

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

    ttk.Button(
        bot_frame, text="+ Přidat další pravidlo", command=pridat_dvojici
    ).pack(anchor=tk.W, pady=(0, 10))

    def ulozit():
        data_radku["mapovani_rules"] = [
            {"src": x["src"].get(), "dst": x["dst"].get()} for x in ui_prvky
        ]
        data_radku["btn_upravit"].config(
            text=f"⚙ Mapování ({len(ui_prvky)} sloupců)"
        )
        popup.destroy()

    ttk.Button(bot_frame, text="Uložit mapování", command=ulozit).pack(
        fill=tk.X, ipady=4
    )


# --- POP-UP OKNO PRO RŮZNÉ DYN. OPERACE ---
def otevrit_popup_operace(data_radku):
    op = data_radku["operace"].get()
    popup = tk.Toplevel(root)
    popup.title(f"Nastavení: {op}")
    popup.geometry("480x300")
    popup.grab_set()

    frame_main = ttk.Frame(popup, padding=15)
    frame_main.pack(fill=tk.BOTH, expand=True)

    if op == "Sečíst duplicitní řádky":
        ttk.Label(
            frame_main,
            text="Sloučit duplicitní řádky podle sloupce u zákazníka:",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor=tk.W, pady=(0, 5))
        cb_single = ttk.Combobox(
            frame_main, values=sloupce_zakaznik, width=42, state="readonly"
        )
        cb_single.set(data_radku.get("klic1", sloupce_zakaznik[0] if sloupce_zakaznik else ""))
        cb_single.pack(anchor=tk.W, pady=(0, 15))

        def ulozit_secist():
            data_radku["klic1"] = cb_single.get()
            data_radku["btn_upravit"].config(
                text=f"⚙ Sloučit podle [{cb_single.get()}]"
            )
            popup.destroy()

        ttk.Button(frame_main, text="Uložit", command=ulozit_secist).pack(
            side=tk.BOTTOM, fill=tk.X, ipady=4
        )

    elif op == "Přičíst sloupec k jinému":
        ttk.Label(
            frame_main,
            text="Zdrojový sloupec:",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor=tk.W, pady=(0, 2))
        cb_s = ttk.Combobox(
            frame_main, values=sloupce_zakaznik, width=42, state="readonly"
        )
        cb_s.set(data_radku.get("klic1", sloupce_zakaznik[0] if sloupce_zakaznik else ""))
        cb_s.pack(anchor=tk.W, pady=(0, 10))

        ttk.Label(
            frame_main,
            text="Cílový sloupec:",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor=tk.W, pady=(0, 2))
        cb_r = ttk.Combobox(
            frame_main, values=sloupce_zakaznik, width=42, state="readonly"
        )
        cb_r.set(data_radku.get("klic2", sloupce_zakaznik[0] if sloupce_zakaznik else ""))
        cb_r.pack(anchor=tk.W, pady=(0, 15))

        def ulozit_pricist():
            data_radku["klic1"] = cb_s.get()
            data_radku["klic2"] = cb_r.get()
            data_radku["btn_upravit"].config(
                text=f"⚙ Přičíst [{cb_s.get()}] ➔ [{cb_r.get()}]"
            )
            popup.destroy()

        ttk.Button(frame_main, text="Uložit", command=ulozit_pricist).pack(
            side=tk.BOTTOM, fill=tk.X, ipady=4
        )

    elif op in ["Vyčistit sloupec", "Naplnit sloupec"]:
        ttk.Label(
            frame_main,
            text="1. V jakém souboru upravit sloupec?",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor=tk.W, pady=(0, 2))
        cb_tabulka = ttk.Combobox(
            frame_main,
            values=["Soubor Eolix", "Soubor zákazníka"],
            width=42,
            state="readonly",
        )
        cb_tabulka.set(data_radku.get("tabulka", "Soubor Eolix"))
        cb_tabulka.pack(anchor=tk.W, pady=(0, 10))

        ttk.Label(
            frame_main, text="2. Vyber sloupec:", font=("Segoe UI", 10, "bold")
        ).pack(anchor=tk.W, pady=(0, 2))
        cb_sloupec = ttk.Combobox(frame_main, width=42, state="readonly")
        cb_sloupec.pack(anchor=tk.W, pady=(0, 10))

        def aktualizovat_sloupce(e=None):
            if cb_tabulka.get() in ["Soubor Eolix", "Náš firemní soubor"]:
                cb_sloupec["values"] = sloupce_nase
                if sloupce_nase and (cb_sloupec.get() not in sloupce_nase):
                    cb_sloupec.set(sloupce_nase[0])
            else:
                cb_sloupec["values"] = sloupce_zakaznik
                if sloupce_zakaznik and (
                    cb_sloupec.get() not in sloupce_zakaznik
                ):
                    cb_sloupec.set(sloupce_zakaznik[0])

        cb_tabulka.bind("<<ComboboxSelected>>", aktualizovat_sloupce)
        aktualizovat_sloupce()
        if data_radku.get("vybrany_sloupec"):
            cb_sloupec.set(data_radku.get("vybrany_sloupec"))

        txt_val = None
        if op == "Naplnit sloupec":
            ttk.Label(
                frame_main,
                text="3. Zadej hodnotu pro naplnění:",
                font=("Segoe UI", 10, "bold"),
            ).pack(anchor=tk.W, pady=(0, 2))
            txt_val = ttk.Entry(frame_main, width=45)
            txt_val.insert(0, data_radku.get("hodnota_naplneni", ""))
            txt_val.pack(anchor=tk.W, pady=(0, 10))

        def ulozit_sloupec():
            data_radku["tabulka"] = cb_tabulka.get()
            data_radku["vybrany_sloupec"] = cb_sloupec.get()
            if op == "Naplnit sloupec" and txt_val:
                data_radku["hodnota_naplneni"] = txt_val.get()
                data_radku["btn_upravit"].config(
                    text=f"⚙ Naplnit [{cb_sloupec.get()}] = '{txt_val.get()}'"
                )
            else:
                data_radku["btn_upravit"].config(
                    text=f"⚙ Vyčistit [{cb_sloupec.get()}]"
                )
            popup.destroy()

        ttk.Button(frame_main, text="Uložit", command=ulozit_sloupec).pack(
            side=tk.BOTTOM, fill=tk.X, ipady=4
        )


# --- STRÁNKA 2: PROFILY (ULOŽIT / NAČÍST) ---
def otevrit_popup_ulozit_profil():
    if not seznam_operaci:
        messagebox.showwarning("Prázdné operace", "Nemáš vytvořené žádné kroky k uložení!")
        return

    popup = tk.Toplevel(root)
    popup.title("Uložit aktuální profil")
    popup.geometry("450x250")
    popup.grab_set()

    f = ttk.Frame(popup, padding=15)
    f.pack(fill=tk.BOTH, expand=True)

    ttk.Label(f, text="Název profilu:", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W, pady=(0, 2))
    txt_nazev = ttk.Entry(f, width=50)
    txt_nazev.pack(anchor=tk.W, pady=(0, 10))

    ttk.Label(f, text="Popis profilu (volitelné):", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W, pady=(0, 2))
    txt_popis = ttk.Entry(f, width=50)
    txt_popis.pack(anchor=tk.W, pady=(0, 15))

    def potvrdit_ulozeni():
        nazev = txt_nazev.get().strip()
        if not nazev:
            messagebox.showwarning("Chyba", "Musíš zadat název profilu!")
            return

        uložené_kroky = []
        for radek in seznam_operaci:
            uložené_kroky.append({
                "operace": radek["operace"].get(),
                "mapovani_rules": radek.get("mapovani_rules", []),
                "klic1": radek.get("klic1", ""),
                "klic2": radek.get("klic2", ""),
                "tabulka": radek.get("tabulka", "Soubor Eolix"),
                "vybrany_sloupec": radek.get("vybrany_sloupec", ""),
                "hodnota_naplneni": radek.get("hodnota_naplneni", ""),
            })

        profily = nacist_profily_ze_souboru()
        profily[nazev] = {
            "popis": txt_popis.get().strip(),
            "kroky": uložené_kroky
        }
        ulozit_profily_do_souboru(profily)

        messagebox.showinfo("Hotovo", f"Profil '{nazev}' byl úspěšně uložen!")
        popup.destroy()

    ttk.Button(f, text="Uložit profil", command=potvrdit_ulozeni).pack(side=tk.BOTTOM, fill=tk.X, ipady=4)


def otevrit_popup_nacist_profil():
    profily = nacist_profily_ze_souboru()
    if not profily:
        messagebox.showinfo("Žádné profily", "Zatím nemáš uložené žádné konfigurace profilů.")
        return

    popup = tk.Toplevel(root)
    popup.title("Správa a načítání profilů")
    # Zvětšeno pro mnohem lepší čitelnost
    popup.geometry("650x500")
    popup.grab_set()

    ttk.Label(popup, text="Uložené konfigurace operací:", font=("Segoe UI", 11, "bold")).pack(anchor=tk.W, padx=15, pady=10)

    canvas = tk.Canvas(popup, borderwidth=0, highlightthickness=0)
    scrollbar = ttk.Scrollbar(popup, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas)

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    pripojit_scrollovani_koleckem(canvas)

    canvas.pack(side="top", fill="both", expand=True, padx=10, pady=5)
    scrollbar.pack(side="right", fill="y")

    def vykreslit_seznam():
        for child in scroll_frame.winfo_children():
            child.destroy()

        aktualni_profily = nacist_profily_ze_souboru()
        if not aktualni_profily:
            ttk.Label(scroll_frame, text="Žádné profily.").pack(pady=20)
            return

        for p_nazev, p_data in aktualni_profily.items():
            box = ttk.LabelFrame(scroll_frame, text=f" {p_nazev} ", padding=10)
            box.pack(fill=tk.X, pady=6, padx=5)

            # Čitelnější název popisu
            popis_text = p_data.get('popis', 'Bez popisu')
            if not popis_text:
                popis_text = "Bez popisu"

            lbl_p = ttk.Label(box, text=f"Popis: {popis_text}", font=("Segoe UI", 9, "italic"), foreground="#444444")
            lbl_p.pack(anchor=tk.W, pady=(0, 8))

            btns_frame = ttk.Frame(box)
            btns_frame.pack(fill=tk.X)

            ttk.Button(
                btns_frame, text="Načíst do programu",
                command=lambda name=p_nazev, data=p_data: aplikovat_profil(name, data, popup)
            ).pack(side=tk.LEFT, padx=(0, 8))

            ttk.Button(
                btns_frame, text="✏ Upravit",
                command=lambda name=p_nazev, data=p_data: upravit_profil_dialog(name, data, vykreslit_seznam)
            ).pack(side=tk.LEFT, padx=4)

            ttk.Button(
                btns_frame, text="✕ Smazat",
                command=lambda name=p_nazev: smazat_profil(name, vykreslit_seznam)
            ).pack(side=tk.RIGHT, padx=4)

    def aplikovat_profil(nazev, data, win):
        for radek in list(seznam_operaci):
            smazat_radek_operace(radek["frame"], radek)

        for krok in data.get("kroky", []):
            radek_data = pridat_radek_operace()
            radek_data["operace"].set(krok.get("operace"))
            zmena_operace(radek_data["operace"], radek_data["btn_upravit"], radek_data)

            radek_data["mapovani_rules"] = krok.get("mapovani_rules", [])
            radek_data["klic1"] = krok.get("klic1", "")
            radek_data["klic2"] = krok.get("klic2", "")
            radek_data["tabulka"] = krok.get("tabulka", "Soubor Eolix")
            radek_data["vybrany_sloupec"] = krok.get("vybrany_sloupec", "")
            radek_data["hodnota_naplneni"] = krok.get("hodnota_naplneni", "")

            op = krok.get("operace")
            if op == "Přesunout/kopírovat data do naší tabulky":
                radek_data["btn_upravit"].config(text=f"⚙ Mapování ({len(radek_data['mapovani_rules'])} sloupců)")
            elif op == "Sečíst duplicitní řádky":
                radek_data["btn_upravit"].config(text=f"⚙ Sloučit podle [{radek_data['klic1']}]")
            elif op == "Přičíst sloupec k jinému":
                radek_data["btn_upravit"].config(text=f"⚙ Přičíst [{radek_data['klic1']}] ➔ [{radek_data['klic2']}]")
            elif op == "Naplnit sloupec":
                radek_data["btn_upravit"].config(text=f"⚙ Naplnit [{radek_data['vybrany_sloupec']}] = '{radek_data['hodnota_naplneni']}'")
            elif op == "Vyčistit sloupec":
                radek_data["btn_upravit"].config(text=f"⚙ Vyčistit [{radek_data['vybrany_sloupec']}]")

        messagebox.showinfo("Profil načten", f"Profil '{nazev}' byl načten!")
        win.destroy()

    def upravit_profil_dialog(stary_nazev, data, refresh_func):
        u_popup = tk.Toplevel(popup)
        u_popup.title(f"Úprava profilu: {stary_nazev}")
        u_popup.geometry("450x220")
        u_popup.grab_set()

        uf = ttk.Frame(u_popup, padding=12)
        uf.pack(fill=tk.BOTH, expand=True)

        ttk.Label(uf, text="Název profilu:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        u_nazev = ttk.Entry(uf, width=48)
        u_nazev.insert(0, stary_nazev)
        u_nazev.pack(anchor=tk.W, pady=(0, 10))

        ttk.Label(uf, text="Popis:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        u_popis = ttk.Entry(uf, width=48)
        u_popis.insert(0, data.get("popis", ""))
        u_popis.pack(anchor=tk.W, pady=(0, 10))

        def ulozit_zmeny():
            novy_nazev = u_nazev.get().strip()
            if not novy_nazev: return
            all_p = nacist_profily_ze_souboru()
            if stary_nazev in all_p:
                del all_p[stary_nazev]
            all_p[novy_nazev] = {
                "popis": u_popis.get().strip(),
                "kroky": data.get("kroky", [])
            }
            ulozit_profily_do_souboru(all_p)
            u_popup.destroy()
            refresh_func()

        ttk.Button(uf, text="Uložit změny", command=ulozit_zmeny).pack(side=tk.BOTTOM, fill=tk.X, ipady=4)

    def smazat_profil(nazev, refresh_func):
        if messagebox.askyesno("Smazat", f"Opravdu chceš smazat profil '{nazev}'?"):
            all_p = nacist_profily_ze_souboru()
            if nazev in all_p:
                del all_p[nazev]
                ulozit_profily_do_souboru(all_p)
                refresh_func()

    vykreslit_seznam()


# --- CHYTRÉ PŘESKLÁDÁNÍ A PŘEPOČET ČÍSEL BEZ DUPLICIT ---
def preskladat_radky_operaci(event=None, meneny_radek=None):
    if meneny_radek:
        try:
            nova_pozice = int(meneny_radek["poradi"].get())
            seznam_operaci.remove(meneny_radek)
            seznam_operaci.insert(nova_pozice - 1, meneny_radek)
        except ValueError:
            pass

    for index, radek in enumerate(seznam_operaci):
        radek["poradi"].set(str(index + 1))
        radek["frame"].config(text=f" Krok {index + 1} ")
        radek["frame"].pack_forget()
        radek["frame"].pack(fill=tk.X, pady=4, padx=5)

    canvas.update_idletasks()
    canvas.configure(scrollregion=canvas.bbox("all"))


# --- STRÁNKA 2: DYNAMICKÝ SEZNAM KROKŮ ---
def zmena_operace(cb_operace, btn_upravit, data_radku):
    op = cb_operace.get()
    if op == "Přesunout/kopírovat data do naší tabulky":
        btn_upravit.config(
            text="⚙ Mapovat sloupce",
            command=lambda: otevrit_popup_mapovani(data_radku),
        )
    else:
        btn_upravit.config(
            text="⚙ Nastavit",
            command=lambda: otevrit_popup_operace(data_radku),
        )


def pridat_radek_operace():
    radek_frame = ttk.LabelFrame(
        scrollable_frame, text=f" Krok {len(seznam_operaci) + 1} ", padding=8
    )
    radek_frame.pack(fill=tk.X, pady=4, padx=5)

    cb_poradi = ttk.Combobox(
        radek_frame,
        values=[str(i) for i in range(1, 31)],
        width=3,
        state="readonly",
    )
    cb_poradi.set(str(len(seznam_operaci) + 1))
    cb_poradi.pack(side=tk.LEFT, padx=(5, 10))

    cb_poradi.bind(
        "<<ComboboxSelected>>",
        lambda e: preskladat_radky_operaci(e, data_radku),
    )

    dostupne_operace = [
        "Přesunout/kopírovat data do naší tabulky",
        "Sečíst duplicitní řádky",
        "Přičíst sloupec k jinému",
        "Vyčistit sloupec",
        "Naplnit sloupec",
    ]
    cb_operace = ttk.Combobox(
        radek_frame, values=dostupne_operace, width=35, state="readonly"
    )
    cb_operace.set("Přesunout/kopírovat data do naší tabulky")
    cb_operace.pack(side=tk.LEFT, padx=5)

    btn_upravit = ttk.Button(radek_frame, text="⚙ Mapovat sloupce")
    btn_upravit.pack(side=tk.LEFT, padx=15)

    data_radku = {
        "frame": radek_frame,
        "poradi": cb_poradi,
        "operace": cb_operace,
        "btn_upravit": btn_upravit,
        "mapovani_rules": [],
        "klic1": "",
        "klic2": "",
        "tabulka": "Soubor Eolix",
        "vybrany_sloupec": "",
        "hodnota_naplneni": "",
    }

    zmena_operace(cb_operace, btn_upravit, data_radku)
    cb_operace.bind(
        "<<ComboboxSelected>>",
        lambda e: zmena_operace(cb_operace, btn_upravit, data_radku),
    )

    btn_smazat = ttk.Button(
        radek_frame,
        text="✕",
        width=3,
        command=lambda: smazat_radek_operace(radek_frame, data_radku),
    )
    btn_smazat.pack(side=tk.RIGHT, padx=5)

    seznam_operaci.append(data_radku)
    preskladat_radky_operaci()
    return data_radku


def smazat_radek_operace(frame, data_radku):
    frame.destroy()
    seznam_operaci.remove(data_radku)
    preskladat_radky_operaci()


# --- MOTOR AUTOMATIZACE (PANDAS) ---
def spustit_konverzi():
    try:
        hdr_zak_idx = int(sp_hdr_zakaznik.get()) - 1
        hdr_nas_idx = int(sp_hdr_nas.get()) - 1

        df_zak = nacist_excel_bezpecne(
            cesta_zakaznik,
            sheet_name=cb_list_zakaznik.get(),
            header=hdr_zak_idx,
        )
        df_nas = nacist_excel_bezpecne(
            cesta_nas, sheet_name=cb_list_nas.get(), header=hdr_nas_idx
        )

        if df_zak is None or df_nas is None:
            messagebox.showerror(
                "Chyba", "Nepodařilo se načíst vstupní Excel soubory."
            )
            return

        for radek in seznam_operaci:
            op = radek["operace"].get()

            if op == "Sečíst duplicitní řádky":
                odkud = radek.get("klic1")
                if not odkud or odkud not in df_zak.columns:
                    continue
                agg_dict = {
                    col: (
                        "sum"
                        if pd.api.types.is_numeric_dtype(df_zak[col])
                        else "first"
                    )
                    for col in df_zak.columns
                    if col != odkud
                }
                df_zak = df_zak.groupby(odkud, as_index=False).agg(agg_dict)

            elif op == "Přičíst sloupec k jinému":
                odkud = radek.get("klic1")
                kam = radek.get("klic2")
                if (
                    not odkud
                    or not kam
                    or odkud not in df_zak.columns
                    or kam not in df_zak.columns
                ):
                    continue
                zdroj = pd.to_numeric(df_zak[odkud], errors="coerce").fillna(0)
                cil = pd.to_numeric(df_zak[kam], errors="coerce").fillna(0)
                df_zak[kam] = cil + zdroj

            elif op == "Přesunout/kopírovat data do naší tabulky":
                pravidla = radek.get("mapovani_rules", [])
                konec_col = cb_konec_col.get()
                if (
                    not pravidla
                    or not konec_col
                    or konec_col not in df_zak.columns
                ):
                    continue

                posledni_index = df_zak[konec_col].dropna().index.max()
                if pd.isna(posledni_index):
                    posledni_index = len(df_zak) - 1

                df_zak_oriznuto = df_zak.loc[0:posledni_index]
                klic_zak, klic_nas = cb_klic_zakaznik.get(), cb_klic_nas.get()

                for rule in pravidla:
                    src_col, dst_col = rule["src"], rule["dst"]
                    if (
                        src_col in df_zak_oriznuto.columns
                        and dst_col in df_nas.columns
                    ):
                        mapa_hodnot = dict(
                            zip(
                                df_zak_oriznuto[klic_zak],
                                df_zak_oriznuto[src_col],
                            )
                        )
                        df_nas[dst_col] = (
                            df_nas[klic_nas]
                            .map(mapa_hodnot)
                            .fillna(df_nas[dst_col])
                        )

            elif op == "Vyčistit sloupec":
                tab = radek.get("tabulka")
                col = radek.get("vybrany_sloupec")
                if tab in ["Soubor Eolix", "Náš firemní soubor"] and col in df_nas.columns:
                    df_nas[col] = None
                elif tab == "Soubor zákazníka" and col in df_zak.columns:
                    df_zak[col] = None

            elif op == "Naplnit sloupec":
                tab = radek.get("tabulka")
                col = radek.get("vybrany_sloupec")
                val = radek.get("hodnota_naplneni", "")
                if tab in ["Soubor Eolix", "Náš firemní soubor"] and col in df_nas.columns:
                    df_nas[col] = val
                elif tab == "Soubor zákazníka" and col in df_zak.columns:
                    df_zak[col] = val

        # Uložení výstupu
        vystupni_cesta = os.path.splitext(cesta_cil)[0] + ".xlsx"
        with pd.ExcelWriter(vystupni_cesta, engine="openpyxl") as writer:
            df_nas.to_excel(writer, index=False, sheet_name=cb_list_nas.get())

        messagebox.showinfo(
            "Hotovo",
            f"Zpracování proběhlo úspěšně!\nUloženo do:\n{os.path.basename(vystupni_cesta)}",
        )
        root.destroy()

    except PermissionError:
        nazev_souboru = os.path.basename(cesta_cil)
        messagebox.showerror(
            "Soubor je otevřený",
            f"Nepodařilo se uložit výsledek!\n\nSoubor '{nazev_souboru}' je právě otevřený v jiném programu (např. v Excelu).\n\nNejdříve ho prosím zavři a zkus to znovu.",
        )
    except Exception as e:
        err_str = str(e)
        if "Permission denied" in err_str or "PermissionError" in err_str:
            nazev_souboru = os.path.basename(cesta_cil)
            messagebox.showerror(
                "Soubor je otevřený",
                f"Nepodařilo se uložit výsledek!\n\nSoubor '{nazev_souboru}' je právě otevřený v jiném programu (např. v Excelu).\n\nNejdříve ho prosím zavři a zkus to znovu.",
            )
        else:
            messagebox.showerror(
                "Chyba při zpracování", f"Něco kleklo v Pandas:\n{err_str}"
            )


def navrat_na_krok_1():
    frame_strana2.pack_forget()
    frame_strana1.pack(fill=tk.BOTH, expand=True)


# ================= HLAVNÍ OKNO =================
root = tk.Tk()
root.title("Automatizace konverze tabulek")
root.geometry("880x640")

style = ttk.Style()
style.theme_use("vista")

# ================= STRÁNKA 1: VSTUPY A SOUBORY =================
frame_strana1 = ttk.Frame(root, padding="20")
frame_strana1.pack(fill=tk.BOTH, expand=True)

ttk.Label(
    frame_strana1,
    text="Krok 1: Výběr vstupních souborů a záhlaví",
    font=("Segoe UI", 12, "bold"),
).pack(anchor=tk.W, pady=(0, 15))

grid_files = ttk.Frame(frame_strana1)
grid_files.pack(fill=tk.X, pady=(0, 15))

# KARTA 1: ZÁKAZNÍK
box_zak = ttk.LabelFrame(grid_files, text=" Soubor od zákazníka ", padding="12")
box_zak.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 8))

ttk.Button(
    box_zak,
    text="Procházet soubor...",
    command=lambda: vybrat_soubor("zakaznik"),
).pack(anchor=tk.W, pady=(0, 5))
lbl_zakaznik = ttk.Label(
    box_zak, text="Není vybráno", font=("Segoe UI", 8, "italic"), foreground="gray"
)
lbl_zakaznik.pack(anchor=tk.W, pady=(0, 10))

ttk.Label(box_zak, text="Vyber list:").pack(anchor=tk.W)
cb_list_zakaznik = ttk.Combobox(box_zak, state="readonly", width=25)
cb_list_zakaznik.pack(anchor=tk.W, pady=(0, 10))

ttk.Label(box_zak, text="Řádek se záhlavím:").pack(anchor=tk.W)
sp_hdr_zakaznik = ttk.Spinbox(box_zak, from_=1, to=100, width=5)
sp_hdr_zakaznik.set(1)
sp_hdr_zakaznik.pack(anchor=tk.W)

# KARTA 2: NAŠE ŠABLONA
box_nas = ttk.LabelFrame(
    grid_files, text=" Náš firemní soubor ", padding="12"
)
box_nas.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(8, 0))

ttk.Button(
    box_nas, text="Procházet soubor...", command=lambda: vybrat_soubor("nas")
).pack(anchor=tk.W, pady=(0, 5))
lbl_nas = ttk.Label(
    box_nas, text="Není vybráno", font=("Segoe UI", 8, "italic"), foreground="gray"
)
lbl_nas.pack(anchor=tk.W, pady=(0, 10))

ttk.Label(box_nas, text="Vyber list:").pack(anchor=tk.W)
cb_list_nas = ttk.Combobox(box_nas, state="readonly", width=25)
cb_list_nas.pack(anchor=tk.W, pady=(0, 10))

ttk.Label(box_nas, text="Řádek se záhlavím:").pack(anchor=tk.W)
sp_hdr_nas = ttk.Spinbox(box_nas, from_=1, to=100, width=5)
sp_hdr_nas.set(1)
sp_hdr_nas.pack(anchor=tk.W)

# KARTA 3: VÝSTUPNÍ SOUBOR
box_out = ttk.LabelFrame(frame_strana1, text=" Výstupní soubor ", padding="12")
box_out.pack(fill=tk.X, pady=(0, 20))

ttk.Button(
    box_out,
    text="Určit kam uložit nový soubor...",
    command=lambda: vybrat_soubor("cil"),
).pack(side=tk.LEFT, padx=(0, 10))
lbl_cil = ttk.Label(
    box_out, text="Není vybráno", font=("Segoe UI", 9, "italic"), foreground="gray"
)
lbl_cil.pack(side=tk.LEFT)

ttk.Button(
    frame_strana1,
    text="Pokračovat k nastavení automatizace ➔",
    command=prejit_na_pracovni_plochu,
).pack(fill=tk.X, side=tk.BOTTOM, ipady=8)


# ================= STRÁNKA 2: PRACOVNÍ PLOCHA A OPERACE =================
frame_strana2 = ttk.Frame(root, padding="15")

# NASTAVENÍ PÁROVÁNÍ NAHOŘE
frame_top_config = ttk.LabelFrame(
    frame_strana2,
    text=" Krok 2: Párování řádků a určování rozsahu ",
    padding="10",
)
frame_top_config.pack(fill=tk.X, pady=(0, 12))

ttk.Label(frame_top_config, text="P/N u Zákazníka:").grid(
    row=0, column=0, sticky=tk.W, padx=5, pady=2
)
cb_klic_zakaznik = ttk.Combobox(frame_top_config, width=28, state="readonly")
cb_klic_zakaznik.grid(row=0, column=1, padx=5, pady=2)

ttk.Label(frame_top_config, text="PN v Eolix:").grid(
    row=0, column=2, sticky=tk.W, padx=(20, 5), pady=2
)
cb_klic_nas = ttk.Combobox(frame_top_config, width=28, state="readonly")
cb_klic_nas.grid(row=0, column=3, padx=5, pady=2)

ttk.Label(frame_top_config, text="Konec dat:").grid(
    row=1, column=0, sticky=tk.W, padx=5, pady=(8, 2)
)
cb_konec_col = ttk.Combobox(frame_top_config, width=28, state="readonly")
cb_konec_col.grid(row=1, column=1, padx=5, pady=(8, 2))

# SEZNAM KROKŮ A LIŠTA PROFILŮ
header_krok3_frame = ttk.Frame(frame_strana2)
header_krok3_frame.pack(fill=tk.X, pady=(5, 5))

ttk.Label(
    header_krok3_frame,
    text="Krok 3: Seznam prováděných operací",
    font=("Segoe UI", 10, "bold"),
).pack(side=tk.LEFT)

# Tlačítka profily nahoře vpravo
btn_profily_frame = ttk.Frame(header_krok3_frame)
btn_profily_frame.pack(side=tk.RIGHT)

ttk.Button(
    btn_profily_frame, text="💾 Uložit profil", command=otevrit_popup_ulozit_profil
).pack(side=tk.LEFT, padx=4)

ttk.Button(
    btn_profily_frame, text="📂 Načíst profil", command=otevrit_popup_nacist_profil
).pack(side=tk.LEFT, padx=4)

canvas = tk.Canvas(frame_strana2, borderwidth=0, highlightthickness=0)
scrollbar = ttk.Scrollbar(
    frame_strana2, orient="vertical", command=canvas.yview
)
scrollable_frame = ttk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)
canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)

pripojit_scrollovani_koleckem(canvas)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

spodek_frame = ttk.Frame(frame_strana2)
spodek_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(12, 0))

# Tlačítka dole - přehlednější uspořádání s velkými mezerami
ttk.Button(
    spodek_frame, text="+ Přidat další krok", command=pridat_radek_operace
).pack(anchor=tk.W, pady=(0, 15))

nav_frame = ttk.Frame(spodek_frame)
nav_frame.pack(fill=tk.X)

ttk.Button(
    nav_frame, text="⮌ Zpět na výběr souborů", command=navrat_na_krok_1
).pack(side=tk.LEFT, ipady=5)

ttk.Button(
    nav_frame, text="Spustit konverzi", command=spustit_konverzi
).pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(20, 0), ipady=5)

root.mainloop()
import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd

# Soubor pro trvalé ukládání profilů
PROFILY_FILE = "konfigurace_profilu.json"

# Globální struktury
seznam_souboru_ui = []   # Seznam UI prvků souborů na Stránce 1
nactene_soubory = {}     # Slovník načtených dat
seznam_operaci = []      # Seznam kroků automatizace
cesta_cil_global = ""


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


# --- HELPER PRO TISKOVOU NÁPOVĚDU V POP-UPECH ---
def pridat_box_napovedy(parent_frame, text_napovedy):
    box = ttk.LabelFrame(parent_frame, text=" 💡 Nápověda ", padding=8)
    box.pack(fill=tk.X, pady=(0, 10))
    lbl = ttk.Label(box, text=text_napovedy, font=("Segoe UI", 8, "italic"), wraplength=440, justify="left", foreground="#333333")
    lbl.pack(anchor=tk.W)


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


# --- STRÁNKA 1: DYNAMICKÁ SPRÁVA SOUBORŮ ---
def pridat_soubor_ui():
    box = ttk.LabelFrame(scroll_frame_soubory, text=f" Soubor {len(seznam_souboru_ui) + 1} ", padding=10)
    box.pack(fill=tk.X, expand=True, pady=5, padx=5)

    f_row1 = ttk.Frame(box)
    f_row1.pack(fill=tk.X, pady=(0, 5))

    btn_browse = ttk.Button(f_row1, text="Procházet...")
    btn_browse.pack(side=tk.LEFT, padx=(0, 10))

    lbl_path = ttk.Label(f_row1, text="Není vybráno", font=("Segoe UI", 8, "italic"), foreground="gray")
    lbl_path.pack(side=tk.LEFT, fill=tk.X, expand=True)

    btn_del = ttk.Button(f_row1, text="✕", width=3, command=lambda: odebrat_soubor_ui(polozka))
    btn_del.pack(side=tk.RIGHT)

    f_row2 = ttk.Frame(box)
    f_row2.pack(fill=tk.X, pady=(5, 0))

    ttk.Label(f_row2, text="Název/Alias:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
    entry_alias = ttk.Entry(f_row2, width=22)
    entry_alias.grid(row=0, column=1, padx=(0, 15), sticky=tk.W)

    ttk.Label(f_row2, text="List:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
    cb_list = ttk.Combobox(f_row2, width=20, state="readonly")
    cb_list.grid(row=0, column=3, padx=(0, 15), sticky=tk.W)

    ttk.Label(f_row2, text="Řádek záhlaví:").grid(row=0, column=4, sticky=tk.W, padx=(0, 5))
    sp_hdr = ttk.Spinbox(f_row2, from_=1, to=100, width=5)
    sp_hdr.set(1)
    sp_hdr.grid(row=0, column=5, sticky=tk.W)

    polozka = {
        "box": box,
        "cesta": "",
        "lbl_path": lbl_path,
        "entry_alias": entry_alias,
        "cb_list": cb_list,
        "sp_hdr": sp_hdr,
    }

    def akce_vybrat():
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls *.xlsm")])
        if path:
            polozka["cesta"] = path
            lbl_path.config(text=os.path.basename(path), foreground="#2e7d32")

            base_name = os.path.splitext(os.path.basename(path))[0]
            entry_alias.delete(0, tk.END)
            entry_alias.insert(0, base_name)

            listy = nacist_listy(path)
            cb_list["values"] = listy
            if listy:
                cb_list.set(listy[0])

    btn_browse.config(command=akce_vybrat)
    seznam_souboru_ui.append(polozka)

    canvas_soubory.update_idletasks()
    canvas_soubory.configure(scrollregion=canvas_soubory.bbox("all"))
    aktualizovat_vystupni_sekci()


def odebrat_soubor_ui(polozka):
    polozka["box"].destroy()
    seznam_souboru_ui.remove(polozka)
    for idx, item in enumerate(seznam_souboru_ui):
        item["box"].config(text=f" Soubor {idx + 1} ")
    canvas_soubory.update_idletasks()
    canvas_soubory.configure(scrollregion=canvas_soubory.bbox("all"))
    aktualizovat_vystupni_sekci()


def vybrat_cilovy_soubor():
    global cesta_cil_global
    path = filedialog.asksaveasfilename(
        filetypes=[("Excel (.xlsx)", "*.xlsx")],
        defaultextension=".xlsx"
    )
    if path:
        cesta_cil_global = path
        lbl_cil_path.config(text=os.path.basename(path), foreground="#2e7d32")


def aktualizovat_vystupni_sekci():
    pocet = len(seznam_souboru_ui)
    if pocet <= 1:
        box_vystup_multi.pack_forget()
        box_vystup_single.pack(fill=tk.X, pady=(10, 0))
    else:
        box_vystup_single.pack_forget()
        box_vystup_multi.pack(fill=tk.X, pady=(10, 0))


def prejit_na_pracovni_plochu():
    global nactene_soubory
    if not seznam_souboru_ui:
        messagebox.showwarning("Chyba", "Musíš přidat alespoň jeden soubor!")
        return

    nactene_soubory.clear()

    for idx, item in enumerate(seznam_souboru_ui):
        cesta = item["cesta"]
        alias = item["entry_alias"].get().strip()
        sheet = item["cb_list"].get()
        try:
            hdr = int(item["sp_hdr"].get()) - 1
        except ValueError:
            hdr = 0

        if not cesta:
            messagebox.showwarning("Chyba", f"Soubor {idx + 1} nemá vybranou cestu!")
            return
        if not alias:
            alias = f"Soubor_{idx + 1}"

        sloupce = nacist_sloupce(cesta, sheet, hdr)
        if not sloupce:
            messagebox.showerror("Chyba", f"Nepodařilo se načíst sloupce ze souboru '{alias}'.")
            return

        nactene_soubory[alias] = {
            "cesta": cesta,
            "sheet": sheet,
            "hdr": hdr,
            "sloupce": sloupce
        }

    if len(nactene_soubory) > 1 and not cesta_cil_global:
        messagebox.showwarning("Chyba", "Při práci s více soubory musíš určit výstupní soubor!")
        return

    frame_strana1.pack_forget()
    frame_strana2.pack(fill=tk.BOTH, expand=True)

    if not seznam_operaci:
        pridat_radek_operace()


# --- POP-UP OKNO PRO KOPÍROVÁNÍ PODLE 1 KLÍČE ---
def otevrit_popup_mapovani(data_radku):
    popup = tk.Toplevel(root)
    popup.title("Mapování sloupců pro přesun/kopírování (1 Klíč)")
    popup.geometry("750x640")
    popup.grab_set()

    f_help = ttk.Frame(popup, padding=(12, 12, 12, 0))
    f_help.pack(fill=tk.X)
    pridat_box_napovedy(f_help, "Tato funkce najde v Cílovém souboru řádky, které mají stejný kód (P/N) jako ve Zdrojovém souboru, a překopíruje z nich hodnoty ze zadaných sloupců.")

    f_top = ttk.Frame(popup, padding=12)
    f_top.pack(fill=tk.X)

    seznam_aliasu = list(nactene_soubory.keys())

    ttk.Label(f_top, text="Zdrojový soubor:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
    cb_src_file = ttk.Combobox(f_top, values=seznam_aliasu, width=28, state="readonly")
    cb_src_file.set(data_radku.get("src_file", seznam_aliasu[0]))
    cb_src_file.grid(row=0, column=1, padx=5, pady=2)

    ttk.Label(f_top, text="Cílový soubor:", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, sticky=tk.W, padx=(20, 5), pady=2)
    cb_dst_file = ttk.Combobox(f_top, values=seznam_aliasu, width=28, state="readonly")
    cb_dst_file.set(data_radku.get("dst_file", seznam_aliasu[-1] if len(seznam_aliasu) > 1 else seznam_aliasu[0]))
    cb_dst_file.grid(row=0, column=3, padx=5, pady=2)

    ttk.Label(f_top, text="Kód ve Zdrojovém s.:", font=("Segoe UI", 9)).grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
    cb_klic_src = ttk.Combobox(f_top, width=28, state="readonly")
    cb_klic_src.grid(row=1, column=1, padx=5, pady=2)

    ttk.Label(f_top, text="Kód v Cílovém s.:", font=("Segoe UI", 9)).grid(row=1, column=2, sticky=tk.W, padx=(20, 5), pady=2)
    cb_klic_dst = ttk.Combobox(f_top, width=28, state="readonly")
    cb_klic_dst.grid(row=1, column=3, padx=5, pady=2)

    ttk.Label(f_top, text="Konec dat podle (Zdroj):", font=("Segoe UI", 9)).grid(row=2, column=0, sticky=tk.W, padx=5, pady=(6, 2))
    cb_konec_col = ttk.Combobox(f_top, width=28, state="readonly")
    cb_konec_col.grid(row=2, column=1, padx=5, pady=(6, 2))

    def aktualizovat_sloupce_souboru(e=None):
        src_cols = nactene_soubory.get(cb_src_file.get(), {}).get("sloupce", [])
        dst_cols = nactene_soubory.get(cb_dst_file.get(), {}).get("sloupce", [])

        cb_klic_src["values"] = src_cols
        if src_cols: cb_klic_src.set(data_radku.get("klic_src", src_cols[0]))

        cb_klic_dst["values"] = dst_cols
        if dst_cols: cb_klic_dst.set(data_radku.get("klic_dst", dst_cols[0]))

        cb_konec_col["values"] = src_cols
        if src_cols: cb_konec_col.set(data_radku.get("konec_col", src_cols[0]))

    cb_src_file.bind("<<ComboboxSelected>>", aktualizovat_sloupce_souboru)
    cb_dst_file.bind("<<ComboboxSelected>>", aktualizovat_sloupce_souboru)
    aktualizovat_sloupce_souboru()

    canvas = tk.Canvas(popup, borderwidth=0, highlightthickness=0)
    scrollbar = ttk.Scrollbar(popup, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas)

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    pop_win_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(pop_win_id, width=e.width))

    canvas.configure(yscrollcommand=scrollbar.set)
    pripojit_scrollovani_koleckem(canvas)

    canvas.pack(side="top", fill="both", expand=True, padx=10, pady=5)
    scrollbar.pack(side="right", fill="y")

    ui_prvky = []

    def pridat_dvojici(src_val=None, dst_val=None):
        f = ttk.LabelFrame(scroll_frame, text=f" Pravidlo {len(ui_prvky) + 1} ", padding=8)
        f.pack(fill=tk.X, expand=True, pady=4, padx=5)

        src_cols = nactene_soubory.get(cb_src_file.get(), {}).get("sloupce", [])
        dst_cols = nactene_soubory.get(cb_dst_file.get(), {}).get("sloupce", [])

        ttk.Label(f, text="Zdrojový sloupec:").grid(row=0, column=0, sticky=tk.W, padx=5)
        cb_src = ttk.Combobox(f, values=src_cols, width=32, state="readonly")
        cb_src.set(src_val if src_val else (src_cols[0] if src_cols else ""))
        cb_src.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(f, text="Cílový sloupec:").grid(row=1, column=0, sticky=tk.W, padx=5)
        cb_dst = ttk.Combobox(f, values=dst_cols, width=32, state="readonly")
        cb_dst.set(dst_val if dst_val else (dst_cols[0] if dst_cols else ""))
        cb_dst.grid(row=1, column=1, padx=5, pady=2)

        polozka = {"frame": f, "src": cb_src, "dst": cb_dst}
        btn_del = ttk.Button(f, text="✕", width=3, command=lambda: odebrat_dvojici(polozka))
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

    ttk.Button(bot_frame, text="+ Přidat další pravidlo", command=pridat_dvojici).pack(anchor=tk.W, pady=(0, 10))

    def ulozit():
        data_radku["src_file"] = cb_src_file.get()
        data_radku["dst_file"] = cb_dst_file.get()
        data_radku["klic_src"] = cb_klic_src.get()
        data_radku["klic_dst"] = cb_klic_dst.get()
        data_radku["konec_col"] = cb_konec_col.get()
        data_radku["mapovani_rules"] = [{"src": x["src"].get(), "dst": x["dst"].get()} for x in ui_prvky]

        data_radku["btn_upravit"].config(text=f"⚙ Mapování [{cb_src_file.get()}] ➔ [{cb_dst_file.get()}] ({len(ui_prvky)})")
        popup.destroy()

    ttk.Button(bot_frame, text="Uložit mapování", command=ulozit).pack(fill=tk.X, ipady=4)


# --- POP-UP OKNO PRO KOPÍROVÁNÍ PODLE 2 KLÍČŮ (UNIVERZÁLNÍ) ---
def otevrit_popup_mapovani_2_klice(data_radku):
    popup = tk.Toplevel(root)
    popup.title("Kopírování podle DVOU klíčů")
    popup.geometry("750x680")
    popup.grab_set()

    f_help = ttk.Frame(popup, padding=(12, 12, 12, 0))
    f_help.pack(fill=tk.X)
    pridat_box_napovedy(f_help, "Páruje řádky podle kombinace DVOU klíčů současně (např. P/N kód + Typ zástupce). Data se překopírují pouze do řádků, kde se shodují oba klíče najednou.")

    f_top = ttk.Frame(popup, padding=12)
    f_top.pack(fill=tk.X)

    seznam_aliasu = list(nactene_soubory.keys())

    ttk.Label(f_top, text="Zdrojový soubor:", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
    cb_src_file = ttk.Combobox(f_top, values=seznam_aliasu, width=28, state="readonly")
    cb_src_file.set(data_radku.get("src_file", seznam_aliasu[0]))
    cb_src_file.grid(row=0, column=1, padx=5, pady=2)

    ttk.Label(f_top, text="Cílový soubor:", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, sticky=tk.W, padx=(20, 5), pady=2)
    cb_dst_file = ttk.Combobox(f_top, values=seznam_aliasu, width=28, state="readonly")
    cb_dst_file.set(data_radku.get("dst_file", seznam_aliasu[-1] if len(seznam_aliasu) > 1 else seznam_aliasu[0]))
    cb_dst_file.grid(row=0, column=3, padx=5, pady=2)

    ttk.Label(f_top, text="Klíč 1 (Zdroj):", font=("Segoe UI", 9)).grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
    cb_k1_src = ttk.Combobox(f_top, width=28, state="readonly")
    cb_k1_src.grid(row=1, column=1, padx=5, pady=2)

    ttk.Label(f_top, text="Klíč 1 (Cíl):", font=("Segoe UI", 9)).grid(row=1, column=2, sticky=tk.W, padx=(20, 5), pady=2)
    cb_k1_dst = ttk.Combobox(f_top, width=28, state="readonly")
    cb_k1_dst.grid(row=1, column=3, padx=5, pady=2)

    ttk.Label(f_top, text="Klíč 2 (Zdroj):", font=("Segoe UI", 9)).grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
    cb_k2_src = ttk.Combobox(f_top, width=28, state="readonly")
    cb_k2_src.grid(row=2, column=1, padx=5, pady=2)

    ttk.Label(f_top, text="Klíč 2 (Cíl):", font=("Segoe UI", 9)).grid(row=2, column=2, sticky=tk.W, padx=(20, 5), pady=2)
    cb_k2_dst = ttk.Combobox(f_top, width=28, state="readonly")
    cb_k2_dst.grid(row=2, column=3, padx=5, pady=2)

    ttk.Label(f_top, text="Konec dat podle (Zdroj):", font=("Segoe UI", 9)).grid(row=3, column=0, sticky=tk.W, padx=5, pady=(6, 2))
    cb_konec_col = ttk.Combobox(f_top, width=28, state="readonly")
    cb_konec_col.grid(row=3, column=1, padx=5, pady=(6, 2))

    def refresh_cols(e=None):
        src_cols = nactene_soubory.get(cb_src_file.get(), {}).get("sloupce", [])
        dst_cols = nactene_soubory.get(cb_dst_file.get(), {}).get("sloupce", [])

        cb_k1_src["values"] = src_cols
        if src_cols: cb_k1_src.set(data_radku.get("k1_src", src_cols[0]))
        cb_k1_dst["values"] = dst_cols
        if dst_cols: cb_k1_dst.set(data_radku.get("k1_dst", dst_cols[0]))

        cb_k2_src["values"] = src_cols
        if src_cols: cb_k2_src.set(data_radku.get("k2_src", src_cols[0]))
        cb_k2_dst["values"] = dst_cols
        if dst_cols: cb_k2_dst.set(data_radku.get("k2_dst", dst_cols[0]))

        cb_konec_col["values"] = src_cols
        if src_cols: cb_konec_col.set(data_radku.get("konec_col", src_cols[0]))

    cb_src_file.bind("<<ComboboxSelected>>", refresh_cols)
    cb_dst_file.bind("<<ComboboxSelected>>", refresh_cols)
    refresh_cols()

    canvas = tk.Canvas(popup, borderwidth=0, highlightthickness=0)
    scrollbar = ttk.Scrollbar(popup, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas)

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    pop_win_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(pop_win_id, width=e.width))

    canvas.configure(yscrollcommand=scrollbar.set)
    pripojit_scrollovani_koleckem(canvas)

    canvas.pack(side="top", fill="both", expand=True, padx=10, pady=5)
    scrollbar.pack(side="right", fill="y")

    ui_prvky = []

    def pridat_dvojici(src_val=None, dst_val=None):
        f = ttk.LabelFrame(scroll_frame, text=f" Pravidlo {len(ui_prvky) + 1} ", padding=8)
        f.pack(fill=tk.X, expand=True, pady=4, padx=5)

        src_cols = nactene_soubory.get(cb_src_file.get(), {}).get("sloupce", [])
        dst_cols = nactene_soubory.get(cb_dst_file.get(), {}).get("sloupce", [])

        ttk.Label(f, text="Zdrojový sloupec:").grid(row=0, column=0, sticky=tk.W, padx=5)
        cb_src = ttk.Combobox(f, values=src_cols, width=32, state="readonly")
        cb_src.set(src_val if src_val else (src_cols[0] if src_cols else ""))
        cb_src.grid(row=0, column=1, padx=5, pady=2)

        ttk.Label(f, text="Cílový sloupec:").grid(row=1, column=0, sticky=tk.W, padx=5)
        cb_dst = ttk.Combobox(f, values=dst_cols, width=32, state="readonly")
        cb_dst.set(dst_val if dst_val else (dst_cols[0] if dst_cols else ""))
        cb_dst.grid(row=1, column=1, padx=5, pady=2)

        polozka = {"frame": f, "src": cb_src, "dst": cb_dst}
        btn_del = ttk.Button(f, text="✕", width=3, command=lambda: odebrat_dvojici(polozka))
        btn_del.grid(row=0, column=2, rowspan=2, padx=10)

        ui_prvky.append(polozka)
        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    def odebrat_dvojici(polozka):
        polozka["frame"].destroy()
        ui_prvky.remove(polozka)

    mapovani = data_radku.get("mapovani_rules", [])
    if mapovani:
        for r in mapovani: pridat_dvojici(r["src"], r["dst"])
    else:
        pridat_dvojici()

    bot_frame = ttk.Frame(popup, padding=10)
    bot_frame.pack(fill=tk.X, side=tk.BOTTOM)

    ttk.Button(bot_frame, text="+ Přidat další pravidlo", command=pridat_dvojici).pack(anchor=tk.W, pady=(0, 10))

    def ulozit():
        data_radku["src_file"] = cb_src_file.get()
        data_radku["dst_file"] = cb_dst_file.get()
        data_radku["k1_src"] = cb_k1_src.get()
        data_radku["k1_dst"] = cb_k1_dst.get()
        data_radku["k2_src"] = cb_k2_src.get()
        data_radku["k2_dst"] = cb_k2_dst.get()
        data_radku["konec_col"] = cb_konec_col.get()
        data_radku["mapovani_rules"] = [{"src": x["src"].get(), "dst": x["dst"].get()} for x in ui_prvky]

        data_radku["btn_upravit"].config(text=f"⚙ [2 Klíče] [{cb_src_file.get()}] ➔ [{cb_dst_file.get()}]")
        popup.destroy()

    ttk.Button(bot_frame, text="Uložit mapování", command=ulozit).pack(fill=tk.X, ipady=4)


# --- POP-UP OKNO PRO PODMÍNĚNÉ ROZDĚLENÍ ---
def otevrit_popup_podminene_rozdeleni(data_radku):
    popup = tk.Toplevel(root)
    popup.title("Podmíněné rozdělení dat (Procenta / Pololetí)")
    popup.geometry("680x680")
    popup.grab_set()

    f = ttk.Frame(popup, padding=15)
    f.pack(fill=tk.BOTH, expand=True)

    pridat_box_napovedy(f, "Sečte zadané sloupce (např. pololetí). Pokud text příjemce obsahuje přesnou zkratku (např. '/MA'), vloží 100 % do IAM. Jinak rozpočítá zadaná procenta mezi OE a OES.")

    seznam_aliasu = list(nactene_soubory.keys())

    box_files = ttk.LabelFrame(f, text=" 1. Soubory a Párovací P/N kód ", padding=10)
    box_files.pack(fill=tk.X, pady=(0, 10))

    ttk.Label(box_files, text="Zdrojový s.:").grid(row=0, column=0, sticky=tk.W)
    cb_src = ttk.Combobox(box_files, values=seznam_aliasu, width=20, state="readonly")
    cb_src.set(data_radku.get("src_file", seznam_aliasu[0]))
    cb_src.grid(row=0, column=1, padx=5, pady=2)

    ttk.Label(box_files, text="Cílový s.:").grid(row=0, column=2, sticky=tk.W, padx=(15, 0))
    cb_dst = ttk.Combobox(box_files, values=seznam_aliasu, width=20, state="readonly")
    cb_dst.set(data_radku.get("dst_file", seznam_aliasu[-1] if len(seznam_aliasu) > 1 else seznam_aliasu[0]))
    cb_dst.grid(row=0, column=3, padx=5, pady=2)

    ttk.Label(box_files, text="P/N Zdroj:").grid(row=1, column=0, sticky=tk.W)
    cb_pn_src = ttk.Combobox(box_files, width=20, state="readonly")
    cb_pn_src.grid(row=1, column=1, padx=5, pady=2)

    ttk.Label(box_files, text="P/N Cíl:").grid(row=1, column=2, sticky=tk.W, padx=(15, 0))
    cb_pn_dst = ttk.Combobox(box_files, width=20, state="readonly")
    cb_pn_dst.grid(row=1, column=3, padx=5, pady=2)

    box_calc = ttk.LabelFrame(f, text=" 2. Vstupní data u Zákazníka ", padding=10)
    box_calc.pack(fill=tk.X, pady=(0, 10))

    ttk.Label(box_calc, text="Sloupec s příjemcem:").grid(row=0, column=0, sticky=tk.W)
    cb_rec_col = ttk.Combobox(box_calc, width=20, state="readonly")
    cb_rec_col.grid(row=0, column=1, padx=5, pady=2)

    ttk.Label(box_calc, text="Hledaný text (interní):").grid(row=0, column=2, sticky=tk.W, padx=(15, 0))
    txt_filter = ttk.Entry(box_calc, width=15)
    txt_filter.insert(0, data_radku.get("filter_text", "/MA"))
    txt_filter.grid(row=0, column=3, padx=5, pady=2)

    ttk.Label(box_calc, text="Sčítané sloupce u Zákazníka (např. pololetí):").grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 2))
    cb_sum1 = ttk.Combobox(box_calc, width=20, state="readonly")
    cb_sum1.grid(row=2, column=0, columnspan=2, sticky=tk.W, padx=5)

    cb_sum2 = ttk.Combobox(box_calc, width=20, state="readonly")
    cb_sum2.grid(row=2, column=2, columnspan=2, sticky=tk.W, padx=5)

    box_dst = ttk.LabelFrame(f, text=" 3. Cílové umístění v Naší tabulce ", padding=10)
    box_dst.pack(fill=tk.X, pady=(0, 10))

    ttk.Label(box_dst, text="Cílový sloupec s kategorií (např. Sloupec S):").grid(row=0, column=0, sticky=tk.W)
    cb_cat_col = ttk.Combobox(box_dst, width=22, state="readonly")
    cb_cat_col.grid(row=0, column=1, padx=5, pady=2)

    ttk.Label(box_dst, text="Cílový hodnotový sloupec (např. Rok 2026 / AT):").grid(row=1, column=0, sticky=tk.W)
    cb_val_col = ttk.Combobox(box_dst, width=22, state="readonly")
    cb_val_col.grid(row=1, column=1, padx=5, pady=2)

    box_pct = ttk.LabelFrame(f, text=" 4. Pravidla procentuálního rozdělení ", padding=10)
    box_pct.pack(fill=tk.X, pady=(0, 10))

    ttk.Label(box_pct, text="Když obsahuje hledaný text ➔ do kategorie:", font=("Segoe UI", 8, "bold")).grid(row=0, column=0, columnspan=2, sticky=tk.W)
    txt_m_cat = ttk.Entry(box_pct, width=10); txt_m_cat.insert(0, data_radku.get("m_cat", "IAM"))
    txt_m_cat.grid(row=0, column=2, padx=2)
    ttk.Label(box_pct, text="Procento:").grid(row=0, column=3, sticky=tk.W)
    txt_m_pct = ttk.Entry(box_pct, width=6); txt_m_pct.insert(0, data_radku.get("m_pct", "100"))
    txt_m_pct.grid(row=0, column=4, padx=2)

    ttk.Label(box_pct, text="Jinak (Ostatní) ➔ Rozdělit do kategorií:", font=("Segoe UI", 8, "bold")).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
    txt_e1_cat = ttk.Entry(box_pct, width=10); txt_e1_cat.insert(0, data_radku.get("e1_cat", "OE"))
    txt_e1_cat.grid(row=1, column=2, padx=2)
    txt_e1_pct = ttk.Entry(box_pct, width=6); txt_e1_pct.insert(0, data_radku.get("e1_pct", "30"))
    txt_e1_pct.grid(row=1, column=3, padx=2)

    txt_e2_cat = ttk.Entry(box_pct, width=10); txt_e2_cat.insert(0, data_radku.get("e2_cat", "OES"))
    txt_e2_cat.grid(row=1, column=4, padx=2)
    txt_e2_pct = ttk.Entry(box_pct, width=6); txt_e2_pct.insert(0, data_radku.get("e2_pct", "70"))
    txt_e2_pct.grid(row=1, column=5, padx=2)

    def refresh_cols(e=None):
        src_cols = nactene_soubory.get(cb_src.get(), {}).get("sloupce", [])
        dst_cols = nactene_soubory.get(cb_dst.get(), {}).get("sloupce", [])

        cb_pn_src["values"] = src_cols
        if src_cols: cb_pn_src.set(data_radku.get("pn_src", src_cols[0]))
        cb_pn_dst["values"] = dst_cols
        if dst_cols: cb_pn_dst.set(data_radku.get("pn_dst", dst_cols[0]))

        cb_rec_col["values"] = src_cols
        if src_cols: cb_rec_col.set(data_radku.get("rec_col", src_cols[0]))

        cb_sum1["values"] = src_cols
        if src_cols: cb_sum1.set(data_radku.get("sum1", src_cols[0]))
        cb_sum2["values"] = ["-- Žádný --"] + src_cols
        if src_cols: cb_sum2.set(data_radku.get("sum2", src_cols[1] if len(src_cols) > 1 else "-- Žádný --"))

        cb_cat_col["values"] = dst_cols
        if dst_cols: cb_cat_col.set(data_radku.get("cat_col", dst_cols[0]))
        cb_val_col["values"] = dst_cols
        if dst_cols: cb_val_col.set(data_radku.get("val_col", dst_cols[0]))

    cb_src.bind("<<ComboboxSelected>>", refresh_cols)
    cb_dst.bind("<<ComboboxSelected>>", refresh_cols)
    refresh_cols()

    def ulozit():
        data_radku["src_file"] = cb_src.get()
        data_radku["dst_file"] = cb_dst.get()
        data_radku["pn_src"] = cb_pn_src.get()
        data_radku["pn_dst"] = cb_pn_dst.get()
        data_radku["rec_col"] = cb_rec_col.get()
        data_radku["filter_text"] = txt_filter.get()
        data_radku["sum1"] = cb_sum1.get()
        data_radku["sum2"] = cb_sum2.get()
        data_radku["cat_col"] = cb_cat_col.get()
        data_radku["val_col"] = cb_val_col.get()
        data_radku["m_cat"] = txt_m_cat.get()
        data_radku["m_pct"] = txt_m_pct.get()
        data_radku["e1_cat"] = txt_e1_cat.get()
        data_radku["e1_pct"] = txt_e1_pct.get()
        data_radku["e2_cat"] = txt_e2_cat.get()
        data_radku["e2_pct"] = txt_e2_pct.get()

        data_radku["btn_upravit"].config(text=f"⚙ [Rozdělit %] {cb_src.get()} ➔ {cb_dst.get()} [{cb_val_col.get()}]")
        popup.destroy()

    ttk.Button(f, text="Uložit nastavení", command=ulozit).pack(side=tk.BOTTOM, fill=tk.X, ipady=4)


# --- POP-UP OKNO PRO SEČTENÍ VÍCE SLOUPCŮ ---
def otevrit_popup_secist_více_sloupců(data_radku):
    popup = tk.Toplevel(root)
    popup.title("Sečíst více sloupců do nového sloupce")
    popup.geometry("520x500")
    popup.grab_set()

    f = ttk.Frame(popup, padding=15)
    f.pack(fill=tk.BOTH, expand=True)

    pridat_box_napovedy(f, "Sečte hodnoty ze zaškrtnutých sloupců řádek po řádku a v daném souboru vytvoří nový sloupec s celkovým součtem.")

    seznam_aliasu = list(nactene_soubory.keys())

    ttk.Label(f, text="1. Vyber soubor:", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W, pady=(0, 2))
    cb_target_file = ttk.Combobox(f, values=seznam_aliasu, width=42, state="readonly")
    cb_target_file.set(data_radku.get("target_file", seznam_aliasu[0]))
    cb_target_file.pack(anchor=tk.W, pady=(0, 10))

    ttk.Label(f, text="2. Název nového sloupce pro součet:", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W, pady=(0, 2))
    txt_new_col = ttk.Entry(f, width=45)
    txt_new_col.insert(0, data_radku.get("new_col_name", "Součet_pololetí"))
    txt_new_col.pack(anchor=tk.W, pady=(0, 10))

    ttk.Label(f, text="3. Vyber sloupce k sečtení:", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W, pady=(0, 5))

    frame_cols = ttk.Frame(f)
    frame_cols.pack(fill=tk.BOTH, expand=True, pady=5)

    check_vars = {}

    def refresh_cols(e=None):
        for child in frame_cols.winfo_children(): child.destroy()
        check_vars.clear()

        cols = nactene_soubory.get(cb_target_file.get(), {}).get("sloupce", [])
        saved_sum = data_radku.get("sum_cols", [])

        for c in cols:
            var = tk.BooleanVar(value=(c in saved_sum))
            chk = ttk.Checkbutton(frame_cols, text=c, variable=var)
            chk.pack(anchor=tk.W, pady=1)
            check_vars[c] = var

    cb_target_file.bind("<<ComboboxSelected>>", refresh_cols)
    refresh_cols()

    def ulozit():
        selected = [col for col, var in check_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("Chyba", "Musíš vybrat alespoň jeden sloupec k sečtení!")
            return
        data_radku["target_file"] = cb_target_file.get()
        data_radku["new_col_name"] = txt_new_col.get().strip()
        data_radku["sum_cols"] = selected
        data_radku["btn_upravit"].config(text=f"⚙ [{cb_target_file.get()}] Sečíst ({len(selected)} sloupců) ➔ [{txt_new_col.get()}]")
        popup.destroy()

    ttk.Button(f, text="Uložit", command=ulozit).pack(side=tk.BOTTOM, fill=tk.X, ipady=4)


# --- POP-UP OKNO PRO OSTATNÍ DYN. OPERACE ---
def otevrit_popup_operace(data_radku):
    op = data_radku["operace"].get()

    if op == "Podmíněné rozdělení dat (Procenta / Pololetí)":
        otevrit_popup_podminene_rozdeleni(data_radku)
        return
    elif op == "Sečíst více sloupců do nového sloupce":
        otevrit_popup_secist_více_sloupců(data_radku)
        return
    elif op == "Přesunout/kopírovat data podle DVOU klíčů":
        otevrit_popup_mapovani_2_klice(data_radku)
        return

    popup = tk.Toplevel(root)
    popup.title(f"Nastavení: {op}")
    popup.geometry("500x440")
    popup.grab_set()

    frame_main = ttk.Frame(popup, padding=15)
    frame_main.pack(fill=tk.BOTH, expand=True)

    if op == "Sečíst duplicitní řádky":
        pridat_box_napovedy(frame_main, "Sloučí duplicitní řádky podle zvoleného sloupce. Číselné hodnoty u stejných položek sečte do jednoho řádku.")
    elif op == "Přičíst sloupec k jinému":
        pridat_box_napovedy(frame_main, "Přičte čísla ze zdrojového sloupce k cílovému sloupci v rámci jednoho souboru.")
    elif op == "Vyčistit sloupec":
        pridat_box_napovedy(frame_main, "Smaže a vyprázdní veškerá data v celém vybraném sloupci.")
    elif op == "Naplnit sloupec":
        pridat_box_napovedy(frame_main, "Vloží zadanou hodnotu/text do vybraného sloupce (lze omezit podle posledního řádku jiného sloupce).")

    seznam_aliasu = list(nactene_soubory.keys())

    ttk.Label(frame_main, text="1. Vyber upravovaný soubor:", font=("Segoe UI", 10, "bold")).pack(anchor=tk.W, pady=(0, 2))
    cb_target_file = ttk.Combobox(frame_main, values=seznam_aliasu, width=45, state="readonly")
    cb_target_file.set(data_radku.get("target_file", seznam_aliasu[0]))
    cb_target_file.pack(anchor=tk.W, pady=(0, 12))

    frame_dynamic = ttk.Frame(frame_main)
    frame_dynamic.pack(fill=tk.BOTH, expand=True)

    def nacist_sloupce_pro_vybrany_soubor():
        return nactene_soubory.get(cb_target_file.get(), {}).get("sloupce", [])

    if op == "Sečíst duplicitní řádky":
        ttk.Label(frame_dynamic, text="Sloučit duplicitní řádky podle sloupce:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
        cb_col = ttk.Combobox(frame_dynamic, width=45, state="readonly")
        cb_col.pack(anchor=tk.W, pady=(0, 10))

        def refresh_cols(e=None):
            cols = nacist_sloupce_pro_vybrany_soubor()
            cb_col["values"] = cols
            if cols: cb_col.set(data_radku.get("klic1", cols[0]))

        cb_target_file.bind("<<ComboboxSelected>>", refresh_cols)
        refresh_cols()

        def ulozit():
            data_radku["target_file"] = cb_target_file.get()
            data_radku["klic1"] = cb_col.get()
            data_radku["btn_upravit"].config(text=f"⚙ [{cb_target_file.get()}] Sloučit podle [{cb_col.get()}]")
            popup.destroy()

        ttk.Button(frame_main, text="Uložit", command=ulozit).pack(side=tk.BOTTOM, fill=tk.X, ipady=4)

    elif op == "Přičíst sloupec k jinému":
        ttk.Label(frame_dynamic, text="Zdrojový sloupec:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
        cb_s = ttk.Combobox(frame_dynamic, width=45, state="readonly")
        cb_s.pack(anchor=tk.W, pady=(0, 8))

        ttk.Label(frame_dynamic, text="Cílový sloupec (kam přičíst):", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
        cb_r = ttk.Combobox(frame_dynamic, width=45, state="readonly")
        cb_r.pack(anchor=tk.W, pady=(0, 8))

        lbl_note = ttk.Label(frame_dynamic, text="", font=("Segoe UI", 9, "italic"), foreground="#2e7d32")
        lbl_note.pack(anchor=tk.W, pady=(5, 10))

        def update_note(e=None):
            lbl_note.config(text=f"💡 Výsledný součet bude uložen ve sloupci: '{cb_r.get()}'")

        def refresh_cols(e=None):
            cols = nacist_sloupce_pro_vybrany_soubor()
            cb_s["values"] = cols
            cb_r["values"] = cols
            if cols:
                cb_s.set(data_radku.get("klic1", cols[0]))
                cb_r.set(data_radku.get("klic2", cols[0]))
            update_note()

        cb_s.bind("<<ComboboxSelected>>", update_note)
        cb_r.bind("<<ComboboxSelected>>", update_note)
        cb_target_file.bind("<<ComboboxSelected>>", refresh_cols)
        refresh_cols()

        def ulozit():
            data_radku["target_file"] = cb_target_file.get()
            data_radku["klic1"] = cb_s.get()
            data_radku["klic2"] = cb_r.get()
            data_radku["btn_upravit"].config(text=f"⚙ [{cb_target_file.get()}] Přičíst [{cb_s.get()}] ➔ [{cb_r.get()}]")
            popup.destroy()

        ttk.Button(frame_main, text="Uložit", command=ulozit).pack(side=tk.BOTTOM, fill=tk.X, ipady=4)

    elif op in ["Vyčistit sloupec", "Naplnit sloupec"]:
        ttk.Label(frame_dynamic, text="Vyber sloupec:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
        cb_col = ttk.Combobox(frame_dynamic, width=45, state="readonly")
        cb_col.pack(anchor=tk.W, pady=(0, 10))

        txt_val = None
        cb_konec = None

        if op == "Naplnit sloupec":
            ttk.Label(frame_dynamic, text="Zadej hodnotu pro naplnění:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
            txt_val = ttk.Entry(frame_dynamic, width=48)
            txt_val.insert(0, data_radku.get("hodnota_naplneni", ""))
            txt_val.pack(anchor=tk.W, pady=(0, 10))

            ttk.Label(frame_dynamic, text="Konec dat podle sloupce (volitelné):", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W, pady=(0, 2))
            cb_konec = ttk.Combobox(frame_dynamic, width=45, state="readonly")
            cb_konec.pack(anchor=tk.W, pady=(0, 10))

        def refresh_cols(e=None):
            cols = nacist_sloupce_pro_vybrany_soubor()
            cb_col["values"] = cols
            if cols: cb_col.set(data_radku.get("vybrany_sloupec", cols[0]))

            if cb_konec:
                cols_with_option = ["-- Celý sloupec --"] + cols
                cb_konec["values"] = cols_with_option
                preset = data_radku.get("konec_col", "-- Celý sloupec --")
                cb_konec.set(preset if preset in cols_with_option else "-- Celý sloupec --")

        cb_target_file.bind("<<ComboboxSelected>>", refresh_cols)
        refresh_cols()

        def ulozit():
            data_radku["target_file"] = cb_target_file.get()
            data_radku["vybrany_sloupec"] = cb_col.get()
            if op == "Naplnit sloupec":
                if txt_val: data_radku["hodnota_naplneni"] = txt_val.get()
                if cb_konec: data_radku["konec_col"] = cb_konec.get()
                data_radku["btn_upravit"].config(text=f"⚙ [{cb_target_file.get()}] Naplnit [{cb_col.get()}] = '{txt_val.get()}'")
            else:
                data_radku["btn_upravit"].config(text=f"⚙ [{cb_target_file.get()}] Vyčistit [{cb_col.get()}]")
            popup.destroy()

        ttk.Button(frame_main, text="Uložit", command=ulozit).pack(side=tk.BOTTOM, fill=tk.X, ipady=4)


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
                "src_file": radek.get("src_file", ""),
                "dst_file": radek.get("dst_file", ""),
                "target_file": radek.get("target_file", ""),
                "klic_src": radek.get("klic_src", ""),
                "klic_dst": radek.get("klic_dst", ""),
                "k1_src": radek.get("k1_src", ""), "k1_dst": radek.get("k1_dst", ""),
                "k2_src": radek.get("k2_src", ""), "k2_dst": radek.get("k2_dst", ""),
                "konec_col": radek.get("konec_col", ""),
                "mapovani_rules": radek.get("mapovani_rules", []),
                "klic1": radek.get("klic1", ""),
                "klic2": radek.get("klic2", ""),
                "vybrany_sloupec": radek.get("vybrany_sloupec", ""),
                "hodnota_naplneni": radek.get("hodnota_naplneni", ""),
                "pn_src": radek.get("pn_src", ""), "pn_dst": radek.get("pn_dst", ""),
                "rec_col": radek.get("rec_col", ""), "filter_text": radek.get("filter_text", ""),
                "sum1": radek.get("sum1", ""), "sum2": radek.get("sum2", ""),
                "cat_col": radek.get("cat_col", ""), "val_col": radek.get("val_col", ""),
                "m_cat": radek.get("m_cat", ""), "m_pct": radek.get("m_pct", ""),
                "e1_cat": radek.get("e1_cat", ""), "e1_pct": radek.get("e1_pct", ""),
                "e2_cat": radek.get("e2_cat", ""), "e2_pct": radek.get("e2_pct", ""),
                "new_col_name": radek.get("new_col_name", ""), "sum_cols": radek.get("sum_cols", [])
            })

        profily = nacist_profily_ze_souboru()
        profily[nazev] = {
            "popis": txt_popis.get().strip(),
            "pouzite_soubory": list(nactene_soubory.keys()),
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
    popup.geometry("560x440")
    popup.grab_set()

    ttk.Label(popup, text="Uložené konfigurace operací:", font=("Segoe UI", 11, "bold")).pack(anchor=tk.W, padx=15, pady=10)

    canvas = tk.Canvas(popup, borderwidth=0, highlightthickness=0)
    scrollbar = ttk.Scrollbar(popup, orient="vertical", command=canvas.yview)
    scroll_frame = ttk.Frame(canvas)

    scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    pop_win_id = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfig(pop_win_id, width=e.width))

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
            box.pack(fill=tk.X, expand=True, pady=6, padx=5)

            popis_text = p_data.get('popis', 'Bez popisu')
            if not popis_text: popis_text = "Bez popisu"

            pouzite_s = ", ".join(p_data.get("pouzite_soubory", []))
            lbl_p = ttk.Label(
                box,
                text=f"Popis: {popis_text}\nOčekávané soubory: {pouzite_s}",
                font=("Segoe UI", 9, "italic"),
                foreground="#444444",
                wraplength=480,
                justify="left"
            )
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

            for key in ["src_file", "dst_file", "target_file", "klic_src", "klic_dst", "k1_src", "k1_dst", "k2_src", "k2_dst",
                        "konec_col", "mapovani_rules", "klic1", "klic2", "vybrany_sloupec", "hodnota_naplneni",
                        "pn_src", "pn_dst", "rec_col", "filter_text", "sum1", "sum2", "cat_col", "val_col",
                        "m_cat", "m_pct", "e1_cat", "e1_pct", "e2_cat", "e2_pct", "new_col_name", "sum_cols"]:
                if key in krok: radek_data[key] = krok[key]

            op = krok.get("operace")
            if op == "Přesunout/kopírovat data mezi soubory":
                radek_data["btn_upravit"].config(text=f"⚙ Mapování [{radek_data.get('src_file')}] ➔ [{radek_data.get('dst_file')}]")
            elif op == "Podmíněné rozdělení dat (Procenta / Pololetí)":
                radek_data["btn_upravit"].config(text=f"⚙ [Rozdělit %] {radek_data.get('src_file')} ➔ {radek_data.get('dst_file')}")
            elif op == "Sečíst více sloupců do nového sloupce":
                radek_data["btn_upravit"].config(text=f"⚙ [{radek_data.get('target_file')}] Sečíst ➔ [{radek_data.get('new_col_name')}]")
            elif op == "Přesunout/kopírovat data podle DVOU klíčů":
                radek_data["btn_upravit"].config(text=f"⚙ [2 Klíče] [{radek_data.get('src_file')}] ➔ [{radek_data.get('dst_file')}]")

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
        u_nazev = ttk.Entry(uf, width=48); u_nazev.insert(0, stary_nazev); u_nazev.pack(anchor=tk.W, pady=(0, 10))

        ttk.Label(uf, text="Popis:", font=("Segoe UI", 9, "bold")).pack(anchor=tk.W)
        u_popis = ttk.Entry(uf, width=48); u_popis.insert(0, data.get("popis", "")); u_popis.pack(anchor=tk.W, pady=(0, 10))

        def ulozit_zmeny():
            novy_nazev = u_nazev.get().strip()
            if not novy_nazev: return
            all_p = nacist_profily_ze_souboru()
            if stary_nazev in all_p: del all_p[stary_nazev]
            all_p[novy_nazev] = {
                "popis": u_popis.get().strip(),
                "pouzite_soubory": data.get("pouzite_soubory", []),
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


# --- STRÁNKA 2: SEZNAM KROKŮ ---
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
        radek["frame"].pack(fill=tk.X, expand=True, pady=4, padx=5)

    canvas_kroky.update_idletasks()
    canvas_kroky.configure(scrollregion=canvas_kroky.bbox("all"))


def zmena_operace(cb_operace, btn_upravit, data_radku):
    op = cb_operace.get()
    if op == "Přesunout/kopírovat data mezi soubory":
        btn_upravit.config(text="⚙ Mapovat sloupce", command=lambda: otevrit_popup_mapovani(data_radku))
    else:
        btn_upravit.config(text="⚙ Nastavit", command=lambda: otevrit_popup_operace(data_radku))


def pridat_radek_operace():
    radek_frame = ttk.LabelFrame(scrollable_frame_kroky, text=f" Krok {len(seznam_operaci) + 1} ", padding=8)
    radek_frame.pack(fill=tk.X, expand=True, pady=4, padx=5)

    cb_poradi = ttk.Combobox(radek_frame, values=[str(i) for i in range(1, 31)], width=3, state="readonly")
    cb_poradi.set(str(len(seznam_operaci) + 1))
    cb_poradi.pack(side=tk.LEFT, padx=(5, 10))

    cb_poradi.bind("<<ComboboxSelected>>", lambda e: preskladat_radky_operaci(e, data_radku))

    dostupne_operace = [
        "Přesunout/kopírovat data mezi soubory",
        "Podmíněné rozdělení dat (Procenta / Pololetí)",
        "Sečíst více sloupců do nového sloupce",
        "Přesunout/kopírovat data podle DVOU klíčů",
        "Sečíst duplicitní řádky",
        "Přičíst sloupec k jinému",
        "Vyčistit sloupec",
        "Naplnit sloupec",
    ]
    cb_operace = ttk.Combobox(radek_frame, values=dostupne_operace, width=38, state="readonly")
    cb_operace.set("Přesunout/kopírovat data mezi soubory")
    cb_operace.pack(side=tk.LEFT, padx=5)

    btn_upravit = ttk.Button(radek_frame, text="⚙ Mapovat sloupce")
    btn_upravit.pack(side=tk.LEFT, padx=15)

    prvni_alias = list(nactene_soubory.keys())[0] if nactene_soubory else ""

    data_radku = {
        "frame": radek_frame,
        "poradi": cb_poradi,
        "operace": cb_operace,
        "btn_upravit": btn_upravit,
        "src_file": prvni_alias, "dst_file": prvni_alias, "target_file": prvni_alias,
        "mapovani_rules": [], "klic_src": "", "klic_dst": "", "konec_col": "",
        "klic1": "", "klic2": "", "vybrany_sloupec": "", "hodnota_naplneni": ""
    }

    zmena_operace(cb_operace, btn_upravit, data_radku)
    cb_operace.bind("<<ComboboxSelected>>", lambda e: zmena_operace(cb_operace, btn_upravit, data_radku))

    btn_smazat = ttk.Button(radek_frame, text="✕", width=3, command=lambda: smazat_radek_operace(radek_frame, data_radku))
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
        dfs = {}
        for alias, info in nactene_soubory.items():
            df = nacist_excel_bezpecne(info["cesta"], sheet_name=info["sheet"], header=info["hdr"])
            if df is None:
                messagebox.showerror("Chyba", f"Nepodařilo se načíst soubor '{alias}'.")
                return
            dfs[alias] = df

        for radek in seznam_operaci:
            op = radek["operace"].get()

            if op == "Sečíst duplicitní řádky":
                t_file = radek.get("target_file")
                odkud = radek.get("klic1")
                if t_file in dfs and odkud in dfs[t_file].columns:
                    agg_dict = {
                        col: ("sum" if pd.api.types.is_numeric_dtype(dfs[t_file][col]) else "first")
                        for col in dfs[t_file].columns if col != odkud
                    }
                    dfs[t_file] = dfs[t_file].groupby(odkud, as_index=False).agg(agg_dict)

            elif op == "Přičíst sloupec k jinému":
                t_file = radek.get("target_file")
                odkud, kam = radek.get("klic1"), radek.get("klic2")
                if t_file in dfs and odkud in dfs[t_file].columns and kam in dfs[t_file].columns:
                    zdroj = pd.to_numeric(dfs[t_file][odkud], errors="coerce").fillna(0)
                    cil = pd.to_numeric(dfs[t_file][kam], errors="coerce").fillna(0)
                    dfs[t_file][kam] = cil + zdroj

            elif op == "Sečíst více sloupců do nového sloupce":
                t_file = radek.get("target_file")
                new_col = radek.get("new_col_name")
                cols = radek.get("sum_cols", [])
                if t_file in dfs and new_col and cols:
                    dfs[t_file][new_col] = dfs[t_file][cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)

            elif op == "Přesunout/kopírovat data mezi soubory":
                s_file, d_file = radek.get("src_file"), radek.get("dst_file")
                klic_s, klic_d = radek.get("klic_src"), radek.get("klic_dst")
                konec_col = radek.get("konec_col")
                pravidla = radek.get("mapovani_rules", [])

                if s_file in dfs and d_file in dfs and klic_s and klic_d and konec_col:
                    df_src, df_dst = dfs[s_file], dfs[d_file]
                    posledni_idx = df_src[konec_col].dropna().index.max()
                    if pd.isna(posledni_idx): posledni_idx = len(df_src) - 1
                    df_src_oriznuto = df_src.loc[0:posledni_idx]

                    for rule in pravidla:
                        src_col, dst_col = rule["src"], rule["dst"]
                        if src_col in df_src_oriznuto.columns and dst_col in df_dst.columns:
                            mapa = dict(zip(df_src_oriznuto[klic_s], df_src_oriznuto[src_col]))
                            df_dst[dst_col] = df_dst[klic_d].map(mapa).fillna(df_dst[dst_col])

            elif op == "Přesunout/kopírovat data podle DVOU klíčů":
                s_file, d_file = radek.get("src_file"), radek.get("dst_file")
                k1_s, k1_d = radek.get("k1_src"), radek.get("k1_dst")
                k2_s, k2_d = radek.get("k2_src"), radek.get("k2_dst")
                konec_col = radek.get("konec_col")
                pravidla = radek.get("mapovani_rules", [])

                if s_file in dfs and d_file in dfs and k1_s and k1_d and k2_s and k2_d:
                    df_src, df_dst = dfs[s_file], dfs[d_file]
                    if konec_col and konec_col in df_src.columns:
                        posledni_idx = df_src[konec_col].dropna().index.max()
                        if pd.isna(posledni_idx): posledni_idx = len(df_src) - 1
                        df_src_oriznuto = df_src.loc[0:posledni_idx]
                    else:
                        df_src_oriznuto = df_src

                    comp_src = df_src_oriznuto[k1_s].astype(str).str.strip().str.lower() + "|||" + df_src_oriznuto[k2_s].astype(str).str.strip().str.lower()
                    comp_dst = df_dst[k1_d].astype(str).str.strip().str.lower() + "|||" + df_dst[k2_d].astype(str).str.strip().str.lower()

                    for rule in pravidla:
                        src_col, dst_col = rule["src"], rule["dst"]
                        if src_col in df_src_oriznuto.columns and dst_col in df_dst.columns:
                            mapa = dict(zip(comp_src, df_src_oriznuto[src_col]))
                            df_dst[dst_col] = comp_dst.map(mapa).fillna(df_dst[dst_col])

            elif op == "Podmíněné rozdělení dat (Procenta / Pololetí)":
                s_file, d_file = radek.get("src_file"), radek.get("dst_file")
                pn_s, pn_d = radek.get("pn_src"), radek.get("pn_dst")
                rec_col = radek.get("rec_col")

                f_text = radek.get("filter_text", "").strip()

                sum1, sum2 = radek.get("sum1"), radek.get("sum2")
                cat_col, val_col = radek.get("cat_col"), radek.get("val_col")

                m_cat = radek.get("m_cat", "IAM").strip().upper()
                m_pct = float(radek.get("m_pct", "100") or 0) / 100.0
                e1_cat = radek.get("e1_cat", "OE").strip().upper()
                e1_pct = float(radek.get("e1_pct", "30") or 0) / 100.0
                e2_cat = radek.get("e2_cat", "OES").strip().upper()
                e2_pct = float(radek.get("e2_pct", "70") or 0) / 100.0

                if s_file in dfs and d_file in dfs and pn_s and pn_d and cat_col and val_col:
                    df_src, df_dst = dfs[s_file], dfs[d_file]

                    dst_pn_clean = df_dst[pn_d].astype(str).str.strip().str.lower()
                    dst_cat_clean = df_dst[cat_col].astype(str).str.strip().str.upper()

                    for idx, row in df_src.iterrows():
                        pn_val = str(row[pn_s]).strip().lower() if pd.notna(row[pn_s]) else ""
                        if not pn_val or pn_val == "nan": continue

                        rec_val = str(row[rec_col]) if rec_col and pd.notna(row[rec_col]) else ""

                        val_sum = 0
                        if sum1 and sum1 in df_src.columns:
                            val_sum += float(pd.to_numeric(row[sum1], errors="coerce") or 0)
                        if sum2 and sum2 != "-- Žádný --" and sum2 in df_src.columns:
                            val_sum += float(pd.to_numeric(row[sum2], errors="coerce") or 0)

                        if f_text and f_text in rec_val:
                            mask = (dst_pn_clean == pn_val) & (dst_cat_clean == m_cat)
                            existing = pd.to_numeric(df_dst.loc[mask, val_col], errors="coerce").fillna(0)
                            df_dst.loc[mask, val_col] = existing + (val_sum * m_pct)
                        else:
                            if e1_cat:
                                mask1 = (dst_pn_clean == pn_val) & (dst_cat_clean == e1_cat)
                                existing1 = pd.to_numeric(df_dst.loc[mask1, val_col], errors="coerce").fillna(0)
                                df_dst.loc[mask1, val_col] = existing1 + (val_sum * e1_pct)
                            if e2_cat:
                                mask2 = (dst_pn_clean == pn_val) & (dst_cat_clean == e2_cat)
                                existing2 = pd.to_numeric(df_dst.loc[mask2, val_col], errors="coerce").fillna(0)
                                df_dst.loc[mask2, val_col] = existing2 + (val_sum * e2_pct)

            elif op == "Vyčistit sloupec":
                t_file = radek.get("target_file")
                col = radek.get("vybrany_sloupec")
                if t_file in dfs and col in dfs[t_file].columns:
                    dfs[t_file][col] = None

            elif op == "Naplnit sloupec":
                t_file = radek.get("target_file")
                col = radek.get("vybrany_sloupec")
                val = radek.get("hodnota_naplneni", "")
                konec_c = radek.get("konec_col", "")

                if t_file in dfs and col in dfs[t_file].columns:
                    if konec_c and konec_c != "-- Celý sloupec --" and konec_c in dfs[t_file].columns:
                        posledni_idx = dfs[t_file][konec_c].dropna().index.max()
                        if pd.notna(posledni_idx):
                            dfs[t_file].loc[0:posledni_idx, col] = val
                        else:
                            dfs[t_file][col] = val
                    else:
                        dfs[t_file][col] = val

        # Uložení výsledku
        if len(nactene_soubory) == 1:
            alias = list(nactene_soubory.keys())[0]
            if var_single_output.get() == "overwrite":
                vystupni_cesta = nactene_soubory[alias]["cesta"]
            else:
                vystupni_cesta = filedialog.asksaveasfilename(
                    filetypes=[("Excel (.xlsx)", "*.xlsx")],
                    defaultextension=".xlsx",
                    initialdir=os.path.dirname(nactene_soubory[alias]["cesta"])
                )
                if not vystupni_cesta: return
            sheet_out = nactene_soubory[alias]["sheet"]
            df_out = dfs[alias]
        else:
            vystupni_cesta = os.path.splitext(cesta_cil_global)[0] + ".xlsx"
            alias_out = list(nactene_soubory.keys())[-1]
            sheet_out = nactene_soubory[alias_out]["sheet"]
            df_out = dfs[alias_out]

        with pd.ExcelWriter(vystupni_cesta, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False, sheet_name=sheet_out)

        messagebox.showinfo(
            "Hotovo",
            f"Zpracování proběhlo úspěšně!\nUloženo do:\n{os.path.basename(vystupni_cesta)}",
        )
        root.destroy()

    except PermissionError:
        messagebox.showerror(
            "Soubor je otevřený",
            "Nepodařilo se uložit výsledek!\nCílový soubor je právě otevřený v jiném programu (např. v Excelu).\nNejdříve ho prosím zavři.",
        )
    except Exception as e:
        messagebox.showerror("Chyba při zpracování", f"Něco kleklo v Pandas:\n{str(e)}")


def navrat_na_krok_1():
    frame_strana2.pack_forget()
    frame_strana1.pack(fill=tk.BOTH, expand=True)


# ================= HLAVNÍ OKNO =================
root = tk.Tk()
root.title("Automatizace konverze tabulek")
root.geometry("960x650")

style = ttk.Style()
style.theme_use("vista")

# ================= STRÁNKA 1: NAČTENÍ SOUBORŮ =================
frame_strana1 = ttk.Frame(root, padding="20")
frame_strana1.pack(fill=tk.BOTH, expand=True)

ttk.Label(
    frame_strana1,
    text="Krok 1: Správa vstupních souborů",
    font=("Segoe UI", 12, "bold"),
).pack(anchor=tk.W, pady=(0, 10))

canvas_soubory = tk.Canvas(frame_strana1, borderwidth=0, highlightthickness=0)
scrollbar_soubory = ttk.Scrollbar(frame_strana1, orient="vertical", command=canvas_soubory.yview)
scroll_frame_soubory = ttk.Frame(canvas_soubory)

scroll_frame_soubory.bind("<Configure>", lambda e: canvas_soubory.configure(scrollregion=canvas_soubory.bbox("all")))
win_soubory_id = canvas_soubory.create_window((0, 0), window=scroll_frame_soubory, anchor="nw")
canvas_soubory.bind("<Configure>", lambda e: canvas_soubory.itemconfig(win_soubory_id, width=e.width))

canvas_soubory.configure(yscrollcommand=scrollbar_soubory.set)
pripojit_scrollovani_koleckem(canvas_soubory)

canvas_soubory.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 10))
scrollbar_soubory.pack(side=tk.RIGHT, fill=tk.Y)

ttk.Button(frame_strana1, text="+ Přidat soubor", command=pridat_soubor_ui).pack(anchor=tk.W, pady=(0, 10))

f_vystup_container = ttk.Frame(frame_strana1)
f_vystup_container.pack(fill=tk.X, pady=(0, 15))

box_vystup_single = ttk.LabelFrame(f_vystup_container, text=" Volba uložení pro 1 soubor ", padding="10")
var_single_output = tk.StringVar(value="copy")
ttk.Radiobutton(box_vystup_single, text="Uložit jako nový soubor (Kopie)", variable=var_single_output, value="copy").pack(anchor=tk.W, pady=2)
ttk.Radiobutton(box_vystup_single, text="Přepsat původní originální soubor", variable=var_single_output, value="overwrite").pack(anchor=tk.W, pady=2)

box_vystup_multi = ttk.LabelFrame(f_vystup_container, text=" Výstupní soubor ", padding="10")
ttk.Button(box_vystup_multi, text="Určit kam uložit nový soubor...", command=vybrat_cilovy_soubor).pack(side=tk.LEFT, padx=(0, 10))
lbl_cil_path = ttk.Label(box_vystup_multi, text="Není vybráno", font=("Segoe UI", 9, "italic"), foreground="gray")
lbl_cil_path.pack(side=tk.LEFT)

pridat_soubor_ui()

ttk.Button(
    frame_strana1,
    text="Pokračovat k nastavení automatizace ➔",
    command=prejit_na_pracovni_plochu,
).pack(fill=tk.X, side=tk.BOTTOM, ipady=8)


# ================= STRÁNKA 2: PRACOVNÍ PLOCHA A OPERACE =================
frame_strana2 = ttk.Frame(root, padding="15")

frame_body = ttk.Frame(frame_strana2)
frame_body.pack(fill=tk.BOTH, expand=True)

frame_left = ttk.Frame(frame_body)
frame_left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

ttk.Label(
    frame_left,
    text="Seznam prováděných kroků automatizace",
    font=("Segoe UI", 11, "bold"),
).pack(anchor=tk.W, pady=(0, 5))

canvas_kroky = tk.Canvas(frame_left, borderwidth=0, highlightthickness=0)
scrollbar_kroky = ttk.Scrollbar(frame_left, orient="vertical", command=canvas_kroky.yview)
scrollable_frame_kroky = ttk.Frame(canvas_kroky)

scrollable_frame_kroky.bind("<Configure>", lambda e: canvas_kroky.configure(scrollregion=canvas_kroky.bbox("all")))
main_canvas_win = canvas_kroky.create_window((0, 0), window=scrollable_frame_kroky, anchor="nw")
canvas_kroky.bind("<Configure>", lambda e: canvas_kroky.itemconfig(main_canvas_win, width=e.width))

canvas_kroky.configure(yscrollcommand=scrollbar_kroky.set)
pripojit_scrollovani_koleckem(canvas_kroky)

canvas_kroky.pack(side="left", fill="both", expand=True)
scrollbar_kroky.pack(side="right", fill="y")


frame_right = ttk.Frame(frame_body, width=200)
frame_right.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))

box_profily = ttk.LabelFrame(frame_right, text=" Profily ", padding="10")
box_profily.pack(fill=tk.X, pady=(0, 15))

ttk.Button(box_profily, text="💾 Uložit profil", command=otevrit_popup_ulozit_profil).pack(fill=tk.X, pady=3, ipady=3)
ttk.Button(box_profily, text="📂 Načíst profil", command=otevrit_popup_nacist_profil).pack(fill=tk.X, pady=3, ipady=3)

box_kroky = ttk.LabelFrame(frame_right, text=" Úprava kroků ", padding="10")
box_kroky.pack(fill=tk.X, pady=(0, 15))

ttk.Button(box_kroky, text="+ Přidat další krok", command=pridat_radek_operace).pack(fill=tk.X, pady=3, ipady=3)

box_akce = ttk.LabelFrame(frame_right, text=" Akce ", padding="10")
box_akce.pack(fill=tk.X, side=tk.BOTTOM)

ttk.Button(box_akce, text="🚀 Spustit konverzi", command=spustit_konverzi).pack(fill=tk.X, pady=(3, 8), ipady=6)
ttk.Button(box_akce, text="⮌ Zpět na výběr souborů", command=navrat_na_krok_1).pack(fill=tk.X, pady=3, ipady=3)

root.mainloop()
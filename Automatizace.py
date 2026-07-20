import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd

# Globální proměnné, do kterých si uložíme cesty a data
cesta_zakaznik = ""
cesta_cil = ""
script_dir = os.path.dirname(os.path.abspath(__file__))


# --- FUNKCE PRO STRÁNKU 1 (VÝBĚR SOUBORŮ) ---
def vybrat_zakaznika():
    global cesta_zakaznik
    soubor = filedialog.askopenfilename(
        filetypes=[("Excel", "*.xlsx *.xls *.xlsm")]
    )
    if soubor:
        cesta_zakaznik = soubor
        lbl_zakaznik.config(text=os.path.basename(soubor))


def vybrat_cil():
    global cesta_cil
    # Zeptáme se, kam a pod jakým názvem to chce uložit (umí vytvořit nový i přepsat starý)
    soubor = filedialog.asksaveasfilename(
        filetypes=[("Excel", "*.xlsx *.xlsm")],
        defaultextension=".xlsx",
        initialdir=os.path.dirname(cesta_zakaznik) if cesta_zakaznik else None,
    )
    if soubor:
        cesta_cil = soubor
        lbl_cil.config(text=os.path.basename(soubor))


def prejit_na_operace():
    if not cesta_zakaznik or not cesta_cil:
        messagebox.showwarning("Chyba", "Musíš vybrat soubor i cíl!")
        return

    # Skryjeme první stránku
    frame_strana1.pack_forget()
    # Zobrazíme druhou stránku
    frame_strana2.pack(fill=tk.BOTH, expand=True)


# --- FUNKCE PRO STRÁNKU 2 (OPERACE A KONVERZE) ---
def spustit_konverzi():
    try:
        # Najdeme naši šablonu
        sablona = os.path.join(script_dir, "nase_sablona.xlsx")
        if not os.path.exists(sablona):
            sablona = os.path.join(script_dir, "nase_sablona.xlsm")

        if not os.path.exists(sablona):
            messagebox.showerror("Chyba", "V adresáři chybí nase_sablona!")
            return

        # Načtení Excelů do Pandas
        df_nase = pd.read_excel(sablona, engine="openpyxl")
        df_zakaznik = pd.read_excel(cesta_zakaznik, engine="openpyxl")

        # 1. HLAVNÍ OPERACE: Spárování podle součástek
        vysledek = pd.merge(
            df_nase,
            df_zakaznik,
            left_on="Cislo Soucastky",
            right_on="ID dilu",
            how="left",
        )

        # TODO: Tady se pak jednoduše přidají další zaškrtnuté operace
        # if chk_secti.get():
        #     vysledek['NovaCena'] = vysledek['Sloupec1'] + vysledek['Sloupec2']

        # Uložení výsledku (zachová makra, pokud ukládáme do .xlsm)
        is_xlsm = cesta_cil.lower().endswith(".xlsm")
        with pd.ExcelWriter(
            cesta_cil,
            engine="openpyxl",
            engine_kwargs={"keep_vba": True} if is_xlsm else {},
        ) as writer:
            vysledek.to_excel(writer, index=False)

        messagebox.showinfo("Hotovo", "Tabulka byla úspěšně vygenerována!")
        root.destroy()  # Zavře program po úspěchu

    except Exception as e:
        messagebox.showerror("Chyba", f"Něco kleklo:\n{str(e)}")


def navrat_zpet():
    # Schováme dvojku, ukážeme jedničku
    frame_strana2.pack_forget()
    frame_frame_strana1.pack(fill=tk.BOTH, expand=True)


# --- HLAVNÍ OKNO ---
root = tk.Tk()
root.title("Automatizace tabulek")
root.geometry("500x350")
root.resizable(False, False)

# Společný styl
style = ttk.Style()
style.theme_use("vista")

# ================= STRÁNKA 1: VÝBĚR =================
frame_strana1 = ttk.Frame(root, padding="20")
frame_strana1.pack(fill=tk.BOTH, expand=True)

ttk.Label(
    frame_strana1, text="1. Vyber soubor od zákazníka", font=("Segoe UI", 10)
).pack(pady=(10, 5))
ttk.Button(frame_strana1, text="Procházet...", command=vybrat_zakaznika).pack()
lbl_zakaznik = ttk.Label(
    frame_strana1, text="Není vybráno", font=("Segoe UI", 9, "italic")
)
lbl_zakaznik.pack(pady=(0, 20))

ttk.Label(frame_strana1, text="2. Kam uložit výsledek?", font=("Segoe UI", 10)).pack(
    pady=(10, 5)
)
ttk.Button(frame_strana1, text="Určit cíl a název...", command=vybrat_cil).pack()
lbl_cil = ttk.Label(
    frame_strana1, text="Není vybráno", font=("Segoe UI", 9, "italic")
)
lbl_cil.pack(pady=(0, 30))

ttk.Button(frame_strana1, text="Pokračovat ➔", command=prejit_na_operace).pack(
    fill=tk.X, ipady=5
)

# ================= STRÁNKA 2: OPERACE =================
frame_strana2 = ttk.Frame(root, padding="20")
# Tuhle stránku teď nespouštíme (.pack() zavolá až funkce)

ttk.Label(
    frame_strana2, text="Nastavení dodatečných operací", font=("Segoe UI", 12, "bold")
).pack(pady=10)

# Místo pro checkboxy na operace
ttk.Label(
    frame_strana2, text="[ Zde pak přidáme zaškrtávátka pro akce ]", foreground="gray"
).pack(pady=30)

# Spodní tlačítka na druhé stránce
btn_frame = ttk.Frame(frame_strana2)
btn_frame.pack(fill=tk.X, side=tk.BOTTOM)

ttk.Button(btn_frame, text="⮌ Zpět", command=navrat_zpet).pack(
    side=tk.LEFT, ipady=5
)
ttk.Button(btn_frame, text="Spustit konverzi", command=spustit_konverzi).pack(
    side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0), ipady=5
)

# Fix pro navigaci zpět
frame_frame_strana1 = frame_strana1

root.mainloop()
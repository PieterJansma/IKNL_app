# iknl_gui_working_fix.py
import os, sys, runpy, traceback
import tkinter as tk
from tkinter import filedialog, messagebox

PROCESSOR_SCRIPT_NAME = "final_iknl_exe_ready.py"   # naam van je verwerkingsscript

def bundled_path(relname: str) -> str:
    """
    Geeft het pad naar een gebundeld bestand (bijv. in de .exe) of,
    als we niet gebundeld zijn, naar een bestand in dezelfde map als dit script.
    """
    if getattr(sys, "frozen", False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relname)

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("IKNL Interface")
        self.geometry("820x460")
        self.resizable(False, False)

        self.path_excel = None      # sleutellijst
        self.path_csv   = None      # CKV/IKNL data
        self.path_dict  = None      # dictionary-Excel
        self.out_dir    = os.path.join(os.getcwd(), "output")
        self.status     = tk.StringVar(value="Klaar. Resultaten staan in de output-map.")

        # Kop
        tk.Label(
            self,
            text="Upload je bestanden en klik op Run",
            font=("Arial", 16, "bold")
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=18, pady=(18, 4))

        tk.Label(
            self,
            text="1) Sleutellijst (Excel)   2) Data (CSV)   3) Dictionary (Excel)   4) Output-map"
        ).grid(row=1, column=0, columnspan=3, sticky="w", padx=18, pady=(0, 10))

        # Sleutellijst
        tk.Label(
            self,
            text="Upload hier je sleutellijst (Excel)",
            font=("Arial", 12, "bold")
        ).grid(row=2, column=0, sticky="w", padx=18, pady=(8, 2))

        tk.Button(self, text="Sleutellijst", width=16, command=self.pick_excel).grid(
            row=3, column=0, sticky="w", padx=18
        )
        self.lbl_excel = tk.Label(self, text="Nog geen bestand gekozen")
        self.lbl_excel.grid(row=4, column=0, columnspan=2, sticky="w", padx=18, pady=(4, 8))

        # Data
        tk.Label(
            self,
            text="Upload hier je data (CSV)",
            font=("Arial", 12, "bold")
        ).grid(row=2, column=1, sticky="w", padx=18, pady=(8, 2))

        tk.Button(self, text="data.csv", width=16, command=self.pick_csv).grid(
            row=3, column=1, sticky="w", padx=18
        )
        self.lbl_csv = tk.Label(self, text="Nog geen bestand gekozen")
        self.lbl_csv.grid(row=4, column=1, columnspan=2, sticky="w", padx=18, pady=(4, 8))

        # Output
        tk.Label(
            self,
            text="Output-map (waar de resultaten komen)",
            font=("Arial", 12, "bold")
        ).grid(row=2, column=2, sticky="w", padx=18, pady=(8, 2))

        tk.Button(self, text="outputmap", width=16, command=self.pick_output_dir).grid(
            row=3, column=2, sticky="w", padx=18
        )
        self.lbl_out = tk.Label(self, text=self.out_dir, wraplength=240, justify="left")
        self.lbl_out.grid(row=4, column=2, sticky="w", padx=18, pady=(4, 8))

        # Dictionary (nieuwe sectie)
        tk.Label(
            self,
            text="Dictionary (Excel met mappings)",
            font=("Arial", 12, "bold")
        ).grid(row=5, column=0, sticky="w", padx=18, pady=(8, 2))

        self.btn_dict = tk.Button(self, text="Dictionary Excel", width=16, command=self.pick_dict)
        self.btn_dict.grid(row=6, column=0, sticky="w", padx=18)

        self.lbl_dict = tk.Label(self, text="Nog geen dictionary gekozen")
        self.lbl_dict.grid(row=7, column=0, columnspan=2, sticky="w", padx=18, pady=(4, 8))

        # Acties (onderste rij)
        self.btn_run = tk.Button(self, text="Run", width=18, command=self.run_clicked, state="disabled")
        self.btn_run.grid(row=8, column=0, sticky="w", padx=18, pady=(18, 8))

        tk.Button(self, text="Open output-map", width=18, command=self.open_output).grid(
            row=8, column=1, sticky="w", padx=18, pady=(18, 8)
        )

        # Statusbalk
        tk.Label(
            self,
            textvariable=self.status,
            anchor="w",
            bd=1,
            relief="sunken"
        ).grid(row=9, column=0, columnspan=3, sticky="we", padx=18, pady=(8, 16))

        for c in range(3):
            self.columnconfigure(c, weight=1)

        if not os.path.isfile(bundled_path(PROCESSOR_SCRIPT_NAME)):
            messagebox.showwarning("Let op", f"{PROCESSOR_SCRIPT_NAME} niet gevonden.")

    # --------- Pickers ---------
    def pick_excel(self):
        p = filedialog.askopenfilename(
            title="Kies sleutellijst (Excel)",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if p:
            self.path_excel = p
            self.lbl_excel.config(text=os.path.basename(p))
            self._update_run()

    def pick_csv(self):
        p = filedialog.askopenfilename(
            title="Kies CKV/IKNL data (CSV)",
            filetypes=[("CSV", "*.csv")]
        )
        if p:
            self.path_csv = p
            self.lbl_csv.config(text=os.path.basename(p))
            self._update_run()

    def pick_dict(self):
        p = filedialog.askopenfilename(
            title="Kies dictionary Excel",
            filetypes=[("Excel", "*.xlsx *.xls")]
        )
        if p:
            self.path_dict = p
            self.lbl_dict.config(text=os.path.basename(p))
            self._update_run()

    def pick_output_dir(self):
        p = filedialog.askdirectory(title="Kies output-map")
        if p:
            self.out_dir = p
            self.lbl_out.config(text=p)

    # --------- Hulpfuncties ---------
    def open_output(self):
        if not os.path.isdir(self.out_dir):
            messagebox.showinfo("Info", "Output-map bestaat nog niet.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(self.out_dir)
            elif sys.platform == "darwin":
                import subprocess
                subprocess.Popen(["open", self.out_dir])
            else:
                import subprocess
                subprocess.Popen(["xdg-open", self.out_dir])
        except Exception as e:
            messagebox.showerror("Fout", f"Kon output map niet openen:\n{e}")

    def _update_run(self):
        # Run pas mogelijk als alle drie bestanden gekozen zijn
        can_run = bool(self.path_excel and self.path_csv and self.path_dict)
        self.btn_run.configure(state=("normal" if can_run else "disabled"))

    # --------- Run-knop ---------
    def run_clicked(self):
        if not (self.path_excel and os.path.isfile(self.path_excel)):
            messagebox.showerror("Fout", "Kies een geldige Excel sleutellijst.")
            return
        if not (self.path_csv and os.path.isfile(self.path_csv)):
            messagebox.showerror("Fout", "Kies een geldige CKV CSV.")
            return
        if not (self.path_dict and os.path.isfile(self.path_dict)):
            messagebox.showerror("Fout", "Kies een geldige dictionary Excel.")
            return

        try:
            os.makedirs(self.out_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Fout", f"Kon output-map niet maken:\n{e}")
            return

        # Geef de paden door aan het verwerkende script via environment variables
        os.environ["IKNL_SLEUTEL_PATH"] = self.path_excel
        os.environ["IKNL_CKV_PATH"]     = self.path_csv
        os.environ["IKNL_DICTS_PATH"]   = self.path_dict   # <-- LET OP: naam matched verwerkingsscript
        os.environ["IKNL_OUTPUT_DIR"]   = self.out_dir

        old_cwd = os.getcwd()
        os.chdir(self.out_dir)

        self.status.set("Bezig met verwerken…")
        self.update_idletasks()
        try:
            runpy.run_path(bundled_path(PROCESSOR_SCRIPT_NAME), run_name="__main__")
            self.status.set("Klaar. Resultaten staan in de output-map.")
            messagebox.showinfo("Gereed", "Verwerking gereed.")
        except SystemExit:
            self.status.set("Klaar.")
            messagebox.showinfo("Gereed", "Verwerking gereed.")
        except Exception as e:
            tb = traceback.format_exc()
            messagebox.showerror("Fout tijdens verwerking", f"{e}\n\n{tb}")
            self.status.set("Fout opgetreden.")
        finally:
            os.chdir(old_cwd)

if __name__ == "__main__":
    App().mainloop()

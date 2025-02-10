# ui.py
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import subprocess
import platform

from logic import ExcelProcessor

class ExcelMatcherApp:
    """
    Updated GUI that lets the user pick:
      - reference key column
      - multiple reference columns to copy
      - offer key column
    """

    def __init__(self):
        self.master = ttk.Window(themename="cyborg")  # Valitse teema tässä
        self.master.title("Tuotenumeroiden Yhdistäjä")
        self.master.geometry("700x850")
        self.master.minsize(550, 600)

        # Configure custom button style
        self.configure_styles()

        # Main frame
        self.main_frame = ttk.Frame(self.master)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Step 1: Reference File
        self.build_reference_section()

        # Step 2: Offer File
        self.build_offer_section()

        # Step 3: Save Location
        self.build_save_location_section()

        # Preview
        self.build_preview_section()

        # Process & progress
        self.build_process_section()

        # Internal state
        self.reference_file = None
        self.offer_file = None
        self.save_location = None

        # We'll store the user-selected reference columns in here
        self.selected_reference_columns = []

        # Our core logic object
        self.processor = ExcelProcessor()

        self.master.mainloop()

    def configure_styles(self):
        style = ttk.Style()
        # Muokkaa olemassa olevaa 'primary.TButton' -tyyliä asettamalla fontti lihavoiduksi, relief ja padding
        style.configure('primary.TButton', font=('Helvetica', 10, 'bold'), padding=(10, 5), relief='raised', borderwidth=2)
        # Voit myös määrittää eri relief-tyylejä, kuten 'sunken' tai 'groove'
        # style.configure('primary.TButton', font=('Helvetica', 10, 'bold'), padding=10, relief='sunken', borderwidth=2)

        # Define new style for white Progressbar
        style.configure("White.Horizontal.TProgressbar",
                        troughcolor="#FFFFFF",  # White background
                        background="#00FF00",   # Green bar (muuta tarvittaessa)
                        bordercolor="#FFFFFF",  # White border
                        lightcolor="#FFFFFF",   # White light effect
                        darkcolor="#FFFFFF")    # White dark effect

        # Define style for selected labels
        style.configure("Selected.TLabel", foreground="green")
        style.configure("Default.TLabel", foreground="gray")

    def build_reference_section(self):
        self.step1_frame = ttk.Labelframe(
            self.main_frame,
            text="Vaihe 1: Valitse referenssitiedosto",
            bootstyle="primary"
        )
        self.step1_frame.pack(fill=tk.X, pady=10)

        ref_button = ttk.Button(
            self.step1_frame,
            text="Valitse Referenssitiedosto",
            bootstyle="primary",  # Käytä 'primary' bootstylea
            command=self.choose_reference_file
        )
        ref_button.pack(pady=5, padx=5, anchor="w", ipadx=10, ipady=5)

        self.ref_label = ttk.Label(self.step1_frame, text="Ei tiedostoa valittu.", style="Default.TLabel")
        self.ref_label.pack(anchor="w", padx=5)

        lbl = ttk.Label(
            self.step1_frame, 
            text="Valitse referenssitiedoston sarake, jossa on 'ulkoinen tunnus':"
        )
        lbl.pack(anchor="w", padx=5, pady=(10, 0))

        self.reference_column_var = tk.StringVar(self.master)
        self.ref_column_menu = ttk.Combobox(
            self.step1_frame,
            textvariable=self.reference_column_var,
            state="readonly",
            bootstyle="secondary"
        )
        self.ref_column_menu.pack(pady=5, padx=5, anchor="w")
        self.ref_column_menu.bind("<<ComboboxSelected>>", self.on_reference_column_selected)

        # A new label for user-chosen columns
        lbl2 = ttk.Label(
            self.step1_frame, 
            text="Valitse haluamasi sarakkeet referenssitiedostosta:"
        )
        lbl2.pack(anchor="w", padx=5, pady=(10, 0))

        # We'll use a Listbox for multi-select
        self.ref_cols_listbox = tk.Listbox(
            self.step1_frame, 
            selectmode=tk.MULTIPLE, 
            height=10, 
            exportselection=False
        )
        self.ref_cols_listbox.pack(pady=5, padx=5, fill=tk.X)
        self.ref_cols_listbox.bind('<<ListboxSelect>>', self.on_columns_selected)

    def build_offer_section(self):
        self.step2_frame = ttk.Labelframe(
            self.main_frame,
            text="Vaihe 2: Valitse tarjoustiedosto",
            bootstyle="primary"
        )
        self.step2_frame.pack(fill=tk.X, pady=10)

        offer_button = ttk.Button(
            self.step2_frame,
            text="Valitse Tarjoustiedosto",
            bootstyle="primary",  # Käytä 'primary' bootstylea
            command=self.choose_offer_file
        )
        offer_button.pack(pady=5, padx=5, anchor="w", ipadx=10, ipady=5)

        self.offer_label = ttk.Label(self.step2_frame, text="Ei tiedostoa valittu.", style="Default.TLabel")
        self.offer_label.pack(anchor="w", padx=5)

        lbl = ttk.Label(
            self.step2_frame, 
            text="Valitse tarjoustiedoston tuotenumerosarake:"
        )
        lbl.pack(anchor="w", padx=5, pady=(10, 0))

        self.offer_column_var = tk.StringVar(self.master)
        self.offer_column_menu = ttk.Combobox(
            self.step2_frame,
            textvariable=self.offer_column_var,
            state="readonly",
            bootstyle="secondary"
        )
        self.offer_column_menu.pack(pady=5, padx=5, anchor="w")

    def build_save_location_section(self):
        self.save_frame = ttk.Labelframe(
            self.main_frame,
            text="Tallennussijainti",
            bootstyle="primary"
        )
        self.save_frame.pack(fill=tk.X, pady=10)

        save_button = ttk.Button(
            self.save_frame, 
            text="Valitse sijainti", 
            bootstyle="primary",  # Käytä 'primary' bootstylea
            command=self.choose_save_location
        )
        save_button.pack(pady=5, padx=5, anchor="w", ipadx=10, ipady=5)

        self.save_label = ttk.Label(self.save_frame, text="Ei sijaintia valittu.", style="Default.TLabel")
        self.save_label.pack(anchor="w", padx=5)

    def build_preview_section(self):
        self.preview_frame = ttk.Labelframe(
            self.main_frame,
            text="Tiedostojen Tiedot",
            bootstyle="primary"
        )
        self.preview_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.info_text = tk.Text(
            self.preview_frame,
            height=10,
            wrap=tk.WORD,
            state=tk.DISABLED,
            bg="#f5f5f5"
        )
        self.info_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def build_process_section(self):
        # Luo kehyksen napuille
        buttons_frame = ttk.Frame(self.main_frame)
        buttons_frame.pack(pady=10, padx=10)

        # "Aloita Prosessi" nappi
        self.process_button = ttk.Button(
            buttons_frame,
            text="Aloita Prosessi",
            bootstyle="primary",  # Käytä 'primary' bootstylea
            command=self.start_process,
            state=tk.DISABLED
        )
        self.process_button.pack(side=LEFT, padx=(0, 10), ipadx=10, ipady=5)

        # "Ohjeet" nappi
        help_button = ttk.Button(
            buttons_frame, 
            text="Ohjeet", 
            bootstyle="primary-outline",  # Käytä outline bootstylea
            command=self.show_help
        )
        help_button.pack(side=LEFT, ipadx=10, ipady=5)

        # Progress bar alla
        self.progress_bar = ttk.Progressbar(
            self.main_frame,
            orient="horizontal",
            length=400,
            mode="determinate",
            style="White.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(pady=5)

    # -----------------------------
    #         UI Callbacks
    # -----------------------------
    def show_help(self):
        help_text = (
            "Tämä ohjelma yhdistää tarjoustiedoston tuotteet referenssitiedostoon.\n\n\n"
            "**Vaihe 1**\n"
            "1. Valitse referenssitiedosto, jossa on ulkoiset ja sisäiset tuotetiedot.\n"
            "2. Valitse sarake, josta löytyy ulkoiset tuotekoodit.\n"
            "3. Valitse kaikki sarakkeet, jotka haluat kopioida lopulliseen tiedostoon.\n\n\n"
            "**Vaihe 2**\n"
            "4. Valitse tarjoustiedosto.\n"
            "5. Valitse sarake, josta löytyy ulkoiset tuotenumerot.\n\n\n"
            "**Vaihe 3**\n"
            "6. Valitse tallennuskansio.\n"
            "7. Klikkaa 'Aloita Prosessi'.\n"
            "Tämän jälkeen ohjelma luo uuden tiedoston valitsemaasi kansioon.\n"
            "Tiedostossa ovat vain valitsemasi sarakkeet referenssitiedostosta.\n"
        )
        messagebox.showinfo("Ohjeet", help_text)

    def choose_reference_file(self):
        self.reference_file = filedialog.askopenfilename(
            filetypes=[("Excel-tiedostot", "*.xlsx *.xls")]
        )
        if self.reference_file:
            try:
                df_reference = pd.read_excel(self.reference_file, dtype=str)
                columns = list(df_reference.columns)

                # Jos sarakkeita on, aseta oletusreferenssiavaimesarake
                if columns:
                    self.reference_column_var.set(columns[0])
                self.update_reference_column_menu(columns)

                # Päivitä listbox, mutta poista referenssiavaimen sarake
                self.ref_cols_listbox.delete(0, tk.END)
                for col in columns:
                    if col != self.reference_column_var.get():  # Poista referenssiavaimen sarake listasta
                        self.ref_cols_listbox.insert(tk.END, col)

                # Lisää vihreä tarkistusmerkki ja muuta tekstin väri
                self.ref_label.config(text=f"✓ {self.reference_file}", style="Selected.TLabel")
                self.update_preview_text("Referenssitiedosto", df_reference, "Referenssi")

            except Exception as e:
                messagebox.showerror("Virhe", f"Virhe tiedoston lukemisessa:\n{str(e)}")

            self.check_ready_to_process()

    def update_reference_column_menu(self, columns):
        self.ref_column_menu['values'] = columns
        if columns:
            self.ref_column_menu.current(0)
            self.on_reference_column_selected(None)

    def on_reference_column_selected(self, event):
        value = self.reference_column_var.get()
        # Päivitä monivalintalistasta poistamalla valittu referenssiavaimen sarake
        self.ref_cols_listbox.delete(0, tk.END)
        try:
            df_reference = pd.read_excel(self.reference_file, dtype=str)
            for col in df_reference.columns:
                if col != value:
                    self.ref_cols_listbox.insert(tk.END, col)
        except Exception as e:
            messagebox.showerror("Virhe", f"Virhe tiedoston lukemisessa:\n{str(e)}")

        self.check_ready_to_process()

    def on_columns_selected(self, event):
        self.check_ready_to_process()

    def choose_offer_file(self):
        self.offer_file = filedialog.askopenfilename(
            filetypes=[("Excel-tiedostot", "*.xlsx *.xls")]
        )
        if self.offer_file:
            try:
                df_offer = pd.read_excel(self.offer_file, dtype=str)
                columns = list(df_offer.columns)

                if columns:
                    self.offer_column_var.set(columns[0])
                self.update_offer_column_menu(columns)

                # Lisää vihreä tarkistusmerkki ja muuta tekstin väri
                self.offer_label.config(text=f"✓ {self.offer_file}", style="Selected.TLabel")
                self.update_preview_text("Tarjoustiedosto", df_offer, "Tarjous")
            except Exception as e:
                messagebox.showerror("Virhe", f"Virhe tiedoston lukemisessa:\n{str(e)}")

            self.check_ready_to_process()

    def update_offer_column_menu(self, columns):
        self.offer_column_menu['values'] = columns
        if columns:
            self.offer_column_menu.current(0)

    def update_preview_text(self, title, dataframe, file_type):
        self.info_text.config(state=tk.NORMAL)
        self.info_text.delete(1.0, tk.END)

        info = (
            f"{title} - Rivit: {len(dataframe)}, Sarakkeet: {len(dataframe.columns)}\n"
            f"Sarakkeet (näkyvistä vain 5): {', '.join(dataframe.columns[:5])}{'...' if len(dataframe.columns) > 5 else ''}\n"
        )
        self.info_text.insert(tk.END, info)
        self.info_text.config(state=tk.DISABLED)

    def choose_save_location(self):
        self.save_location = filedialog.askdirectory()
        if self.save_location:
            # Lisää vihreä tarkistusmerkki ja muuta tekstin väri
            self.save_label.config(text=f"✓ {self.save_location}", style="Selected.TLabel")
            self.check_ready_to_process()

    def check_ready_to_process(self):
        """
        Varmistaa, että referenssi-, tarjous- ja tallennuskansiotiedostot on valittu,
        sekä että vähintään yksi sarake on valittu referenssitiedostosta.
        """
        if self.reference_file and self.offer_file and self.save_location:
            selected_indices = self.ref_cols_listbox.curselection()
            if len(selected_indices) > 0:
                self.process_button.config(state=tk.NORMAL)
                return
        self.process_button.config(state=tk.DISABLED)

    def validate_selection(self):
        ref_key = self.reference_column_var.get()
        offer_key = self.offer_column_var.get()
        if not ref_key or not offer_key:
            messagebox.showerror("Virhe", "Valitse molemmat avainsarakkeet ennen prosessin aloittamista.")
            return False
        
        if not self.selected_reference_columns:
            messagebox.showerror("Virhe", "Valitse vähintään yksi referenssitiedoston sarake.")
            return False
        
        try:
            df_reference = pd.read_excel(self.reference_file, dtype=str)
            df_offer = pd.read_excel(self.offer_file, dtype=str)

            # Tarkista referenssitiedoston avainsarake
            if ref_key not in df_reference.columns:
                messagebox.showerror("Virhe", f"Sarake '{ref_key}' ei löydy referenssitiedostosta.")
                return False

            # Jos on kyse MATCHED-tiedostosta, etsitään vaihtoehtoisia sarakenimiä
            if "MATCHED_" in self.offer_file:
                if offer_key not in df_offer.columns and (offer_key + " (MATCHED)") not in df_offer.columns:
                    messagebox.showerror("Virhe", f"Sarake '{offer_key}' tai '{offer_key} (MATCHED)' ei löydy tarjoustiedostosta.")
                    return False
            else:
                if offer_key not in df_offer.columns:
                    messagebox.showerror("Virhe", f"Sarake '{offer_key}' ei löydy tarjoustiedostosta.")
                    return False

            # Tarkista, että valitut sarakkeet ovat olemassa referenssitiedostossa
            for col in self.selected_reference_columns:
                if col not in df_reference.columns:
                    messagebox.showerror("Virhe", f"Sarake '{col}' ei löydy referenssitiedostosta.")
                    return False
            return True
        except Exception as e:
            messagebox.showerror("Virhe", f"Virhe tiedostojen lukemisessa:\n{str(e)}")
            return False

    def start_process(self):
        self.progress_bar["value"] = 0
        self.master.update_idletasks()

        # Kerää käyttäjän valitsemat sarakkeet ennen validointia
        selected_indices = self.ref_cols_listbox.curselection()
        self.selected_reference_columns = [self.ref_cols_listbox.get(i) for i in selected_indices]

        if not self.validate_selection():
            return  # Lopeta prosessi, jos validointi epäonnistui

        try:
            self.process_files()
        except Exception as e:
            messagebox.showerror("Virhe", f"Virhe tiedoston käsittelyssä:\n{str(e)}")
        finally:
            self.progress_bar["value"] = 0

    def process_files(self):
        # 1) Get the chosen reference key column
        reference_column = self.reference_column_var.get()

        # 2) Get the chosen offer key column
        competitor_column = self.offer_column_var.get()

        # Pass these to the logic
        self.processor.ref_key_column = reference_column
        self.processor.offer_key_column = competitor_column
        self.processor.selected_ref_columns = self.selected_reference_columns

        try:
            total_steps = 3
            increment = 100 // total_steps

            # Step 1
            self.progress_bar["value"] += increment
            self.master.update_idletasks()

            # Now we call process_files in logic with the final columns
            output_path, missing_count = self.processor.process_files(
                self.reference_file,
                self.offer_file,
                reference_column,
                competitor_column
            )

            # Step 2 – Move file to chosen directory
            final_output_path = self.move_output_file(output_path)
            self.progress_bar["value"] += increment
            self.master.update_idletasks()

            # Step 3 – Reveal file in Explorer (on Windows)
            final_output_path = os.path.abspath(final_output_path)
            self.reveal_in_file_explorer(final_output_path)

            self.progress_bar["value"] = 100
            self.master.update_idletasks()

            messagebox.showinfo(
                "Valmis!",
                f"Uusi tiedosto luotu:\n{final_output_path}\n\nRivejä ilman vastaavuutta: {missing_count}"
            )
        except Exception as e:
            raise e

    def move_output_file(self, output_path):
        output_filename = os.path.basename(output_path)
        final_output_path = os.path.join(self.save_location, output_filename)

        # Tarkistetaan, ovatko tiedoston nykyinen sijainti ja tallennuskansio samat
        if os.path.abspath(os.path.dirname(output_path)) == os.path.abspath(self.save_location):
            # Jos kansiot ovat samat, ei tarvitse tehdä siirtoa
            return output_path

        # Jos tiedosto on jo olemassa kohdekansiossa, poistetaan se automaattisesti.
        if os.path.exists(final_output_path):
            os.remove(final_output_path)
    
        os.rename(output_path, final_output_path)
        return final_output_path

    def reveal_in_file_explorer(self, path):
        if platform.system() == "Windows":
            subprocess.run(f'explorer /select,"{path}"', shell=True)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", "--", path])
        else:  # Linux
            subprocess.run(["xdg-open", os.path.dirname(path)])

if __name__ == "__main__":
    ExcelMatcherApp()

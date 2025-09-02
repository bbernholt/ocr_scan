import tkinter as tk
from tkinter import ttk 
from PIL import Image, ImageTk
import os
import cv2
import pytesseract
import threading
import openpyxl
from pygrabber.dshow_graph import FilterGraph

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("Warenein- und -ausgang")
            
            # Bildschirmgröße holen
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            self.geometry(f"{screen_width}x{screen_height//2}+0+0")
            
            # Maße des Logos
            breite = 120   # frei wählbar
            hoehe = 60     # frei wählbar

            # Logo laden und scalieren
            self.load_and_scale_logo(breite, hoehe)

            # Maße der Bilder (Piktogramm für Wareneingang udn Ausgang) definieren
            breite = 450
            hoehe = 450
            
            # Bilder laden und skalieren
            self.load_and_scale_images(breite, hoehe)
            
            # Frames (Seiten)
            self.startseite = tk.Frame(self)
            self.wareneingang_seite = tk.Frame(self)
            self.warenausgang_seite = tk.Frame(self)
            
            # Artikel-Daten
            self.artikel_dict = {}         # Excel-Daten
            self.detected_articles = []    # Erkannte Artikel
            
            # Startseite aufbauen
            self.build_startseite()
            
            # Wareneingang-Seite aufbauen
            self.build_wareneingang()
            
            # Warenausgang-Seite aufbauen
            self.build_warenausgang()
            
            # Startseite anzeigen
            self.show_startseite()

            # Clean Exit
            #self.protocol("WM_DELETE_WINDOW", self.on_close)

        # FUNKTION: Läd und skaliert die Wareneingang- und Warenausgang-Bilder
        def load_and_scale_images(self, breite, hoehe):
            """Lädt und skaliert die Wareneingang- und Warenausgang-Bilder"""
            try:
                eingang_img = Image.open("../bilder/wareneingang.png").resize((breite, hoehe), Image.Resampling.LANCZOS)
                ausgang_img = Image.open("../bilder/warenausgang.png").resize((breite, hoehe), Image.Resampling.LANCZOS)
                self.eingang_photo = ImageTk.PhotoImage(eingang_img)
                self.ausgang_photo = ImageTk.PhotoImage(ausgang_img)
            except FileNotFoundError as e:
                print(f"Bilddatei nicht gefunden: {e}")
            except Exception as e:
                print(f"Fehler beim Laden der Bilder: {e}")

        # FUNKTION: Läd und skaliert das Logo
        def load_and_scale_logo(self, breite, hoehe):
            """Lädt und skaliert das Logo"""
            try:
                logo_img = Image.open("../bilder/logo.png").resize((breite, hoehe), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img)
            except FileNotFoundError as e:
                print(f"Logo-Datei nicht gefunden: {e}")
            except Exception as e:
                print(f"Fehler beim Laden des Logos: {e}")

        # FUNKTION: Erstelle Layout der Startseite
        def build_startseite(self):
            # Erstellt Button für Wareneingang
            btn_eingang = tk.Button(
                self.startseite, image=self.eingang_photo,
                command=self.show_wareneingang, borderwidth=0
            )
            # Erstellt Button für Warenausgang
            btn_ausgang = tk.Button(
                self.startseite, image=self.ausgang_photo,
                command=self.show_warenausgang, borderwidth=0
            )
            # Positioniert die Buttons
            btn_eingang.pack(side="left", expand=True, fill="both")
            btn_ausgang.pack(side="right", expand=True, fill="both")

            # Fügt das Logo hinzu
            self.add_logo(self.startseite)

        # FUNKTION: Zeige das Layout der Startseite an
        def show_startseite(self):
            self.wareneingang_seite.pack_forget()  # Versteckt die Wareneingang-Seite
            self.warenausgang_seite.pack_forget()  # Versteckt die Warenausgang-Seite
            self.startseite.pack(expand=True, fill="both")  # Zeigt die Startseite an

        # FUNKTION: Fügt das Logo in das gegebene Frame ein
        def add_logo(self, parent_frame):
            """Fügt unten rechts im gegebenen Frame das Logo ein"""
            logo_label = tk.Label(parent_frame, image=self.logo_photo, borderwidth=0)
            logo_label.image = self.logo_photo  # Referenz behalten
            logo_label.pack(side="right", anchor="se", padx=10, pady=10)

        # FUNKTION: Erstelle Layout des Wareneingangs
        def build_wareneingang(self):
            # Überschrift
            tk.Label(self.wareneingang_seite, text="Wareneingang", font=("Arial", 30)).pack(pady=20)

            # Innerer Frame für linke und rechte Hälfte
            main_frame = tk.Frame(self.wareneingang_seite)
            main_frame.pack(expand=True, fill="both")

            # Linke Hälfte
            left_frame = tk.Frame(main_frame)
            left_frame.pack(side="left", expand=True, fill="both", padx=20, pady=20)

            # Frame für Button + Dropdown nebeneinander
            dropdown_frame = tk.Frame(left_frame)
            dropdown_frame.pack(pady=10)


            # Dropdown für Excel-Dateien rechts vom Button
            self.dropdown_var = tk.StringVar()
            self.dropdown = ttk.Combobox(dropdown_frame, textvariable=self.dropdown_var, width=30)
            self.dropdown.pack(side="left")

            # Label "Erfasste Artikel"
            tk.Label(left_frame, text="Erfasste Artikel", font=("Arial", 14)).pack(pady=(20,5))

            # Tabelle für Artikel
            columns = ("artikelnummer", "menge", "karton", "beutel", "status")
            self.tree = ttk.Treeview(left_frame, columns=columns, show="headings", height=15)

            # Spaltenüberschriften setzen
            self.tree.heading("artikelnummer", text="Artikelnummer")
            self.tree.heading("menge", text="Menge")
            self.tree.heading("karton", text="Karton")
            self.tree.heading("beutel", text="Beutel")
            self.tree.heading("status", text="Status")

            # Spaltenbreite
            self.tree.column("artikelnummer", width=120, anchor="center")
            self.tree.column("menge", width=80, anchor="center")
            self.tree.column("karton", width=80, anchor="center")
            self.tree.column("beutel", width=80, anchor="center")
            self.tree.column("status", width=80, anchor="center")

            self.tree.pack(expand=True, fill="both")

            # Rechte Hälfte
            self.right_frame = tk.Frame(main_frame)
            self.right_frame.pack(side="right", expand=True, fill="both", padx=20, pady=20)

            # Webcam starten
            #self.start_webcam()

            # Frame für Buttons unten
            button_frame = tk.Frame(self.wareneingang_seite)
            button_frame.pack(pady=10)

            # "Drucken"-Button
            tk.Button(button_frame, text="Drucken").pack(side="left", padx=5)
            # "Zurück"-Button
            tk.Button(button_frame, text="Zurück", command=self.show_startseite).pack(side="left", padx=5)

            # Excel-Dateien laden
            #self.load_excel_files()

            # add Logo
            self.add_logo(self.wareneingang_seite)

        # FUNKTION: Wareneingang anzeigen
        def show_wareneingang(self):
            # Navigation: Alle anderen Seiten ausblenden
            self.startseite.pack_forget()           # Startseite verstecken
            self.warenausgang_seite.pack_forget()   # Warenausgang-Seite verstecken
            
            # Wareneingang-Seite anzeigen und Layout füllen
            self.wareneingang_seite.pack(expand=True, fill="both")
            
            # Tastatur-Shortcuts für Wareneingang-Seite definieren
            self.bind("<Return>", lambda e: self.drucken())        # Enter-Taste → Drucken-Funktion
            self.bind("<Escape>", lambda e: self.show_startseite()) # Escape-Taste → Zurück zur Startseite

        # FUNKTION: Erstelle Layout des Warenausgangs
        def build_warenausgang(self):
            tk.Label(self.warenausgang_seite, text="Warenausgang", font=("Arial", 30)).pack(pady=20)
            tk.Button(self.warenausgang_seite, text="Zurück", command=self.show_startseite).pack(pady=10)

            # add Logo
            self.add_logo(self.warenausgang_seite)

        # FUNKTION: Warenausgang anzeigen
        def show_warenausgang(self):
            # Navigation: Alle anderen Seiten ausblenden  
            self.startseite.pack_forget()         # Startseite verstecken
            self.wareneingang_seite.pack_forget() # Wareneingang-Seite verstecken
            
            # Warenausgang-Seite anzeigen und Layout füllen
            self.warenausgang_seite.pack(expand=True, fill="both")
            
            # Tastatur-Shortcuts für Wareneingang-Seite definieren
            self.bind("<Return>", lambda e: self.drucken())        # Enter-Taste → Drucken-Funktion
            self.bind("<Escape>", lambda e: self.show_startseite()) # Escape-Taste → Zurück zur Startseite

if __name__ == "__main__":
    app = App()
    #app.test()
    app.mainloop()
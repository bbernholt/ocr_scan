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
            self.artikel_dict_eingang = []    # Excel-Daten Wareneingang
            self.artikel_dict_ausgang = []    # Excel-Daten Warenausgang
            self.detected_articles = []       # Erkannte Artikel
            
            # Startseite aufbauen
            self.build_startseite()
            
            # Wareneingang-Seite aufbauen
            self.build_wareneingang()
            
            # Warenausgang-Seite aufbauen
            self.build_warenausgang()
            
            # Startseite anzeigen
            self.show_startseite()

            # Webcam suchen und initialisieren (in separatem Thread)
            self.cap = None
            threading.Thread(target=self.initialize_webcam, daemon=True).start()

            # Clean Exit
            #self.protocol("WM_DELETE_WINDOW", self.on_close)

        #------------------------------------------------------- GUI ---------------------------------------------------------

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
            # Webcam-Stream stoppen falls aktiv
            self.stop_webcam_stream()
            
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

            # Dropdown für Excel-Dateien aus Eingang-Verzeichnis
            self.dropdown_var_eingang = tk.StringVar()
            self.dropdown_eingang = ttk.Combobox(left_frame, textvariable=self.dropdown_var_eingang, width=30)
            # Initial laden
            excel_files_eingang = self.load_excel_files("../eingang")
            self.dropdown_eingang['values'] = excel_files_eingang
            # Erste Datei automatisch auswählen und laden
            if excel_files_eingang:
                self.dropdown_eingang.set(excel_files_eingang[0])
                # Erste Datei automatisch laden
                filepath = os.path.join("../eingang", excel_files_eingang[0])
                self.load_excel_data(filepath, "eingang")
            # Event-Binding für automatische Aktualisierung und Excel-Laden
            self.dropdown_eingang.bind('<Button-1>', self.refresh_dropdown_eingang)
            self.dropdown_eingang.bind('<<ComboboxSelected>>', self.on_excel_select_eingang)
            self.dropdown_eingang.pack(pady=10)

            # Label "Erfasste Artikel"
            tk.Label(left_frame, text="Erfasste Artikel", font=("Arial", 14)).pack(pady=(20,5))

            # Tabelle für Artikel (bleibt leer bis Artikel erkannt werden)
            columns = ("artikelnummer", "menge", "karton", "beutel", "status")
            self.tree_eingang = ttk.Treeview(left_frame, columns=columns, show="headings", height=15)

            # Spaltenüberschriften setzen
            self.tree_eingang.heading("artikelnummer", text="Artikelnummer")
            self.tree_eingang.heading("menge", text="Menge")
            self.tree_eingang.heading("karton", text="Karton")
            self.tree_eingang.heading("beutel", text="Beutel")
            self.tree_eingang.heading("status", text="Status")

            # Spaltenbreite
            self.tree_eingang.column("artikelnummer", width=120, anchor="center")
            self.tree_eingang.column("menge", width=80, anchor="center")
            self.tree_eingang.column("karton", width=80, anchor="center")
            self.tree_eingang.column("beutel", width=80, anchor="center")
            self.tree_eingang.column("status", width=80, anchor="center")

            self.tree_eingang.pack(expand=True, fill="both")

            # Rechte Hälfte - Webcam-Bereich (GEÄNDERT: expand=False)
            self.right_frame_eingang = tk.Frame(main_frame, bg="black")
            self.right_frame_eingang.pack(side="right", expand=False, fill="y", padx=20, pady=20)

            # Frame für Buttons unten
            button_frame = tk.Frame(self.wareneingang_seite)
            button_frame.pack(pady=10)

            # "Drucken"-Button
            tk.Button(button_frame, text="Drucken").pack(side="left", padx=5)
            # "Zurück"-Button
            tk.Button(button_frame, text="Zurück", command=self.show_startseite).pack(side="left", padx=5)

            # add Logo
            self.add_logo(self.wareneingang_seite)

        # FUNKTION: Wareneingang anzeigen
        def show_wareneingang(self):
            # Navigation: Alle anderen Seiten ausblenden
            self.startseite.pack_forget()           # Startseite verstecken
            self.warenausgang_seite.pack_forget()   # Warenausgang-Seite verstecken
            
            # Wareneingang-Seite anzeigen und Layout füllen
            self.wareneingang_seite.pack(expand=True, fill="both")
            
            # Webcam-Verfügbarkeit prüfen
            self.check_webcam_for_page()
            
            # Webcam-Stream starten
            self.start_webcam_stream(self.right_frame_eingang)
            
            # Tastatur-Shortcuts für Wareneingang-Seite definieren
            self.bind("<Return>", lambda e: self.drucken())        # Enter-Taste → Drucken-Funktion
            self.bind("<Escape>", lambda e: self.show_startseite()) # Escape-Taste → Zurück zur Startseite

        # FUNKTION: Erstelle Layout des Warenausgangs
        def build_warenausgang(self):
            # Überschrift
            tk.Label(self.warenausgang_seite, text="Warenausgang", font=("Arial", 30)).pack(pady=20)

            # Innerer Frame für linke und rechte Hälfte
            main_frame = tk.Frame(self.warenausgang_seite)
            main_frame.pack(expand=True, fill="both")

            # Linke Hälfte
            left_frame = tk.Frame(main_frame)
            left_frame.pack(side="left", expand=True, fill="both", padx=20, pady=20)

            # Dropdown für Excel-Dateien aus Ausgang-Verzeichnis
            self.dropdown_var_ausgang = tk.StringVar()
            self.dropdown_ausgang = ttk.Combobox(left_frame, textvariable=self.dropdown_var_ausgang, width=30)
            # Initial laden
            excel_files_ausgang = self.load_excel_files("../ausgang")
            self.dropdown_ausgang['values'] = excel_files_ausgang
            # Erste Datei automatisch auswählen und laden
            if excel_files_ausgang:
                self.dropdown_ausgang.set(excel_files_ausgang[0])
                # Erste Datei automatisch laden
                filepath = os.path.join("../ausgang", excel_files_ausgang[0])
                self.load_excel_data(filepath, "ausgang")
            # Event-Binding für automatische Aktualisierung und Excel-Laden
            self.dropdown_ausgang.bind('<Button-1>', self.refresh_dropdown_ausgang)
            self.dropdown_ausgang.bind('<<ComboboxSelected>>', self.on_excel_select_ausgang)
            self.dropdown_ausgang.pack(pady=10)

            # Label "Erfasste Artikel"
            tk.Label(left_frame, text="Erfasste Artikel", font=("Arial", 14)).pack(pady=(20,5))

            # Tabelle für Artikel (bleibt leer bis Artikel erkannt werden)
            columns = ("artikelnummer", "menge", "karton", "beutel", "empfaenger", "status")
            self.tree_ausgang = ttk.Treeview(left_frame, columns=columns, show="headings", height=15)

            # Spaltenüberschriften setzen
            self.tree_ausgang.heading("artikelnummer", text="Artikelnummer")
            self.tree_ausgang.heading("menge", text="Menge")
            self.tree_ausgang.heading("karton", text="Karton")
            self.tree_ausgang.heading("beutel", text="Beutel")
            self.tree_ausgang.heading("empfaenger", text="Empfänger")
            self.tree_ausgang.heading("status", text="Status")

            # Spaltenbreite
            self.tree_ausgang.column("artikelnummer", width=70, anchor="center")
            self.tree_ausgang.column("menge", width=70, anchor="center")
            self.tree_ausgang.column("karton", width=70, anchor="center")
            self.tree_ausgang.column("beutel", width=70, anchor="center")
            self.tree_ausgang.column("empfaenger", width=70, anchor="center")
            self.tree_ausgang.column("status", width=50, anchor="center")

            self.tree_ausgang.pack(expand=True, fill="both")

            # Rechte Hälfte - Webcam-Bereich (GEÄNDERT: expand=False)
            self.right_frame_ausgang = tk.Frame(main_frame, bg="black")
            self.right_frame_ausgang.pack(side="right", expand=False, fill="y", padx=20, pady=20)

            # Frame für Buttons unten
            button_frame = tk.Frame(self.warenausgang_seite)
            button_frame.pack(pady=10)

            # "Drucken"-Button
            tk.Button(button_frame, text="Drucken").pack(side="left", padx=5)
            # "Zurück"-Button
            tk.Button(button_frame, text="Zurück", command=self.show_startseite).pack(side="left", padx=5)

            # add Logo
            self.add_logo(self.warenausgang_seite)

        # FUNKTION: Warenausgang anzeigen
        def show_warenausgang(self):
            # Navigation: Alle anderen Seiten ausblenden  
            self.startseite.pack_forget()         # Startseite verstecken
            self.wareneingang_seite.pack_forget() # Wareneingang-Seite verstecken
            
            # Warenausgang-Seite anzeigen und Layout füllen
            self.warenausgang_seite.pack(expand=True, fill="both")
            
            # Webcam-Verfügbarkeit prüfen
            self.check_webcam_for_page()
            
            # Webcam-Stream starten
            self.start_webcam_stream(self.right_frame_ausgang)
            
            # Tastatur-Shortcuts für Warenausgang-Seite definieren
            self.bind("<Return>", lambda e: self.drucken())        # Enter-Taste → Drucken-Funktion
            self.bind("<Escape>", lambda e: self.show_startseite()) # Escape-Taste → Zurück zur Startseite

        #------------------------------------------------------- FUNKTIONALITÄT ---------------------------------------------------------

        # FUNKTION: Lädt Excel-Dateien aus einem Verzeichnis
        def load_excel_files(self, directory):
            """Lädt alle Excel-Dateien aus dem angegebenen Verzeichnis"""
            excel_files = []
            try:
                if os.path.exists(directory):
                    for file in os.listdir(directory):
                        if file.endswith(('.xlsx', '.xls')):
                            excel_files.append(file)
                else:
                    print(f"Verzeichnis nicht gefunden: {directory}")
            except Exception as e:
                print(f"Fehler beim Laden der Excel-Dateien: {e}")
            return excel_files

        # FUNKTION: Lädt Excel-Datei und speichert Inhalt
        def load_excel_data(self, filepath, page_type="eingang"):
            """Lädt Excel-Datei und speichert Daten in artikel_dict"""
            try:
                # Excel-Datei öffnen
                workbook = openpyxl.load_workbook(filepath)
                sheet = workbook.active
                
                # Spaltentitel aus Zeile 1 lesen
                headers = []
                for cell in sheet[1]:
                    if cell.value:
                        headers.append(cell.value)
                    else:
                        break
                
                # Datenzeilen ab Zeile 2 lesen
                data_rows = []
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    if any(row):  # Nur nicht-leere Zeilen
                        row_dict = {}
                        for i, value in enumerate(row[:len(headers)]):
                            row_dict[headers[i]] = value if value is not None else ""
                        data_rows.append(row_dict)
                
                # Je nach Seite in entsprechende Variable speichern
                if page_type == "eingang":
                    self.artikel_dict_eingang = data_rows
                    print(f"Wareneingang: {len(data_rows)} Artikel aus Excel geladen")
                    print(f"Spaltentitel: {headers}")
                    print("Erste 3 Datensätze:")
                    for i, row in enumerate(data_rows[:3]):
                        print(f"  Zeile {i+2}: {row}")
                    if len(data_rows) > 3:
                        print(f"  ... und {len(data_rows)-3} weitere Datensätze")
                else:
                    self.artikel_dict_ausgang = data_rows
                    print(f"Warenausgang: {len(data_rows)} Artikel aus Excel geladen")
                    print(f"Spaltentitel: {headers}")
                    print("Erste 3 Datensätze:")
                    for i, row in enumerate(data_rows[:3]):
                        print(f"  Zeile {i+2}: {row}")
                    if len(data_rows) > 3:
                        print(f"  ... und {len(data_rows)-3} weitere Datensätze")
                
                print("-" * 50)  # Trennlinie für bessere Lesbarkeit
                
                workbook.close()
                return data_rows
                
            except Exception as e:
                print(f"Fehler beim Laden der Excel-Datei: {e}")
                return []

        # FUNKTION: Aktualisiert Dropdown-Inhalte für Wareneingang
        def refresh_dropdown_eingang(self, event=None):
            """Aktualisiert Dropdown-Inhalte beim Klicken - Wareneingang"""
            excel_files = self.load_excel_files("../eingang")
            self.dropdown_eingang['values'] = excel_files

        # FUNKTION: Aktualisiert Dropdown-Inhalte für Warenausgang
        def refresh_dropdown_ausgang(self, event=None):
            """Aktualisiert Dropdown-Inhalte beim Klicken - Warenausgang"""
            excel_files = self.load_excel_files("../ausgang")
            self.dropdown_ausgang['values'] = excel_files

        # FUNKTION: Event-Handler für Dropdown-Auswahl Wareneingang
        def on_excel_select_eingang(self, event=None):
            """Wird aufgerufen wenn Excel-Datei im Wareneingang ausgewählt wird"""
            selected_file = self.dropdown_var_eingang.get()
            if selected_file:
                filepath = os.path.join("../eingang", selected_file)
                self.load_excel_data(filepath, "eingang")
                # KEINE Tabellen-Aktualisierung!

        # FUNKTION: Event-Handler für Dropdown-Auswahl Warenausgang
        def on_excel_select_ausgang(self, event=None):
            """Wird aufgerufen wenn Excel-Datei im Warenausgang ausgewählt wird"""
            selected_file = self.dropdown_var_ausgang.get()
            if selected_file:
                filepath = os.path.join("../ausgang", selected_file)
                self.load_excel_data(filepath, "ausgang")
                # KEINE Tabellen-Aktualisierung!

        # FUNKTION: Sucht nach der Logitech C920 Webcam
        def find_logitech_c920(self, show_popup=False):
            """Sucht nach der Logitech C920 Webcam über den Gerätenamen"""
            try:
                # FilterGraph verwenden um alle verfügbaren Kameras zu finden
                graph = FilterGraph()
                devices = graph.get_input_devices()
                print(f"Gefundene Geräte: {devices}")
                
                # Nach Logitech C920 suchen
                for device_index, device_name in enumerate(devices):
                    if "c920" in device_name.lower():
                        print(f"Logitech C920 gefunden: {device_name} (Index: {device_index})")
                        return device_index
                
                # Wenn nicht gefunden und Pop-up erwünscht
                if show_popup:
                    self.show_camera_not_found_popup()
                return None
                
            except Exception as e:
                print(f"Fehler beim Suchen der Webcam: {e}")
                if show_popup:
                    self.show_camera_not_found_popup()
                return None

        # FUNKTION: Webcam initialisieren
        def initialize_webcam(self):
            """Initialisiert die Webcam falls gefunden"""
            camera_index = self.find_logitech_c920(show_popup=False)  # Kein Pop-up beim Start
            if camera_index is not None:
                try:
                    self.cap = cv2.VideoCapture(camera_index)
                    if self.cap.isOpened():
                        print(f"Webcam erfolgreich initialisiert (Index: {camera_index})")
                        return True
                    else:
                        print("Webcam konnte nicht geöffnet werden")
                        return False
                except Exception as e:
                    print(f"Fehler beim Öffnen der Webcam: {e}")
                    return False
            return False

        # FUNKTION: Prüft Webcam-Verfügbarkeit für Wareneingang/Warenausgang
        def check_webcam_for_page(self):
            """Prüft ob Webcam verfügbar ist und zeigt Pop-up falls nicht"""
            if self.cap is None or not self.cap.isOpened():
                # Erneut nach Webcam suchen mit Pop-up
                camera_index = self.find_logitech_c920(show_popup=True)
                if camera_index is not None:
                    self.cap = cv2.VideoCapture(camera_index)

        # FUNKTION: Startet den Webcam-Livestream
        def start_webcam_stream(self, frame):
            """Startet den Webcam-Livestream im gegebenen Frame"""
            # Label für Webcam-Stream erstellen
            self.webcam_label = tk.Label(frame, text="Webcam wird geladen...", 
                                       font=("Arial", 14), bg="lightgray")
            self.webcam_label.pack(expand=True, fill="both")
            
            # Webcam-Stream starten
            self.update_webcam_stream()

        # FUNKTION: Aktualisiert den Webcam-Stream kontinuierlich
        def update_webcam_stream(self):
            """Aktualisiert das Webcam-Bild kontinuierlich"""
            if self.cap is not None and self.cap.isOpened():
                ret, frame = self.cap.read()
                if ret:
                    # Frame für Display vorbereiten
                    frame = cv2.flip(frame, 1)  # Horizontal spiegeln (Spiegel-Effekt)
                    frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                    
                    # Frame skalieren für bessere Darstellung
                    height, width = frame_rgb.shape[:2]
                    max_width, max_height = 640, 480
                    
                    # Seitenverhältnis beibehalten
                    if width > max_width or height > max_height:
                        scale = min(max_width/width, max_height/height)
                        new_width = int(width * scale)
                        new_height = int(height * scale)
                        frame_rgb = cv2.resize(frame_rgb, (new_width, new_height))
                    
                    # In PhotoImage konvertieren
                    img = Image.fromarray(frame_rgb)
                    photo = ImageTk.PhotoImage(img)
                    
                    # Label aktualisieren
                    if hasattr(self, 'webcam_label'):
                        self.webcam_label.configure(image=photo, text="")
                        self.webcam_label.image = photo  # Referenz behalten
                else:
                    # Fehler beim Lesen des Frames
                    if hasattr(self, 'webcam_label'):
                        self.webcam_label.configure(text="Webcam-Fehler", image="")
            else:
                # Webcam nicht verfügbar
                if hasattr(self, 'webcam_label'):
                    self.webcam_label.configure(text="Keine Webcam gefunden", image="")
            
            # Nächstes Update planen (ca. 30 FPS)
            if hasattr(self, 'webcam_label'):
                self.after(33, self.update_webcam_stream)

        # FUNKTION: Stoppt den Webcam-Stream
        def stop_webcam_stream(self):
            """Stoppt den Webcam-Stream"""
            if hasattr(self, 'webcam_label'):
                self.webcam_label.destroy()
                delattr(self, 'webcam_label')

        # FUNKTION: Zeigt Pop-up an wenn Kamera nicht gefunden wurde
        def show_camera_not_found_popup(self):
            """Zeigt ein Pop-up an, wenn die Logitech C920 nicht gefunden wurde"""
            popup = tk.Toplevel(self)
            popup.title("Kamera nicht gefunden")
            popup.geometry("300x150")
            popup.resizable(False, False)
            
            # Pop-up zentrieren
            popup.transient(self)
            popup.grab_set()
            
            # Text anzeigen
            message_label = tk.Label(popup, text="Logitech C920 nicht gefunden", 
                                   font=("Arial", 12), pady=20)
            message_label.pack()
            
            # OK Button
            ok_button = tk.Button(popup, text="OK", command=popup.destroy, 
                                width=10, pady=5)
            ok_button.pack(pady=10)
            
            # Fokus auf Pop-up setzen
            popup.focus_set()

        # FUNKTION: Placeholder für Drucken-Funktion
        def drucken(self):
            """Placeholder für Drucken-Funktionalität"""
            print("Drucken-Funktion aufgerufen")

        # FUNKTION: Placeholder für Find Printer-Funktion
        def find_printer(self):
            """Placeholder für Find Printer-Funktionalität"""
            print("Find Printer-Funktion aufgerufen")

if __name__ == "__main__":
    app = App()
    #app.test()
    app.mainloop()
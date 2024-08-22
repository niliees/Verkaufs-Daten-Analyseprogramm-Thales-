import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.ensemble import GradientBoostingRegressor
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from dateutil.relativedelta import relativedelta
import threading
import json
import os

class VerkaufsprognoseApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Verkaufsprognose für Nachtsichtbrillen")

        # Setze das Theme
        self.style = ttk.Style()
        self.style.theme_use("clam")  # Verwende "clam" Theme als Basis

        # Fenstergröße und -position
        self.master.geometry("600x500")
        self.master.configure(bg="#2e2e2e")  # Dunkler Hintergrund
        self.master.eval('tk::PlaceWindow . center')

        # Custom Styles für das Obsidian-ähnliche Design
        self.style.configure("TFrame", background="#2e2e2e")
        self.style.configure("TLabel", background="#2e2e2e", foreground="#c9d1d9", font=("Helvetica", 12))
        self.style.configure("TButton", background="#3a3a3a", foreground="#c9d1d9", font=("Helvetica", 12), padding=10)
        self.style.map("TButton", background=[("active", "#1f6e91")], foreground=[("active", "#ffffff")])

        # Hauptframe
        self.main_frame = ttk.Frame(master, padding="20")
        self.main_frame.grid(row=0, column=0, sticky="NSEW")

        # Datei-Auswahl Label
        self.label = ttk.Label(self.main_frame, text="Wähle die Excel-Datei mit den Verkaufsdaten:")
        self.label.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="W")

        # Dropdown-Menü für den Verlauf
        self.recent_files = self.load_recent_files()
        self.selected_file = tk.StringVar(value="Datei auswählen")

        self.dropdown = ttk.Combobox(self.main_frame, textvariable=self.selected_file, values=self.recent_files)
        self.dropdown.grid(row=1, column=0, columnspan=2, padx=10, pady=10, sticky="EW")
        self.dropdown.bind("<<ComboboxSelected>>", self.on_file_select)

        # Datei auswählen Button
        self.browse_button = ttk.Button(self.main_frame, text="Andere Datei auswählen", command=self.browse_file)
        self.browse_button.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="EW")

        # Fortschrittsanzeige
        self.progress = ttk.Progressbar(self.main_frame, mode='indeterminate')
        self.progress.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="EW")
        self.progress.grid_remove()  # Fortschrittsanzeige zunächst verstecken

        # Prognose erstellen Button
        self.predict_button = ttk.Button(self.main_frame, text="Prognose erstellen", command=self.start_prediction, state=tk.DISABLED)
        self.predict_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="EW")

        # Prognose anzeigen Button
        self.show_prediction_button = ttk.Button(self.main_frame, text="Prognose anzeigen", command=self.show_prediction, state=tk.DISABLED)
        self.show_prediction_button.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="EW")

        # Einzelne Tagesprognose Button
        self.single_day_button = ttk.Button(self.main_frame, text="Prognose für bestimmten Tag", command=self.predict_single_day, state=tk.DISABLED)
        self.single_day_button.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky="EW")

        # Beenden Button
        self.quit_button = ttk.Button(self.main_frame, text="Beenden", command=self.save_and_quit)
        self.quit_button.grid(row=7, column=0, columnspan=2, padx=10, pady=10, sticky="EW")

        # Kommandozeilen-Eingabe
        self.command_entry = ttk.Entry(self.main_frame, font=("Helvetica", 12))
        self.command_entry.grid(row=8, column=0, columnspan=2, padx=10, pady=10, sticky="EW")
        self.command_entry.bind("<Return>", self.handle_command)

        # Fenstergröße anpassen
        master.grid_columnconfigure(0, weight=1)
        master.grid_rowconfigure(0, weight=1)

        self.filepath = None
        self.daten = None
        self.modell = None
        self.gesamt_daten = None  # Variable zum Speichern der prognostizierten Daten

        # Load configuration
        self.config = self.load_config()

        if self.config:
            # Überprüfen, ob die automatische Prognose aktiviert ist
            self.auto_prognose = self.config.get('auto_prognose', False)
            print(f"Auto Prognose: {self.auto_prognose}")  # Debug-Ausgabe, um den Wert von auto_prognose zu überprüfen
            if self.auto_prognose is True:
                self.predict_button.grid_remove()  # Entferne den "Prognose erstellen" Button
        else:
            self.auto_prognose = False

    def load_recent_files(self):
        """Lädt die zuletzt verwendeten Dateien aus einer JSON-Datei."""
        if os.path.exists('recent_files.json'):
            with open('recent_files.json', 'r') as f:
                return json.load(f)
        return []

    def save_recent_file(self, filepath):
        """Speichert die zuletzt verwendeten Dateien in einer JSON-Datei."""
        if filepath not in self.recent_files:
            self.recent_files.insert(0, filepath)
            self.recent_files = self.recent_files[:5]  # Begrenze den Verlauf auf 5 Einträge
            with open('recent_files.json', 'w') as f:
                json.dump(self.recent_files, f)

    def on_file_select(self, event):
        """Wird aufgerufen, wenn eine Datei aus dem Verlauf ausgewählt wird."""
        self.filepath = self.selected_file.get()
        self.load_data()
        
        # Auto-Prognose Check
        if self.auto_prognose is True:
            self.predict_button.grid_remove()  # Entferne den "Prognose erstellen" Button
            self.start_prediction()  # Starte die Prognose automatisch
        else:
            self.predict_button.config(state=tk.NORMAL)
        
        self.single_day_button.config(state=tk.NORMAL)

    def browse_file(self):
        """Öffnet den Dateidialog, um eine neue Datei auszuwählen."""
        self.filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.filepath:
            self.save_recent_file(self.filepath)
            self.load_data()

            # Auto-Prognose Check
            if self.auto_prognose:
                self.predict_button.grid_remove()  # Entferne den "Prognose erstellen" Button
                self.start_prediction()  # Starte die Prognose automatisch
            else:
                self.predict_button.config(state=tk.NORMAL)
            
            self.single_day_button.config(state=tk.NORMAL)

    def load_data(self):
        try:
            self.daten = pd.read_excel(self.filepath)

            # Datenbereinigung: Fehlende Werte entfernen
            self.daten.dropna(inplace=True)

            # Nachricht anzeigen, dass Daten erfolgreich geladen wurden
            messagebox.showinfo("Erfolg", "Daten erfolgreich geladen und bereinigt.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Datei: {e}")
            print(f"Fehler beim Laden der Datei: {e}")  # Debug-Ausgabe

    def start_prediction(self):
        self.progress.grid()
        self.progress.start()
        threading.Thread(target=self.make_prediction).start()

    def make_prediction(self):
        try:
            # Schritt 2: Datenvorbereitung
            self.daten['Datum'] = pd.to_datetime(self.daten['Datum'])
            self.daten.set_index('Datum', inplace=True)

            print("Shape der historischen Daten:", self.daten.shape)  # Debug-Ausgabe

            # Erstellen von Features für das Modell
            self.daten['Monat'] = self.daten.index.month
            self.daten['Jahr'] = self.daten.index.year

            X = self.daten[['Monat', 'Jahr']]
            y = self.daten['Verkaufte Menge']

            # Modellanpassung
            self.modell = GradientBoostingRegressor(n_estimators=100, learning_rate=0.1, max_depth=3, random_state=42)
            self.modell.fit(X, y)

            # Prognose erstellen für 12 Monate (1 Jahr)
            monate_in_zukunft = 12
            forecast_index = pd.date_range(start=self.daten.index[-1] + relativedelta(months=1), 
                                           periods=monate_in_zukunft, 
                                           freq='M')
            future_features = pd.DataFrame({
                'Monat': forecast_index.month,
                'Jahr': forecast_index.year
            })

            forecast = self.modell.predict(future_features)
            forecast_series = pd.Series(forecast, index=forecast_index, name='Vorhersage')

            # Sicherstellen, dass alle Datenreihen die gleiche Dimension haben
            self.gesamt_daten = pd.concat([self.daten[['Verkaufte Menge']], forecast_series], axis=1)

            # Debugging: Überprüfen der Dimensionen und Werte
            print("Forecasted Values:")
            print(forecast_series)
            print("Combined Data (gesamt_daten):")
            print(self.gesamt_daten)

            # Aktiviert den Button "Prognose anzeigen"
            self.master.after(0, lambda: self.show_prediction_button.config(state=tk.NORMAL))

        except Exception as e:
            print(f"Fehler bei der Prognose: {e}")  # Debug-Ausgabe
            messagebox.showerror("Fehler", f"Fehler bei der Prognose: {e}")
        finally:
            self.progress.stop()
            self.progress.grid_remove()

    def show_prediction(self):
        """Zeigt die zuvor erstellte Prognose an."""
        if self.gesamt_daten is not None:
            self.plot_results(self.gesamt_daten)
        else:
            messagebox.showerror("Fehler", "Bitte erstellen Sie zuerst die Prognose.")

    def plot_results(self, gesamt_daten):
        try:
            plt.figure(figsize=tuple(self.config.get('figure_size', [10, 6])))

            # Plotten der historischen Daten
            plt.plot(gesamt_daten.index, 
                     gesamt_daten['Verkaufte Menge'], 
                     label='Verkaufte Menge', 
                     color=self.config.get('line_color', '#7289da'), 
                     linestyle=self.config.get('line_style', '-'))

            # Plotten der Vorhersage
            plt.plot(gesamt_daten.index, 
                     gesamt_daten['Vorhersage'], 
                     label='Vorhersage', 
                     color=self.config.get('prediction_color', '#f04747'), 
                     linestyle=self.config.get('prediction_style', '--'), 
                     linewidth=self.config.get('line_width', 2.5))

            plt.xlabel(self.config.get('xlabel', 'Monat'), color='#c9d1d9')
            plt.ylabel(self.config.get('ylabel', 'Verkaufte Menge'), color='#c9d1d9')
            plt.title(self.config.get('title', 'Verkaufsprognose für Nachtsichtbrillen'), color='#c9d1d9')

            if self.config.get('show_legend', True):
                plt.legend(loc=self.config.get('legend_location', 'best'))

            plt.grid(self.config.get('grid', True))

            # Manuelle Anpassung der y-Achse
            y_min = self.config.get('y_axis_min', gesamt_daten.min().min() - 10)
            y_max = self.config.get('y_axis_max', gesamt_daten.max().max() + 10)
            plt.ylim([y_min, y_max])

            if self.config.get('save_plot', False):
                plt.savefig(self.config.get('save_path', 'prognose_plot.png'))

            plt.show()

            # Zeige eine Nachricht an, wenn die Vorhersage abgeschlossen ist
            messagebox.showinfo("Abgeschlossen", "Die Verkaufsprognose für 1 Jahr wurde erfolgreich abgeschlossen.")

            # Plot der Originaldaten für Analyse
            self.plot_original_data()

        except Exception as e:
            print(f"Fehler beim Plotten der Ergebnisse: {e}")  # Debug-Ausgabe
            messagebox.showerror("Fehler", f"Fehler beim Plotten der Ergebnisse: {e}")

    def plot_original_data(self):
        try:
            plt.figure(figsize=tuple(self.config.get('figure_size', [10, 6])))
            plt.plot(self.daten.index, self.daten['Verkaufte Menge'], 
                     label='Verkaufte Menge', 
                     color=self.config.get('line_color', '#7289da'))
            plt.xlabel(self.config.get('xlabel', 'Datum'), color='#c9d1d9')
            plt.ylabel(self.config.get('ylabel', 'Verkaufte Menge'), color='#c9d1d9')
            plt.title(self.config.get('title', 'Verkaufte Menge über die Zeit'), color='#c9d1d9')

            if self.config.get('show_legend', True):
                plt.legend(loc=self.config.get('legend_location', 'best'))

            plt.grid(self.config.get('grid', True))
            plt.show()
        except Exception as e:
            print(f"Fehler beim Plotten der Originaldaten: {e}")  # Debug-Ausgabe
            messagebox.showerror("Fehler", f"Fehler beim Plotten der Originaldaten: {e}")

    def predict_single_day(self):
        """Öffnet einen Dialog zur Eingabe eines Datums und gibt die Prognose für diesen Tag aus."""
        if not self.modell:
            messagebox.showerror("Fehler", "Bitte erstellen Sie zuerst die Prognose für das Jahr.")
            return

        # Datum eingeben
        date_str = simpledialog.askstring("Datumseingabe", "Geben Sie ein Datum ein (YYYY-MM-DD):")
        if not date_str:
            return  # Benutzer hat den Dialog abgebrochen

        try:
            target_date = pd.to_datetime(date_str)
        except ValueError:
            messagebox.showerror("Fehler", "Ungültiges Datum. Bitte geben Sie das Datum im Format YYYY-MM-DD ein.")
            return

        # Überprüfen, ob das Datum nach dem letzten Datum der historischen Daten liegt
        if target_date <= self.daten.index[-1]:
            messagebox.showerror("Fehler", "Das Datum liegt vor oder am letzten bekannten Datum. Bitte wählen Sie ein späteres Datum.")
            return

        # Features für das Ziel-Datum erstellen
        target_features = pd.DataFrame({
            'Monat': [target_date.month],
            'Jahr': [target_date.year]
        })

        # Prognose für den eingegebenen Tag
        prediction = self.modell.predict(target_features)[0]
        messagebox.showinfo("Prognose", f"Die Verkaufsprognose für den {target_date.strftime('%Y-%m-%d')} beträgt {prediction:.2f} Einheiten.")

    def handle_command(self, event):
        """Verarbeitet die Eingabe von Befehlen in der Kommandozeile."""
        command = self.command_entry.get().strip().lower()
        if command == "vda prognose":
            if self.gesamt_daten is not None:
                self.plot_results(self.gesamt_daten)
            else:
                messagebox.showerror("Fehler", "Bitte erstellen Sie zuerst die Prognose.")
        elif command.startswith("vda prognose "):
            try:
                date_str = command.split(" ")[-1]
                self.command_entry.delete(0, tk.END)  # Kommandozeile löschen
                self.predict_single_day_with_date(date_str)
            except Exception as e:
                messagebox.showerror("Fehler", f"Ungültiger Befehl: {e}")
        elif command == "vda data":
            if self.filepath:
                messagebox.showinfo("Datei", f"Die zuletzt verwendete Datei ist: {self.filepath}")
            else:
                messagebox.showerror("Fehler", "Keine Datei wurde zuletzt verwendet.")
        else:
            messagebox.showerror("Fehler", "Ungültiger Befehl.")

    def predict_single_day_with_date(self, date_str):
        """Führt die Vorhersage für ein spezifisches Datum durch."""
        if not self.modell:
            messagebox.showerror("Fehler", "Bitte erstellen Sie zuerst die Prognose für das Jahr.")
            return

        try:
            target_date = pd.to_datetime(date_str)
        except ValueError:
            messagebox.showerror("Fehler", "Ungültiges Datum. Bitte geben Sie das Datum im Format YYYY-MM-DD ein.")
            return

        if target_date <= self.daten.index[-1]:
            messagebox.showerror("Fehler", "Das Datum liegt vor oder am letzten bekannten Datum. Bitte wählen Sie ein späteres Datum.")
            return

        target_features = pd.DataFrame({
            'Monat': [target_date.month],
            'Jahr': [target_date.year]
        })

        prediction = self.modell.predict(target_features)[0]
        messagebox.showinfo("Prognose", f"Die Verkaufsprognose für den {target_date.strftime('%Y-%m-%d')} beträgt {prediction:.2f} Einheiten.")

    def save_and_quit(self):
        """Speichert den Verlauf und schließt die Anwendung."""
        if self.filepath:
            self.save_recent_file(self.filepath)
        self.master.quit()

    def load_config(self):
        try:
            with open('config.json', 'r') as f:
                config = json.load(f)
            print("Konfiguration geladen:", config)  # Debug-Ausgabe
            return config
        except FileNotFoundError:
            messagebox.showerror("Fehler", "Konfigurationsdatei 'config.json' nicht gefunden.")
            return None  # Rückgabe von None, wenn die Datei nicht gefunden wird
        except json.JSONDecodeError as e:
            messagebox.showerror("Fehler", f"Fehler beim Parsen der Konfigurationsdatei: {e}")
            return None  # Rückgabe von None, wenn ein JSON-Parsing-Fehler auftritt

if __name__ == "__main__":
    root = tk.Tk()
    app = VerkaufsprognoseApp(root)
    root.mainloop()

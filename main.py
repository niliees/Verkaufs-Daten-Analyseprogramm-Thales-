import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.ensemble import GradientBoostingRegressor
import tkinter as tk
from tkinter import filedialog, messagebox
from dateutil.relativedelta import relativedelta
import threading
import json
import sys

class VerkaufsprognoseApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Verkaufsprognose für Nachtsichtbrillen")

        self.label = tk.Label(master, text="Wähle die Excel-Datei mit den Verkaufsdaten:")
        self.label.pack(pady=10)

        self.browse_button = tk.Button(master, text="Datei auswählen", command=self.browse_file)
        self.browse_button.pack(pady=10)

        self.predict_button = tk.Button(master, text="Prognose erstellen", command=self.start_prediction, state=tk.DISABLED)
        self.predict_button.pack(pady=10)

        self.quit_button = tk.Button(master, text="Beenden", command=master.quit)
        self.quit_button.pack(pady=10)

        self.filepath = None
        self.daten = None

        # Load configuration
        self.config = self.load_config()

    def load_config(self):
        try:
            with open('config.json', 'r') as f:
                config = json.load(f)
            print("Konfiguration geladen:", config)  # Debug-Ausgabe
            return config
        except FileNotFoundError:
            messagebox.showerror("Fehler", "Konfigurationsdatei 'config.json' nicht gefunden.")
            sys.exit(1)  # Programm beenden, wenn die Konfigurationsdatei nicht gefunden wird
        except json.JSONDecodeError as e:
            messagebox.showerror("Fehler", f"Fehler beim Parsen der Konfigurationsdatei: {e}")
            sys.exit(1)

    def browse_file(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.filepath:
            self.load_data()
            self.predict_button.config(state=tk.NORMAL)

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
            modell = GradientBoostingRegressor(n_estimators=100, learning_rate=0.1, max_depth=3, random_state=42)
            modell.fit(X, y)

            # Prognose erstellen für 12 Monate (1 Jahr)
            monate_in_zukunft = 12
            forecast_index = pd.date_range(start=self.daten.index[-1] + relativedelta(months=1), 
                                           periods=monate_in_zukunft, 
                                           freq='M')
            future_features = pd.DataFrame({
                'Monat': forecast_index.month,
                'Jahr': forecast_index.year
            })

            forecast = modell.predict(future_features)
            forecast_series = pd.Series(forecast, index=forecast_index, name='Vorhersage')

            # Sicherstellen, dass alle Datenreihen die gleiche Dimension haben
            gesamt_daten = pd.concat([self.daten[['Verkaufte Menge']], forecast_series], axis=1)

            # Debugging: Überprüfen der Dimensionen und Werte
            print("Forecasted Values:")
            print(forecast_series)
            print("Combined Data (gesamt_daten):")
            print(gesamt_daten)

            # Plotten der Ergebnisse im Hauptthread
            self.master.after(0, self.plot_results, gesamt_daten)

        except Exception as e:
            print(f"Fehler bei der Prognose: {e}")  # Debug-Ausgabe
            messagebox.showerror("Fehler", f"Fehler bei der Prognose: {e}")

    def plot_results(self, gesamt_daten):
        try:
            plt.figure(figsize=tuple(self.config.get('figure_size', [10, 6])))

            # Plotten der historischen Daten
            plt.plot(gesamt_daten.index, 
                     gesamt_daten['Verkaufte Menge'], 
                     label='Verkaufte Menge', 
                     color=self.config.get('line_color', 'blue'), 
                     linestyle=self.config.get('line_style', '-'))

            # Plotten der Vorhersage
            plt.plot(gesamt_daten.index, 
                     gesamt_daten['Vorhersage'], 
                     label='Vorhersage', 
                     color=self.config.get('prediction_color', 'red'), 
                     linestyle=self.config.get('prediction_style', '--'), 
                     linewidth=self.config.get('line_width', 2.5))

            plt.xlabel(self.config.get('xlabel', 'Monat'))
            plt.ylabel(self.config.get('ylabel', 'Verkaufte Menge'))
            plt.title(self.config.get('title', 'Verkaufsprognose für Nachtsichtbrillen'))

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
                     color=self.config.get('line_color', 'blue'))
            plt.xlabel(self.config.get('xlabel', 'Datum'))
            plt.ylabel(self.config.get('ylabel', 'Verkaufte Menge'))
            plt.title(self.config.get('title', 'Verkaufte Menge über die Zeit'))

            if self.config.get('show_legend', True):
                plt.legend(loc=self.config.get('legend_location', 'best'))

            plt.grid(self.config.get('grid', True))
            plt.show()
        except Exception as e:
            print(f"Fehler beim Plotten der Originaldaten: {e}")  # Debug-Ausgabe
            messagebox.showerror("Fehler", f"Fehler beim Plotten der Originaldaten: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = VerkaufsprognoseApp(root)
    root.mainloop()

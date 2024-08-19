import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.statespace.sarimax import SARIMAX
import tkinter as tk
from tkinter import filedialog, messagebox
from dateutil.relativedelta import relativedelta
import threading

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

    def browse_file(self):
        self.filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.filepath:
            self.load_data()
            self.predict_button.config(state=tk.NORMAL)

    def load_data(self):
        try:
            self.daten = pd.read_excel(self.filepath)
            messagebox.showinfo("Erfolg", "Daten erfolgreich geladen.")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Datei: {e}")

    def start_prediction(self):
        # Starte die Vorhersage in einem eigenen Thread, um die GUI nicht zu blockieren
        threading.Thread(target=self.make_prediction).start()

    def make_prediction(self):
        try:
            # Schritt 2: Datenvorbereitung
            self.daten['Datum'] = pd.to_datetime(self.daten['Datum'])
            self.daten.set_index('Datum', inplace=True)

            # Modellbildung mit SARIMA (Seasonal ARIMA)
            modell = SARIMAX(self.daten['Verkaufte Menge'], 
                             order=(1, 1, 1), 
                             seasonal_order=(1, 1, 1, 12), 
                             enforce_stationarity=False, 
                             enforce_invertibility=False)
            modell_fit = modell.fit(disp=False)

            # Prognose erstellen für 12 Monate (1 Jahr)
            monate_in_zukunft = 12
            forecast = modell_fit.get_forecast(steps=monate_in_zukunft)
            forecast_index = pd.date_range(start=self.daten.index[-1] + relativedelta(months=1), 
                                           periods=monate_in_zukunft, 
                                           freq='M')
            forecast_series = pd.Series(forecast.predicted_mean, index=forecast_index)

            # Sicherstellen, dass alle Datenreihen die gleiche Dimension haben
            forecast_series = forecast_series.to_frame(name='Verkaufte Menge')
            gesamt_daten = pd.concat([self.daten[['Verkaufte Menge']], forecast_series])

            # Debugging: Überprüfen der Dimensionen
            print(f"Shape der historischen Daten: {self.daten['Verkaufte Menge'].shape}")
            print(f"Shape der Vorhersage-Daten: {forecast_series.shape}")
            print(f"Shape der kombinierten Daten: {gesamt_daten.shape}")

            # Plotten der Ergebnisse im Hauptthread
            self.master.after(0, self.plot_results, gesamt_daten)

        except Exception as e:
            print(f"Fehler bei der Prognose: {e}")
            messagebox.showerror("Fehler", f"Fehler bei der Prognose: {e}")

    def plot_results(self, gesamt_daten):
        # Ergebnisse visualisieren
        plt.figure(figsize=(10, 6))
        plt.plot(gesamt_daten.index, gesamt_daten['Verkaufte Menge'], label='Verkaufte Menge')

        plt.xlabel('Monat')
        plt.ylabel('Verkaufte Menge')
        plt.title('Verkaufsprognose für Nachtsichtbrillen')
        plt.legend()
        plt.grid(True)
        plt.show()

        # Zeige eine Nachricht an, wenn die Vorhersage abgeschlossen ist
        messagebox.showinfo("Abgeschlossen", "Die Verkaufsprognose für 1 Jahr wurde erfolgreich abgeschlossen.")

if __name__ == "__main__":
    root = tk.Tk()
    app = VerkaufsprognoseApp(root)
    root.mainloop()

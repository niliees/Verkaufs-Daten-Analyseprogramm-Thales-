import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from statsmodels.tsa.arima.model import ARIMA
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

            # Schritt 3: Modellbildung mit ARIMA
            modell = ARIMA(self.daten['Verkaufte Menge'], order=(5, 1, 0))  # (p, d, q) - order kann angepasst werden
            modell_fit = modell.fit()

            # Schritt 4: Prognose erstellen für 12 Monate (1 Jahr)
            monate_in_zukunft = 12
            forecast = modell_fit.forecast(steps=monate_in_zukunft)

            zukuenftige_monate_datum = [self.daten.index[-1] + relativedelta(months=i) for i in range(1, monate_in_zukunft + 1)]

            # Kombiniere historische Daten und Vorhersage
            gesamt_daten = pd.concat([self.daten, pd.Series(forecast, index=zukuenftige_monate_datum, name='Verkaufte Menge')])

            # Plotten der Ergebnisse im Hauptthread
            self.master.after(0, self.plot_results, gesamt_daten)

        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler bei der Prognose: {e}")

    def plot_results(self, gesamt_daten):
        # Ergebnisse visualisieren
        plt.plot(gesamt_daten.index, gesamt_daten, label='Verkaufte Menge')

        plt.xlabel('Monat')
        plt.ylabel('Verkaufte Menge')
        plt.title('Verkaufsprognose für Nachtsichtbrillen')
        plt.legend()
        plt.show()

        # Zeige eine Nachricht an, wenn die Vorhersage abgeschlossen ist
        messagebox.showinfo("Abgeschlossen", "Die Verkaufsprognose für 1 Jahr wurde erfolgreich abgeschlossen.")

if __name__ == "__main__":
    root = tk.Tk()
    app = VerkaufsprognoseApp(root)
    root.mainloop()

#für mehr gehe auf meinen Linktree
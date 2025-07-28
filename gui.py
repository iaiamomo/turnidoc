import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from pathlib import Path
from main import process_turni

class ExcelProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Turni File Processor")
        self.root.geometry("600x400")
        
        # Variabili per i percorsi dei file
        self.input_file = ""
        self.output_file = ""
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame principale
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configurazione griglia
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Titolo
        title_label = ttk.Label(main_frame, 
                                text="Turni File Processor",
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Informazioni sull'app
        ttk.Label(main_frame, 
                  text="Ricorda di utilizzare celle singole e che l'ordine delle ultime colonne sia amb, 104, agg, ferie."
                  ).grid(row=1, column=0, columnspan=3, pady=(0, 10))
        
        # Sezione selezione file input
        ttk.Label(main_frame, text="File di input:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.input_label = ttk.Label(main_frame, text="Nessun file selezionato", foreground="gray")
        self.input_label.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        
        input_btn = ttk.Button(main_frame, text="Seleziona File", command=self.select_input_file)
        input_btn.grid(row=2, column=2, padx=(10, 0), pady=5)
        
        # Sezione selezione file output
        ttk.Label(main_frame, text="Salva come:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.output_label = ttk.Label(main_frame, text="Nessuna destinazione selezionata", foreground="gray")
        self.output_label.grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(10, 0), pady=5)
        
        output_btn = ttk.Button(main_frame, text="Scegli Destinazione", command=self.select_output_file)
        output_btn.grid(row=3, column=2, padx=(10, 0), pady=5)
        
        # Separatore
        separator = ttk.Separator(main_frame, orient='horizontal')
        separator.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=20)
        
        # Pulsante di processamento
        process_btn = ttk.Button(main_frame, text="Processa File", command=self.process_file)
        process_btn.grid(row=6, column=0, columnspan=3, pady=20)
        
        # Area di log
        log_frame = ttk.LabelFrame(main_frame, text="Log", padding="5")
        log_frame.grid(row=8, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(8, weight=1)
        
        self.log_text = tk.Text(log_frame, height=8)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar per il log
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.configure(yscrollcommand=scrollbar.set)
    
    def log_message(self, message):
        """Aggiunge un messaggio al log"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def select_input_file(self):
        """Seleziona il file Excel di input"""
        file_path = filedialog.askopenfilename(
            title="Seleziona file Excel",
            filetypes=[
                ("File Excel", "*.xlsx *.xls"),
                ("Tutti i file", "*.*")
            ]
        )
        
        if file_path:
            self.input_file = file_path
            self.input_label.config(text=os.path.basename(file_path), foreground="black")
            self.log_message(f"File di input selezionato: {file_path}")
    
    def select_output_file(self):
        """Seleziona dove salvare il file processato"""
        file_path = filedialog.asksaveasfilename(
            title="Salva file come",
            defaultextension=".xlsx",
            filetypes=[
                ("File Excel", "*.xlsx"),
                ("File CSV", "*.csv"),
                ("Tutti i file", "*.*")
            ]
        )
        
        if file_path:
            self.output_file = file_path
            self.output_label.config(text=os.path.basename(file_path), foreground="black")
            self.log_message(f"File di output selezionato: {file_path}")
    
    def process_file(self):
        """Processa il file Excel"""        
        if not self.input_file:
            messagebox.showerror("Errore", "Seleziona un file di input!")
            return
        
        if not self.output_file:
            messagebox.showerror("Errore", "Seleziona una destinazione per il file di output!")
            return
        
        try:
            self.log_message("Inizio processamento del file...")
            
            # Leggi il file Excel
            self.log_message("Lettura del file Excel...")
            df = pd.read_excel(self.input_file)
            self.log_message(f"File letto con successo. Righe: {len(df)}, Colonne: {len(df.columns)}")
            
            # Processamento del file (personalizza questa sezione)
            df_processed = self.apply_processing(df)
            
            # Salva il file processato
            self.log_message("Salvataggio del file processato...")
            
            # Determina il formato di output dal nome del file
            output_ext = Path(self.output_file).suffix.lower()
            
            if output_ext == '.csv':
                df_processed.to_csv(self.output_file, index=False, header=False)
            else:
                df_processed.to_excel(self.output_file, index=False, header=False)
            
            self.log_message(f"File salvato con successo: {self.output_file}")
            messagebox.showinfo("Successo", "File processato e salvato con successo!")
            
        except Exception as e:
            error_msg = f"Errore durante il processamento: {str(e)}"
            self.log_message(error_msg)
            messagebox.showerror("Errore", error_msg)
    
    def apply_processing(self, df):
        """Applica le trasformazioni al DataFrame"""
        
        df_processed = process_turni(df)
        
        return df_processed

def main():
    root = tk.Tk()
    app = ExcelProcessor(root)
    root.mainloop()

if __name__ == "__main__":
    main()
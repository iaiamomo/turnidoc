# TURNIDOC

1. Installazione delle dipendenze

    Prima di tutto, crea un ambiente virtuale e installa le librerie necessarie:
    ```bash
    # Crea un ambiente virtuale
    python -m venv excel_processor_env

    # Attiva l'ambiente virtuale
    # Su Windows:
    excel_processor_env\Scripts\activate
    # Su macOS/Linux:
    source excel_processor_env/bin/activate

    # Installa le dipendenze
    pip install pandas openpyxl pyinstaller tk
    ```

2. File requirements.txt

    Crea un file requirements.txt con:
    ```bash
    pandas>=2.0.0
    openpyxl>=3.1.0
    pyinstaller>=5.0.0
    ```

3. Conversione in EXE con PyInstaller

    Comando base:
    ```bash
    pyinstaller --onefile --windowed gui.py
    ```
    Comando avanzato con icona e nome personalizzato:
    ```bash
    pyinstaller --onefile --windowed --name "TurniProcessor" --icon=icon.ico gui.py
    ```
    Opzioni PyInstaller utili:
    ```
    --onefile: Crea un singolo file .exe
    --windowed: Nasconde la console (importante per app GUI)
    --name: Nome personalizzato per l'exe
    --icon: Icona personalizzata (.ico)
    --add-data: Aggiunge file extra se necessari
    ```

4. Per ridurre la dimensione dell'.exe
    ```bash
    pyinstaller --onefile --windowed --optimize=2 --name "TurniProcessor" --icon=icon.ico gui.py
    ```
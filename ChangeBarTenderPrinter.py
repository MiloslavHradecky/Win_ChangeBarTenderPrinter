#!/usr/bin/env python3

import os
import logging
import subprocess
import win32com.client
import configparser

__version__ = '2.0.0.0'


class PrinterChanger:
    """
    Třída pro změnu tiskárny u Bartender souborů.
    Obsahuje kontrolu instalace BarTenderu.

    - Načítá složku nebo složky s etiketami z 'config.ini'
    - Prochází soubory '.btw' a nastavuje správné tiskárny
    - Ukládá změny zpět do souboru
    """

    def __init__(self, config_file='config.ini'):
        """
        Inicializuje 'PrinterChanger' a načte konfiguraci.

        :param config_file: Cesta ke konfiguračnímu souboru ('config.ini')
        """
        config = configparser.ConfigParser()
        config.optionxform = str
        config.read(config_file)

        self.bartender_path = config.get('Paths', 'bartender_path')

        # 📌 Načteme složky a rozdělíme podle středníku (';')
        self.labels_folders = config.get('Paths', 'labels_folders').split('; ')

        # 📌 Odstraníme mezery kolem cest
        self.labels_folders = [folder.strip() for folder in self.labels_folders]

        # 📌 Převod 'PrinterMapping' z INI na slovník v Pythonu
        self.prefix_printer_map = {key: value for key, value in config.items('PrinterMapping')}

        self.logger = LoggerManager()

        # 📌 Kontrola, zda je BarTender nainstalovaný
        if not self.is_bartender_installed():
            self.logger.log('Error', '❌ BarTender není nainstalován! Zkontrolujte instalaci před spuštěním skriptu.')
            exit(1)  # ✅ Ukončí skript s chybovým kódem

    def check_paths(self):
        """ 📌 Ověří, zda všechny cesty uvedené v klíči existují. """
        paths = self.labels_folders
        result = {path: os.path.exists(path) for path in paths}  # ✅ Ověříme existenci každé cesty
        return result

    def is_bartender_installed(self):
        """
        Ověří, zda existuje 'bartender.exe' v zadané cestě.

        :return: 'True', pokud soubor existuje, jinak 'False'
        """
        return os.path.exists(self.bartender_path)

    def kill_bartender_processes(self):
        """ Ukončí všechny běžící instance BarTender (Cmdr.exe a bartend.exe). """
        try:
            subprocess.run('taskkill /f /im cmdr.exe 1>nul 2>nul', shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
            subprocess.run('taskkill /f /im bartend.exe 1>nul 2>nul', shell=True, creationflags=subprocess.CREATE_NO_WINDOW)

        except subprocess.CalledProcessError as e:
            self.logger.log('Error', f'❗ Chyba při ukončování BarTender procesů: {e}')

    def change_printer_for_files(self):
        """
        Prochází soubory '.btw' ve více složkách a nastavuje správnou tiskárnu.

        - Ověří existenci složek před spuštěním procesu
        - Otevře Bartender aplikaci
        - Pro každý '.btw' soubor nastaví tiskárnu podle prefixu
        - Uloží změny a zaloguje výsledek
        """

        # 📌 Ověříme existenci všech složek z 'labels_folders'
        paths_status = self.check_paths()

        missing_paths = [path for path, exists in paths_status.items() if not exists]
        if missing_paths:
            self.logger.log('Error', f'❌ Chyba následující složky neexistují: {", ".join(missing_paths)}')
            exit(1)  # ✅ Ukončíme skript s chybovým kódem

        # 📌 Zavřeme všechny bartender a commander procesy
        self.kill_bartender_processes()

        bt_app = win32com.client.Dispatch('BarTender.Application')
        bt_app.Visible = False

        self.logger.start_logging_session()

        # 📌 Projdeme všechny složky, které jsme načetli z configu
        for folder_path in self.labels_folders:
            if os.path.exists(folder_path):
                self.logger.log('Info', f'📂 Zpracovává se složka: {folder_path}')
                self.process_folder(bt_app, folder_path)
            else:
                self.logger.log('Warning', f'⚠ Složka neexistuje: {folder_path}')

        bt_app.Quit(1)  # ✅ btDoNotSaveChanges

    def process_folder(self, bt_app, folder_path):
        """
        Změní tiskárnu pouze pro soubory '.btw', které mají povolený prefix.
        """
        for filename in os.listdir(folder_path):
            if filename.endswith('.btw'):
                # 📌 Ověříme, zda soubor začíná některým z povolených prefixů
                prefix = next((p for p in self.prefix_printer_map if filename.startswith(p)), None)

                if prefix:  # ✅ Pokud soubor začíná povoleným prefixem
                    file_path = os.path.join(folder_path, filename)
                    try:
                        bt_format = bt_app.Formats.Open(file_path, False, '')
                        if bt_format:
                            printer_name = self.prefix_printer_map[prefix]  # ✅ Načteme správnou tiskárnu z configu
                            bt_format.Printer = printer_name
                            bt_format.Save()
                            bt_format.Close(1)  # ✅ btDoNotSaveChanges
                            self.logger.log('Info', f'ℹ️ Tiskárna "{printer_name}" úspěšně změněna pro soubor: {filename}')
                        else:
                            self.logger.log('Error', f'❗ Selhalo otevření souboru: {filename}')
                    except Exception as e:
                        self.logger.log('Error', f'❗ Chyba při zpracování souboru {filename}: {e}')
                else:
                    pass


class LoggerManager:
    """
    Třída pro správu logování aplikace.

    - Nastavuje 'logging' s časovým razítkem
    - Přidává prázdný řádek pouze při spuštění skriptu
    - Umožňuje logování různých úrovní ('Info', 'Warning', 'Error')
    """

    def __init__(self, config_file='config.ini'):
        """
        Inicializuje 'LoggerManager' a nastaví konfiguraci logování.

        :param config_file: Cesta ke konfiguračnímu souboru ('config.ini')
        """
        config = configparser.ConfigParser()
        config.optionxform = str
        config.read(config_file)

        self.log_file_path = os.path.abspath(config.get('Paths', 'log_file_path'))

        # 📌 Vytvoření adresáře pro log soubor, pokud neexistuje
        log_dir = os.path.dirname(self.log_file_path)
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)

        logging.basicConfig(
            filename=self.log_file_path,
            level=logging.INFO,
            encoding='utf-8',
            format='%(asctime)s_%(levelname)-7s: %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

        self.logger = logging.getLogger(__name__)

    def start_logging_session(self):
        """
        Přidá prázdný řádek při spuštění skriptu, aby oddělil každé spuštění od předchozího.
        """
        with open(self.log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write('\n')

    def log(self, level, message):
        """ Zaloguje zprávu podle zvolené úrovně. """
        if level == 'Info':
            self.logger.info(message)
        elif level == 'Warning':
            self.logger.warning(message)
        elif level == 'Error':
            self.logger.error(message)


# 📌 Spuštění procesu
if __name__ == '__main__':
    printer_changer = PrinterChanger()
    printer_changer.change_printer_for_files()

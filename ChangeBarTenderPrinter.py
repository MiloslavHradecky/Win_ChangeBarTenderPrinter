#!/usr/bin/env python3

import os
import logging
import win32com.client
import configparser

__version__ = '1.0.0.0'


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
        self.labels_folders = config.get('Paths', 'labels_folder').split('; ')

        # 📌 Odstraníme mezery kolem cest
        self.labels_folders = [folder.strip() for folder in self.labels_folders]

        # 📌 Převod 'PrinterMapping' z INI na slovník v Pythonu
        self.prefix_printer_map = {key: value for key, value in config.items('PrinterMapping')}

        self.logger = LoggerManager()

        # 📌 Kontrola, zda je BarTender nainstalovaný
        if not self.is_bartender_installed():
            self.logger.log('Error', '❌ BarTender není nainstalován! Zkontrolujte instalaci před spuštěním skriptu.')
            exit(1)  # ✅ Ukončí skript s chybovým kódem

    def is_bartender_installed(self):
        """
        Ověří, zda existuje 'bartender.exe' v zadané cestě.

        :return: 'True', pokud soubor existuje, jinak 'False'
        """
        return os.path.exists(self.bartender_path)

    def change_printer_for_files(self):
        """
        Prochází soubory '.btw' ve více složkách a nastavuje správnou tiskárnu.

        - Otevře Bartender aplikaci
        - Pro každý '.btw' soubor nastaví tiskárnu podle prefixu
        - Uloží změny a zaloguje výsledek
        """
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
        Změní tiskárnu pro všechny soubory '.btw' v dané složce.
        """
        for filename in os.listdir(folder_path):
            if filename.endswith('.btw'):
                file_path = os.path.join(folder_path, filename)
                try:
                    bt_format = bt_app.Formats.Open(file_path, False, '')
                    if bt_format:
                        # 📌 Dynamicky načítáme tiskárnu z 'config.ini'
                        printer_name = self.prefix_printer_map.get(filename[:filename.index('_')], 'Default Printer')
                        bt_format.Printer = printer_name
                        bt_format.Save()
                        bt_format.Close(1)  # ✅ btDoNotSaveChanges
                        self.logger.log('Info', f'Tiskárna "{printer_name}" úspěšně změněna pro soubor: {filename}')
                    else:
                        self.logger.log('Error', f'Selhalo otevření souboru: {filename}')
                except Exception as e:
                    self.logger.log('Error', f'Chyba při zpracování souboru {filename}: {e}')


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

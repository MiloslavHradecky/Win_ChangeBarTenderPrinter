#!/usr/bin/env python3

import os
import logging
import win32com.client
import configparser

__version__ = '1.0.0.0'


class PrinterChanger:
    """
    Třída pro změnu tiskárny u Bartender souborů.

    - Načítá složku s etiketami z 'config.ini'
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

        self.folder_path = config.get('Paths', 'labels_folder')

        # 📌 Převod 'PrinterMapping' z INI na slovník v Pythonu
        self.prefix_printer_map = {key: value for key, value in config.items('PrinterMapping')}

        self.logger = LoggerManager()

    def change_printer_for_files(self):
        """
        Prochází soubory '.btw' a nastavuje správnou tiskárnu.

        - Otevře Bartender aplikaci
        - Pro každý '.btw' soubor nastaví tiskárnu podle prefixu
        - Uloží změny a zaloguje výsledek
        """
        bt_app = win32com.client.Dispatch('BarTender.Application')
        bt_app.Visible = False

        self.logger.start_logging_session()

        # 📌 Projdeme všechny soubory ve složce
        for filename in os.listdir(self.folder_path):
            if filename.endswith('.btw'):
                for prefix, printer_name in self.prefix_printer_map.items():
                    if filename.startswith(prefix):
                        file_path = os.path.join(self.folder_path, filename)
                        try:
                            bt_format = bt_app.Formats.Open(file_path, False, '')
                            if bt_format:
                                bt_format.Printer = printer_name
                                bt_format.Save()
                                bt_format.Close(1)  # ✅ btDoNotSaveChanges
                                self.logger.log('Info', f'Tiskárna "{printer_name}" úspěšně změněna pro soubor: {filename}')
                            else:
                                self.logger.log('Error', f'Selhalo otevření souboru: {filename}')
                        except Exception as e:
                            self.logger.log('Error', f'Chyba při zpracování souboru {filename}: {e}')

        bt_app.Quit(1)  # ✅ btDoNotSaveChanges


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

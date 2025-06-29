#!/usr/bin/env python3

import os
import logging
import subprocess
import win32com.client
import configparser

__version__ = '2.0.0.0'


class PrinterChanger:
    """
    A class for changing default printers in BarTender .btw files.
    Checks BarTender installation and applies printer settings based on filename prefixes.

    - Loads label folders from 'config.ini'
    - Iterates through '.btw' files and assigns the correct printer
    - Saves changes directly to the label file
    """

    def __init__(self, config_file='config.ini'):
        """
        Initializes PrinterChanger and loads configuration from the INI file.

        :param config_file: Path to the configuration file ('config.ini')
        """
        config = configparser.ConfigParser()
        config.optionxform = str
        config.read(config_file)

        self.bartender_path = config.get('Paths', 'bartender_path')

        # 📌 Load folders and split by semicolon (Načteme složky a rozdělíme podle středníku (';'))
        self.labels_folders = config.get('Paths', 'labels_folders').split('; ')

        # 📌 Remove trailing spaces (Odstraníme mezery kolem cest)
        self.labels_folders = [folder.strip() for folder in self.labels_folders]

        # 📌 Convert 'PrinterMapping' section into a dictionary (Převod 'PrinterMapping' z INI na slovník v Pythonu)
        self.prefix_printer_map = {key: value for key, value in config.items('PrinterMapping')}

        self.logger = LoggerManager()

        # 📌 Verify BarTender installation (Kontrola, zda je BarTender nainstalovaný)
        if not self.is_bartender_installed():
            self.logger.log('Error', '❌ BarTender není nainstalován! Zkontrolujte instalaci před spuštěním skriptu.')
            exit(1)

    def check_paths(self):
        """ 📌 Checks whether each folder path listed in config exists. """
        paths = self.labels_folders
        result = {path: os.path.exists(path) for path in paths}
        return result

    def is_bartender_installed(self):
        """
        Verifies that 'bartender.exe' exists at the configured path.

        :return: True if the executable is found, False otherwise.
        """
        return os.path.exists(self.bartender_path)

    def kill_bartender_processes(self):
        """ Terminates all running BarTender and Commander processes. """
        try:
            subprocess.run('taskkill /f /im cmdr.exe 1>nul 2>nul', shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
            subprocess.run('taskkill /f /im bartend.exe 1>nul 2>nul', shell=True, creationflags=subprocess.CREATE_NO_WINDOW)

        except subprocess.CalledProcessError as e:
            self.logger.log('Error', f'❗ Chyba při ukončování BarTender procesů: {e}')

    def change_printer_for_files(self):
        """
        Updates the assigned printer in all valid '.btw' files across configured directories.

        - Validates all label paths
        - Closes any active BarTender processes
        - Opens each '.btw' file and assigns the correct printer based on prefix
        - Logs success or failure for each file
        """

        # 📌 Verify all folders from 'labels_folders' exist (Ověříme existenci všech složek z 'labels_folders')
        paths_status = self.check_paths()
        missing_paths = [path for path, exists in paths_status.items() if not exists]

        if missing_paths:
            self.logger.log('Error', f'❌ Chyba následující složky neexistují: {", ".join(missing_paths)}')
            exit(1)

        self.kill_bartender_processes()

        bt_app = win32com.client.Dispatch('BarTender.Application')
        bt_app.Visible = False

        self.logger.start_logging_session()

        for folder_path in self.labels_folders:
            if os.path.exists(folder_path):
                self.logger.log('Info', f'📂 Zpracovává se složka: {folder_path}')
                self.process_folder(bt_app, folder_path)
            else:
                self.logger.log('Warning', f'⚠ Složka neexistuje: {folder_path}')

        bt_app.Quit(1)  # ✅ btDoNotSaveChanges

    def process_folder(self, bt_app, folder_path):
        """
        Processes each '.btw' file in the given folder and applies the correct printer.

        :param bt_app: BarTender COM application instance
        :param folder_path: Full path to the directory
        """
        for filename in os.listdir(folder_path):
            if filename.endswith('.btw'):
                # 📌 Verify that the file begins with one of the allowed prefixes (Ověříme, zda soubor začíná některým z povolených prefixů)
                prefix = next((p for p in self.prefix_printer_map if filename.startswith(p)), None)

                if prefix:
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
     A logging helper class for tracking activity and errors.

    - Initializes structured logging with timestamps
    - Starts each session with a blank line for visual separation
    - Allows Info, Warning, and Error log entries
    """

    def __init__(self, config_file='config.ini'):
        """
        Initializes logging settings from the configuration file.

        :param config_file: Path to the config file ('config.ini')
        """
        config = configparser.ConfigParser()
        config.optionxform = str
        config.read(config_file)

        self.log_file_path = os.path.abspath(config.get('Paths', 'log_file_path'))

        # 📌 Create a directory for the log file if it does not exist (Vytvoření adresáře pro log soubor, pokud neexistuje)
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
        """ Appends a blank line at the start of a new logging session. """
        with open(self.log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write('\n')

    def log(self, level, message):
        """ Logs a message at the specified log level. """
        if level == 'Info':
            self.logger.info(message)
        elif level == 'Warning':
            self.logger.warning(message)
        elif level == 'Error':
            self.logger.error(message)


# 📌 Entry point (Spuštění procesu)
if __name__ == '__main__':
    printer_changer = PrinterChanger()
    printer_changer.change_printer_for_files()

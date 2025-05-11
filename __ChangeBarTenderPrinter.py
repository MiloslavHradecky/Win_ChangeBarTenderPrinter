#!/usr/bin/env python3

import os
import logging
import win32com.client

__version__ = '1.0.0.0'

# Nastavení logování s časovým razítkem.
logging.basicConfig(filename='./log/log.txt', level=logging.INFO, encoding='utf-8',
                    format='%(asctime)s_%(levelname)-7s: %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')

# Vytvoření loggeru.
logger = logging.getLogger(__name__)


# Funkce pro přidání prázdného řádku a následného logovacího záznamu
def log_with_empty_line():
    # Zkontroluje, zda soubor existuje
    if not os.path.exists('./log/log.txt'):
        open('./log/log.txt', 'w', encoding='utf-8').close()  # Vytvoří soubor, pokud neexistuje

    with open('./log/log.txt', 'a', encoding='utf-8') as log_file:
        log_file.write('\n')  # Přidá prázdný řádek


# Záznam zpráv různých úrovní
# logging.debug('Toto je debug zpráva')
# logging.info('Toto je info zpráva')
# logging.warning('Toto je varování')
# logging.error('Toto je chyba')
# logging.critical('Toto je kritická chyba')

########################################################################################################
# Skript pro změnu tiskárny
########################################################################################################

def change_printer_for_files(folder_path, prefix_printer_map):
    bt_app = win32com.client.GetActiveObject('BarTender.Application')
    bt_app.Visible = False

    log_with_empty_line()

    # Projdeme všechny soubory ve složce
    for filename in os.listdir(folder_path):
        if filename.endswith('.btw'):
            for prefix, printer_name in prefix_printer_map.items():
                if filename.startswith(prefix):
                    file_path = os.path.join(folder_path, filename)
                    try:
                        # Otevření specifického formátu Bartenderu
                        bt_format = bt_app.Formats.Open(file_path, False, '')
                        if bt_format:
                            # Nastavení nové tiskárny
                            bt_format.Printer = printer_name
                            # Uložení změn do souboru
                            bt_format.Save()
                            # Zavření formátu
                            bt_format.Close(1)  # btDoNotSaveChanges
                            logging.info(f'Tiskárna "{printer_name}" úspěšně změněna pro soubor: {filename}')
                            # print(f'Tiskárna "{printer_name}" úspěšně změněna pro soubor: {filename}')
                        else:
                            logging.error(f'Selhalo otevření souboru: {filename}')
                            # print(f'Selhalo otevření souboru: {filename}')
                    except Exception as e:
                        logging.error(f'Chyba při zpracování souboru {filename}: {e}')
                        # print(f'Chyba při zpracování souboru {filename}: {e}')

    # Ukončení aplikace Bartender
    bt_app.Quit(1)  # btDoNotSaveChanges


# Použití funkce
folder_path = r'T:\Prikazy\DataTPV\ManualLabelPrint_DfA\Etikety'
prefix_printer_map = {
    '25x10_': 'OneNote (Desktop)',
    '50x30_': 'Microsoft Print to PDF'
    # přidej další prefixy a tiskárny podle potřeby
}
change_printer_for_files(folder_path, prefix_printer_map)

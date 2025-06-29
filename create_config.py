import configparser
from io import StringIO

config = configparser.ConfigParser()
config.optionxform = str  # ✅ It retains the original capital letters! (Zachová původní velká písmena!)

config['Paths'] = {
    'log_file_path': './log/app.log',
    'labels_folders': 'T:/Prikazy/DataTPV/ManualLabelPrint_DfA/Etikety; T:/Prikazy/DataTPV/ManualLabelPrint/Etikety',
    'python_path': 'C:/Users/hradecky/AppData/Local/Programs/Python/Python313/python.exe',
    'python_script_path': 'C:/GitWork/Windows/Win_ChangeBarTenderPrinter/ChangeBarTenderPrinter.py',
    'bartender_path': 'C:/Program Files (x86)/Seagull/BarTender Suite/bartend.exe'
}

config['PrinterMapping'] = {
    '25x10_': '420t',
    '50x30_': '50x30_430t'
}

# 📌 Write configuration to StringIO for testing
configfile = StringIO()
config.write(configfile)

# 📌 Output StringIO contents to verify functionality
print(configfile.getvalue())

with open('config.ini', mode='w') as file:
    file.write(configfile.getvalue())

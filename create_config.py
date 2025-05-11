import configparser
from io import StringIO

config = configparser.ConfigParser()
config.optionxform = str  # ✅ Zachová původní velká písmena!

config['Paths'] = {
    'log_file_path': './log/app.log',
    'labels_folder': 'T:/Prikazy/DataTPV/ManualLabelPrint_DfA/Etikety',
    'python_path': 'C:/Users/hradecky/AppData/Local/Programs/Python/Python313/python.exe',
    'python_script_path': 'C:/GitWork/Windows/Win_ChangeBarTenderPrinter/ChangeBarTenderPrinter.py'
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

import configparser
from io import StringIO

config = configparser.ConfigParser()
config.optionxform = str  # ✅ Zachová původní velká písmena!

config['Paths'] = {
    'log_file_path': './log/app.log',
    'python_path': 'C:/Users/Home/AppData/Local/Programs/Python/Python313/python.exe',
    'python_script_path': 'C:/Users/Home/Documents/Coding/Windows/Win_ChangeBarTenderPrinter/ChangeBarTenderPrinter.py',
    'csv_output_file': 'T:/Prikazy/DataTPV/ManualLabelPrint/Etikety/label.csv',
    'bartender_path': 'C:/Program Files (x86)/Seagull/BarTender Suite/bartend.exe',
    'label_folder': 'T:/Prikazy/DataTPV/ManualLabelPrint/Etikety/'
}

# 📌 Write configuration to StringIO for testing
configfile = StringIO()
config.write(configfile)

# 📌 Output StringIO contents to verify functionality
print(configfile.getvalue())

with open('config.ini', mode='w') as file:
    file.write(configfile.getvalue())

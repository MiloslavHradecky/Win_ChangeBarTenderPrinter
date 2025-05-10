import configparser
from io import StringIO

config = configparser.ConfigParser()
config.optionxform = str  # ✅ Zachová původní velká písmena!

config['Paths'] = {
    'log_file_path': './log/app.log',
    'szv_input_file': 'T:/Prikazy/DataTPV/SZV.dat',
    'csv_input_file': 'T:/Prikazy/DataTPV/ManualLabelPrint/Databaze/MLP.csv',
    'csv_output_file': 'T:/Prikazy/DataTPV/ManualLabelPrint/Etikety/label.csv',
    'bartender_path': 'C:/Program Files (x86)/Seagull/BarTender Suite/bartend.exe',
    'label_folder': 'T:/Prikazy/DataTPV/ManualLabelPrint/Etikety/'
}

config['Products'] = {
    'allowed_values': '9159010'
}

# 📌 Write configuration to StringIO for testing
configfile = StringIO()
config.write(configfile)

# 📌 Output StringIO contents to verify functionality
print(configfile.getvalue())

with open('config.ini', mode='w') as file:
    file.write(configfile.getvalue())

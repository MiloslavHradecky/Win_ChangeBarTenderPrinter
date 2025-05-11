import configparser
from io import StringIO

config = configparser.ConfigParser()
config.optionxform = str  # ✅ Zachová původní velká písmena!

config['Paths'] = {
    'log_file_path': './log/app.log',
    'python_path': 'C:/Users/Home/AppData/Local/Programs/Python/Python313/python.exe',
    'python_script_path': 'C:/Users/Home/Documents/Coding/Windows/Win_ChangeBarTenderPrinter/ChangeBarTenderPrinter.py'
}

config['PrinterMapping'] = {
    '25x10_': 'OneNote (Desktop)',
    '50x30_': 'Microsoft Print to PDF'
}

# 📌 Write configuration to StringIO for testing
configfile = StringIO()
config.write(configfile)

# 📌 Output StringIO contents to verify functionality
print(configfile.getvalue())

with open('config.ini', mode='w') as file:
    file.write(configfile.getvalue())

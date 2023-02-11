import re, pandas as pd, subprocess, sys, os

column_names = {
    'Plugin': 'Plugin',
    'VP': 'Vulnerability Parameter',
    'CVE': 'CVE',
    'PN': 'Plugin Name',
    'Status': 'Status',
    'Date': 'Date',
    'NCF': 'New/Carried forward',
    'CD': 'Close Date',
    'HO': 'How old',
    'SBD': 'SLA Breached / Day',
    'Severity': 'Severity',
    'Entity': 'Entity',
    'Host': 'Host|(Ip Address)',
    'NBN': 'NetBIOS Name',
    'Description': 'Description',
    'Solution': 'Solution',
    'DD': '(Date discovered)|(First Discovered)',
    'CCD': 'Closing/Current Date',
    'SBDC': 'SLA Breached / Day Count'
}

# from https://stackoverflow.com/questions/31836104/pyinstaller-and-onefile-how-to-include-an-image-in-the-exe-file

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def sub_process(type='open', title='Select Master sheet', initial=''):
    return subprocess.run(['python', resource_path('file_picker.py'), '--title=' + title, '--type=' + type, '--initial=' + initial], capture_output=True, text=True).stdout.strip()

def convert_bytes(num):
    for x in ['bytes', 'KB', 'MB', 'GB', 'TB']:
        if num < 1024.0:
            return "%3.1f %s" % (num, x)
        num /= 1024.0

def _input_ (title):
        return input(title).strip()

def get_column(sheet, title):
    for col in sheet.iter_cols(1, sheet.max_column):
        value = col[0].value
        if value and re.search(title, value, re.IGNORECASE):
            return chr(64 + col[0].column)


def get_column_data(sheet, column):
    if not column:
        return tuple([])
    all = list(sheet[column])
    all.pop(0)
    return tuple(all)

def print_bound(text: str, lines: int=100):
        print('\n' + '-'*lines)
        print(text)
        print('-'*lines + '\n')

def _dir_(path: str):
    split = path.split('/')
    split.pop()
    return '/'.join(split)

def ext(path: str):
    split = path.split('.')
    return '.' + split[len(split) - 1]



def to_excel(path):
    splitted = path.split('/')
    name_with_ext = splitted[len(splitted) - 1]
    if name_with_ext.endswith('.xlsx'):
        return path
    name = name_with_ext.split('.')[0]
    csv = pd.read_csv(path)
    tmp_path = resource_path('__tmp__\\' + name + '.xlsx')
    writer = pd.ExcelWriter(tmp_path)
    csv.to_excel(writer, index=False)
    writer.save()
    return tmp_path

def del_tmp_files():
    folder = resource_path('__tmp__')
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if filename != 'dont-delete.txt' and (os.path.isfile(file_path) or os.path.islink(file_path)):
                os.unlink(file_path)
        except:
            pass
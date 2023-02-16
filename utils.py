import pandas as pd, sys, os, webbrowser, subprocess, win32com.client as win32
from playsound import playsound

column_names = {
    'Plugin': 'Plugin~&~&~Plugin ID',
    'VP': 'Vulnerability Parameter~&~&~internal/external/Scan type',
    'CVE': 'CVE',
    'PN': 'Plugin Name~&~&~Name',
    'Status': 'Status',
    'Date': 'Date',
    'NCF': 'New/Carried forward',
    'CD': 'Close Date',
    'HO': 'How old',
    'SBD': 'SLA Breached / Day',
    'Severity': 'Severity',
    'Entity': 'Entity',
    'Host': 'Host~&~&~Ip Address',
    'NBN': 'NetBIOS Name',
    'Description': 'Description',
    'Solution': 'Solution',
    'DD': 'Date discovered~&~&~First Discovered~&~&~First found',
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

def sub_process(type='open'):
    return subprocess.run([resource_path('assets\\' + type + '.bat')], capture_output=True, text=True).stdout.replace('"', '').replace('\\', '/').strip()

def convert_bytes(num):
    for x in ['bytes', 'KB', 'MB', 'GB', 'TB']:
        if num < 1024.0:
            return "%3.1f %s" % (num, x)
        num /= 1024.0

def _input_ (title = ''):
        value = input(title).strip()
        test_value = value.lower()
        if test_value == '--x':
            sys.exit()
        if test_value == '--r':
            webbrowser.open('https://github.com/julius-ek-hub/excel-automation')
            return _input_(title)
        return value


def print_bound(text: str, lines: int=100, type='info'):
        cprint('\n' + '-'*lines, type, False)
        cprint(text, type, False)
        cprint('-'*lines + '\n', type, False)

def _dir_(path: str):
    split = path.split('/')
    split.pop()
    return '/'.join(split)


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
    writer.close()
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

def cprint(value: str = '', type: str = 'info', lable=True):
    print({
        "error": "\033[91m {}\033[00m",
        "info": "\033[96m {}\033[00m",
        "success": "\033[92m {}\033[00m",
        "warn": "\033[93m {}\033[00m"
    }[type].format(('[' + type.upper() + ']: ' if lable else '') + value))

def beep(play_sound):
    try:
        if play_sound:
            playsound(resource_path('assets\\beep.mp3'))
    except:
        pass

def send_error(body, subject):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'julius.ekane@beaconred.ae'
    mail.Subject = subject
    mail.HtmlBody = body

    mail.Display(True)

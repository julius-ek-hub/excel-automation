import os
from validator import *
from utils import _dir_, sub_process, convert_bytes, _input_, resource_path, cprint


class Collector:

    def __init__(self) -> None:
        self.scans = []

    def ask(self, question, answer_is_ok):
        q = _input_(question).lower()
        if answer_is_ok(q):
            return q
        cprint('Sorry, I do not understand your choice!', lable=False, type='warn')
        return self.ask(question, answer_is_ok)

    def get_path_to_open(self, name: str, sufix=""):

        path = _input_('\nFull path to ' + name + ' or hit Enter with no input to open file dialog' + sufix + ': ')
        path = path.replace('"', '').replace('\'', '').replace('\\', '/').strip()

        if not path:
            path = sub_process('open')
        if not sheet_path_open_is_ok(path):
            cprint(path + ' --> Invalid file (' + name + '), try another.', 'error')
            return self.get_path_to_open(name, sufix)
        else:
            size = os.stat(path).st_size
            if size/1024/1024 > 2:
                cprint('File size larger than 2MB, Reduce file size by checking and removing excessive blanc rows' + ((' and/or convert your file to .xlsx manually (might reduce the size)' if not path.endswith('.xlsx') else '') + ' otherwise the scanning might take a while.'), 'warn')
            
            cprint(path + ' --> OK (' + convert_bytes(size) + ')', 'success')
            return path

    def get_path_to_save(self, default: str = '', sufix: str ='', ms_path: str = ''):

        new_ms_name = _input_(sufix + open(resource_path('help\\save.guide.txt'), 'r').read())
        new_ms_name = new_ms_name.replace('"', '').replace('\'', '').replace('\\', '/').strip()

        if new_ms_name.lower() == '--rp':
            new_ms_name = ms_path

        if new_ms_name and len(new_ms_name.split(':/')) == 1:
            new_ms_name = _dir_(ms_path) + '/' + new_ms_name

        if not new_ms_name:
            new_ms_name = sub_process('save')
        
        if new_ms_name and len(new_ms_name.split(':/')) == 0:
            new_ms_name = _dir_(ms_path) + '/' + new_ms_name

        if not sheet_path_save_is_ok(new_ms_name):
            return self.get_path_to_save(default, 'Failed to save!\n', ms_path)

        new_ms_name = new_ms_name.replace('.xlsx', '') + '.xlsx'
        confirm = self.ask('\nSave to ' + new_ms_name + '? n = No, y = Yes (default = y): ', lambda ans: ans in ['yes', 'y', 'no', 'n', ''])
        if confirm in ['n', 'no']:
            return self.get_path_to_save(default=default, ms_path=ms_path)

        return new_ms_name


    def get_text(self, name: str, default: str, validator, help=''):

        _help = ''
        if help:
            _help = ', Type --h for help.'

        value = _input_('\n' + name + ' (default = ' + str(default) + _help + '): ')

        if help and value.lower() == '--h':
            print(open(resource_path(help), 'r').read())
            return self.get_text(name, default, validator, help)

        if not value:
            value = default
        if not validator(value):
            cprint(str(value) + ' --> Invalid! ' + name + ', try again.', 'error')
            return self.get_text(name, default, validator)
        else:
            cprint(str(value) + ' --> OK', 'success')
            return value

    def get_text_from_options(self, options, label):

        print('\n' + label + ', type only the letter that corresponds to your choice. (default = a): ')
        for _key in options:
            print(_key + ' = ' + options[_key]) 

        key = _input_().lower()

        if not key:
            key = 'a'
        value = options.get(key)

        if not value:
            cprint(str(value) + ' --> Invalid! ' + key + ' does not match any option.', 'error')
            return self.get_text_from_options(options, label)
        else:
            cprint(value + ' --> OK', 'success')
            return value
        
    def collect_scans(self):

        new_scan_index = str(len(self.scans) + 1)

        cprint('\nScan ' + new_scan_index + '.\n------------------', lable=False)

        ss_path = self.get_path_to_open('Scan sheet.', sufix=" (Mistakenly pressed enter from previous scan? Just enter a random scan and delete it in the Grand Confirmation)")
        ss_target_sheet = self.get_text('Worksheet target', default=None, validator=lambda v: True, help='help\\target.sheet.txt')
        
        scan_date = self.get_text('Scan date in DD/MM/YY', default=datetime.datetime.today().strftime('%d/%m/%Y'), validator=scan_date_is_ok)
        entity = self.get_text_from_options({
            "a": "EDGE",
            "b": "ADSB",
            "c": "ADASI",
            "d": "BEACON RED",
            "e": "LAHAB",
            "f": "NIMR",
            "g": "D14",
            "h": "GAL",
            "i": "HALCON",
            "j": "EARTH",
            "k": "AL HOSN",
            "l": "AL JASOOR",
            "m": "AL TARIQ",
            "n": "APT",
            "o": "EPI",
            "p": "CARACAL",
            "q": "KNOWLEDGE POINT",
            "r": "EDIC",
            "s": "AMMROC",
            "t": "HORIZON",
            "u": "JAHEZIYA",
            "v": "SIM",
            "w": "EDGE-DT"
        }, 'Entity')
        vulnerability_param = self.get_text_from_options({
            "a": "Internal",
            "b": "External",
        }, 'Vulnerability parameter')

        cprint('\nConfirm scan ' + new_scan_index + '!\n------------------', lable=False)
        cprint('Scan sheet: ' + ss_path + '\nScan date: ' + scan_date + '\nEntity: ' + entity + '\nVulnerability parameter: ' + vulnerability_param + '\nTarget worksheet: ' + str(ss_target_sheet if ss_target_sheet else 'Active') + '\n', 'success', False)

        confirm = self.ask('Correct? n = No, y = Yes and collect next scan, --done = Yes and move to next stage (default = y): ', lambda ans: ans in ['n', 'no', 'y', 'yes', '--done', ''])

        if confirm in ['n', 'no']:
            return self.collect_scans()
        
        self.scans.append({
            "path": ss_path,
            "target": ss_target_sheet,
            "date": scan_date,
            "entity": entity,
            "vp": vulnerability_param
        })

        if confirm == '--done':
            return self.scans

        return self.collect_scans()
import os
from validator import *
from utils import _dir_, sub_process, convert_bytes, _input_


class Collector:

    def get_path_to_open(self, name: str, var):

        path = _input_('\nFull path to ' + name + ' (or press enter without typing, to open file dialog): ')
        path = ''.join(path.split('"')).strip()
        if path.lower() == '--x': return exit()

        if not path:
            path = sub_process('open', 'Select ' + name)
        if not sheet_path_open_is_ok(path):
            print(path, ' --> Invalid file (' + name + '), try another.')
            return self.get_path_to_open(name, var)
        else:
            size = os.stat(path).st_size
            print(path, '--> OK (' + convert_bytes(size) + ')')
            return path

    def get_path_to_save(self, default: str = '', sufix: str =''):

        new_ms_name = _input_(sufix + 'Type new filename and/or press Enter to save updated master sheet: ')
        new_ms_name = ''.join(new_ms_name.split('"')).strip()
        if new_ms_name.lower() == '--x': return exit()

        if not new_ms_name:
            new_ms_name = default
        else:
            new_ms_name = _dir_(default) + '/' + new_ms_name

        path = sub_process('save', 'Save updated master sheet as', new_ms_name)

        if not sheet_path_save_is_ok(path):
            return self.get_path_to_save(default, 'Failed to save!\n')
        return path


    def get_text(self, name: str, default: str, validator):
        value = _input_('\n' + name + ' (or press enter without typing, to use ' + default + '): ')
        if value.lower() == '--x': return exit()

        if not value:
            value = default
        if not validator(value):
            print(value, ' --> Invalid! ' + name + ', try again.')
            return self.get_text(name, default, validator)
        else:
            print(value, '--> OK')
            return value

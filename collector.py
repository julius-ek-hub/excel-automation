import os
from validator import *
from utils import _dir_, sub_process, convert_bytes, _input_, resource_path, cprint


class Collector:

    def get_path_to_open(self, name: str):

        path = _input_('\nFull path to ' + name + ' (or hit Enter with no input to open file dialog): ')

        if not path:
            path = sub_process('open')
        if not sheet_path_open_is_ok(path):
            cprint(path + ' --> Invalid file (' + name + '), try another.', 'error')
            return self.get_path_to_open(name)
        else:
            size = os.stat(path).st_size
            cprint(path + ' --> OK (' + convert_bytes(size) + ')', 'success')
            return path

    def get_path_to_save(self, default: str = '', sufix: str ='', ms_path: str = ''):

        new_ms_name = _input_(sufix + open(resource_path('save.guide.txt'), 'r').read())
        new_ms_name = new_ms_name.replace('"', '').replace('\\', '/').strip()

        if new_ms_name.lower() == '--rp':
            new_ms_name = ms_path

        if new_ms_name and len(new_ms_name.split(':/')) == 1:
            new_ms_name = _dir_(ms_path) + '/' + new_ms_name

        if not new_ms_name:
            new_ms_name = sub_process('save')

        if not sheet_path_save_is_ok(new_ms_name):
            return self.get_path_to_save(default, 'Failed to save!\n', ms_path)

        new_ms_name = new_ms_name.replace('.xlsx', '') + '.xlsx'
        confirm = _input_('Save to ' + new_ms_name + '? n = No, anything else = Yes: ')
        if confirm.lower() == 'n':
            return self.get_path_to_save(default=default, ms_path=ms_path)

        return new_ms_name


    def get_text(self, name: str, default: str, validator):
        value = _input_('\n' + name + ' (or hit Enter with no input to use ' + str(default) + '): ')

        if not value:
            value = default
        if not validator(value):
            cprint(str(value) + ' --> Invalid! ' + name + ', try again.', 'error')
            return self.get_text(name, default, validator)
        else:
            cprint(str(value) + ' --> OK', 'success')
            return value

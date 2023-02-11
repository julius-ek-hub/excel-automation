import os

def scan_date_is_ok(date: str):
    return len(date.split('/')) == 3

def entity_is_ok(value: str):
    return len(value) > 0

def vp_is_ok(value: str):
    return len(value) > 0

def sheet_path_open_is_ok(path: str):
    return os.path.isfile(path) and (path.endswith('.xlsx') or path.endswith('.csv')) and os.stat(path).st_size > 0

def sheet_path_save_is_ok(path: str):
    _break = path.split('/')
    if len(_break) <= 1:
        _break = path.split('\\')
    _break.pop()
    return os.path.isdir('/'.join(_break))
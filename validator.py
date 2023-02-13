import os, datetime

def scan_date_is_ok(date: str):
    try:
        return datetime.datetime.strptime(date, '%d/%m/%Y')
    except:
        return False

def entity_is_ok(value: str):
    return value.lower() in [
        'edge', 
        'beacon red',
        'katim',
        'sign4l'
    ]

def vp_is_ok(value: str):
    return value.lower() in ['internal', 'external']

def target_sheet_is_ok(value):
    return True

def sheet_path_open_is_ok(path: str):
    return os.path.isfile(path) and (path.endswith('.xlsx') or path.endswith('.csv')) and os.stat(path).st_size > 0

def sheet_path_save_is_ok(path: str):
    _break = path.split('/')
    _break.pop()
    return os.path.isdir('/'.join(_break))
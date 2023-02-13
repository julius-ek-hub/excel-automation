import os, webbrowser, traceback, sys, datetime

from validator import *
from scanner import Scanner
from utils import _input_, print_bound, resource_path, cprint
from collector import Collector

def confirm_restart():
    confirm = _input_('Would you like to restart the process? n = No, anything else = Yes: ')
    if confirm.lower() == 'n':
        sys.exit()
    else:
        runProgram()

def runProgram():

    os.system('cls')

    print_bound(open(resource_path('intro.txt'), 'r').read())

    try:
        col = Collector()

        ms_path = col.get_path_to_open('Master sheet')
        ss_path = col.get_path_to_open('Scan sheet')

        print_bound('Choose target sheets, Leave as None if the file has only one sheet or\n you want to use the active sheet (usually the first)')
        ms_target_sheet = col.get_text('Master sheet target', default=None, validator=target_sheet_is_ok)
        ss_target_sheet = col.get_text('Scan sheet target', default=None, validator=target_sheet_is_ok)
        
        scan_date = col.get_text('Scan date in DD/MM/YY', default=datetime.datetime.today().strftime('%d/%m/%Y'), validator=scan_date_is_ok)
        entity = col.get_text_from_options({
            "a": "EDGE",
            "b": "ADSB",
            "c": "KATIM",
            "d": "BEACON RED",
            "e": "SIGN4L"
        }, 'Entity')
        vulnerability_param = col.get_text_from_options({
            "a": "Internal",
            "b": "External",
        }, 'Vulnerability parameter')

        cprint('\nConfirm!\n------------------\n')
        cprint('Master sheet: ' + ms_path + '\nScan sheet: ' + ss_path + '\nScan date: ' + scan_date + '\nEntity: ' + entity + '\nVulnerability parameter: ' + vulnerability_param + '\n', 'success')
        confirm = _input_('Correct? n = No, anything else = Yes: ')
        if confirm.lower() == 'n':
            return runProgram()

        scanner = Scanner(ss_path, ms_path, scan_date, entity, vulnerability_param, ms_target_sheet, ss_target_sheet)
        scanner.scan()

        if scanner.total_update == 0 and scanner.total_new == 0:
            print('No updates or new vulnerabilities were added to the Mastersheet')
            return confirm_restart()

        scanner.save(col.get_path_to_save(default=ms_path, ms_path=ms_path))

        print_bound('ALL GOOD!!', 20, type='success')

        confirm_restart()
    except Exception as e:
        print_bound('\nAn error occured, please try again.\n'+ str(e) +'\nIf proplem persists, enter --e to send me an email with the error trace.', type='error')

        inp = _input_('Hit enter to start over or --e to send error: ')
        if inp.lower() == '--e':
            webbrowser.open('mailto:?to=julius.ekane@beaconred.ae&subject=Excel%20Atomation%20Error%20-%20' + str(e).replace(' ', '%20') + '&body=' + str(traceback.format_exc()).replace(' ', '%20'))
        else:
            runProgram()


runProgram()
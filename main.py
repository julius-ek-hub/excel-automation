import os, webbrowser, traceback, sys, openpyxl as ope, time

from validator import *
from scanner import Scanner
from utils import _input_, print_bound, resource_path, cprint, to_excel, del_tmp_files
from collector import Collector

def confirm_restart():
    confirm = _input_('Would you like to restart the process? n = No, anything else = Yes: ')
    if confirm.lower() in ['n', 'no']:
        sys.exit()
    else:
        runProgram()

def runProgram():

    os.system('cls')

    print_bound(open(resource_path('intro.txt'), 'r').read())

    try:
        col = Collector()

        ms_path = col.get_path_to_open('Master sheet')
        ms_target_sheet = col.get_text('Master sheet target', default=None, validator=target_sheet_is_ok)

        print_bound('Input scans (As many as you want)')
        scans = col.collect_scans()

        can_move_to_next_step = False

        while not can_move_to_next_step:

            cprint('\nFinal Confirmation!\n-----------------------')

            scans_str = ''

            for i, val in enumerate(scans):
                scans_str += '\n\nScan ' + str(i + 1) + '\nScan sheet ' + str(i + 1) + ' path: ' + val['path'] + '\nTarget sheet: ' + str(val['target']) + '\nDate: ' + val['date'] + '\nEntity: ' + val['entity'] + '\nVulnerability parameter: ' + val['vp']

            cprint('Master sheet path: ' + ms_path + '\nTarget sheet: ' + str(ms_target_sheet) + scans_str, 'success')

            confirm = _input_('Correct? n = No, --rm=n = Removes scan number n, eg --rm=1 removes scan 1, anything else = Yes: ').lower()
            if confirm.lower() in ['n', 'no']:
                return runProgram()
            
            check_rm = confirm.split('--rm=')

            if len(check_rm) == 2:
                try:
                    rm_index = int(check_rm[1])
                    scans.pop(rm_index - 1)
                    cprint('Scan ' + str(rm_index) + ' removed, scans re-ordered, ', 'success')
                except Exception as e:
                    cprint('Failed to delete scan, ' + str(e), 'error')
            else:
                can_move_to_next_step = True
            
        
        cprint('\nOn it......\n')
        cprint('Loading Mastersheet.....')

        ms_workbook = ope.load_workbook(to_excel(path=ms_path))

        cprint('Done!', 'success')

                
        time_start = time.time()
        total_update = 0
        total_new = 0

        for i, s in enumerate(scans):
            scanner = Scanner(s['path'], ms_workbook, s['date'], s['entity'], s['vp'], ms_target_sheet, s['target'], i)
            scanner.scan()
            total_update += scanner.total_update
            total_new += scanner.total_new

        time_stop = time.time()
        time_diff = time_stop - time_start

        if time_diff > 60:
            time_diff = str('%2.f' % (time_diff/60.0)) + ' minute(s)'
        else:
            time_diff = str('%2.f' % time_diff) + ' seconds(s)'

        print_bound('SCANNING AND UPDATE COMPLETE!\n\nTotal cells updated =  ' + str(total_update) + '\nNew vulnerabilities added = ' + str(total_new) + '\nTime spent = ' + time_diff, 40, 'success')
        del_tmp_files()


        if total_update == 0 and total_new == 0:
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
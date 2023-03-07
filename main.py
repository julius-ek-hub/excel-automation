import os, traceback, sys, openpyxl as ope, time

from validator import *
from scanner import Scanner
from utils import _input_, print_bound, resource_path, cprint, to_excel, del_tmp_files, beep, send_error
from collector import Collector

def confirm_restart(play_sound):
    beep(play_sound)
    confirm = _input_('Would you like to restart the process? n = No, y = Yes, (default = y): ').lower()
    if not confirm in ['yes', 'y', '', 'n', 'no']:
        cprint('I do not understand your choice!', 'warning', False)
        return confirm_restart(play_sound)
    
    if confirm in ['n', 'no']:
        sys.exit()
    else:
        runProgram()

def runProgram():

    os.system('cls')

    print_bound(open(resource_path('help\intro.txt'), 'r').read())

    play_sound = False

    try:
        col = Collector()

        ms_path = col.get_path_to_open('Master file')
        ms_target_sheet = col.get_text('Worksheet', default=None, validator=lambda v: True, help='help\\target.sheet.txt')

        print_bound('Input scans')
        scans = col.collect_scans()

        ips_path = None
        ips_target_sheet = None
        has_external = False

        if any(scan['vp'] == 'External' for scan in scans):
            has_external = True
            ips_path = col.get_path_to_open('All IPs/Entities file', sufix='. (Enter --d to use old/default)', important=False)
            if ips_path:
                ips_target_sheet = col.get_text('All IPs/Entities worksheet', default=None, validator=lambda v: True, help='help\\target.sheet.txt')


        can_move_to_next_step = False

        while not can_move_to_next_step:

            cprint('\nGrand Confirmation!\n-----------------------', lable=False)

            scans_str = '\n'

            for i, val in enumerate(scans):
                scans_str += '\nScan ' + str(i + 1) + '\nPath: ' + val['path'] + '\nWorksheet: ' + str(val['target'] if val['target'] else 'Active') + '\nDate: ' + val['date'] + '\nEntity: ' + val['entity'] + '\nVulnerability parameter: ' + val['vp'] + '\n'

            cprint('Master file\nPath: ' + ms_path + '\nWorksheet: ' + str(ms_target_sheet if ms_target_sheet else 'Active'), 'success', lable=False)

            if has_external:
                cprint('\nAll IPs/Entities\nPath: ' + (ips_path if ips_path else './assets/ips.xlxs') + '\nWorksheet: ' + str(ips_target_sheet if ips_target_sheet else 'Active'), 'success', lable=False)

            cprint(scans_str, lable=False, type='success')

            confirm = col.ask('Correct? n = No, y = Yes, --rm=n = Remove scan n, eg --rm=1 removes scan 1 (default = y): ', lambda v: (v in ['yes', 'no', 'y', 'n', ''] or (len(v.split('--rm=')) == 2 and not v.split('--rm=')[0] and v.split('--rm=')[1].isnumeric)))
            
            if confirm.lower() in ['n', 'no']:
                return runProgram()
            
            check_rm = confirm.split('--rm=')

            if len(check_rm) == 2:
                try:
                    rm_index = int(check_rm[1])
                    if len(scans) == 1:
                        raise Exception('Only 1 scan remaining.')
                    scans.pop(rm_index - 1)
                    cprint('Scan ' + str(rm_index) + ' removed, scans re-ordered, ', 'success')
                except Exception as e:
                    cprint('Failed to delete, ' + str(e), 'error')
            else:
                can_move_to_next_step = True
            
        play_sound = col.ask('\nMake a beep sound if your attention is needed? y = Yes, n = No (default = y): ', lambda ans: ans in ['yes', 'y', '', 'n', 'no']).lower() in ['', 'y', 'yes']

        cprint('On it.... You can move on with other things' + (' cuz this might take a while' if len(scans) > 2 else '')  + '\n')
        cprint('Loading Mastersheet.....')

        ms_workbook = ope.load_workbook(to_excel(path=ms_path))
        ms_col_ids = None

        cprint('Done!', 'success')

        time_start = time.time()
        total_updates = {"New": 0, "Carried Forward": 0, "Closed": 0}

        for i, s in enumerate(scans):
            scanner = Scanner(s['path'], ms_workbook, s['date'], s['entity'], s['vp'], ms_target_sheet, s['target'], i)
            scanner.ms_col_ids = ms_col_ids if ms_col_ids else {}
            scanner.ips_path = ips_path
            scanner.ips_target_sheet = ips_target_sheet
            scanner.play_sound = play_sound
            scanner.total_updates = total_updates
            scanner.scan()
            ms_col_ids = scanner.ms_col_ids
            total_updates = scanner.total_updates

        time_stop = time.time()
        time_diff = time_stop - time_start

        if time_diff > 60:
            time_diff = str(round(time_diff/60, 1)) + ' minute(s)'
        else:
            time_diff = str(round(time_diff, 1)) + ' seconds(s)'

        total_updates_str = ''

        for tu in total_updates:
            total_updates_str += (tu + ': ' + str(total_updates[tu]) + '\n')
        total_updates_str += ('Time spent = ' + time_diff)
        
        print_bound('SCANNING AND UPDATE COMPLETE!\n\n' + total_updates_str, 40, 'success')
        del_tmp_files()


        if not any(value > 0 for value in total_updates.values()):
            print('No updates or new vulnerabilities were added to the Mastersheet')
            return confirm_restart(play_sound)
        
        beep(play_sound)

        scanner.save(col.get_path_to_save(default=ms_path, ms_path=ms_path))

        print_bound('ALL GOOD!!', 20, type='success')

        confirm_restart(play_sound)

    except Exception as e:
        beep(play_sound)
        del_tmp_files()
        print_bound('\nAn error occured, please try again.\n'+ str(e) +'\nIf proplem persists, enter --e to send me an email with the error trace.', type='error')
        inp = _input_('Hit enter to start over or --e to send error: ')
        if inp.lower() == '--e':
            send_error(body=traceback.format_exc(), subject='Excel Atomation Error - ' + str(e))
        else:
            runProgram()


runProgram()
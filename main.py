from datetime import datetime
import os, webbrowser, traceback

from validator import *
from scanner import Scanner
from utils import _input_, print_bound
from collector import Collector


def main():
    os.system('cls')

    print_bound("""
EXCEL AUTOMATION v1.0.0 by Julius Ekane.\n\n
Try to make sure the necessary column names exist in both files for the best result.
Make sure files to use are not open while program is running.
CSV files will be converted to Excel
Enter --r anytime to open the source code repo with documentaion on github
Enter --x anytime to exit
Reach out if any issue -> julius.ekane@beaconred.ae.
""")

    try:

        col = Collector()

        ms_path = col.get_path_to_open('Master sheet', 'mss')
        ss_path = col.get_path_to_open('Scan sheet', 'sss')
        scan_date = col.get_text('Scan date in DD/MM/YY', datetime.today().strftime('%d/%m/%Y'), scan_date_is_ok)
        entity = col.get_text('Entity', 'BEACON RED', entity_is_ok)
        vulnerability_param = col.get_text('Vulnerability parameter', 'Internal', vp_is_ok)

        confirm = _input_('\nPlease Confirm!\n------------------\nMaster sheet: ' + ms_path + '\nScan sheet: ' + ss_path + '\nScan date: ' + scan_date + '\nEntity: ' + entity + '\nVulnerability parameter: ' + vulnerability_param + '\nCorrect? n = No, anything else = yes: ')
        if confirm.lower() == 'n':
            return main()

        scanner = Scanner(ss_path, ms_path, scan_date, entity, vulnerability_param)
        scanner.scan()
        scanner.save(col.get_path_to_save(ms_path))

        print_bound('ALL GOOD!!', 20)

        confirm_restart = _input_('Would you like to restart the process? y = yes, anything else = exit: ')
        if confirm_restart.lower() == 'y':
            main()
        else:
            exit()
    except Exception as e:
        print_bound("""
An error occured!!
{error}
Please make sure selected files are not currently open. Else close all and start again.
If proplem persists, enter --e to send me an email with the error trace.
""".format(error=str(e)))

        inp = _input_('Press enter to start over or --e to send error: ')
        if inp.lower() == '--e':
            webbrowser.open('mailto:?to=julius.ekane@beaconred.ae&subject=Excel%20Atomation ' + str(e).replace(' ', '%20') + '&body=' + str(traceback.format_exc()).replace(' ', '%20'))
        else:
            main()


main()
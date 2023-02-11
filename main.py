from datetime import datetime
import os

from validator import *
from scanner import Scanner
from utils import _input_, print_bound
from collector import Collector


def main():
    os.system('cls')

    print_bound("""
EXCEL AUTOMATION v1.0.0 by Julius Ekane.\n\n
Try to make sure the necessary column names exist in both files for the best result.
Enter --r anytime to open the source code repo on github
Enter --x anytime to exit
Reach out if any issue -> julius.ekane@beaconred.ae.
""")

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


main()
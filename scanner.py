import openpyxl as ex, re, time, os
from utils import to_excel, print_bound, del_tmp_files, column_names, _input_

class Scanner:
    def __init__(self, ss_path, ms_path, scan_date, entity, vulnerability_param):
        self.ss_path = ss_path
        self.ms_path = ms_path
        self.scan_date =scan_date 
        self.entity = entity 
        self.vulnerability_param = vulnerability_param
        self.ms_default_cols = column_names
        self.ss_default_cols = column_names
        self.total_update = 0
        self.total_new = 0

    @staticmethod
    def get_column(sheet, search_value):
        for col in sheet.iter_cols(1, sheet.max_column):
            value = col[0].value
            if value and re.search(search_value, value, re.IGNORECASE):
                return chr(64 + col[0].column)

    @staticmethod
    def get_column_data(sheet, column):
        if not column:
            return tuple([])
        all = list(sheet[column])
        all.pop(0)
        return tuple(all)
    
    def get_column_by_all_means(self, label: str, key: str, sheet: str, default_cols, sheet_name: str = 'Master sheet', important: bool = True):
        col = self.get_column(sheet, default_cols[key])
        
        if not col and important:
            from_user = _input_(label + ' column doesn\'t exist for value ' + default_cols[key] + ', check the ' + sheet_name + ' and enter the value for ' + label + ' column title: ')
            if from_user:
                default_cols[key] = from_user
            return self.get_column_by_all_means(label, key, sheet, default_cols, sheet_name)
        return col

    def get_columns(self):

        print('Identifying columns....')

        self.ms_host_column = self.get_column_by_all_means(sheet=self.ms, key='Host', default_cols=self.ms_default_cols, label=column_names['Host'])
        self.ms_plugin_column = self.get_column_by_all_means(sheet=self.ms, key='Plugin', default_cols=self.ms_default_cols, label=column_names['Plugin'])
        self.ms_date_column = self.get_column_by_all_means(sheet=self.ms, key='Date', default_cols=self.ms_default_cols, label=column_names['Date'])
        self.ms_status_column = self.get_column_by_all_means(sheet=self.ms, key='Status', default_cols=self.ms_default_cols, label=column_names['Status'])
        self.ms_ncf_column = self.get_column_by_all_means(sheet=self.ms, key='NCF', default_cols=self.ms_default_cols, label=column_names['NCF'])

        self.ss_host_column = self.get_column_by_all_means(sheet=self.ss, key='Host', default_cols=self.ss_default_cols, label=column_names['Host'], sheet_name='Scan sheet')
        self.ss_plugin_column = self.get_column_by_all_means(sheet=self.ss, key='Plugin', default_cols=self.ss_default_cols, label=column_names['Plugin'], sheet_name='Scan sheet')


    def check_master_sheet_with_scan_sheet(self):

        for msRow in self.get_column_data(sheet=self.ms, column=self.ms_plugin_column):

            ms_host_value_num = str(msRow.row)
            ms_plugin_value = msRow.value
            ms_host_value = self.ms[self.ms_host_column + ms_host_value_num].value

            host_and_plugin_matched = False

            for ssRow in self.get_column_data(sheet=self.ss, column=self.ss_plugin_column):

                ss_row_num = str(ssRow.row)
                ss_plugin_value = ssRow.value
                ss_host_value = self.ss[self.ss_host_column + ss_row_num].value

                if (ms_plugin_value == ss_plugin_value and ms_host_value == ss_host_value):
                    (host_and_plugin_matched) = True
                    break

            ms_status_cell = self.ms[self.ms_status_column + ms_host_value_num]
            ms_ncf_cell = self.ms[self.ms_ncf_column + ms_host_value_num]
            ms_sd_cell = self.ms[self.ms_date_column + ms_host_value_num]

            cf = 'Carried Forward'
            patched = 'Patched'

            if (host_and_plugin_matched):
                if not re.search(cf, str(ms_status_cell.value).strip(), re.IGNORECASE):
                    ms_status_cell.value = cf
                    self.total_update = self.total_update + 1
                    print('Status (' + self.ms_status_column + ms_host_value_num + ') updated to ' + cf)

                if not re.search(cf, str(ms_ncf_cell.value).strip(), re.IGNORECASE):
                    ms_ncf_cell.value = cf
                    self.total_update = self.total_update + 1
                    print('New/' + cf + ' (' + self.ms_ncf_column + ms_host_value_num + ') updated to ' + cf)

            else:
                if self.scan_date != str(ms_sd_cell.value).strip():
                    ms_sd_cell.value = self.scan_date
                    self.total_update = self.total_update + 1
                    print('Date (' + self.ms_date_column + ms_host_value_num + ') updated to ' + self.scan_date)

                if not re.search(patched, str(ms_status_cell.value).strip(), re.IGNORECASE):
                    ms_status_cell.value = patched
                    self.total_update = self.total_update + 1
                    print('Status (' + self.ms_status_column + ms_host_value_num + ') updated to Patched')


    def check_scan_sheet_with_master_sheet(self):
        
        for ssRow in self.get_column_data(sheet=self.ss, column=self.ms_plugin_column):

            ss_row_num = str(ssRow.row)
            ss_plugin_value = ssRow.value
            ss_host_value = self.ss[self.ss_host_column + ss_row_num].value

            host_and_plugin_exists = False

            for msRow in self.get_column_data(sheet=self.ms, column=self.ms_plugin_column):

                ms_host_value_num = str(msRow.row)
                ms_plugin_value = msRow.value
                ms_host_value = self.ms[self.ms_host_column + ms_host_value_num].value

                if (ms_plugin_value == ss_plugin_value and ms_host_value == ss_host_value):
                    host_and_plugin_exists = True
                    break

            if host_and_plugin_exists:
                continue
            else:
                print('New vulnerability found!')
                print('[Updating]: Adding new vulnerability to mastersheet')
                ms_last_empty_row = str(len(self.ms['A']) + 1)

                ms_vp_column = self.get_column_by_all_means(sheet=self.ms, key='VP', label=column_names['VP'], default_cols=self.ms_default_cols)
                ms_entity_column = self.get_column_by_all_means(sheet=self.ms, key='Entity', label=column_names['Entity'], default_cols=self.ms_default_cols)

                self.ms[ms_vp_column + ms_last_empty_row].value = self.vulnerability_param
                self.ms[self.ms_status_column + ms_last_empty_row].value = 'pending'
                self.ms[self.ms_date_column + ms_last_empty_row].value = self.scan_date
                self.ms[ms_entity_column + ms_last_empty_row].value = self.entity
                self.ms[self.ms_ncf_column + ms_last_empty_row].value = 'New'

                new = ['Plugin', 'Host', 'PN', 'Severity',
                       'NBN', 'Description', 'Solution', 'DD', 'CVE']

                for n in new:
                    ss_column = self.get_column_by_all_means(sheet=self.ss, key=n, label=column_names[n], sheet_name='Scan sheet', default_cols=self.ss_default_cols, important=False)
                    ms_column = self.get_column_by_all_means(sheet=self.ms, key=n, label=column_names[n], default_cols=self.ms_default_cols)
                    if (not ss_column):
                        continue
                    self.ms[ms_column + ms_last_empty_row].value = self.ss[ss_column + ss_row_num].value
            print('[Update]: Added new vulnerability to mastersheet (row ' + ms_last_empty_row + ')')
            self.total_new = self.total_new + 1

    def scan(self):

        print_bound('SCANNING & UPDATING............')

        time_start = time.time()

        ms_path = to_excel(path=self.ms_path)
        ss_path = to_excel(path=self.ss_path)

        print('Loading Mastersheet.....')
        self.workbook_ms = ex.load_workbook(ms_path)
        print('Done!')

        print('Loading Scansheet.....')
        workbook_ss = ex.load_workbook(ss_path)
        print('Done!')

        self.ss = workbook_ss.active
        self.ms = self.workbook_ms.active

        # Identify columns
        self.get_columns()

        print('Done!')

        # Check mastersheet with scansheet
        print('[Scanning & updating]: Mastersheet with scansheet.....')
        self.check_master_sheet_with_scan_sheet()

        # Check scansheet with mastersheet
        print('[Scanning for new vulnerabilities]: Scansheet with mastersheet.....')
        self.check_scan_sheet_with_master_sheet()

        time_stop = time.time()
        time_diff = time_stop - time_start

        if time_diff > 60:
            time_diff = str('%2.f' % (time_diff/60.0)) + ' minute(s)'
        else:
            time_diff = str('%2.f' % time_diff) + ' seconds(s)'

        print_bound('SCANNING AND UPDATE COMPLETE! (Cells updated =  ' + str(self.total_update) + ', New vulnerabilities = ' + str(self.total_new) + ', Time spent = ' + time_diff + ')', 110)
        del_tmp_files()

    def save(self, path):
        _path = path + '.xlsx'
        print('Saving to ', _path, ' ....')
        self.workbook_ms.save(_path)
        try:
            os.unlink(path)
        except:
            pass
        print('Done!')
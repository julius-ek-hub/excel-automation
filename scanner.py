import openpyxl as ex, time, os
from utils import to_excel, column_names, _input_, cprint, reason, beep

class Scanner:
    def __init__(self, ss_path, workbook_ms, scan_date, entity, vulnerability_param, ms_target_sheet, ss_target_sheet, scan_index):
        self.ss_path = ss_path
        self.scan_index = str(scan_index + 1)
        self.workbook_ms = workbook_ms
        self.scan_date =scan_date 
        self.ms_target_sheet = ms_target_sheet
        self.ss_target_sheet = ss_target_sheet
        self.entity = entity 
        self.vulnerability_param = vulnerability_param
        self.ms_default_cols = column_names.copy()
        self.ss_default_cols = column_names.copy()
        self.total_update = 0
        self.total_new = 0
        self.ms_plugin_host_severity_pairs = []
        self.play_sound = False

    @staticmethod
    def get_column(sheet, search_value):
        for col in sheet.iter_cols(1, sheet.max_column):
            value = str(col[0].value).strip().lower()
            if value and any(name.lower() == value for name in search_value.split('~&~&~')):
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
            if self.play_sound:
                beep()
            from_user = _input_(sheet_name + ' ' + label + ' column doesn\'t exist for value ' + ' or '.join(default_cols[key].split('~&~&~')) + ', check ' + sheet_name + ' and enter the value for ' + label + ' column title: ')
            if from_user:
                default_cols[key] = from_user
            return self.get_column_by_all_means(label, key, sheet, default_cols, sheet_name)
        return col

    def get_columns(self):

        cprint('Identifying columns....')

        self.ms_host_column = self.get_column_by_all_means(sheet=self.ms, key='Host', default_cols=self.ms_default_cols, label='Host')
        self.ms_plugin_column = self.get_column_by_all_means(sheet=self.ms, key='Plugin', default_cols=self.ms_default_cols, label='Plugin')
        self.ms_date_column = self.get_column_by_all_means(sheet=self.ms, key='Date', default_cols=self.ms_default_cols, label='Date')
        self.ms_status_column = self.get_column_by_all_means(sheet=self.ms, key='Status', default_cols=self.ms_default_cols, label='Status')
        self.ms_ncf_column = self.get_column_by_all_means(sheet=self.ms, key='NCF', default_cols=self.ms_default_cols, label='New/Carried forward')
        self.ms_vp_column = self.get_column_by_all_means(sheet=self.ms, key='VP', default_cols=self.ms_default_cols, label='Vulnerability parameter')
        self.ms_entity_column = self.get_column_by_all_means(sheet=self.ms, key='Entity', default_cols=self.ms_default_cols, label='Entity')
        self.ms_severity_column = self.get_column_by_all_means(sheet=self.ms, key='Severity', default_cols=self.ms_default_cols, label='Severity')
        self.ms_cd_column = self.get_column_by_all_means(sheet=self.ms, key='CD', default_cols=self.ms_default_cols, label='Close date')

        self.ss_host_column = self.get_column_by_all_means(sheet=self.ss, key='Host', default_cols=self.ss_default_cols, label='Host', sheet_name='Scan sheet ' + self.scan_index)
        self.ss_plugin_column = self.get_column_by_all_means(sheet=self.ss, key='Plugin', default_cols=self.ss_default_cols, label='Plugin', sheet_name='Scan sheet ' + self.scan_index)
        self.ss_severity_column = self.get_column_by_all_means(sheet=self.ss, key='Severity', default_cols=self.ss_default_cols, label='Severity', sheet_name='Scan sheet ' + self.scan_index)


    def check_mastersheet_with_scansheet(self):

        for ms_row in self.get_column_data(sheet=self.ms, column=self.ms_plugin_column):

            ms_row_str = str(ms_row.row)

            ms_host_cell_address = self.ms_host_column + ms_row_str
            ms_cd_cell_address = self.ms_cd_column + ms_row_str
            # ms_plugin_address = self.ms_plugin_column + ms_row_str

            ms_cd_cell = self.ms[ms_cd_cell_address]

            ms_plugin_value = ms_row.value
            ms_host_value = self.ms[ms_host_cell_address].value
            ms_severity_value = self.ms[self.ms_severity_column + ms_row_str].value

            # Check if vulnerability has been closed (Has a Close date value)
            closed = str(ms_cd_cell.value).strip()

            target_vp = self.vulnerability_param == str(self.ms[self.ms_vp_column + ms_row_str].value).strip()
            target_entity = self.entity == str(self.ms[self.ms_entity_column + ms_row_str].value).strip()

            # If vulnerabilty parameter and entity are not what user provided, then skip.
            if not (target_vp and target_entity) or closed:
                continue

            # host_and_plugin_matched = False

            for ss_row in self.get_column_data(sheet=self.ss, column=self.ss_plugin_column):

                ss_row_str = str(ss_row.row)

                ss_host_address = self.ss_host_column + ss_row_str
                ss_plugin_address = self.ss_plugin_column + ss_row_str

                ss_plugin_value = ss_row.value
                ss_host_value = self.ss[ss_host_address].value
                ss_severity_value = self.ss[self.ss_severity_column + ss_row_str].value

                same_plugin = ms_plugin_value == ss_plugin_value
                same_host = ms_host_value == ss_host_value
                same_severity = ms_severity_value == ss_severity_value

                cf = 'Carried Forward'
                pcd = 'Patched'

                if (same_host and same_plugin and same_severity):

                    ms_status_cell = self.ms[self.ms_status_column + ms_row_str]
                    ms_ncf_address = self.ms_ncf_column + ms_row_str
                    ms_ncf_cell = self.ms[ms_ncf_address]

                    carried_forward = str(ms_ncf_cell.value).strip().lower() == cf.lower()

                    if not carried_forward:
                        ms_ncf_cell.value = cf
                        self.total_update = self.total_update + 1
                        print(
                            'MS New/' + cf + ' (' + ms_ncf_address + 
                            ') updated to \'' + cf + '\'' + 
                            reason(' matches with\n[SS: Host (' + 
                            ss_host_address + ') = ' + str(ss_host_value) + 
                            ', Plugin (' + ss_plugin_address + 
                            ') = ' + str(ss_plugin_value) + ']'
                            ))

                else:
                    ms_cd_cell.value = self.scan_date
                    self.total_update = self.total_update + 1
                    print(
                        'MS Date (' + ms_cd_cell_address  + 
                        ') updated to \'' + self.scan_date + '\'' +
                          reason()
                    )

                    if pcd != str(ms_status_cell.value).strip().lower():
                        ms_status_cell.value = pcd
                        self.total_update = self.total_update + 1
                        print(
                            'MS Status (' + self.ms_status_column + ms_row_str  + 
                            ') updated to \'Patched\'' + reason()
                        )


    def check_scansheet_with_mastersheet(self):
        
        for ss_row in self.get_column_data(sheet=self.ss, column=self.ms_plugin_column):

            ss_row_str = str(ss_row.row)
            ss_plugin_value = ss_row.value
            ss_plugin_address = self.ms_plugin_column  + ss_row_str
            ss_host_address = self.ss_host_column + ss_row_str
            ss_host_value = self.ss[ss_host_address].value
            ss_severity_value = self.ss[self.ss_severity_column + ss_row_str].value

            vulnerabily_exists = False

            for ms_row in self.get_column_data(sheet=self.ms, column=self.ms_plugin_column):

                ms_row_str = str(ms_row.row)
                ms_plugin_value = ms_row.value
                ms_host_value = self.ms[self.ms_host_column + ms_row_str].value

                target_vp = self.vulnerability_param == str(self.ms[self.ms_vp_column + ms_row_str].value).strip()
                target_entity = self.entity == str(self.ms[self.ms_entity_column + ms_row_str].value).strip()
                closed = str(self.ms[self.ms_cd_column + ms_row_str].value).strip()
                same_plugin = ms_plugin_value == ss_plugin_value 
                same_host = ms_host_value == ss_host_value
                same_severity = self.ms[self.ms_severity_column + ms_row_str].value == ss_severity_value

               # If vulnerabilty parameter and entity are not what user provided, then skip.
                if not (target_vp and target_entity) or closed:
                   continue

                if (same_plugin and same_host & same_severity):
                    vulnerabily_exists = True
                    break

            if vulnerabily_exists:
                continue
            else:
                print(
                    'New vulnerability detected because... \nSS [Host (' + 
                    ss_host_address + ') = ' + str(ss_host_value) + ', Plugin (' + 
                    ss_plugin_address + ') = ' + str(ss_plugin_value) + '] did not match any in MS'
                )
                cprint('[Updating]: Adding new vulnerability to mastersheet')
                ms_last_empty_row = str(len(self.ms['A']) + 1)

                ms_entity_column = self.get_column_by_all_means(sheet=self.ms, key='Entity', label=column_names['Entity'], default_cols=self.ms_default_cols)

                self.ms[self.ms_vp_column + ms_last_empty_row].value = self.vulnerability_param
                self.ms[self.ms_status_column + ms_last_empty_row].value = 'pending'
                self.ms[self.ms_date_column + ms_last_empty_row].value = self.scan_date
                self.ms[ms_entity_column + ms_last_empty_row].value = self.entity
                self.ms[self.ms_ncf_column + ms_last_empty_row].value = 'New'

                new = ['Plugin', 'Host', 'PN', 'Severity',
                       'NBN', 'Description', 'Solution', 'DD', 'CVE']

                for n in new:
                    ss_column = self.get_column_by_all_means(sheet=self.ss, key=n, label=column_names[n], sheet_name='Scan sheet ' + self.scan_index, default_cols=self.ss_default_cols, important=False)
                    ms_column = self.get_column_by_all_means(sheet=self.ms, key=n, label=column_names[n], default_cols=self.ms_default_cols)
                    if (not ss_column):
                        continue
                    self.ms[ms_column + ms_last_empty_row].value = self.ss[ss_column + ss_row_str].value
            cprint('[Update]: Added new vulnerability to mastersheet (row ' + ms_last_empty_row + ')', 'success')
            self.total_new = self.total_new + 1

    def scan(self):

        ss_path = to_excel(path=self.ss_path)

        cprint('Loading Scansheet ' + self.scan_index + '.....')
        workbook_ss = ex.load_workbook(ss_path)
        cprint('Done!', 'success')

        self.ss = workbook_ss.active
        self.ms = self.workbook_ms.active

        if self.ms_target_sheet:
            self.ms = self.workbook_ms[self.ms_target_sheet]
        if self.ss_target_sheet:
            self.ss = self.workbook_ss[self.ss_target_sheet]

        # Identify columns
        self.get_columns()

        cprint('Done!', 'success')

        # # Check mastersheet with scansheet
        cprint('[Scanning & updating]: Mastersheet with scansheet ' + self.scan_index + '.....')
        self.check_mastersheet_with_scansheet()

        # Check scansheet with mastersheet for new vulnerabilities
        cprint('[Scanning for new vulnerabilities]: Scansheet with mastersheet ' + self.scan_index + '.....')
        self.check_scansheet_with_mastersheet()


    def save(self, path: str):

        cprint('Saving to ' + path + ' ....')

        if(os.path.exists(path)):
            os.unlink(path)
        self.workbook_ms.save(path)

        # Sometimes, an additional invalid file with no extension is created. So...
        try:
            os.unlink(path.replace('.xlsx', ''))
        except:
            pass
        cprint('Done!', 'success')
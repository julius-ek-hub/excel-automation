import openpyxl as ex, os
from utils import to_excel, column_names, _input_, cprint, beep

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
        self.non_mandatory_columns = {"ss": {}, "ms": {}}
        self.nm_column_keys = ['Plugin', 'Host', 'PN', 'Severity','NBN', 'Description', 'Solution', 'DD', 'CVE']
        self.total_updates = {"New": 0, "Newly Carried Forward": 0, "Closed": 0}
        self.ms_existing_vulnerability_rows = []
        self.play_sound = False

    @staticmethod
    def trim(value):
        return str(value).strip()
    
    @staticmethod
    def get_column_data(sheet, column):
        if not column:
            return tuple([])
        all = list(sheet[column])
        all.pop(0)
        return tuple(all)

    def get_column(self, sheet, search_value):
        for col in sheet.iter_cols(1, sheet.max_column):
            value = self.trim(col[0].value).lower()
            if value and any(name.lower() == value for name in search_value.split('~&~&~')):
                return chr(64 + col[0].column)


    def get_column_by_all_means(self, label: str, key: str, sheet: str, default_cols, sheet_name: str = 'Master sheet', important: bool = True):
        col = self.get_column(sheet, default_cols[key])
        
        if not col and important:
            beep(self.play_sound)
            from_user = _input_(sheet_name + ' ' + label + ' column doesn\'t exist for value ' + ' or '.join(default_cols[key].split('~&~&~')) + ', check ' + sheet_name + ' and enter the value for ' + label + ' column title: ')
            if from_user:
                default_cols[key] = from_user
            return self.get_column_by_all_means(label, key, sheet, default_cols, sheet_name)
        return col

    def get_columns(self):

        cprint('Identifying columns....')

        # Mandatory columns

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

        # Non mandatory columns

        for key in self.nm_column_keys:
            self.non_mandatory_columns["ss"][key] = self.get_column_by_all_means(sheet=self.ss, key=key, label=column_names[key], sheet_name='Scan sheet ' + self.scan_index, default_cols=self.ss_default_cols, important=False)
            self.non_mandatory_columns["ms"][key] = self.get_column_by_all_means(sheet=self.ms, key=key, label=column_names[key], default_cols=self.ms_default_cols, important=False)


    def check_mastersheet_with_scansheet(self):

        cf = 'Carried Forward'
        pcd = 'Patched'

        for ms_row in self.get_column_data(sheet=self.ms, column=self.ms_plugin_column):

            ms_row_str = str(ms_row.row)

            ms_cd_cell = self.ms[self.ms_cd_column + ms_row_str]
            ms_status_cell = self.ms[self.ms_status_column + ms_row_str]

            ms_plugin_value = self.trim(ms_row.value).lower()
            ms_host_value = self.trim(self.ms[self.ms_host_column + ms_row_str].value).lower()
            ms_severity_value = self.trim(self.ms[self.ms_severity_column + ms_row_str].value).lower()

            closed = ms_cd_cell.value

            target_vp = self.vulnerability_param == self.trim(self.ms[self.ms_vp_column + ms_row_str].value)
            target_entity = self.entity == self.trim(self.ms[self.ms_entity_column + ms_row_str].value)

            self.ms_existing_vulnerability_rows.append(ms_row_str)

            if not (target_entity and target_vp) or closed:
                continue

            vulnerability_match = False

            for ss_row in self.get_column_data(sheet=self.ss, column=self.ss_plugin_column):

                if vulnerability_match: continue

                ss_row_str = str(ss_row.row)

                ss_plugin_value = self.trim(ss_row.value).lower()
                ss_host_value = self.trim(self.ss[self.ss_host_column + ss_row_str].value).lower()
                ss_severity_value = self.trim(self.ss[self.ss_severity_column + ss_row_str].value).lower()

                same_plugin = ms_plugin_value == ss_plugin_value
                same_host = ms_host_value == ss_host_value
                same_severity = ms_severity_value == ss_severity_value

                vulnerability_match = same_host and same_plugin and same_severity

                if (vulnerability_match):

                    ms_ncf_cell = self.ms[self.ms_ncf_column + ms_row_str]

                    carried_forward = self.trim(ms_ncf_cell.value).lower() == cf.lower()

                    if not carried_forward:
                        ms_ncf_cell.value = cf
                        self.total_updates["Newly Carried Forward"] += 1
                    break
                    

            if not vulnerability_match:
                ms_cd_cell.value = self.scan_date
                self.total_updates["Closed"] += 1

                if pcd != self.trim(ms_status_cell.value).lower():
                    ms_status_cell.value = pcd


    def check_scansheet_with_mastersheet(self):
        
        for ss_row in self.get_column_data(sheet=self.ss, column=self.ss_plugin_column):

            ss_row_str = str(ss_row.row)
            ss_plugin_value = self.trim(ss_row.value).lower()
            ss_host_value = self.trim(self.ss[self.ss_host_column + ss_row_str].value).lower()
            ss_severity_value = self.trim(self.ss[self.ss_severity_column + ss_row_str].value).lower()

            vulnerabily_exists = False

            for ms_row in self.ms_existing_vulnerability_rows:
                ms_plugin_value = self.trim(self.ms[self.ms_plugin_column + ms_row].value).lower()
                ms_host_value = self.trim(self.ms[self.ms_host_column + ms_row].value).lower()

                same_plugin = ms_plugin_value == ss_plugin_value 
                same_host = ms_host_value == ss_host_value
                same_severity = self.trim(self.ms[self.ms_severity_column + ms_row].value).lower() == ss_severity_value

                if same_plugin and same_host and same_severity:
                    vulnerabily_exists = True
                    break

            if not vulnerabily_exists:

                ms_last_empty_row = str(len(self.ms['A']) + 1)

                self.ms[self.ms_vp_column + ms_last_empty_row].value = self.vulnerability_param
                self.ms[self.ms_status_column + ms_last_empty_row].value = 'pending'
                self.ms[self.ms_date_column + ms_last_empty_row].value = self.scan_date
                self.ms[self.ms_entity_column + ms_last_empty_row].value = self.entity
                self.ms[self.ms_ncf_column + ms_last_empty_row].value = 'New'

                for n in self.nm_column_keys:
                    ss_column = self.non_mandatory_columns['ss'][n]
                    ms_column = self.non_mandatory_columns['ms'][n]
                    if not (ss_column and ms_column):
                        continue
                    self.ms[self.non_mandatory_columns['ms'][n] + ms_last_empty_row].value = self.ss[self.non_mandatory_columns['ss'][n] + ss_row_str].value

                self.total_updates["New"] += 1

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
        cprint('Scanning & updating Mastersheet with scansheet ' + self.scan_index + '.....')
        self.check_mastersheet_with_scansheet()
        cprint('Done!', 'success')

        # Check scansheet with mastersheet for new vulnerabilities
        cprint('Scanning for new vulnerabilities with scansheet ' + self.scan_index + '.....')
        self.check_scansheet_with_mastersheet()
        cprint('Done!', 'success')

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
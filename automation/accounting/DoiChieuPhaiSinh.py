from os.path import dirname, join
from tqdm import tqdm
from automation.accounting import *


# Client code
def run(
        run_time=dt.datetime(2022,1,25)
):
    report = Report(run_time)
    for func in tqdm([report.runRDT0121], ncols=70):
        func()


class Report:

    def __init__(self, run_time: dt.datetime):
        info = get_info('daily', run_time)
        period = info['period']
        t0_date = info['end_date'].replace('.', '-')
        folder_name = info['folder_name']

        # create folder
        if not os.path.isdir(join(dept_folder, folder_name, period)):
            os.mkdir((join(dept_folder, folder_name, period)))

        self.bravoFolder = join(dirname(dept_folder), 'FileFromBravo')
        self.bravoDateString = run_time.strftime('%Y.%m.%d')
        self.file_date = dt.datetime.strptime(t0_date, '%Y-%m-%d').strftime('%d.%m.%Y')
        self.file_name = f'Đối Chiếu Phái Sinh {self.file_date}.xlsx'
        self.writer = pd.ExcelWriter(
            join(dept_folder, folder_name, period, self.file_name),
            engine='xlsxwriter',
            engine_kwargs={'options': {'nan_inf_to_errors': True}}
        )
        self.workbook = self.writer.book
        self.info_format = self.workbook.add_format(
            {
                'bold': True,
                'align': 'left',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Arial',
            }
        )
        self.FDS_title_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 12,
                'font_name': 'Arial',
                'bg_color': '#FFC000'
            }
        )
        self.bravo_title_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 12,
                'font_name': 'Arial',
                'bg_color': '#00B0F0'
            }
        )
        self.diff_title_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 12,
                'font_name': 'Arial',
                'bg_color': '#FFFF00'
            }
        )
        self.headers_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Arial',
                'text_wrap': True
            }
        )
        self.stt_format = self.workbook.add_format(
            {
                'border': 1,
                'align': 'righ',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Arial',
            }
        )
        self.text_left_format = self.workbook.add_format(
            {
                'border': 1,
                'align': 'left',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Arial',
            }
        )
        self.money_fds_format = self.workbook.add_format(
            {
                'border': 1,
                'align': 'right',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Arial',
                'num_format': '#,##0_);(#,##0)',
            }
        )
        self.money_bravo_diff_format = self.workbook.add_format(
            {
                'border': 1,
                'align': 'right',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Arial',
                'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            }
        )
        self.sum_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'left',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Arial',
            }
        )
        self.money_sum_fds_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'right',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Arial',
                'num_format': '#,##0_);(#,##0)',
            }
        )
        self.money_sum_bravo_diff_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'right',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Arial',
                'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            }
        )

    def __del__(self):
        self.writer.close()

    def runRDT0121(self):
        TaiKhoan_338804 = pd.read_excel(
            join(self.bravoFolder, f'{self.bravoDateString}', f'Sổ tổng hợp công nợ 338804_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            names=['SoTaiKhoan', 'TenKhachHang', 'DuNoDau338804', 'DuCoDau338804',
                   'PhatSinhNo338804', 'PhatSinhCo338804', 'DuNoCuoi338804', 'DuCoCuoi338804']
        )
        TaiKhoan_3243 = pd.read_excel(
            join(self.bravoFolder, f'{self.bravoDateString}', f'Sổ tổng hợp công nợ 3243_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            names=['SoTaiKhoan', 'TenKhachHang', 'DuNoDau3243', 'DuCoDau3243',
                   'PhatSinhNo3243', 'PhatSinhCo3243', 'DuNoCuoi3243', 'DuCoCuoi3243']
        )

        RDT0121 = pd.read_sql(
            f"""
                    SELECT
                        [relationship].[branch_id] [MaChiNhanh],
                        [r].[account_code] [SoTaiKhoan],
                        [account].[customer_name] [TenKhachHangFlex],
                        [r].[cash_balance_at_phs] [TienTaiPHS],
                        [r].[cash_balance_at_vsd] [TienTaiVSD]
                    FROM [rdt0121] [r]
                    LEFT JOIN [relationship] ON [relationship].[account_code] = [r].[account_code] 
                    AND [relationship].[date] = [r].[date]
                    LEFT JOIN [account] ON [account].[account_code] = [r].[account_code]
                    WHERE [r].[date] = '{self.bravoDateString}'
                    """,
            connect_DWH_PhaiSinh
        )

        table = RDT0121.merge(
            TaiKhoan_3243[['SoTaiKhoan', 'DuCoCuoi3243']], how='outer', on='SoTaiKhoan'
        ).merge(
            TaiKhoan_338804[['SoTaiKhoan', 'DuCoCuoi338804']], how = 'outer', on = 'SoTaiKhoan'
        )
        table = table.fillna(0)

        table['TienTaiPHSDiff'] = table['TienTaiPHS'] - table['DuCoCuoi3243']
        table['TienQuyVSDDiff'] = table['TienTaiVSD'] - table['DuCoCuoi338804']

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('RDT0121')
        worksheet.set_column('A:A', 4)
        worksheet.set_column('B:B', 13)
        worksheet.set_column('C:C', 16)
        worksheet.set_column('D:D', 29)
        worksheet.set_column('E:E', 19)
        worksheet.set_column('F:F', 22)
        worksheet.set_column('G:G', 17)
        worksheet.set_column('H:H', 21)
        worksheet.set_column('I:I', 17)
        worksheet.set_column('J:J', 20)

        worksheet.write('A1', 'BÁO CÁO SỐ DƯ TIỀN NHÀ ĐẦU TƯ', self.info_format)
        worksheet.write('A2', f'Từ ngày {self.file_date} đến ngày {self.file_date}', self.info_format)
        worksheet.merge_range('A3:F3', 'FDS', self.FDS_title_format)
        worksheet.merge_range('G3:H3', 'Bravo', self.bravo_title_format)
        worksheet.merge_range('I3:J3', 'Chênh lệch', self.diff_title_format)

        worksheet.write_row(
            'A4',
            [
                'STT',
                'Mã chi nhánh',
                'Tài khoản ký quỹ',
                'Tên khách hàng',
                'Số tiền tại công ty',
                'Số tiền ký quỹ tại VSD',
                'Số tiền tại công ty',
                'Số tiền ký quỹ tại VSD',
                'Số tiền tại công ty',
                'Số tiền ký quỹ tại VSD'
            ],
            self.headers_format
        )
        worksheet.write_column('A5', np.arange(table.shape[0]) + 1, self.stt_format)
        worksheet.write_column('B5', table['MaChiNhanh'], self.text_left_format)
        worksheet.write_column('C5', table['SoTaiKhoan'], self.text_left_format)
        worksheet.write_column('D5', table['TenKhachHangFlex'], self.text_left_format)
        worksheet.write_column('E5', table['TienTaiPHS'], self.money_fds_format)
        worksheet.write_column('F5', table['TienTaiVSD'], self.money_fds_format)
        worksheet.write_column('G5', table['DuCoCuoi3243'], self.money_bravo_diff_format)
        worksheet.write_column('H5', table['DuCoCuoi338804'], self.money_bravo_diff_format)
        worksheet.write_column('I5', table['TienTaiPHSDiff'], self.money_bravo_diff_format)
        worksheet.write_column('J5', table['TienQuyVSDDiff'], self.money_bravo_diff_format)

        sum_start_row = table.shape[0] + 5
        worksheet.merge_range(f'A{sum_start_row}:B{sum_start_row}', 'Tổng cộng:', self.sum_format)
        for col in 'EFGHIJ':
            if col == 'EF':
                worksheet.write(f'{col}{sum_start_row}',f'=SUM({col}5:{col}{sum_start_row - 1})',self.money_sum_fds_format)
            else:
                worksheet.write(f'{col}{sum_start_row}',f'=SUM({col}5:{col}{sum_start_row - 1})',self.money_sum_bravo_diff_format)

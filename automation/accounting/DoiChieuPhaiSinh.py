from tqdm import tqdm
from automation.accounting import *


# Client code
def run(
        run_time=dt.datetime(2022, 1, 25)
):
    report = Report(run_time)
    for func in tqdm([report.runRDT0121, report.runRDT0141, report.runRDT0127, report.runRDO0002], ncols=70):
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
        # date in sub title in Excel
        self.sub_title_date = dt.datetime.strptime(t0_date, '%Y-%m-%d').strftime('%d/%m/%Y')

        self.file_name = f'Đối Chiếu Phái Sinh {self.file_date}.xlsx'
        self.writer = pd.ExcelWriter(
            join(dept_folder, folder_name, period, self.file_name),
            engine='xlsxwriter',
            engine_kwargs={'options': {'nan_inf_to_errors': True}}
        )
        self.workbook = self.writer.book
        self.company_name_format = self.workbook.add_format(
            {
                'bold': True,
                'align': 'left',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'text_wrap': True
            }
        )
        self.company_info_format = self.workbook.add_format(
            {
                'align': 'left',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'text_wrap': True
            }
        )
        self.empty_row_format = self.workbook.add_format(
            {
                'bottom': 1,
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
            }
        )
        self.sheet_title_format = self.workbook.add_format(
            {
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 14,
                'font_name': 'Times New Roman',
                'text_wrap': True
            }
        )
        self.sub_title_format = self.workbook.add_format(
            {
                'bold': True,
                'italic': True,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'text_wrap': True
            }
        )
        self.headers_root_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'text_wrap': True
            }
        )
        self.headers_bravo_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'text_wrap': True,
                'bg_color': '#DAEEF3',
            }
        )
        self.headers_fds_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'text_wrap': True,
                'bg_color': '#EBF1DE',
            }
        )
        self.headers_diff_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'text_wrap': True,
                'bg_color': '#FFFFCC',
            }
        )
        self.text_root_format = self.workbook.add_format(
            {
                'border': 1,
                'align': 'center',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
            }
        )
        self.money_bravo_format = self.workbook.add_format(
            {
                'border': 1,
                'align': 'right',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
                'bg_color': '#DAEEF3',
            }
        )
        self.money_fds_format = self.workbook.add_format(
            {
                'border': 1,
                'align': 'right',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'num_format': '#,##0_);(#,##0)',
                'bg_color': '#EBF1DE',
            }
        )
        self.money_diff_format = self.workbook.add_format(
            {
                'border': 1,
                'align': 'right',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
                'bg_color': '#FFFFCC',
            }
        )
        self.money_sum_fds_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'right',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'num_format': '#,##0_);(#,##0)',
            }
        )
        self.money_sum_bravo_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'right',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            }
        )
        self.money_sum_diff_format = self.workbook.add_format(
            {
                'border': 1,
                'bold': True,
                'align': 'right',
                'valign': 'vcenter',
                'font_size': 10,
                'font_name': 'Times New Roman',
                'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)',
            }
        )

    def __del__(self):
        self.writer.close()

    def runRDT0121(self):
        TaiKhoan_3243 = pd.read_excel(
            join(self.bravoFolder, f'{self.bravoDateString}', f'Sổ tổng hợp công nợ 3243_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            usecols=('1', '2', '8')
        ).rename(columns={'1': 'SoTaiKhoan', '2': 'TenKhachHang3243', '8': 'DuCoCuoi3243'})

        TaiKhoan_338804 = pd.read_excel(
            join(self.bravoFolder, f'{self.bravoDateString}',
                 f'Sổ tổng hợp công nợ 338804_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            usecols=('1', '2', '8')
        ).rename(columns={'1': 'SoTaiKhoan', '2': 'TenKhachHang338804', '8': 'DuCoCuoi338804'})

        RDT0121 = pd.read_sql(
            f"""
            SELECT
                [relationship].[branch_id] [MaChiNhanh],
                [r].[account_code] [SoTaiKhoan],
                [account].[customer_name] [TenKhachHangFlex],
                [r].[cash_balance_at_phs] [TienTaiPHS],
                [r].[cash_balance_at_vsd] [TienTaiVSD]
            FROM [rdt0121] [r]
            LEFT JOIN [relationship] 
            ON [relationship].[account_code] = [r].[account_code] AND [relationship].[date] = [r].[date]
            LEFT JOIN [account] ON [account].[account_code] = [r].[account_code]
            WHERE [r].[date] = '{self.bravoDateString}'
            """,
            connect_DWH_PhaiSinh
        )

        table = RDT0121.merge(
            TaiKhoan_3243, how='outer', on='SoTaiKhoan'
        ).merge(
            TaiKhoan_338804, how='outer', on='SoTaiKhoan'
        )
        table['TenKhachHang'] = table['TenKhachHang3243'].fillna(table['TenKhachHang338804']).fillna(table['TenKhachHangFlex']).fillna('')
        table['MaChiNhanh'] = table['MaChiNhanh'].fillna('')
        table = table.fillna(0)

        table['TienTaiPHSDiff'] = table['TienTaiPHS'] - table['DuCoCuoi3243']
        table['TienQuyVSDDiff'] = table['TienTaiVSD'] - table['DuCoCuoi338804']

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('RDT0121')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('E11')
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

        worksheet.merge_range('A1:J1', CompanyName, self.company_name_format)
        worksheet.merge_range('A2:J2', CompanyAddress, self.company_info_format)
        worksheet.merge_range('A3:J3', CompanyPhoneNumber, self.company_info_format)
        worksheet.write_row('A4', [''] * 10, self.empty_row_format)

        worksheet.merge_range('A6:J6', 'BÁO CÁO SỐ DƯ TIỀN NHÀ ĐẦU TƯ', self.sheet_title_format)
        worksheet.merge_range('A7:J7', f'Date: {self.sub_title_date}', self.sub_title_format)
        worksheet.merge_range('E9:F9', 'FDS', self.headers_fds_format)
        worksheet.merge_range('G9:H9', 'Bravo', self.headers_bravo_format)
        worksheet.merge_range('I9:J9', 'Chênh lệch', self.headers_diff_format)

        worksheet.write_row('A10', ['STT', 'Mã chi nhánh', 'Tài khoản ký quỹ', 'Tên khách hàng'], self.headers_root_format)
        worksheet.write_row('E10', ['Số tiền tại công ty', 'Số tiền ký quỹ tại VSD'], self.headers_fds_format)
        worksheet.write_row('G10', ['Số tiền tại công ty', 'Số tiền ký quỹ tại VSD'], self.headers_bravo_format)
        worksheet.write_row('I10', ['Số tiền tại công ty', 'Số tiền ký quỹ tại VSD'], self.headers_diff_format)

        worksheet.write_column('A11', np.arange(table.shape[0]) + 1, self.text_root_format)
        worksheet.write_column('B11', table['MaChiNhanh'], self.text_root_format)
        worksheet.write_column('C11', table['SoTaiKhoan'], self.text_root_format)
        worksheet.write_column('D11', table['TenKhachHang'], self.text_root_format)
        worksheet.write_column('E11', table['TienTaiPHS'], self.money_fds_format)
        worksheet.write_column('F11', table['TienTaiVSD'], self.money_fds_format)
        worksheet.write_column('G11', table['DuCoCuoi3243'], self.money_bravo_format)
        worksheet.write_column('H11', table['DuCoCuoi338804'], self.money_bravo_format)
        worksheet.write_column('I11', table['TienTaiPHSDiff'], self.money_diff_format)
        worksheet.write_column('J11', table['TienQuyVSDDiff'], self.money_diff_format)

        sum_start_row = table.shape[0] + 11
        worksheet.merge_range(f'A{sum_start_row}:B{sum_start_row}', 'Tổng cộng:', self.headers_root_format)
        worksheet.write_row(f'C{sum_start_row}', [''] * 2, self.text_root_format)
        worksheet.write(f'E{sum_start_row}', f'=SUM(E11:E{sum_start_row - 1})',self.money_sum_fds_format)
        worksheet.write(f'F{sum_start_row}', f'=SUM(F11:F{sum_start_row - 1})',self.money_sum_fds_format)
        worksheet.write(f'G{sum_start_row}', f'=SUM(G11:G{sum_start_row - 1})',self.money_sum_bravo_format)
        worksheet.write(f'H{sum_start_row}', f'=SUM(H11:H{sum_start_row - 1})',self.money_sum_bravo_format)
        worksheet.write(f'I{sum_start_row}', f'=SUM(I11:I{sum_start_row - 1})',self.money_sum_diff_format)
        worksheet.write(f'J{sum_start_row}', f'=SUM(J11:J{sum_start_row - 1})',self.money_sum_diff_format)

    def runRDT0141(self):
        TaiKhoan_13504 = pd.read_excel(
            join(self.bravoFolder, f'{self.bravoDateString}', f'Sổ tổng hợp công nợ 13504_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            usecols=('1', '2', '5', '6', '7')
        ).rename(columns={'1': 'SoTaiKhoan', '2': 'TenKhachHang13504','5': 'PhatSinhNo13504', '6': 'PhatSinhCo13504', '7': 'DuNoCuoi13504'})

        TaiKhoan_13505 = pd.read_excel(
            join(self.bravoFolder, f'{self.bravoDateString}', f'Sổ tổng hợp công nợ 13505_{self.bravoDateString}.xlsx'),
            skiprows=8,
            skipfooter=1,
            usecols=('1', '2', '5', '6', '7')
        ).rename(columns={'1': 'SoTaiKhoan', '2': 'TenKhachHang13505', '5': 'PhatSinhNo13505', '6': 'PhatSinhCo13505', '7': 'DuNoCuoi13505'})

        RDT0141 = pd.read_sql(
            f"""
            SELECT
                [account].[account_code] [SoTaiKhoan],
                [r].[sub_account] [SoTieuKhoan],
                [account].[customer_name] [TenKhachHangFlex],
                [relationship].[branch_id] [MaChiNhanh],
                [r].[deferred_payment_amount_opening] [KhoanChamTraDauKy],
                [r].[deferred_payment_fee_opening] [PhiChamTraDauKy],
                ([r].[deferred_payment_amount_opening]+[r].[deferred_payment_fee_opening]) [TongTienChamDauKy],
                [r].[deferred_payment_amount_increase] [KhoanChamTraPSTangTrongKy],
                [r].[deferred_payment_fee_increase] [PhiChamTraPSTangTrongKy],
                [r].[deferred_payment_amount_decrease] [KhoanChamTraPSGiamTrongKy],
                [r].[deferred_payment_fee_decrease] [PhiChamTraPSGiamTrongKy],
                [r].[deferred_payment_amount_closing] [KhoanChamTraCuoiKy],
                [r].[deferred_payment_fee_closing] [PhiChamTraCuoiKy],
                ([r].[deferred_payment_amount_closing] + [r].[deferred_payment_fee_closing]) [TongTienChamCuoiKy]
            FROM [rdt0141] [r]
            LEFT JOIN [relationship]
            ON [relationship].[sub_account] = [r].[sub_account] AND [relationship].[date] = [r].[date]
            LEFT JOIN [account] ON [account].[account_code] = [relationship].[account_code]
            WHERE [r].[date] = '{self.bravoDateString}'
            """,
            connect_DWH_PhaiSinh
        )

        table = RDT0141.merge(
            TaiKhoan_13504, how='outer', on='SoTaiKhoan'
        ).merge(
            TaiKhoan_13505, how='outer', on='SoTaiKhoan'
        )
        table['TenKhachHang'] = table['TenKhachHang13504'].fillna(table['TenKhachHang13505']).fillna(table['TenKhachHangFlex']).fillna('')
        table['SoTieuKhoan'] = table['SoTieuKhoan'].fillna('')
        table['MaChiNhanh'] = table['MaChiNhanh'].fillna('')
        table = table.fillna(0)

        table['KhoanChamTraCuoiKyDiff'] = table['KhoanChamTraCuoiKy'] - table['DuNoCuoi13504']
        table['PhiChamTraCuoiKyDiff'] = table['PhiChamTraCuoiKy'] - table['DuNoCuoi13505']
        table['KhoanChamTraPSTangTrongKyDiff'] = table['KhoanChamTraPSTangTrongKy'] - table['PhatSinhNo13504']
        table['PhiChamTraPSTangTrongKyDiff'] = table['PhiChamTraPSTangTrongKy'] - table['PhatSinhNo13505']
        table['KhoanChamTraPSGiamTrongKyDiff'] = table['KhoanChamTraPSGiamTrongKy'] - table['PhatSinhCo13504']
        table['PhiChamTraPSGiamTrongKyDiff'] = table['PhiChamTraPSGiamTrongKy'] - table['PhatSinhCo13505']

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('RDT0141')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('E13')
        worksheet.set_column('A:B', 10)
        worksheet.set_column('C:C', 26)
        worksheet.set_column('D:Z', 13)

        worksheet.merge_range('A1:Z1', CompanyName, self.company_name_format)
        worksheet.merge_range('A2:Z2', CompanyAddress, self.company_info_format)
        worksheet.merge_range('A3:Z3', CompanyPhoneNumber, self.company_info_format)
        worksheet.write_row('A4', [''] * 26, self.empty_row_format)
        worksheet.merge_range('A6:Z6', 'BÁO CÁO SỐ DƯ KHOẢN CHẬM TRẢ', self.sheet_title_format)
        worksheet.merge_range('A7:Z7', f'Date: {self.sub_title_date}', self.sub_title_format)
        worksheet.merge_range('E9:N9', 'FDS', self.headers_fds_format)
        worksheet.merge_range('O9:T9', 'Bravo', self.headers_bravo_format)
        worksheet.merge_range('U9:Z9', 'Chênh lệch', self.headers_diff_format)
        worksheet.merge_range('E10:G11', 'Đầu kỳ', self.headers_fds_format)
        worksheet.merge_range('H10:K10', 'Trong kỳ', self.headers_fds_format)
        worksheet.merge_range('L10:N11', 'Cuối kỳ', self.headers_fds_format)
        worksheet.merge_range('O10:P10', 'Cuối kỳ', self.headers_bravo_format)
        worksheet.merge_range('Q10:T10', 'Trong kỳ', self.headers_bravo_format)
        worksheet.merge_range('U10:V10', 'Cuối kỳ', self.headers_diff_format)
        worksheet.merge_range('W10:Z10', 'Trong kỳ', self.headers_diff_format)
        worksheet.merge_range('H11:I11', 'Phát sinh tăng', self.headers_fds_format)
        worksheet.merge_range('J11:K11', 'Phát sinh giảm', self.headers_fds_format)
        worksheet.merge_range('Q11:R11', 'PS tăng', self.headers_bravo_format)
        worksheet.merge_range('S11:T11', 'PS giảm', self.headers_bravo_format)
        worksheet.merge_range('W11:X11', 'PS tăng', self.headers_diff_format)
        worksheet.merge_range('Y11:Z11', 'PS giảm', self.headers_diff_format)
        worksheet.merge_range('A10:A12', 'Tài khoản ký quỹ', self.headers_root_format)
        worksheet.merge_range('B10:B12', 'Tài khoản giao dịch', self.headers_root_format)
        worksheet.merge_range('C10:C12', 'Tên khách hàng', self.headers_root_format)
        worksheet.merge_range('D10:D12', 'Chi nhánh', self.headers_root_format)

        worksheet.write_row('O11', ['']*2, self.headers_bravo_format)
        worksheet.write_row('U11', ['']*2, self.headers_diff_format)
        worksheet.write_row(
            'E12',
            [
                'Khoản chậm trả',
                'Phí chậm trả',
                'Tổng số tiền chậm trả',
                'Khoản chậm trả',
                'Phí chậm trả',
                'Khoản chậm trả',
                'Phí chậm trả',
                'Khoản chậm trả',
                'Phí chậm trả',
                'Tổng số tiền chậm trả',
            ],
            self.headers_fds_format
        )
        worksheet.write_row(
            'O12',
            [
                'Khoản chậm trả',
                'Phí Khoản chậm trả',
                'Khoản chậm trả',
                'Phí chậm trả',
                'Khoản chậm trả',
                'Phí chậm trả'

            ] * 2,
            self.headers_bravo_format
        )
        worksheet.write_row(
            'U12',
            [
                'Khoản chậm trả',
                'Phí Khoản chậm trả',
                'Khoản chậm trả',
                'Phí chậm trả',
                'Khoản chậm trả',
                'Phí chậm trả'
            ],
            self.headers_diff_format
        )

        worksheet.write_column('A13', table['SoTaiKhoan'], self.text_root_format)
        worksheet.write_column('B13', table['SoTieuKhoan'], self.text_root_format)
        worksheet.write_column('C13', table['TenKhachHang'], self.text_root_format)
        worksheet.write_column('D13', table['MaChiNhanh'], self.text_root_format)
        worksheet.write_column('E13', table['KhoanChamTraDauKy'], self.money_fds_format)
        worksheet.write_column('F13', table['PhiChamTraDauKy'], self.money_fds_format)
        worksheet.write_column('G13', table['TongTienChamDauKy'], self.money_fds_format)
        worksheet.write_column('H13', table['KhoanChamTraPSTangTrongKy'], self.money_fds_format)
        worksheet.write_column('I13', table['PhiChamTraPSTangTrongKy'], self.money_fds_format)
        worksheet.write_column('J13', table['KhoanChamTraPSGiamTrongKy'], self.money_fds_format)
        worksheet.write_column('K13', table['PhiChamTraPSGiamTrongKy'], self.money_fds_format)
        worksheet.write_column('L13', table['KhoanChamTraCuoiKy'], self.money_fds_format)
        worksheet.write_column('M13', table['PhiChamTraCuoiKy'], self.money_fds_format)
        worksheet.write_column('N13', table['TongTienChamCuoiKy'], self.money_fds_format)
        worksheet.write_column('O13', table['DuNoCuoi13504'], self.money_bravo_format)
        worksheet.write_column('P13', table['DuNoCuoi13505'], self.money_bravo_format)
        worksheet.write_column('Q13', table['PhatSinhNo13504'], self.money_bravo_format)
        worksheet.write_column('R13', table['PhatSinhNo13505'], self.money_bravo_format)
        worksheet.write_column('S13', table['PhatSinhCo13504'], self.money_bravo_format)
        worksheet.write_column('T13', table['PhatSinhCo13505'], self.money_bravo_format)
        worksheet.write_column('U13', table['KhoanChamTraCuoiKyDiff'], self.money_diff_format)
        worksheet.write_column('V13', table['PhiChamTraCuoiKyDiff'], self.money_diff_format)
        worksheet.write_column('W13', table['KhoanChamTraPSTangTrongKyDiff'], self.money_diff_format)
        worksheet.write_column('X13', table['PhiChamTraPSTangTrongKyDiff'], self.money_diff_format)
        worksheet.write_column('Y13', table['KhoanChamTraPSGiamTrongKyDiff'], self.money_diff_format)
        worksheet.write_column('Z13', table['PhiChamTraPSGiamTrongKyDiff'], self.money_diff_format)

        sum_start_row = table.shape[0] + 13
        worksheet.merge_range(f'A{sum_start_row}:B{sum_start_row}', 'Tổng cộng:', self.headers_root_format)
        worksheet.write_row(f'B{sum_start_row}', [''] * 3, self.headers_root_format)

        for col in 'EFGHIJKLMNOPQRSTUVWXYZ':
            if col in 'EFGHIJKLMN':
                worksheet.write(f'{col}{sum_start_row}', f'=SUM({col}13:{col}{sum_start_row - 1})',
                                self.money_sum_fds_format)
            elif col in 'OPQRST':
                worksheet.write(f'{col}{sum_start_row}', f'=SUM({col}13:{col}{sum_start_row - 1})',
                                self.money_sum_bravo_format)
            else:
                worksheet.write(f'{col}{sum_start_row}', f'=SUM({col}13:{col}{sum_start_row - 1})',
                                self.money_sum_diff_format)

    def runRDT0127(self):
        TaiKhoan_33353 = pd.read_excel(
            join(self.bravoFolder, f'{self.bravoDateString}', f'BẢNG KÊ CTU TK 33353 THANG {self.bravoDateString[5:7]}.{self.bravoDateString[0:4]}.xlsx'),
            skiprows=7,
            skipfooter=1,
            usecols=('Tiền', 'Mã đối tượng\n(chi tiết)')
        ).rename(columns={'Tiền': 'ThueTNCN_Bravo', 'Mã đối tượng\n(chi tiết)': 'SoTaiKhoan'})
        # Trong file mẫu phần PIVOT từ FDS ko thấy lấy tài khoản GO0065 (CỤC THUẾ TPHCM)
        TaiKhoan_33353 = TaiKhoan_33353.loc[TaiKhoan_33353['SoTaiKhoan'] != 'GO0065']

        RDT0127 = pd.read_sql(
            f"""
            SELECT
                [r].[SoTaiKhoan],
                SUM([r].[ThueTNCN]) [ThueTNCN_FDS]
            FROM [RDT0127] [r]
            WHERE [r].[Ngay] = '{self.bravoDateString}'
            GROUP BY [r].[SoTaiKhoan]
            ORDER BY [r].[SoTaiKhoan]
            """,
            connect_DWH_PhaiSinh
        )

        table = RDT0127.merge(TaiKhoan_33353, how='outer', on='SoTaiKhoan')
        table = table.fillna(0)

        table['ThueTNCN_Diff'] = table['ThueTNCN_FDS'] - table['ThueTNCN_Bravo']

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('RDT0127')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('B11')
        worksheet.set_column('A:A', 10)
        worksheet.set_column('B:D', 16)
        worksheet.set_row(1, 24)

        worksheet.merge_range('A1:D1', CompanyName, self.company_name_format)
        worksheet.merge_range('A2:D2', CompanyAddress, self.company_info_format)
        worksheet.merge_range('A3:D3', CompanyPhoneNumber, self.company_info_format)
        worksheet.write_row('A4', [''] * 4, self.empty_row_format)
        worksheet.merge_range('A6:D6', 'BÁO CÁO ĐỐI CHIẾU THUẾ TNCN', self.sheet_title_format)
        worksheet.merge_range('A7:D7', f'Date: {self.sub_title_date}', self.sub_title_format)
        worksheet.write('B9', 'FDS', self.headers_fds_format)
        worksheet.write('C9', 'Bravo', self.headers_bravo_format)
        worksheet.write('D9', 'Chênh lệch', self.headers_diff_format)

        worksheet.write('A10', 'Số tài khoản', self.headers_root_format)
        worksheet.write('B10', 'Thuế TNCN', self.headers_fds_format)
        worksheet.write('C10', 'Thuế TNCN', self.headers_bravo_format)
        worksheet.write('D10', '', self.headers_diff_format)

        worksheet.write_column('A11', table['SoTaiKhoan'], self.text_root_format)
        worksheet.write_column('B11', table['ThueTNCN_FDS'], self.money_fds_format)
        worksheet.write_column('C11', table['ThueTNCN_Bravo'], self.money_bravo_format)
        worksheet.write_column('D11', table['ThueTNCN_Diff'], self.money_diff_format)

        sum_start_row = table.shape[0] + 11
        worksheet.write(f'A{sum_start_row}', 'Grand Total', self.headers_root_format)
        worksheet.write(f'B{sum_start_row}', f'=SUM(B11:B{sum_start_row - 1})', self.money_sum_fds_format)
        worksheet.write(f'C{sum_start_row}', f'=SUM(C11:C{sum_start_row - 1})', self.money_sum_bravo_format)
        worksheet.write(f'D{sum_start_row}', f'=SUM(D11:D{sum_start_row - 1})', self.money_sum_diff_format)

    def runRDO0002(self):
        TaiKhoan_5115104 = pd.read_excel(
            join(self.bravoFolder, f'{self.bravoDateString}',
                 f'BẢNG KÊ CTU TK 5115104 THANG {self.bravoDateString[5:7]}.{self.bravoDateString[0:4]}.xlsx'),
            skiprows=8,
            skipfooter=1,
            usecols=('Tiền', 'Mã đối tượng\n(chi tiết)')
        ).rename(columns={'Tiền': 'PhiGD_Bravo', 'Mã đối tượng\n(chi tiết)': 'SoTaiKhoan'})

        RDO0002 = pd.read_sql(
            f"""
            SELECT
                [relationship].[account_code] [SoTaiKhoan],
                SUM([r].[fee]) [PhiGD_FDS]
            FROM [rdo0002] [r]
            LEFT JOIN [relationship]
            ON [relationship].[sub_account] = [r].[sub_account] AND [relationship].[date] = [r].[date]
            WHERE [r].[date] = '{self.bravoDateString}'
            GROUP BY [relationship].[account_code]
            ORDER BY [SoTaiKhoan]
            """,
            connect_DWH_PhaiSinh
        )

        table = RDO0002.merge(TaiKhoan_5115104, how='outer', on='SoTaiKhoan')
        table = table.fillna(0)

        table['PhiGD_Diff'] = table['PhiGD_FDS'] - table['PhiGD_Bravo']

        ###################################################
        ###################################################
        ###################################################

        worksheet = self.workbook.add_worksheet('RDO0002')
        worksheet.hide_gridlines(option=2)
        worksheet.freeze_panes('B11')
        worksheet.set_column('A:A', 10)
        worksheet.set_column('B:D', 16)
        worksheet.set_row(1, 24)

        worksheet.merge_range('A1:D1', CompanyName, self.company_name_format)
        worksheet.merge_range('A2:D2', CompanyAddress, self.company_info_format)
        worksheet.merge_range('A3:D3', CompanyPhoneNumber, self.company_info_format)
        worksheet.write_row('A4', [''] * 4, self.empty_row_format)
        worksheet.merge_range('A6:D6', 'SAO KÊ LỆNH KHỚP', self.sheet_title_format)
        worksheet.merge_range('A7:D7', f'Date: {self.sub_title_date}', self.sub_title_format)
        worksheet.write('B9', 'FDS', self.headers_fds_format)
        worksheet.write('C9', 'Bravo', self.headers_bravo_format)
        worksheet.write('D9', 'Chênh lệch', self.headers_diff_format)
        worksheet.write('A10', 'Tài khoản ký quỹ', self.headers_root_format)
        worksheet.write('B10', 'Phí GD', self.headers_fds_format)
        worksheet.write('C10', 'Phí GD', self.headers_bravo_format)
        worksheet.write('D10', '', self.headers_diff_format)
        worksheet.write_column('A11', table['SoTaiKhoan'], self.text_root_format)
        worksheet.write_column('B11', table['PhiGD_FDS'], self.money_fds_format)
        worksheet.write_column('C11', table['PhiGD_Bravo'], self.money_bravo_format)
        worksheet.write_column('D11', table['PhiGD_Diff'], self.money_diff_format)

        sum_start_row = table.shape[0] + 11
        worksheet.write(f'A{sum_start_row}', 'Grand Total', self.headers_root_format)
        worksheet.write(f'B{sum_start_row}', f'=SUM(B11:B{sum_start_row - 1})', self.money_sum_fds_format)
        worksheet.write(f'C{sum_start_row}', f'=SUM(C11:C{sum_start_row - 1})', self.money_sum_bravo_format)
        worksheet.write(f'D{sum_start_row}', f'=SUM(D11:D{sum_start_row - 1})', self.money_sum_diff_format)

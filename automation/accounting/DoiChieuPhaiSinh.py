import calendar
from os.path import dirname, join
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
                'font_name': 'Arial'
            }
        )
        self.headers_wrap_format = self.workbook.add_format(
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
                'align': 'right',
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
        worksheet.write('A2', f'Từ ngày {self.sub_title_date} đến ngày {self.sub_title_date}', self.info_format)
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
        worksheet.write_column('D5', table['TenKhachHang'], self.text_left_format)
        worksheet.write_column('E5', table['TienTaiPHS'], self.money_fds_format)
        worksheet.write_column('F5', table['TienTaiVSD'], self.money_fds_format)
        worksheet.write_column('G5', table['DuCoCuoi3243'], self.money_bravo_diff_format)
        worksheet.write_column('H5', table['DuCoCuoi338804'], self.money_bravo_diff_format)
        worksheet.write_column('I5', table['TienTaiPHSDiff'], self.money_bravo_diff_format)
        worksheet.write_column('J5', table['TienQuyVSDDiff'], self.money_bravo_diff_format)

        sum_start_row = table.shape[0] + 5
        worksheet.merge_range(f'A{sum_start_row}:B{sum_start_row}', 'Tổng cộng:', self.sum_format)
        worksheet.write_row(f'C{sum_start_row}', [''] * 2, self.sum_format)
        for col in 'EFGHIJ':
            if col == 'EF':
                worksheet.write(f'{col}{sum_start_row}', f'=SUM({col}5:{col}{sum_start_row - 1})',
                                self.money_sum_fds_format)
            else:
                worksheet.write(f'{col}{sum_start_row}', f'=SUM({col}5:{col}{sum_start_row - 1})',
                                self.money_sum_bravo_diff_format)

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
        worksheet.set_column('A:B', 17)
        worksheet.set_column('C:C', 43)
        worksheet.set_column('D:Z', 17)

        worksheet.write('A1', 'BÁO CÁO SỐ DƯ KHOẢN CHẬM TRẢ', self.info_format)
        worksheet.write('A2', f'Từ ngày {self.sub_title_date} đến ngày {self.sub_title_date}', self.info_format)
        worksheet.merge_range('A3:N3', 'FDS', self.FDS_title_format)
        worksheet.merge_range('O3:T3', 'Bravo', self.bravo_title_format)
        worksheet.merge_range('U3:Z3', 'Chênh lệch', self.diff_title_format)
        worksheet.merge_range('E4:G5', 'Đầu kỳ', self.headers_format)
        worksheet.merge_range('H4:K4', 'Trong kỳ', self.headers_format)
        worksheet.merge_range('L4:N5', 'Cuối kỳ', self.headers_format)
        worksheet.merge_range('O4:P4', 'Cuối kỳ', self.headers_format)
        worksheet.merge_range('Q4:T4', 'Trong kỳ', self.headers_format)
        worksheet.merge_range('U4:V4', 'Cuối kỳ', self.headers_format)
        worksheet.merge_range('W4:Z4', 'Trong kỳ', self.headers_format)
        worksheet.merge_range('H5:I5', 'Phát sinh tăng', self.headers_format)
        worksheet.merge_range('J5:K5', 'Phát sinh giảm', self.headers_format)
        worksheet.merge_range('Q5:R5', 'PS tăng', self.headers_format)
        worksheet.merge_range('S5:T5', 'PS giảm', self.headers_format)
        worksheet.merge_range('W5:X5', 'PS tăng', self.headers_format)
        worksheet.merge_range('Y5:Z5', 'PS giảm', self.headers_format)

        worksheet.write_row('A4', ['']*4, self.headers_format)
        worksheet.write_row('A5', ['']*4, self.headers_format)
        worksheet.write_row('O5', ['']*2, self.headers_format)
        worksheet.write_row('U5', ['']*2, self.headers_format)
        worksheet.write_row(
            'A6',
            [
                'Tài khoản ký quỹ',
                'Tài khoản giao dịch',
                'Tên Khách Hàng',
                'Chi nhánh',
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
                'Khoản chậm trả',
                'Phí Khoản chậm trả',
                'Khoản chậm trả',
                'Phí chậm trả',
                'Khoản chậm trả',
                'Phí chậm trả',
                'Khoản chậm trả',
                'Phí Khoản chậm trả',
                'Khoản chậm trả',
                'Phí chậm trả',
                'Khoản chậm trả',
                'Phí chậm trả'
            ],
            self.headers_wrap_format
        )

        worksheet.write_column('A7', table['SoTaiKhoan'], self.text_left_format)
        worksheet.write_column('B7', table['SoTieuKhoan'], self.text_left_format)
        worksheet.write_column('C7', table['TenKhachHang'], self.text_left_format)
        worksheet.write_column('D7', table['MaChiNhanh'], self.text_left_format)
        worksheet.write_column('E7', table['KhoanChamTraDauKy'], self.money_fds_format)
        worksheet.write_column('F7', table['PhiChamTraDauKy'], self.money_fds_format)
        worksheet.write_column('G7', table['TongTienChamDauKy'], self.money_fds_format)
        worksheet.write_column('H7', table['KhoanChamTraPSTangTrongKy'], self.money_fds_format)
        worksheet.write_column('I7', table['PhiChamTraPSTangTrongKy'], self.money_fds_format)
        worksheet.write_column('J7', table['KhoanChamTraPSGiamTrongKy'], self.money_fds_format)
        worksheet.write_column('K7', table['PhiChamTraPSGiamTrongKy'], self.money_fds_format)
        worksheet.write_column('L7', table['KhoanChamTraCuoiKy'], self.money_fds_format)
        worksheet.write_column('M7', table['PhiChamTraCuoiKy'], self.money_fds_format)
        worksheet.write_column('N7', table['TongTienChamCuoiKy'], self.money_fds_format)
        worksheet.write_column('O7', table['DuNoCuoi13504'], self.money_bravo_diff_format)
        worksheet.write_column('P7', table['DuNoCuoi13505'], self.money_bravo_diff_format)
        worksheet.write_column('Q7', table['PhatSinhNo13504'], self.money_bravo_diff_format)
        worksheet.write_column('R7', table['PhatSinhNo13505'], self.money_bravo_diff_format)
        worksheet.write_column('S7', table['PhatSinhCo13504'], self.money_bravo_diff_format)
        worksheet.write_column('T7', table['PhatSinhCo13505'], self.money_bravo_diff_format)
        worksheet.write_column('U7', table['KhoanChamTraCuoiKyDiff'], self.money_bravo_diff_format)
        worksheet.write_column('V7', table['PhiChamTraCuoiKyDiff'], self.money_bravo_diff_format)
        worksheet.write_column('W7', table['KhoanChamTraPSTangTrongKyDiff'], self.money_bravo_diff_format)
        worksheet.write_column('X7', table['PhiChamTraPSTangTrongKyDiff'], self.money_bravo_diff_format)
        worksheet.write_column('Y7', table['KhoanChamTraPSGiamTrongKyDiff'], self.money_bravo_diff_format)
        worksheet.write_column('Z7', table['PhiChamTraPSGiamTrongKyDiff'], self.money_bravo_diff_format)

        sum_start_row = table.shape[0] + 7
        worksheet.merge_range(f'A{sum_start_row}:B{sum_start_row}', 'Tổng cộng:', self.sum_format)
        worksheet.write_row(f'B{sum_start_row}', [''] * 3, self.sum_format)

        for col in 'EFGHIJKLMNOPQRSTUVWXYZ':
            if col == 'EFGHIJKLMN':
                worksheet.write(f'{col}{sum_start_row}', f'=SUM({col}5:{col}{sum_start_row - 1})',
                                self.money_sum_fds_format)
            else:
                worksheet.write(f'{col}{sum_start_row}', f'=SUM({col}5:{col}{sum_start_row - 1})',
                                self.money_sum_bravo_diff_format)

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
        worksheet.set_column('A:B', 16)
        worksheet.set_column('C:D', 18)

        worksheet.write('A1', 'BÁO CÁO ĐỐI CHIẾU THUẾ TNCN', self.info_format)
        worksheet.write('A2', f'Từ ngày/from : {self.sub_title_date} Đến ngày/to : {self.sub_title_date}', self.info_format)
        worksheet.merge_range('A3:B3', 'FDS', self.FDS_title_format)
        worksheet.write('C3', 'Bravo', self.bravo_title_format)
        worksheet.write('D3', 'Chênh lệch', self.diff_title_format)

        worksheet.write_row(
            'A4',
            [
                'Số tài khoản',
                'Thuế TNCN',
                'Thuế TNCN',
                ''
            ],
            self.headers_format
        )
        worksheet.write_column('A5', table['SoTaiKhoan'], self.text_left_format)
        worksheet.write_column('B5', table['ThueTNCN_FDS'], self.money_bravo_diff_format)
        worksheet.write_column('C5', table['ThueTNCN_Bravo'], self.money_bravo_diff_format)
        worksheet.write_column('D5', table['ThueTNCN_Diff'], self.money_bravo_diff_format)

        sum_start_row = table.shape[0] + 5
        worksheet.write(f'A{sum_start_row}', 'Grand Total', self.sum_format)
        for col in 'BCD':
            worksheet.write(
                f'{col}{sum_start_row}',
                f'=SUM({col}5:{col}{sum_start_row-1})',
                self.money_sum_bravo_diff_format
            )

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
        worksheet.set_column('A:B', 16)
        worksheet.set_column('C:D', 18)

        worksheet.write('A1', 'SAO KÊ LỆNH KHỚP', self.info_format)
        worksheet.write('A2',f'Từ ngày/from : {self.sub_title_date} Đến ngày/to : {self.sub_title_date}',self.info_format)
        worksheet.merge_range('A3:B3', 'FDS', self.FDS_title_format)
        worksheet.write('C3', 'Bravo', self.bravo_title_format)
        worksheet.write('D3', 'Chênh lệch', self.diff_title_format)

        worksheet.write_row(
            'A4',
            [
                'Tài khoản ký quỹ',
                'Phí GD',
                'Phí GD',
                ''
            ],
            self.headers_format
        )
        worksheet.write_column('A5', table['SoTaiKhoan'], self.text_left_format)
        worksheet.write_column('B5', table['PhiGD_FDS'], self.money_bravo_diff_format)
        worksheet.write_column('C5', table['PhiGD_Bravo'], self.money_bravo_diff_format)
        worksheet.write_column('D5', table['PhiGD_Diff'], self.money_bravo_diff_format)

        sum_start_row = table.shape[0] + 5
        worksheet.write(f'A{sum_start_row}', 'Grand Total', self.sum_format)
        for col in 'BCD':
            worksheet.write(
                f'{col}{sum_start_row}',
                f'=SUM({col}5:{col}{sum_start_row-1})',
                self.money_sum_bravo_diff_format)

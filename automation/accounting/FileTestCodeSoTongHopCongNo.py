from automation.accounting import *


def run():
    start = time.time()
    t_date = '2022.01.25'

    bravoFolder = join(dirname(dept_folder), 'FileFromBravo')
    # query data RLN0006 (Margin outstanding)
    query_rln06 = pd.read_sql(
        f"""
            SELECT
                [margin_outstanding].[account_code] [SoTaiKhoan],
                MAX([account].[customer_name]) [TenKhachHangFlex],
                SUM([margin_outstanding].[interest_outstanding]) [DuCuoiNoFlex]
            FROM [margin_outstanding]
            LEFT JOIN [account] ON [account].[account_code] = [margin_outstanding].[account_code]
            WHERE [margin_outstanding].[date] = '{t_date}'
                AND  [margin_outstanding].[type] <> N'Ứng trước cổ tức'
            GROUP BY [margin_outstanding].[account_code]
            """,
        connect_DWH_CoSo
    )

    col_name = ['SoTaiKhoan', 'TenKhachHangBravo', 'DuDauNoBravo', 'DuDauCoBravo', 'PhatSinhNoBravo',
                'PhatSinhCoBravo', 'DuCuoiNoBravo', 'DuCuoiCoBravo']

    # process data by pandas - sheet 1231
    df_1231 = pd.read_excel(
        join(bravoFolder, f'{t_date}', f'Sổ tổng hợp công nợ 1231_{t_date}.xlsx'),
        skiprows=8,
        skipfooter=1,
        names=col_name
    )

    table_1231 = pd.merge(df_1231, query_rln06, how='outer', on='SoTaiKhoan')
    table_1231['TenKhachHang'] = table_1231['TenKhachHangBravo'].fillna(table_1231['TenKhachHangFlex'])
    table_1231 = table_1231.fillna(0)
    table_1231['DuCuoiNoDiff'] = table_1231['DuCuoiNoBravo'] - table_1231['DuCuoiNoFlex']

    # process data by pandas - sheet 13226
    df_13226 = pd.read_excel(
        join(bravoFolder, f'{t_date}', f'Sổ tổng hợp công nợ 13226_{t_date}.xlsx'),
        skiprows=8,
        skipfooter=1,
        names=col_name
    )

    table_13226 = pd.merge(df_13226, query_rln06, how='outer', on='SoTaiKhoan')

    table_13226['TenKhachHang'] = table_13226['TenKhachHangBravo'].fillna(table_13226['TenKhachHangFlex'])
    table_13226 = table_13226.fillna(0)

    table_13226 = table_13226.sort_values('SoTaiKhoan', ignore_index=True)
    table_13226['CL'] = table_13226['DuNoCuoiBravo'] - table_13226['interest_outstanding']


def run_ps_1():
    start = time.time()
    t_date = '2022.01.25'
    bravoFolder = join(dirname(dept_folder), 'FileFromBravo')

    TaiKhoan_3243 = pd.read_excel(
        join(bravoFolder, f'{t_date}', f'Sổ tổng hợp công nợ 3243_{t_date}.xlsx'),
        skiprows=8,
        skipfooter=1,
        usecols=('1', '2', '8')
    ).rename(columns={'1': 'SoTaiKhoan', '2': 'TenKhachHang3243', '8': 'DuCoCuoi3243'})

    TaiKhoan_338804 = pd.read_excel(
        join(bravoFolder, f'{t_date}', f'Sổ tổng hợp công nợ 338804_{t_date}.xlsx'),
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
        LEFT JOIN [relationship] ON [relationship].[account_code] = [r].[account_code] 
        AND [relationship].[date] = [r].[date]
        LEFT JOIN [account] ON [account].[account_code] = [r].[account_code]
        WHERE [r].[date] = '{t_date}'
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


def run_ps_2():
    start = time.time()
    end_date = '2022.01.25'
    bravoFolder = join(dirname(dept_folder), 'FileFromBravo')

    TaiKhoan_13504 = pd.read_excel(
        join(bravoFolder, f'{end_date}', f'Sổ tổng hợp công nợ 13504_{end_date}.xlsx'),
        skiprows=8,
        skipfooter=1,
        usecols=('1', '2', '5', '6', '7')
    ).rename(columns={'1': 'SoTaiKhoan', '2': 'TenKhachHang13504', '5': 'PhatSinhNo13504', '6': 'PhatSinhCo13504',
                      '7': 'DuNoCuoi13504'})
    TaiKhoan_13505 = pd.read_excel(
        join(bravoFolder, f'{end_date}', f'Sổ tổng hợp công nợ 13505_{end_date}.xlsx'),
        skiprows=8,
        skipfooter=1,
        usecols=('1', '2', '5', '6', '7')
    ).rename(columns={'1': 'SoTaiKhoan', '2': 'TenKhachHang13505', '5': 'PhatSinhNo13505', '6': 'PhatSinhCo13505',
                      '7': 'DuNoCuoi13505'})

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
        WHERE [r].[date] = '{end_date}'
        ORDER BY [SoTaiKhoan]
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


def run_ps_3():
    run_time = dt.datetime(2022, 1, 25)
    t_date = run_time.strftime('%Y.%m.%d')
    sod = dt.datetime(run_time.year, run_time.month, 1).strftime('%Y.%m.%d')
    eod = dt.datetime(
        run_time.year, run_time.month, calendar.monthrange(run_time.year, run_time.month)[1]
    ).strftime('%Y.%m.%d')

    bravoFolder = join(dirname(dept_folder), 'FileFromBravo')

    TaiKhoan_33353 = pd.read_excel(
        join(bravoFolder, f'{t_date}', f'BẢNG KÊ CTU TK 33353 THANG {t_date[5:7]}.{t_date[0:4]}.xlsx'),
        skiprows=7,
        skipfooter=1,
        usecols=('Tiền', 'Mã đối tượng\n(chi tiết)')
    ).rename(columns={'Tiền': 'ThueTNCN_Bravo', 'Mã đối tượng\n(chi tiết)': 'SoTaiKhoan'})
    TaiKhoan_33353 = TaiKhoan_33353.loc[TaiKhoan_33353['SoTaiKhoan'] != 'GO0065']

    RDT0127 = pd.read_sql(
        f"""
        SELECT
            [r].[SoTaiKhoan],
            SUM([r].[ThueTNCN]) [ThueTNCN_FDS]
        FROM [RDT0127] [r]
        WHERE [r].[Ngay] = '{t_date}'
        GROUP BY [r].[SoTaiKhoan]
        ORDER BY [r].[SoTaiKhoan]
        """,
        connect_DWH_PhaiSinh
    )

    table = RDT0127.merge(TaiKhoan_33353, how='outer', on='SoTaiKhoan')
    table = table.fillna(0)

    table['ThueTNCN_Diff'] = table['ThueTNCN_FDS'] - table['ThueTNCN_Bravo']


def run_ps_4():
    run_time = dt.datetime(2022, 1, 25)
    t_date = run_time.strftime('%Y.%m.%d')
    sod = dt.datetime(run_time.year, run_time.month, 1).strftime('%Y.%m.%d')
    eod = dt.datetime(
        run_time.year, run_time.month, calendar.monthrange(run_time.year, run_time.month)[1]
    ).strftime('%Y.%m.%d')

    bravoFolder = join(dirname(dept_folder), 'FileFromBravo')

    TaiKhoan_5115104 = pd.read_excel(
        join(bravoFolder, f'{t_date}', f'BẢNG KÊ CTU TK 5115104 THANG {t_date[5:7]}.{t_date[0:4]}.xlsx'),
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
        WHERE [r].[date] = '{t_date}'
        GROUP BY [relationship].[account_code]
        ORDER BY [SoTaiKhoan]
        """,
        connect_DWH_PhaiSinh
    )

    table = RDO0002.merge(TaiKhoan_5115104, how='outer', on='SoTaiKhoan')
    table = table.fillna(0)

    table['PhiGD_Diff'] = table['PhiGD_FDS'] - table['PhiGD_Bravo']


def run_PLK():
    # df['CL'] = np.where(df['DuNoCuoi']>0, df['DuNoCuoi'] - df['Flex'], df['DuCoCuoi'] - df['Flex'])
    run_time = dt.datetime(2022, 1, 25)
    t_date = run_time.strftime('%Y.%m.%d')
    sod = dt.datetime(run_time.year, run_time.month, 1).strftime('%Y.%m.%d')
    eod = dt.datetime(
        run_time.year, run_time.month, calendar.monthrange(run_time.year, run_time.month)[1]
    ).strftime('%Y.%m.%d')

    bravoFolder = join(dirname(dept_folder), 'FileFromBravo')
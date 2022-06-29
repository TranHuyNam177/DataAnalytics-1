from automation.accounting import *


def run():
    start = time.time()
    t_date = '2022.01.25'

    # query data RLN0006 (Margin outstanding)
    query_rln06 = pd.read_sql(
        f"""
            SELECT
                [r].[account_code] [SoTaiKhoan],
                SUM([r].[principal_outstanding]) [principal_outstanding],
                SUM([r].[interest_outstanding]) [interest_outstanding]
            FROM [margin_outstanding] [r]
            WHERE [r].[date] = '{t_date}'
            AND  [r].[type] <> N'Ứng trước cổ tức'
            GROUP BY [r].[account_code]
            """,
        connect_DWH_CoSo
    )

    col_name = ['SoTaiKhoan', 'TenKhachHang', 'DuNoDauBravo', 'DuCoDauBravo',
                'PhatSinhNoBravo', 'PhatSinhCoBravo', 'DuNoCuoiBravo', 'DuCoCuoiBravo']

    # process data by pandas - sheet 1231
    df_1231 = pd.read_excel(
        join(dirname(__file__), 'file', f'Sổ tổng hợp công nợ 1231_{t_date}.xlsx'),
        skiprows=8,
        skipfooter=1,
        names=col_name
    )

    table_1231 = df_1231.merge(query_rln06[['SoTaiKhoan', 'principal_outstanding']], on='SoTaiKhoan', how='outer')

    table_1231['TenKhachHang'] = table_1231['TenKhachHang'].fillna('')
    table_1231.iloc[2:] = table_1231.iloc[2:].fillna(0)

    table_1231 = table_1231.sort_values('SoTaiKhoan', ignore_index=True)
    table_1231['CL'] = table_1231['DuNoCuoiBravo'] - table_1231['principal_outstanding']

    # process data by pandas - sheet 13226
    df_13226 = pd.read_excel(
        join(dirname(__file__), 'file', f'Sổ tổng hợp công nợ 13226_{t_date}.xlsx'),
        skiprows=8,
        skipfooter=1,
        names=col_name
    )

    table_13226 = df_13226.merge(query_rln06[['SoTaiKhoan', 'interest_outstanding']], on='SoTaiKhoan', how='outer')

    table_13226['TenKhachHang'] = table_13226['TenKhachHang'].fillna('')
    table_13226.iloc[2:] = table_13226.iloc[2:].fillna(0)

    table_13226 = table_13226.sort_values('SoTaiKhoan', ignore_index=True)
    table_13226['CL'] = table_13226['DuNoCuoiBravo'] - table_13226['interest_outstanding']


def run_ps():
    start = time.time()
    t_date = '2022.01.25'
    bravoFolder = join(dirname(dept_folder), 'FileFromBravo')

    TaiKhoan_3243 = pd.read_excel(
        join(bravoFolder, f'{t_date}', f'Sổ tổng hợp công nợ 3243_{t_date}.xlsx'),
        skiprows=8,
        skipfooter=1,
        usecols=('1', '8')
    ).rename(columns={'1': 'SoTaiKhoan', '8': 'DuCoCuoi3243'})

    TaiKhoan_338804 = pd.read_excel(
        join(bravoFolder, f'{t_date}', f'Sổ tổng hợp công nợ 338804_{t_date}.xlsx'),
        skiprows=8,
        skipfooter=1,
        usecols=('1', '8')
    ).rename(columns={'1': 'SoTaiKhoan', '8': 'DuCoCuoi338804'})

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
        names=['SoTaiKhoan', 'TenKhachHang', 'DuNoDau13504', 'DuCoDau13504',
               'PhatSinhNo13504', 'PhatSinhCo13504', 'DuNoCuoi13504', 'DuCoCuoi13504']
    )
    TaiKhoan_13505 = pd.read_excel(
        join(bravoFolder, f'{end_date}', f'Sổ tổng hợp công nợ 13505_{end_date}.xlsx'),
        skiprows=8,
        skipfooter=1,
        names=['SoTaiKhoan', 'TenKhachHang', 'DuNoDau13505', 'DuCoDau13505',
               'PhatSinhNo13505', 'PhatSinhCo13505', 'DuNoCuoi13505', 'DuCoCuoi13505']
    )

    RDT0141 = pd.read_sql(
        f"""
        SELECT
            [account].[account_code] [SoTaiKhoan],
            [r].[sub_account] [SoTieuKhoan],
            [account].[customer_name] [TenKhachHang],
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
        TaiKhoan_13504[['SoTaiKhoan','DuNoCuoi13504','PhatSinhNo13504','PhatSinhCo13504']],how='outer',on='SoTaiKhoan'
    ).merge(
        TaiKhoan_13505[['SoTaiKhoan','DuNoCuoi13505','PhatSinhNo13505','PhatSinhCo13505']],how='outer',on='SoTaiKhoan'
    )
    table = table.fillna(0)

    table['KhoanChamTraCuoiKyDiff'] = table['KhoanChamTraCuoiKy'] - table['DuNoCuoi13504']
    table['PhiChamTraCuoiKyDiff'] = table['PhiChamTraCuoiKy'] - table['DuNoCuoi13505']
    table['KhoanChamTraPSTangTrongKyDiff'] = table['KhoanChamTraPSTangTrongKy'] - table['PhatSinhNo13504']
    table['PhiChamTraPSTangTrongKyDiff'] = table['PhiChamTraPSTangTrongKy'] - table['PhatSinhNo13505']
    table['KhoanChamTraPSGiamTrongKyDiff'] = table['KhoanChamTraPSGiamTrongKy'] - table['PhatSinhCo13504']
    table['PhiChamTraPSGiamTrongKyDiff'] = table['PhiChamTraPSGiamTrongKy'] - table['PhatSinhCo13505']


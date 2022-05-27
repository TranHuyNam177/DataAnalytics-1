from automation.risk_management import *
from request.stock import *
from datawarehouse import *

"""
- bảng vmr0104 1 trong 2 cột "Sum of Room hệ thống đã sử dụng", "Sum of Room đặc biệt sử dụng" bằng 0,
loại trường hợp nào mà cả 2 cột đều bằng 0 
- bảng rmr0062 chỉ quan tâm có vay (cột loan_type = 1)
- bảng vmr9004 loại các trường hợp mà KL margin = 0
"""


def run(  # chạy hàng ngày
        run_time=dt.datetime.now()
):
    start = time.time()
    info = get_info('daily', run_time)
    period = info['period']
    t0_date = info['end_date']
    t1_date = BDATE(t0_date, -1)
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ###################################################
    ###################################################
    ###################################################

    # "Vốn chủ sở hữu của CTCK (4)"
    equity_CTCK = 1565140

    # Get and process data
    tickers = internal.mlist()
    tup_tickers = tuple(tickers)

    table = pd.read_sql(
        f"""
        WITH [rmr62] AS (
            SELECT
                [rmr0062].[account_code],
                [rmr0062].[cash]
            FROM [rmr0062]
            WHERE [rmr0062].[date] = '{t1_date}'
            AND [rmr0062].[loan_type] = 1
        ),
        [rln06_rmr62] AS (
            SELECT
                [margin_outstanding].[account_code],
                [margin_outstanding].[type],
                CASE
                    WHEN ([margin_outstanding].[principal_outstanding]-[rmr62].[cash]) > 0
                        THEN [margin_outstanding].[principal_outstanding]-[rmr62].[cash]
                    ELSE 0
                END [net_cash]
            FROM [margin_outstanding]
            LEFT JOIN [rmr62] 
            ON [rmr62].[account_code] = [margin_outstanding].[account_code]
            WHERE [margin_outstanding].[date] = '{t1_date}'
            AND [margin_outstanding].[type] IN (N'Margin', N'Trả chậm', N'Bảo lãnh')
        ),
        [v1] AS (
            SELECT
                [vmr0104].[ticker],
                [vmr0104].[sub_account],
                [vmr0104].[mark_type],
                [vmr0104].[used_system_room],
                [vmr0104].[special_room],
                ([vmr0104].[used_system_room]+[vmr0104].[special_room]) [total_used_room]
            FROM [vmr0104]
            WHERE [vmr0104].[date] = '{t1_date}'
            AND ([vmr0104].[used_system_room] <> 0 OR [vmr0104].[special_room] <> 0)
        ),
        [v2] AS (
            SELECT
                [vmr9004].[ticker],
                [vmr9004].[sub_account],
                [vmr9004].[margin_volume]
            FROM [vmr9004]
            WHERE [vmr9004].[date] = '{t1_date}'
            AND [vmr9004].[margin_volume] <> 0
        ),
        [sub_acc] AS (
            SELECT
                [relationship].[account_code],
                [relationship].[sub_account]
            FROM [relationship]
            WHERE [relationship].[date] = '{t1_date}'
        ),
        [market_price] AS (
            SELECT
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ticker],
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Close]
            FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
            WHERE [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Date] = '{t1_date}'
        ),
        [exchange] AS (
            SELECT
                [DWH-ThiTruong].[dbo].[DanhSachMa].[Ticker],
                [DWH-ThiTruong].[dbo].[DanhSachMa].[Exchange]
            FROM [DWH-ThiTruong].[dbo].[DanhSachMa]
            WHERE [DWH-ThiTruong].[dbo].[DanhSachMa].[Date] = '{t1_date}'
        ),
        [table] AS (
            SELECT
                [sub_acc].[account_code],
                [v1].[sub_account],
                [v1].[ticker],
                [v1].[total_used_room],
                ([v2].[margin_volume]*1000*[market_price].[Close]) [asset_val_by_stock],
                [rln06_rmr62].[net_cash]
            FROM [v1]
            LEFT JOIN [sub_acc] ON [sub_acc].[sub_account] = [v1].[sub_account]
            LEFT JOIN [v2]
            ON [v1].[sub_account] = [v2].[sub_account] AND [v1].[ticker] = [v2].[ticker]
            LEFT JOIN [market_price] ON [market_price].[Ticker] = [v1].[ticker]
            LEFT JOIN [rln06_rmr62] ON [rln06_rmr62].[account_code] = [sub_acc].[account_code]
            WHERE [v2].[margin_volume] IS NOT NULL
            AND [rln06_rmr62].[net_cash] IS NOT NULL
        ),
        [table_2] AS (
            SELECT
                [table].[ticker],
                [table].[total_used_room],
                {equity_CTCK} AS [equity_CTCK],
                [table].[asset_val_by_stock] / (SUM(asset_val_by_stock) OVER (PARTITION BY [account_code])) [ty_trong],
                [table].[net_cash]
            FROM
                [table]
        ),
        [final_table] AS (
            SELECT
                [table_2].[ticker],
                CASE
                    WHEN SUM(([table_2].[net_cash] * [table_2].[ty_trong])/1000000) > MAX([table_2].[equity_CTCK])*0.1
                        THEN MAX([table_2].[equity_CTCK])*0.1
                    ELSE SUM(([table_2].[net_cash] * [table_2].[ty_trong])/1000000)
                END [DuNoChoVayGDKQ],
                SUM([table_2].[total_used_room])/1000 [SoLuongCKChoVayCTCK],
                MAX([table_2].[equity_CTCK]) [VonChuSoHuuCTCK]
            FROM [table_2]
            GROUP BY [table_2].[ticker]
        )
        SELECT
            [exchange].[ticker],
            ISNULL([final_table].[DuNoChoVayGDKQ],0) [DuNoChoVayGDKQ],
            ISNULL([final_table].[SoLuongCKChoVayCTCK],0) [SoLuongCKChoVayCTCK],
            ISNULL([final_table].[VonChuSoHuuCTCK],{equity_CTCK}) [VonChuSoHuuCTCK],
            ISNULL(ROUND(([final_table].[DuNoChoVayGDKQ] / [final_table].VonChuSoHuuCTCK),5),0) [DuNo/VCSH],
            [exchange].[Exchange] [san]
        FROM [final_table]
        FULL OUTER JOIN [exchange]
        ON [exchange].[Ticker] = [final_table].[ticker]
        WHERE [exchange].[Ticker] IN {tup_tickers}
        ORDER BY [ticker]
        """,
        connect_DWH_CoSo
    )
    kl_tsdb = table.groupby('san')['SoLuongCKChoVayCTCK'].sum()
    du_no_ma = table.groupby('san')['DuNoChoVayGDKQ'].sum()

    ###################################################
    ###################################################
    ###################################################

    t0_day = t1_date[8:10]
    t0_month = t1_date[5:7]
    t0_year = t1_date[0:4]
    file_date = t0_day + t0_month + t0_year
    file_name = f'180426__RMD_SCMS_Bao cao ngay truoc 8AM {file_date} New.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder, folder_name, period, file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # Format
    title_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 13,
            'font_name': 'Times New Roman',
        }
    )
    sub_title_format = workbook.add_format(
        {
            'border': 1,
            'italic': True,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 13,
            'font_name': 'Times New Roman',
        }
    )
    headers_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'text_wrap': True
        }
    )
    headers_2_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    TCNY_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '#,##0'
        }
    )
    vcsh_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 13,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    total_vcsh_text_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '#,##0'
        }
    )
    total_vcsh_number_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '#,##0'
        }
    )
    percent_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '0.000%'
        }
    )
    bold_format = workbook.add_format(
        {
            'bold': True,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    bold_number_format = workbook.add_format(
        {
            'bold': True,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '#,##0'
        }
    )
    number_format = workbook.add_format(
        {
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )

    ###################################################
    ###################################################
    ###################################################

    # WRITE EXCEL
    title_sheet = 'TÌNH HÌNH GIAO DỊCH KÝ QUỸ'
    sub_title = 'Đơn vị : nghìn cổ phiếu,triệu đồng'

    headers = [
        'STT',
        'Mã CK',
        'Dư nợ cho vay GDKQ (1)',
        'Số lượng chứng khoán cho vay của CTCK (2)',
        'Số lượng chứng khoán niêm yết của TCNY (3)',
        'Vốn chủ sở hữu của CTCK (4)',
        'Tỷ lệ dư nợ/VCSH (1)/(4)',
        'Tỷ lệ CK cho vay/CKNY (2)/(3)',
        'Sàn'
    ]
    headers_2 = [
        'Sàn',
        'Khối lượng TSĐB',
        'Dư nợ mã'
    ]

    worksheet = workbook.add_worksheet('Bao cao')

    worksheet.set_column('A:B', 8)
    worksheet.set_column('C:C', 10)
    worksheet.set_column('D:D', 15)
    worksheet.set_column('E:E', 16)
    worksheet.set_column('F:G', 12)
    worksheet.set_column('H:H', 14)
    worksheet.set_column('I:I', 9)
    worksheet.set_column('K:K', 15)
    worksheet.set_column('L:L', 20)
    worksheet.set_column('M:M', 11)

    worksheet.set_row(3, 66)

    worksheet.merge_range('A1:H1', title_sheet, title_format)
    worksheet.merge_range('A2:H2', sub_title, sub_title_format)
    worksheet.write_row('A4', headers, headers_format)
    worksheet.write_column('A5', np.arange(table.shape[0]) + 1, text_center_format)
    worksheet.write_column('B5', table['ticker'], text_center_format)
    worksheet.write_column('C5', table['DuNoChoVayGDKQ'], money_format)
    worksheet.write_column('D5', table['SoLuongCKChoVayCTCK'], money_format)
    worksheet.write_column('E5', '', TCNY_format)
    worksheet.write_column('F5', table['VonChuSoHuuCTCK'], vcsh_format)
    worksheet.write_column('G5', table['DuNo/VCSH'], percent_format)
    worksheet.write_column('H5', '', percent_format)
    worksheet.write_column('I5', table['san'], text_left_format)
    sum_row = table.shape[0] + 5
    worksheet.write(f'C{sum_row}', table['DuNoChoVayGDKQ'].sum(), total_vcsh_number_format)
    worksheet.write(f'D{sum_row}', table['SoLuongCKChoVayCTCK'].sum(), total_vcsh_number_format)
    worksheet.write_row('K4', headers_2, headers_2_format)
    worksheet.write_column('K5', ['HOSE', 'HNX'], text_left_format)
    worksheet.write('K7', 'Total', total_vcsh_text_format)
    worksheet.write('L5', kl_tsdb['HOSE'], TCNY_format)
    worksheet.write('L6', kl_tsdb['HNX'], TCNY_format)
    worksheet.write('L7', kl_tsdb.sum(), total_vcsh_number_format)
    worksheet.write('L8', '200% VCSH', bold_format)
    worksheet.write('M5', du_no_ma['HOSE'], TCNY_format)
    worksheet.write('M6', du_no_ma['HNX'], TCNY_format)
    worksheet.write('M7', du_no_ma.sum(), total_vcsh_number_format)
    worksheet.write('M8', equity_CTCK * 2, bold_number_format)
    worksheet.write('M9', du_no_ma.sum() - equity_CTCK * 2, number_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')

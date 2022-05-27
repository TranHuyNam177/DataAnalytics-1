from automation.risk_management import *
from request.stock import *
from datawarehouse import *

"""
Mã bị lệch cột Used general room (cột system_used_room trên SQL trong table [230007])
1. SHB
2. PDR
3. HPG
4. DXG
Note: FLC tuy không có trong DB nhưng vẫn còn dư nợ của cty nên bạn Huy thêm số total outstanding bằng tay
"""


def run(  # chạy hàng ngày
    run_time=None
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

    # read file excel which has companies name
    file_path = join(dirname(__file__), 'excel_file', 'Tổng 2 sàn.xlsx')
    company_name = pd.read_excel(
        file_path,
        usecols=['Mã chứng khoán thực hiện giao dịch ký quỹ','Tên Công ty (ENG)']
    )
    company_name = company_name.set_index('Mã chứng khoán thực hiện giao dịch ký quỹ')

    # convert list to tuple
    def convert(lst):
        return tuple(ele for ele in lst)

    # Get and process data
    lst_ticker = internal.mlist()
    tickers = convert(lst_ticker)

    table = pd.read_sql(
        f"""
        WITH [thitruong] AS (
            SELECT
                [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Ticker] [ticker],
                ([DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Close] * 1000) [closed_price]
            FROM [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay]
            WHERE [DWH-ThiTruong].[dbo].[DuLieuGiaoDichNgay].[Date] = '{t1_date}'
        ),
        [res] AS (
            SELECT
                [230007].[date],
                [230007].[ticker],
                [230007].[system_used_room],
                [230007].[used_special_room],
                ([230007].[system_used_room] + [230007].[used_special_room]) [total_used_room],
                [thitruong].[closed_price],
                [vpr0109].[margin_max_price] [max_price],
                CASE
                    WHEN [thitruong].[closed_price] < [vpr0109].[margin_max_price]
                    THEN [thitruong].[closed_price]
                    ELSE [vpr0109].[margin_max_price]
                END [min_value],
                [vpr0109].[margin_ratio] [ratio]
            FROM [230007]
            LEFT JOIN [vpr0109] ON [vpr0109].[ticker_code] = [230007].[ticker] AND [vpr0109].[date] = [230007].[date]
            LEFT JOIN [thitruong] ON [thitruong].[ticker] = [230007].[ticker]
            WHERE [230007].[date] = '{t1_date}'
            AND [230007].[ticker] IN {tickers}
            AND (
                [vpr0109].[room_code] LIKE N'CL101%' 
                OR [vpr0109].[room_code] LIKE N'TC01%'
            )
        )
        SELECT
            [res].[ticker],
            [res].[system_used_room],
            [res].[used_special_room],
            [res].[total_used_room],
            [res].[closed_price],
            [res].[max_price],
            [res].[ratio],
            ([res].[total_used_room] * [res].[min_value] * [res].[ratio]/100) [total_outstanding]
        FROM [res]
        ORDER BY [total_outstanding] DESC
        """,
        connect_DWH_CoSo,
        index_col='ticker'
    )
    table['name'] = company_name['Tên Công ty (ENG)']
    table = table.reset_index()

    ###################################################
    ###################################################
    ###################################################

    t0_day = t0_date[8:10]
    t0_month = int(t0_date[5:7])
    t0_month = calendar.month_name[t0_month]
    t0_year = t0_date[0:4]
    file_date = t0_month + ' ' + t0_day + ' ' + t0_year
    file_name = f'Top 30 biggest outstanding {file_date}.xlsx'
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
            'bold': True,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 14,
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
            'font_name': 'Arial',
            'text_wrap': True
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial'
        }
    )
    number_format = workbook.add_format(
        {
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial'
        }
    )
    rat_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial'
        }
    )
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Arial',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )

    ###################################################
    ###################################################
    ###################################################

    # WRITE EXCEL
    title_sheet = 'TOP  30 BIGGEST OUTSTANDING'
    sub_date = dt.datetime.strptime(t0_date, '%Y.%m.%d').strftime('%d/%m/%Y')

    headers = [
        'No',
        'Code stock',
        'Name',
        'Used general room',
        'Used special room',
        'Total used room (previous session)',
        'Closed price (previous session)',
        'Max Price',
        'Ratio',
        'Total outstanding'
    ]

    worksheet = workbook.add_worksheet('Top 30')
    worksheet.hide_gridlines(option=2)

    worksheet.set_column('A:A', 8)
    worksheet.set_column('B:B', 12)
    worksheet.set_column('C:C', 59)
    worksheet.set_column('D:I', 0)
    worksheet.set_column('J:J', 23)
    worksheet.set_row(3, 50)

    for i in range(34,table.shape[0]+4):
        worksheet.set_row(i, 0)

    worksheet.write('C1',title_sheet,title_format)
    worksheet.write('C2',sub_date,title_format)
    worksheet.write_row('A4',headers,headers_format)
    worksheet.write_column('A5',np.arange(table.shape[0])+1,number_format)
    worksheet.write_column('B5',table['ticker'],text_left_format)
    worksheet.write_column('C5',table['name'],text_left_format)
    worksheet.write_column('D5',table['system_used_room'],money_format)
    worksheet.write_column('E5',table['used_special_room'],money_format)
    worksheet.write_column('F5',table['total_used_room'],money_format)
    worksheet.write_column('G5',table['closed_price'],money_format)
    worksheet.write_column('H5',table['max_price'],money_format)
    worksheet.write_column('I5',table['ratio'],rat_format)
    worksheet.write_column('J5',table['total_outstanding'],money_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')

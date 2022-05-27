from automation.risk_management import *
from datawarehouse import *


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

    table_1 = pd.read_sql(
        f"""
        SELECT
            CASE
                WHEN [branch].[branch_id] = '0001' THEN N'Headquarter'
                WHEN [branch].[branch_id] = '0101' THEN N'D3'
                WHEN [branch].[branch_id] = '0102' THEN N'PMH T.F'
                WHEN [branch].[branch_id] = '0104' THEN N'D7 T.O'
                WHEN [branch].[branch_id] = '0105' THEN N'TB'
                WHEN [branch].[branch_id] = '0111' THEN N'Institutional Business 01'
                WHEN [branch].[branch_id] = '0113' THEN N'IB'
                WHEN [branch].[branch_id] = '0117' THEN N'D1'
                WHEN [branch].[branch_id] = '0118' THEN N'AMD-03'
                WHEN [branch].[branch_id] = '0119' THEN N'Institutional Business 02'
                WHEN [branch].[branch_id] = '0201' THEN N'HN'
                WHEN [branch].[branch_id] = '0202' THEN N'TX'
                WHEN [branch].[branch_id] = '0203' THEN N'CG'
                WHEN [branch].[branch_id] = '0301' THEN N'HP'
                ELSE
                    [branch].[branch_name]
            END [SUMMARY],
            [RMR0067_1].[CreditLine],
            [RMR0067_1].[TotalOutstanding],
            [RMR0067_1].[OutstandingAndFee],
            [RMR0067_1].[TotalMortgagedAssets],
            [RMR0067_1].[TotalAsset],
            [RMR0067_1].[Surplus]
        FROM [RMR0067_1]
        LEFT JOIN [branch] ON [branch].[branch_id] = [RMR0067_1].[BranchID]
        WHERE [RMR0067_1].[Date] = '{t1_date}' 
        """,
        connect_DWH_CoSo
    )
    table_2 = pd.read_sql(
        f"""
        WITH [q] AS (
            SELECT
            CASE
                WHEN [RMR0067_2].[ValueRange] = N'From0bTo5b' THEN N'Total outstanding (TO) <= 05 bil. dong'
                WHEN [RMR0067_2].[ValueRange] = N'From0bTo1b' THEN N'TO <= 01 bil. dong'
                WHEN [RMR0067_2].[ValueRange] = N'From1bTo3b' THEN N'01 bil. dong < TO <= 03 bil. dong'
                WHEN [RMR0067_2].[ValueRange] = N'From3bTo5b' THEN N'03 bil. dong < TO <= 05 bil. dong'
                WHEN [RMR0067_2].[ValueRange] = N'GreaterThan5b' THEN N'Total outstanding > 05 bil. dong'
            ELSE [RMR0067_2].[ValueRange]
            END [valRange],
            [RMR0067_2].[TotalCustomer],
            [RMR0067_2].[TotalOutstandingAndFee],
            [RMR0067_2].[Ratio] AS [rat],
            CASE
                WHEN [RMR0067_2].[ValueRange] = N'From0bTo5b' THEN 'x'
                WHEN [RMR0067_2].[ValueRange] = N'GreaterThan5b' THEN 'x'
            END [filter]
            FROM [RMR0067_2]
            WHERE [RMR0067_2].[Date] = '{t1_date}'
        ),
        [r] AS (
            SELECT
                [q].[valRange],
                [q].[TotalCustomer],
                [q].[TotalOutstandingAndFee],
                [q].[rat],
                CASE
                    WHEN [q].[valRange] = N'Total outstanding (TO) <= 05 bil. dong' THEN 2
                    WHEN [q].[valRange] = N'TO <= 01 bil. dong' THEN 3
                    WHEN [q].[valRange] = N'01 bil. dong < TO <= 03 bil. dong' THEN 4
                    WHEN [q].[valRange] = N'03 bil. dong < TO <= 05 bil. dong' THEN 5
                    ELSE 6
                END [stt]
            FROM [q]
            UNION ALL (
                SELECT
                    '' AS [valRange],
                    SUM([q].[TotalCustomer]) [TotalCustomer],
                    SUM([q].[TotalOutstandingAndFee]) [TotalOutstandingAndFee],
                    NULL AS [rat],
                    1 AS [stt]
                FROM [q]
                GROUP BY [filter]
                HAVING MAX([filter]) = 'x'
            )
        )
        SELECT
            [r].[valRange],
            [r].[TotalCustomer],
            [r].[TotalOutstandingAndFee],
            [r].[rat]
        FROM [r]
        ORDER BY [r].[stt]
        """,
        connect_DWH_CoSo
    ).fillna('')

    table_3 = pd.read_sql(
        f"""
        SELECT
            SUM([RMR0067_1].[TotalCall]) [TotalCall]
        FROM [RMR0067_1]
        WHERE [RMR0067_1].[Date] = '{t1_date}'
        GROUP BY [Date]
        """,
        connect_DWH_CoSo
    )

    table_4 = pd.read_sql(
        f"""
        SELECT
            CASE
                WHEN [branch].[branch_id] = '0001' THEN N'Headquarter'
                WHEN [branch].[branch_id] = '0101' THEN N'D3'
                WHEN [branch].[branch_id] = '0102' THEN N'PMH T.F'
                WHEN [branch].[branch_id] = '0104' THEN N'D7 T.O'
                WHEN [branch].[branch_id] = '0105' THEN N'TB'
                WHEN [branch].[branch_id] = '0111' THEN N'Institutional Business 01'
                WHEN [branch].[branch_id] = '0113' THEN N'IB'
                WHEN [branch].[branch_id] = '0117' THEN N'D1'
                WHEN [branch].[branch_id] = '0118' THEN N'AMD-03'
                WHEN [branch].[branch_id] = '0119' THEN N'Institutional Business 02'
                WHEN [branch].[branch_id] = '0201' THEN N'HN'
                WHEN [branch].[branch_id] = '0202' THEN N'TX'
                WHEN [branch].[branch_id] = '0203' THEN N'CG'
                WHEN [branch].[branch_id] = '0301' THEN N'HP'
                ELSE
                    [branch].[branch_name]
            END [SUMMARY],
            [RMR0067_1].[MarginOutstanding],
            [RMR0067_1].[Quota],
            [RMR0067_1].[T+]
        FROM [RMR0067_1]
        LEFT JOIN [branch] ON [branch].[branch_id] = [RMR0067_1].[BranchID]
        WHERE [RMR0067_1].[Date] = '{t1_date}'
        """,
        connect_DWH_CoSo
    )
    table_5 = pd.read_sql(
        f"""
        SELECT
            CASE
                WHEN [branch].[branch_id] = '0001' THEN N'Headquarter'
                WHEN [branch].[branch_id] = '0101' THEN N'D3'
                WHEN [branch].[branch_id] = '0102' THEN N'PMH T.F'
                WHEN [branch].[branch_id] = '0104' THEN N'D7 T.O'
                WHEN [branch].[branch_id] = '0105' THEN N'TB'
                WHEN [branch].[branch_id] = '0111' THEN N'Institutional Business 01'
                WHEN [branch].[branch_id] = '0113' THEN N'IB'
                WHEN [branch].[branch_id] = '0117' THEN N'D1'
                WHEN [branch].[branch_id] = '0118' THEN N'AMD-03'
                WHEN [branch].[branch_id] = '0119' THEN N'Institutional Business 02'
                WHEN [branch].[branch_id] = '0201' THEN N'HN'
                WHEN [branch].[branch_id] = '0202' THEN N'TX'
                WHEN [branch].[branch_id] = '0203' THEN N'CG'
                WHEN [branch].[branch_id] = '0301' THEN N'HP'
                ELSE
                    [branch].[branch_name]
            END [SUMMARY],
            [RMR0067_1].[CustomerWithCreditline],
            [RMR0067_1].[CustomerWithoutCreditline],
            [RMR0067_1].[NumberOfActiveMarginCustomers],
            [RMR0067_1].[PercentOfActivesOverCustomersWithCreditline]
        FROM [RMR0067_1]
        LEFT JOIN [branch] ON [branch].[branch_id] = [RMR0067_1].[BranchID]
        WHERE [RMR0067_1].[Date] = '{t1_date}'
        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    file_date = dt.datetime.strptime(t0_date, '%Y.%m.%d').strftime('%d.%m.%Y')
    file_name = f'RMR0067 {file_date}.xlsx'
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
    sheet_title_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 14,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    sub_title_date_format = workbook.add_format(
        {
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    company_info_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'text_wrap': True
        }
    )
    headers_format = workbook.add_format(
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
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    text_left_bold_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    money_bold_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    rat_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0.000'
        }
    )
    rat_bold_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0.000'
        }
    )
    number_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0.00'
        }
    )
    sum_money_format = workbook.add_format(
        {
            'bold': True,
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
            'num_format': '#,##0'
        }
    )
    empty_row_format = workbook.add_format(
        {
            'bottom': 1,
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman',
        }
    )
    footer_format = workbook.add_format(
        {
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )
    footer_bold_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 10,
            'font_name': 'Times New Roman'
        }
    )

    ###################################################
    ###################################################
    ###################################################

    # WRITE EXCEL
    headers_1 = [
        'SUMMARY',
        'Credit line',
        'Total Outstanding',
        'Outstanding + Fee',
        'Total Mortgaged Assets',
        'Total asset',
        'Surplus'
    ]
    headers_2 = [
        'Total customer',
        'Total outstanding + Fee',
        ''
    ]
    headers_3 = [
        'Margin Outstanding',
        'Quota',
        'T+'
    ]
    headers_4 = [
        'No of active margin customers',
        '% of actives/customers with creditline'
    ]
    headers_5 = [
        'Customers with creditline',
        'Customers without creditline'
    ]

    sheet_title_name = 'SUMMARY REPORT'
    sub_date = dt.datetime.strptime(t0_date, '%Y.%m.%d').strftime('%d/%m/%Y')
    sub_title_name = f'Date: {sub_date}'
    companyAddress = '21st Floor, Phu My Hung Tower, 08 Hoang Van Thai street, Tan Phu Ward, District 7, HCMC'
    worksheet = workbook.add_worksheet(f'{period}')
    worksheet.hide_gridlines(option=2)
    worksheet.insert_image('A1', join(dirname(__file__), 'img', 'phs_logo.png'), {'x_scale': 0.84, 'y_scale': 0.71})

    worksheet.set_column('A:B', 20)
    worksheet.set_column('C:C', 18)
    worksheet.set_column('D:E', 20)
    worksheet.set_column('F:G', 19)

    worksheet.merge_range('C2:G2', companyAddress, company_info_format)
    worksheet.merge_range('A6:G6', sheet_title_name, sheet_title_format)
    worksheet.merge_range('A7:G7', sub_title_name, sub_title_date_format)

    worksheet.write_row('A4', [''] * len(headers_1), empty_row_format)
    worksheet.write_row('A10', headers_1, headers_format)
    worksheet.write_column('A11', table_1['SUMMARY'], text_left_format)
    worksheet.write_column('B11', table_1['CreditLine'], money_format)
    worksheet.write_column('C11', table_1['TotalOutstanding'], money_format)
    worksheet.write_column('D11', table_1['OutstandingAndFee'], money_format)
    worksheet.write_column('E11', table_1['TotalMortgagedAssets'], money_format)
    worksheet.write_column('F11', table_1['TotalAsset'], money_format)
    worksheet.write_column('G11', table_1['Surplus'], money_format)
    # row grand total table_1
    row_sum_table_1 = table_1.shape[0] + 11
    worksheet.write(f'A{row_sum_table_1}', 'Grand Total', headers_format)
    worksheet.write(f'B{row_sum_table_1}', table_1['CreditLine'].sum(), sum_money_format)
    worksheet.write(f'C{row_sum_table_1}', table_1['TotalOutstanding'].sum(), sum_money_format)
    worksheet.write(f'D{row_sum_table_1}', table_1['OutstandingAndFee'].sum(), sum_money_format)
    worksheet.write(f'E{row_sum_table_1}', table_1['TotalMortgagedAssets'].sum(), sum_money_format)
    worksheet.write(f'F{row_sum_table_1}', table_1['TotalAsset'].sum(), sum_money_format)
    worksheet.write(f'G{row_sum_table_1}', table_1['Surplus'].sum(), sum_money_format)

    # start row of table 2
    table_2_row = table_1.shape[0] + 15
    worksheet.merge_range(f'A{table_2_row}:B{table_2_row}', '', headers_format)
    worksheet.write_row(f'C{table_2_row}', headers_2, headers_format)

    for i in range(6):
        valRange = table_2.loc[table_2.index[i], 'valRange']
        totalCus = table_2.loc[table_2.index[i], 'TotalCustomer']
        totalOutFee = table_2.loc[table_2.index[i], 'TotalOutstandingAndFee']
        totalRat = table_2.loc[table_2.index[i], 'rat']
        if i in (1, 5) or i in (0, 1, 5):
            fmt = text_left_bold_format
            fmt_money = money_bold_format
            fmt_rat = rat_bold_format
        else:
            fmt = text_left_format
            fmt_money = money_format
            fmt_rat = rat_format

        worksheet.merge_range(f'A{table_2_row + i + 1}:B{table_2_row + i + 1}', valRange, fmt)
        worksheet.write(f'C{table_2_row + i + 1}', totalCus, fmt_money)
        worksheet.write(f'D{table_2_row + i + 1}', totalOutFee, fmt_money)
        worksheet.write(f'E{table_2_row + i + 1}', totalRat, fmt_rat)

    # start row of table 3
    table_3_row = table_2_row + table_2.shape[0] + 4
    worksheet.write(f'B{table_3_row}', 'Total call', headers_format)
    worksheet.write(f'C{table_3_row}', table_3['TotalCall'], money_bold_format)

    # start row of table 4
    table_4_row = table_3_row + 4
    worksheet.merge_range(f'A{table_4_row}:A{table_4_row + 1}', 'SUMMARY', text_left_bold_format)
    worksheet.merge_range(f'B{table_4_row}:D{table_4_row}', 'Total Outstanding', headers_format)
    worksheet.write_row(f'B{table_4_row + 1}', headers_3, headers_format)
    worksheet.write_column(f'A{table_4_row + 2}', table_4['SUMMARY'], text_left_format)
    worksheet.write_column(f'B{table_4_row + 2}', table_4['MarginOutstanding'], money_format)
    worksheet.write_column(f'C{table_4_row + 2}', table_4['Quota'], money_format)
    worksheet.write_column(f'D{table_4_row + 2}', table_4['T+'], money_format)
    # row total table_4
    row_sum_table_4 = table_4_row + 2 + table_4.shape[0]
    worksheet.write(f'A{row_sum_table_4}', 'Total', text_left_bold_format)
    worksheet.write(f'B{row_sum_table_4}', table_4['MarginOutstanding'].sum(), money_bold_format)
    worksheet.write(f'C{row_sum_table_4}', table_4['Quota'].sum(), money_bold_format)
    worksheet.write(f'D{row_sum_table_4}', table_4['T+'].sum(), money_bold_format)
    # non-interest row
    non_interest_row = row_sum_table_4 + 1
    worksheet.merge_range(f'A{non_interest_row}:B{non_interest_row}', 'Non-interest beared amount',
                          text_left_bold_format)
    worksheet.merge_range(f'A{non_interest_row + 1}:B{non_interest_row + 1}', 'Interest beared amount',
                          text_left_bold_format)
    worksheet.merge_range(f'A{non_interest_row + 2}:B{non_interest_row + 2}', 'Grand Total', text_left_bold_format)
    worksheet.merge_range(f'C{non_interest_row}:D{non_interest_row}', '', money_bold_format)
    worksheet.merge_range(f'C{non_interest_row + 1}:D{non_interest_row + 1}', '', money_bold_format)
    worksheet.merge_range(f'C{non_interest_row + 2}:D{non_interest_row + 2}', '', money_bold_format)

    # start row of table 5
    table_5_row = non_interest_row + 6
    worksheet.write(f'A{table_5_row}', 'SUMMARY', text_left_bold_format)
    worksheet.merge_range(f'B{table_5_row}:C{table_5_row}', 'No of margin customers', headers_format)
    worksheet.write_row(f'D{table_5_row}', headers_4, headers_format)
    worksheet.write(f'A{table_5_row + 1}', '', headers_format)
    worksheet.write_row(f'B{table_5_row + 1}', headers_5, headers_format)
    worksheet.write_row(f'D{table_5_row + 1}', [''] * 2, headers_format)
    worksheet.write_column(f'A{table_5_row + 2}', table_5['SUMMARY'], text_left_format)
    worksheet.write_column(f'B{table_5_row + 2}', table_5['CustomerWithCreditline'], money_format)
    worksheet.write_column(f'C{table_5_row + 2}', table_5['CustomerWithoutCreditline'], money_format)
    worksheet.write_column(f'D{table_5_row + 2}', table_5['NumberOfActiveMarginCustomers'], money_format)
    worksheet.write_column(f'E{table_5_row + 2}', table_5['PercentOfActivesOverCustomersWithCreditline'], number_format)

    # row grand total table_5
    row_sum_table_5 = table_5_row + 2 + table_5.shape[0]
    worksheet.write(f'A{row_sum_table_5}', 'Grand Total', text_left_bold_format)
    worksheet.write(f'B{row_sum_table_5}', table_5['CustomerWithCreditline'].sum(), money_bold_format)
    worksheet.write(f'C{row_sum_table_5}', table_5['CustomerWithoutCreditline'].sum(), money_bold_format)
    worksheet.write(f'D{row_sum_table_5}', table_5['NumberOfActiveMarginCustomers'].sum(), money_bold_format)
    worksheet.write(f'E{row_sum_table_5}', table_5['PercentOfActivesOverCustomersWithCreditline'].sum(),
                    money_bold_format)

    # start row of footer
    footer_row = row_sum_table_5 + 7
    t0_day = t0_date[8:10]
    t0_month = int(t0_date[5:7])
    t0_month = calendar.month_name[t0_month]
    t0_year = t0_date[0:4]
    footer_date = t0_day + '-' + t0_month + '-' + t0_year
    worksheet.write(f'G{footer_row}', f'HCM city, {footer_date}', footer_format)
    worksheet.write(f'G{footer_row + 2}', f'The creator', footer_bold_format)
    worksheet.write(f'G{footer_row + 6}', '', footer_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')

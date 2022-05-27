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
    # t1_date = '2022-05-13'
    folder_name = info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ###################################################
    ###################################################
    ###################################################

    # list các tài khoản nợ xấu cố định và loại bỏ thêm 1 tài khoản tự doanh
    acc_lst = [
        '022C078252', '022C012620', '022C012621', '022C012622', '022C089535',
        '022C089950', '022C089957', '022C050302', '022C006827', '022P002222'
    ]
    account = tuple(acc_lst)

    table_sheet_2 = pd.read_sql(
        f"""
        WITH [vcf] AS (
            SELECT
                [vcf0051].[sub_account]
            FROM [vcf0051]
            WHERE [vcf0051].[date] = '{t1_date}'
            AND [vcf0051].[contract_type] LIKE N'MR%'
        ),
        [r] AS (
            SELECT
                [relationship].[account_code],
                [relationship].[sub_account],
                [relationship].[branch_id]
            FROM [relationship]
            JOIN [vcf] ON [vcf].[sub_account] = [relationship].[sub_account]
            WHERE [relationship].[date] = '{t1_date}'
        ),
        [rmr62] AS (
            SELECT
                [rmr0062].[account_code],
                [rmr0062].[cash],
                [rmr0062].[margin_value]
            FROM [rmr0062]
            WHERE [rmr0062].[date] = '{t1_date}'
            AND [rmr0062].[loan_type] = 1
        ),
        [rmr15] AS (
            SELECT
                [rmr0015].[sub_account],
                SUM([rmr0015].[market_value]) [total_asset_val]
            FROM [rmr0015]
            WHERE [rmr0015].[date] = '{t1_date}'
            GROUP BY [sub_account]
        ),
        [rln06] AS (
            SELECT
                [margin_outstanding].[account_code],
                SUM([margin_outstanding].[principal_outstanding]) [principal_outstanding],
                SUM([margin_outstanding].[interest_outstanding]) [interest_outstanding],
                (SUM([principal_outstanding])+SUM([interest_outstanding])+SUM([fee_outstanding])) [total_loan]
            FROM [margin_outstanding]
            WHERE [margin_outstanding].[date] = '{t1_date}'
            AND [margin_outstanding].[type] IN (N'Margin', N'Trả chậm', N'Bảo lãnh')
            GROUP BY [margin_outstanding].[account_code]
        ),
        [table] AS (
            SELECT
                [branch].[branch_name],
                [rln06].[account_code],
                [rln06].[principal_outstanding] [original_loan],
                [rln06].[interest_outstanding] [interest],
                [rln06].[total_loan],
                ISNULL([rmr62].[cash],0) [total_cash],
                ISNULL([rmr62].[margin_value],0) [total_margin_val],
                ISNULL([rmr15].[total_asset_val],0) [total_asset_val]
            FROM [rln06]
            LEFT JOIN [r] ON [r].[account_code] = [rln06].[account_code]
            LEFT JOIN [rmr62] ON [rmr62].[account_code] = [r].[account_code]
            LEFT JOIN [rmr15] ON [rmr15].[sub_account] = [r].[sub_account]
            LEFT JOIN [branch] ON [branch].[branch_id] = [r].[branch_id]
            WHERE [rln06].[account_code] NOT IN {account}
            AND [rln06].[principal_outstanding] <> 0
        ),
        [final_table] AS (
            SELECT
                [table].[branch_name] [location],
                [table].[account_code] [custody],
                [table].[original_loan],
                [table].[interest],
                [table].[total_loan],
                [table].[total_cash],
                [table].[total_margin_val],
                [table].[total_asset_val],
                CASE
                    WHEN ([table].[total_cash]-[table].[total_loan]) > 0 THEN 100
                    ELSE
                        CASE WHEN ([table].[total_loan]-[table].[total_cash]) > 0
                        THEN ISNULL(ROUND(([table].[total_margin_val] - ([table].[total_loan]-[table].[total_cash])) / NULLIF([table].[total_margin_val],0)*100,2),0)
                        ELSE ISNULL(ROUND(([table].[total_margin_val] - 0) / NULLIF([table].[total_margin_val],0)*100,2),0)
                    END
                END [MMR_MarginableAsset],
                CASE
                    WHEN ([table].[total_cash]-[table].[total_loan]) > 0 THEN 100
                    ELSE
                        CASE WHEN ([table].[total_loan]-[table].[total_cash]) > 0
                        THEN ISNULL(ROUND(([table].[total_asset_val] - ([table].[total_loan]-[table].[total_cash])) / NULLIF([table].[total_asset_val],0)*100,2),0)
                        ELSE ISNULL(ROUND(([table].[total_asset_val] - 0) / NULLIF([table].[total_asset_val],0)*100,2),0)
                    END
                END [MMR_TotalAsset]
            FROM [table]
        )
        SELECT
            [final_table].[location],
            [final_table].[custody],
            [final_table].[original_loan],
            [final_table].[interest],
            [final_table].[total_loan],
            [final_table].[total_cash],
            [final_table].[total_margin_val],
            [final_table].[total_asset_val],
            [final_table].[MMR_MarginableAsset],
            [final_table].[MMR_TotalAsset],
            CASE
                WHEN 0 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 10
                THEN '0<Market Pressure<10%'
                WHEN 10 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 15
                THEN '10%<=Market Pressure<15%'
                WHEN 15 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 20
                THEN '15%<=Market Pressure<20%'
                WHEN 20 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 25
                THEN '20%<=Market Pressure<25%'
                WHEN 25 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 30
                THEN '25%<=Market Pressure<30%'
                WHEN 30 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 35
                THEN '30%<=Market Pressure<35%'
                WHEN 35 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 40
                THEN '35%<=Market Pressure<40%'
                WHEN 40 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 45
                THEN '40%<=Market Pressure<45%'
                WHEN 45 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 50
                THEN '45%<=Market Pressure<50%'
                WHEN 50 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 55
                THEN '50%<=Market Pressure<55%'
                WHEN 55 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 60
                THEN '55%<=Market Pressure<60%'
                WHEN 60 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 65
                THEN '60%<=Market Pressure<65%'
                WHEN 65 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 70
                THEN '65%<=Market Pressure<70%'
                WHEN 70 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 75
                THEN '70%<=Market Pressure<75%'
                WHEN 75 <= [final_table].[MMR_TotalAsset] AND [final_table].[MMR_TotalAsset] < 80
                THEN '75%<=Market Pressure<80%'
                ELSE
                    'Market Pressure>=80%'
            END [condition]
        FROM [final_table]
        """,
        connect_DWH_CoSo
    )
    c_s = table_sheet_2.groupby('condition')['original_loan'].agg(['count', 'sum'])
    c_s['sum'] = c_s['sum'] / 1000000
    criteria = {
        'criteria': [
            '0<Market Pressure<10%',
            '10%<=Market Pressure<15%',
            '15%<=Market Pressure<20%',
            '20%<=Market Pressure<25%',
            '25%<=Market Pressure<30%',
            '30%<=Market Pressure<35%',
            '35%<=Market Pressure<40%',
            '40%<=Market Pressure<45%',
            '45%<=Market Pressure<50%',
            '50%<=Market Pressure<55%',
            '55%<=Market Pressure<60%',
            '60%<=Market Pressure<65%',
            '65%<=Market Pressure<70%',
            '70%<=Market Pressure<75%',
            '75%<=Market Pressure<80%',
            'Market Pressure>=80%'
        ]
    }
    table_sheet_1 = pd.DataFrame(criteria)
    table_sheet_1 = table_sheet_1.merge(c_s, how='outer', left_on='criteria', right_index=True).fillna(0)
    table_sheet_1['%TotalOutstanding'] = table_sheet_1['sum'] / table_sheet_1['sum'].sum() * 100

    ###################################################
    ###################################################
    ###################################################

    t0_day = t0_date[8:10]
    t0_month = int(t0_date[5:7])
    t0_month = calendar.month_name[t0_month]
    t0_year = t0_date[0:4]
    file_name = f'RMD_Market Pressure _end of {t0_day}.{t0_month} {t0_year}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder, folder_name, period, file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # Sheet Summary
    # Format
    cell_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 12,
            'font_name': 'Calibri'
        }
    )
    title1_red_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vbottom',
            'font_size': 12,
            'font_name': 'Calibri',
            'color': '#FF0000'
        }
    )
    title_2_format = workbook.add_format(
        {
            'bold': True,
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    title_2_color_format = workbook.add_format(
        {
            'bold': True,
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'color': '#FF0000'
        }
    )
    title_3_format = workbook.add_format(
        {
            'bold': True,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    headers_format = workbook.add_format(
        {
            'bold': True,
            'text_wrap': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    text_left_merge_format = workbook.add_format(
        {
            'border': 1,
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
    text_left_color_format = workbook.add_format(
        {
            'border': 1,
            'bold': True,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Calibri',
            'color': '#FF0000'
        }
    )
    num_right_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    sum_num_format = workbook.add_format(
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
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    money_2_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '0.00'
        }
    )
    percent_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '0.00'
        }
    )
    # WRITE EXCEL
    headers = [
        'Criteria',
        'No of accounts',
        'Outstanding',
        '% Total Oustanding'
    ]
    title_2 = f'Data is as at end {t0_day}.{t0_month} {t0_year} (it is not inculde 08 accounts that belong to Accumulated Negative Value)'
    title_3 = 'C. Market Pressure (%) is used to indicate the breakeven point of loan with assumption that whole portfolio may drop at same percentage.'

    summary_sheet = workbook.add_worksheet('Summary')
    summary_sheet.set_column('A:A', 31)
    summary_sheet.set_column('B:B', 15)
    summary_sheet.set_column('C:C', 14)
    summary_sheet.set_column('D:D', 11)
    summary_sheet.set_column('E:E', 19)
    summary_sheet.set_column('F:F', 21)
    summary_sheet.set_column('G:G', 0)

    summary_sheet.set_row(5, 30)
    summary_sheet.merge_range('A1:I1', "", cell_format)
    summary_sheet.write_rich_string(
        'A1', 'SUMMARY RISK REPORT FOR ', title1_red_format, 'Market Pressure (%)', cell_format
    )
    summary_sheet.merge_range('A2:F2', title_2, title_2_format)
    summary_sheet.merge_range('A3:F3', "", cell_format)
    summary_sheet.write_rich_string(
        'A3', 'Unit for Outstanding: ', title_2_color_format, 'million dong', title_2_format
    )
    summary_sheet.write('A4', title_3, title_3_format)
    summary_sheet.write_row('A6', headers, headers_format)
    summary_sheet.merge_range('A7:A8', '0 < Market Pressure < 10%', text_left_merge_format)
    summary_sheet.write_rich_string('A9', '10%<= Market Pressure', text_left_color_format, ' < 15%', text_left_format)
    summary_sheet.write_rich_string('A10', '15%<= Market Pressure', text_left_color_format, ' < 20%', text_left_format)
    summary_sheet.write_rich_string('A11', '20%<= Market Pressure', text_left_color_format, ' < 25%', text_left_format)
    summary_sheet.write_rich_string('A12', '25%<= Market Pressure', text_left_color_format, ' < 30%', text_left_format)
    summary_sheet.write_rich_string('A13', '30%<= Market Pressure', text_left_color_format, ' < 35%', text_left_format)
    summary_sheet.write_rich_string('A14', '35%<= Market Pressure', text_left_color_format, ' < 40%', text_left_format)
    summary_sheet.write_rich_string('A15', '40%<= Market Pressure', text_left_color_format, ' < 45%', text_left_format)
    summary_sheet.write_rich_string('A16', '45%<= Market Pressure', text_left_color_format, ' < 50%', text_left_format)
    summary_sheet.write_rich_string('A17', '50%<= Market Pressure', text_left_color_format, ' < 55%', text_left_format)
    summary_sheet.write_rich_string('A18', '55%<= Market Pressure', text_left_color_format, ' < 60%', text_left_format)
    summary_sheet.write_rich_string('A19', '60%<= Market Pressure', text_left_color_format, ' < 65%', text_left_format)
    summary_sheet.write_rich_string('A20', '65%<= Market Pressure', text_left_color_format, ' < 70%', text_left_format)
    summary_sheet.write_rich_string('A21', '70%<= Market Pressure', text_left_color_format, ' < 75%', text_left_format)
    summary_sheet.write_rich_string('A22', '75%<= Market Pressure', text_left_color_format, ' < 80%', text_left_format)
    summary_sheet.write_rich_string('A23', 'Market Pressure', text_left_color_format, ' >= 80%', text_left_format)

    summary_sheet.merge_range('B7:B8', table_sheet_1['count'][0], num_right_format)
    # summary_sheet.merge_range('C7:C8',table_sheet_1['sum'][0],money_2_format)
    summary_sheet.merge_range('D7:D8', table_sheet_1['%TotalOutstanding'][0], percent_format)
    summary_sheet.write_column('B9', table_sheet_1['count'][1:], num_right_format)
    # for a,b in enumerate(table_sheet_1.loc[table_sheet_1.index[1:],'sum']):
    for a, b in enumerate(table_sheet_1['sum']):
        if b > 100 or b == 0:
            fmt = money_format
        else:
            fmt = money_2_format
        if a == 0:
            summary_sheet.merge_range('C7:C8', b, fmt)
        else:
            summary_sheet.write(7 + a, 2, b, fmt)

    summary_sheet.write_column('D9', table_sheet_1['%TotalOutstanding'][1:], percent_format)
    sum_row = table_sheet_1.shape[0] + 8
    summary_sheet.write(f'A{sum_row}', 'Total', headers_format)
    summary_sheet.write(f'B{sum_row}', table_sheet_1['count'].sum(), sum_num_format)
    summary_sheet.write(f'C{sum_row}', table_sheet_1['sum'].sum(), sum_num_format)
    summary_sheet.write(f'D{sum_row}', '', sum_num_format)

    ###################################################
    ###################################################
    ###################################################

    # Sheet Detail
    # Format
    sum_num_color_format = workbook.add_format(
        {
            'bold': True,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '#,##0',
            'color': '#FF0000'
        }
    )
    sum_num_format = workbook.add_format(
        {
            'bold': True,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '#,##0',
        }
    )
    header_1_format = workbook.add_format(
        {
            'bold': True,
            'text_wrap': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'bg_color': '#FFC000'
        }
    )
    header_2_format = workbook.add_format(
        {
            'bold': True,
            'text_wrap': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'bg_color': '#FFF2CC'
        }
    )
    header_3_format = workbook.add_format(
        {
            'bold': True,
            'text_wrap': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'color': '#FF0000',
            'bg_color': '#FFF2CC'
        }
    )
    header_4_format = workbook.add_format(
        {
            'bold': True,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    text_center_format = workbook.add_format(
        {
            'border': 1,
            'text_wrap': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    text_left_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman'
        }
    )
    money_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    percent_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vbottom',
            'font_size': 11,
            'font_name': 'Times New Roman',
            'num_format': '0.00'
        }
    )
    # WRITE EXCEL
    headers_1 = [
        'No.',
        'Location',
        'Custody',
        'Original Loan',
        'interest',
        'Total Loan',
    ]
    headers_2 = [
        'Total Cash & PIA (MR0062 có vay xuất cuối ngày làm việc)',
        'Total Margin value (RMR0062)',
        'Total Asset Value (RMR0015 with market price)'
    ]
    headers_3 = [
        'MMR (base on Marginable Asset)',
        'MMR (base on Total Asset)'
    ]

    worksheet = workbook.add_worksheet('Detail')

    worksheet.set_column('A:A', 0)
    worksheet.set_column('B:B', 5.5)
    worksheet.set_column('C:C', 11.5)
    worksheet.set_column('D:D', 17)
    worksheet.set_column('E:G', 19)
    worksheet.set_column('H:K', 0)
    worksheet.set_column('L:L', 14)
    worksheet.set_column('M:M', 16)
    worksheet.set_column('N:N', 0)
    worksheet.set_column('O:O', 9)

    worksheet.set_row(1, 52)
    worksheet.write('A2', 'Bad Loans', header_4_format)
    worksheet.write('B1', table_sheet_2.shape[0], sum_num_color_format)
    worksheet.write('E1', table_sheet_2['original_loan'].sum(), sum_num_color_format)
    worksheet.write('F1', table_sheet_2['original_loan'].sum() / pow(10, 6), sum_num_format)
    worksheet.write_row('B2', headers_1, header_1_format)
    worksheet.write_row('H2', headers_2, header_2_format)
    worksheet.write_row('K2', headers_3, header_3_format)
    worksheet.write('M2', 'Group/deal', header_2_format)
    worksheet.write('N2', 'bad loan', header_4_format)
    worksheet.write('O2', 'Note', header_4_format)
    worksheet.write_column('B3', np.arange(table_sheet_2.shape[0]) + 1, text_center_format)
    worksheet.write_column('C3', table_sheet_2['location'], text_center_format)
    worksheet.write_column('D3', table_sheet_2['custody'], text_left_format)
    worksheet.write_column('E3', table_sheet_2['original_loan'], money_format)
    worksheet.write_column('F3', table_sheet_2['interest'], money_format)
    worksheet.write_column('G3', table_sheet_2['total_loan'], money_format)
    worksheet.write_column('H3', table_sheet_2['total_cash'], money_format)
    worksheet.write_column('I3', table_sheet_2['total_margin_val'], money_format)
    worksheet.write_column('J3', table_sheet_2['total_asset_val'], money_format)
    worksheet.write_column('K3', table_sheet_2['MMR_MarginableAsset'], percent_format)
    worksheet.write_column('L3', table_sheet_2['MMR_TotalAsset'], percent_format)
    worksheet.write_column('M3', [0] * table_sheet_2.shape[0], money_format)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')

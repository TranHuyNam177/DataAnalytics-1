from automation.accounting import *


def run(
    run_time=dt.datetime.now()
):
    start = time.time()
    t_date = '2022.01.25'
    excel_path = join(dirname(__file__), 'file', f'Sổ tổng hợp công nợ 1231_{t_date}.xlsx')
    df = pd.read_excel(
        excel_path,
        sheet_name='SỔ TỔNG HỢP CÔNG NỢ',
        skiprows=8
    )
    df = df.iloc[:-1]
    df = df.rename(columns={
        '1': 'account_code',
        '2': 'customer_name',
        '3': 'du_no_dau',
        '4': 'du_co_dau',
        '5': 'ps_no',
        '6': 'ps_co',
        '7': 'du_no_cuoi',
        '8': 'du_co_cuoi'
    })
    query_rln06 = pd.read_sql(
        f"""
        SELECT
            [r].[account_code],
            SUM([r].[principal_outstanding]) [flex]
        FROM [margin_outstanding] [r]
        WHERE [r].[date] = '{t_date}'
        AND  [r].[type] <> N'Ứng trước cổ tức'
        GROUP BY [r].[account_code]
        """,
        connect_DWH_CoSo
    )
    table = df.merge(query_rln06, on='account_code', how='left')

    table['customer_name'] = table['customer_name'].fillna('')
    table['flex'] = table['flex'].fillna(0)

    table = table.sort_values('account_code', ignore_index=True)
    table['CL'] = table['du_no_cuoi'] - table['flex']

    ################################################################
    ################################################################
    ################################################################

    # WRITE SHEET 1231
    reportDay = t_date[-2:]
    reportMonth = t_date[5:7]
    reportYear = t_date[:4]
    file_name = f'1.1 DOI CHIEU CO SO - {reportDay}.{reportMonth}.{reportYear}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder, file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook = writer.book

    # Set Format
    info_format = workbook.add_format({
        'bold': True,
        'align': 'left',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Times New Roman'
    })
    title_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 11,
        'font_name': 'Times New Roman',
    })
    sub_title_1_format = workbook.add_format({
        'bold': True,
        'italic': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Times New Roman'
    })
    sub_title_2_format = workbook.add_format({
        'bold': True,
        'align': 'center',
        'valign': 'top',
        'font_size': 10,
        'font_name': 'Times New Roman'
    })
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'border': True,
        'align': 'center',
        'valign': 'vcenter',
        'font_size': 10,
        'font_name': 'Times New Roman',
        'bg_color': '#FFFFF0'
    })
    text_left_format = workbook.add_format({
        'align': 'left',
        'valign': 'top',
        'font_size': 10,
        'font_name': 'Times New Roman'
    })
    customer_format = workbook.add_format({
        'text_wrap': True,
        'align': 'left',
        'valign': 'top',
        'font_size': 10,
        'font_name': 'Times New Roman'
    })
    money_format = workbook.add_format({
        'align': 'right',
        'valign': 'top',
        'font_size': 10,
        'font_name': 'Times New Roman',
        'num_format': '#,##0;(#,##0);'
    })
    du_co_cuoi_val_format = workbook.add_format({
        'right': True,
        'align': 'right',
        'valign': 'top',
        'font_size': 10,
        'font_name': 'Times New Roman',
        'num_format': '#,##0;(#,##0); '
    })
    flex_format = workbook.add_format({
        'align': 'right',
        'valign': 'top',
        'font_size': 10,
        'font_name': 'Arial',
        'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
    })
    tong_format = workbook.add_format({
        'bold': True,
        'bottom': True,
        'align': 'left',
        'valign': 'top',
        'font_size': 10,
        'font_name': 'Times New Roman',
        'bg_color': '#FFFFD2'
    })
    sum_format = workbook.add_format({
        'bold': True,
        'bottom': True,
        'align': 'right',
        'valign': 'top',
        'font_size': 10,
        'font_name': 'Times New Roman',
        'num_format': '#,##0;(#,##0); ',
        'bg_color': '#FFFFD2'
    })
    sum2_format = workbook.add_format({
        'bold': True,
        'right': True,
        'bottom': True,
        'align': 'right',
        'valign': 'top',
        'font_size': 10,
        'font_name': 'Times New Roman',
        'num_format': '#,##0;(#,##0); ',
        'bg_color': '#FFFFD2'
    })
    color1_format = workbook.add_format({
        'align': 'left',
        'valign': 'top',
        'font_size': 10,
        'font_name': 'Arial',
        'bg_color': '#FFFF00'
    })
    color2_format = workbook.add_format({
        'align': 'left',
        'valign': 'top',
        'font_size': 10,
        'font_name': 'Arial',
        'bg_color': '#E6B8B7'
    })
    header = [
        'Mã đối tượng',
        'Tên đối tượng',
        'Dư nợ đầu',
        'Dư có đầu',
        'Ps nợ',
        'Ps có',
        'Dư nợ cuối',
        'Dư có cuối',
        'FLEX',
        'CL'
    ]
    worksheet = workbook.add_worksheet('1231')
    # Set Columns & Rows
    worksheet.set_column('A:H', 19)
    worksheet.set_column('I:I', 24)
    worksheet.set_column('J:J', 19)
    # Write
    sub_title_1 = f'Từ ngày {reportDay}/{reportMonth}/{reportYear[-2:]} đến ngày {reportDay}/{reportMonth}/{reportYear[-2:]}'

    worksheet.write('A1', CompanyName, info_format)
    worksheet.write('A2', CompanyAddress, info_format)
    worksheet.merge_range('A3:H3', 'SỔ TỔNG HỢP CÔNG NỢ', title_format)
    worksheet.merge_range('A4:H4', sub_title_1, sub_title_1_format)
    worksheet.merge_range('A5:H5', 'Tài khoản: 1231 - Cho vay hoạt động Margin', sub_title_2_format)
    worksheet.write('I6', 'Dữ liệu từ Flex RLN0006', color1_format)
    worksheet.write('J6', 'Đối chiếu', color2_format)
    worksheet.write_row('A7', header, header_format)
    worksheet.write_column('A8', table['account_code'], text_left_format)
    worksheet.write_column('B8', table['customer_name'], customer_format)
    worksheet.write_column('C8', table['du_no_dau'], money_format)
    worksheet.write_column('D8', table['du_co_dau'], money_format)
    worksheet.write_column('E8', table['ps_no'], money_format)
    worksheet.write_column('F8', table['ps_co'], money_format)
    worksheet.write_column('G8', table['du_no_cuoi'], money_format)
    worksheet.write_column('H8', table['du_co_cuoi'], du_co_cuoi_val_format)
    worksheet.write_column('I8', table['flex'], flex_format)
    worksheet.write_column('J8', table['CL'], money_format)

    sum_start_row = table.shape[0] + 8
    worksheet.write(f'A{sum_start_row}', '', tong_format)
    worksheet.write(f'B{sum_start_row}', 'Tổng cộng:', tong_format)
    for col in 'CDEFGHIJ':
        if col == 'H':
            worksheet.write(f'{col}{sum_start_row}',f'=SUBTOTAL(9,{col}8:{col}{sum_start_row - 1})',sum2_format)
        else:
            worksheet.write(f'{col}{sum_start_row}',f'=SUBTOTAL(9,{col}8:{col}{sum_start_row-1})',sum_format)

    ################################################################
    ################################################################
    ################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')



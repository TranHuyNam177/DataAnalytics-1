from automation.accounting import *


def run(
    run_time=dt.datetime.now()
):
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

    col_name = ['SoTaiKhoan','TenKhachHang','DuNoDauBravo','DuCoDauBravo',
                'PhatSinhNoBravo','PhatSinhCoBravo','DuNoCuoiBravo','DuCoCuoiBravo']

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

    ################################################################
    ################################################################
    ################################################################

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

    ################################################################
    ################################################################
    ################################################################

    # WRITE SHEET 1231
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
    worksheet.write_column('A8', table_1231['SoTaiKhoan'], text_left_format)
    worksheet.write_column('B8', table_1231['TenKhachHang'], customer_format)
    worksheet.write_column('C8', table_1231['DuNoDauBravo'], money_format)
    worksheet.write_column('D8', table_1231['DuCoDauBravo'], money_format)
    worksheet.write_column('E8', table_1231['PhatSinhNoBravo'], money_format)
    worksheet.write_column('F8', table_1231['PhatSinhCoBravo'], money_format)
    worksheet.write_column('G8', table_1231['DuNoCuoiBravo'], money_format)
    worksheet.write_column('H8', table_1231['DuCoCuoiBravo'], du_co_cuoi_val_format)
    worksheet.write_column('I8', table_1231['principal_outstanding'], flex_format)
    worksheet.write_column('J8', table_1231['CL'], money_format)

    sum_start_row = table_1231.shape[0] + 8
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

    # WRITE SHEET 13226
    worksheet = workbook.add_worksheet('13226')
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
    worksheet.merge_range('A5:H5', 'Tài khoản: 13226 - Dự thu tiền lãi - Nghiệp vụ Margin', sub_title_2_format)
    worksheet.write('I6', 'Dữ liệu từ Flex RLN0006', color1_format)
    worksheet.write('J6', 'Đối chiếu', color2_format)
    worksheet.write_row('A7', header, header_format)
    worksheet.write_column('A8', table_13226['SoTaiKhoan'], text_left_format)
    worksheet.write_column('B8', table_13226['TenKhachHang'], customer_format)
    worksheet.write_column('C8', table_13226['DuNoDauBravo'], money_format)
    worksheet.write_column('D8', table_13226['DuCoDauBravo'], money_format)
    worksheet.write_column('E8', table_13226['PhatSinhNoBravo'], money_format)
    worksheet.write_column('F8', table_13226['PhatSinhCoBravo'], money_format)
    worksheet.write_column('G8', table_13226['DuNoCuoiBravo'], money_format)
    worksheet.write_column('H8', table_13226['DuCoCuoiBravo'], du_co_cuoi_val_format)
    worksheet.write_column('I8', table_13226['interest_outstanding'], flex_format)
    worksheet.write_column('J8', table_13226['CL'], money_format)

    sum_start_row = table_13226.shape[0] + 8
    worksheet.write(f'A{sum_start_row}', '', tong_format)
    worksheet.write(f'B{sum_start_row}', 'Tổng cộng:', tong_format)
    for col in 'CDEFGHIJ':
        if col == 'H':
            worksheet.write(f'{col}{sum_start_row}', f'=SUBTOTAL(9,{col}8:{col}{sum_start_row - 1})', sum2_format)
        else:
            worksheet.write(f'{col}{sum_start_row}', f'=SUBTOTAL(9,{col}8:{col}{sum_start_row - 1})', sum_format)

    ################################################################
    ################################################################
    ################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')



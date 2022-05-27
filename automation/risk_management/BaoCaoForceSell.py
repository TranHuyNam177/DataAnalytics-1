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

    table = pd.read_sql(
        f"""
        WITH [r] AS (
            SELECT
                [relationship].[date],
                [relationship].[sub_account],
                [relationship].[account_code],
                [relationship].[branch_id],
                [relationship].[broker_id],
                [branch].[branch_name],
                [broker].[broker_name],
                [account].[customer_name]
            FROM [relationship]
            LEFT JOIN [branch] ON [branch].[branch_id] = [relationship].[branch_id]
            LEFT JOIN [broker] ON [broker].[branch_id] = [relationship].[broker_id]
            LEFT JOIN [account] ON [account].[account_code] = [relationship].[account_code]
        ),
        [rln06] AS (
            SELECT
                [date],
                [account_code],
                ([principal_outstanding] + [interest_outstanding] + [fee_outstanding]) [total_outstanding]
            FROM [margin_outstanding]
            WHERE [margin_outstanding].[type] = N'Trả chậm'
        ),
        [vmr02] AS (
            SELECT
                [VMR0002].[Ngay],
                [VMR0002].[TieuKhoan],
                [VMR0002].[DTMG],
                [VMR0002].[EmailMG]
            FROM [VMR0002]
        )
        SELECT
            [r].[branch_name],
            [r].[account_code],
            '' AS [tinh_trang_force_sell],
            '' AS [xu_ly_force_sell],
            [r].[customer_name],
            [vmr0003].[actual_mr_ratio] [TL_ThucTe_Mr],
            [vmr0003].[actual_dp_ratio] [TL_ThucTe_TC],
            CASE 
                WHEN [vmr0003].[date_of_first_call] IS NULL THEN '' 
                ELSE CONVERT(varchar(50), [vmr0003].[date_of_first_call], 103) 
            END [date_of_first_call],
            [vmr0003].[selling_value] [TienMat_ve_100],
            [vmr0003].[executing_amount] [TienMat_ve_85],
            CASE
                WHEN ([vmr0003].[actual_asset_to_guarantee] - [vmr0003].[converted_asset]) > 0
                THEN [vmr0003].[actual_asset_to_guarantee] - [vmr0003].[converted_asset]
                ELSE 0
            END [TienMat_ve_Rtt_DP],
            [vmr0003].[guarantee_debt] [NoHanMuc],
            [vmr0003].[mr_dp_overdue_amount] [NoMRTC_QuaHan],
            [r].[broker_name],
            [vmr0003].[mr_dp_due_amount] [No_MRTC_DenHan],
            [vmr0003].[contract_type] [maLoaiHinh],
            [vmr0003].[total_cash],
            ISNULL([rln06].[total_outstanding], 0) [DP],
            [vmr0003].[actual_ratio] [TL_ThucTe],
            [vmr0003].[actual_t0_ratio] [TL_ThucTe_T0],
            [vmr0003].[depository_fee_debt] [no_phi_LK],
            0 AS [ChoBan],
            CASE 
                WHEN [vmr0003].[date_of_last_sms] IS NULL THEN '' 
                ELSE CONVERT(varchar(50), [vmr0003].[date_of_last_sms], 103) 
            END [date_of_last_sms],
            CASE 
                WHEN [vmr0003].[time_of_last_sms] IS NULL THEN '' 
                ELSE CONVERT(varchar(50), [vmr0003].[time_of_last_sms], 108) 
            END [time_of_last_sms],
            [vmr0003].[days_maintain_call],
            [vmr0003].[days_warning],
            CASE 
                WHEN [vmr0003].[date_of_last_call] IS NULL THEN '' 
                ELSE CONVERT(varchar(50), [vmr0003].[date_of_last_call], 103) 
            END [date_of_last_call],
            CASE 
                WHEN [vmr0003].[date_of_trigger] IS NULL THEN '' 
                ELSE CONVERT(varchar(50), [vmr0003].[date_of_trigger], 103) 
            END [date_of_trigger],
            [vmr0003].[additional_deposit_amount],
            [vmr0003].[selling_amount],
            [vmr0003].[executing_overdue_amount],
            [vmr0003].[remain_cash_after_sell],
            [vmr0003].[days_execution],
            [vmr0003].[sub_account],
            [vmr0003].[sub_account_type] [tenLoaiHinh],
            '' AS [DT_LienLac],
            '' AS [email],
            [vmr0003].[careby] [DiemHoTro],
            [vmr02].[EmailMG] [emailMG],
            [vmr02].[DTMG] AS [DTMG],
            [r].[branch_id],
            [vmr0003].[force_sell_ordered_value] [TongGTLenhGiaiChap],
            [vmr0003].[force_sell_matched_value] [TongGTKhopBanGiaiChap],
            [vmr0003].[rate_add],
            [vmr0003].[days_add],
            [vmr0003].[right_event],
            [vmr0003].[days_base_call],
            [vmr0003].[safe_rate],
            [vmr0003].[total_loan_amount] [TongDuNoVay],
            [vmr0003].[converted_asset] [TaiSanVayQuiDoi],
            [vmr0003].[actual_asset_to_guarantee] [TSThucBaoDamTKLQ_DuyTri]
        FROM [vmr0003]
        LEFT JOIN [r]
        ON [r].[sub_account] = [vmr0003].[sub_account] AND [r].[date] = [vmr0003].[date]
        LEFT JOIN [vmr02]
        ON [vmr02].[TieuKhoan] = [vmr0003].[sub_account] AND [vmr02].[Ngay] = [vmr0003].[date]
        LEFT JOIN [rln06]
        ON [rln06].[date] = [vmr0003].[date] AND [rln06].[account_code] = [r].[account_code]
        WHERE [vmr0003].[date] = '{t1_date}'
        """,
        connect_DWH_CoSo
    )
    table.loc[table['right_event'] == False, 'right_event'] = ''

    # list các tài khoản cố định (chị Thu Anh dặn)
    acc_lst = [
        '022P002222', '022C006827', '022C012621', '022C012620', '022C012622',
        '022C089535', '022C050302', '022C089950', '022C089957'
    ]
    # loc ra các tài khoản trong list tài khoản cố định
    account_loc = table.loc[table['account_code'].isin(acc_lst)].sort_values('account_code', ascending=False)
    # loc ra các tài khoản không nằm trong list tài khoản cố định
    table_loc = table.loc[~table['account_code'].isin(acc_lst)].sort_values('account_code', ascending=True)

    ###################################################
    ###################################################
    ###################################################

    t0_day = t0_date[8:10]
    t0_month = t0_date[5:7]
    t0_year = t0_date[0:4]
    file_date = t0_day + t0_month + t0_year
    file_name = f'Force sell {file_date}.xlsx'
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
    headers_format = workbook.add_format(
        {
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'text_wrap': True
        }
    )
    text_left_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'top',
            'font_size': 11,
            'font_name': 'Calibri'
        }
    )
    money_format = workbook.add_format(
        {
            'align': 'right',
            'valign': 'top',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '#,##0'
        }
    )
    number_format = workbook.add_format(
        {
            'align': 'right',
            'valign': 'top',
            'font_size': 11,
            'font_name': 'Calibri',
        }
    )
    date_format = workbook.add_format(
        {
            'align': 'left',
            'valign': 'top',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': 'dd/mm/yyyy'
        }
    )

    ###################################################
    ###################################################
    ###################################################

    # WRITE EXCEL
    headers = [
        'Tên chi nhánh', 'Số TK lưu ký', 'Tình Trạng Force sell', 'Xử lý force sell', 'Tên khách hàng', 'TL thực tế MR',
        'TL thực tế TC', 'Ngày bắt đầu Call', 'Tiền mặt nộp về 100%', 'Tiền mặt nộp về 85%', 'Tiền mặt nộp về Rtt_DP',
        'Nợ hạn mức', 'Nợ MR + TC quá hạn', 'Tên MG', 'Nợ MR + TC đến hạn', 'Mã loại hình', 'Tổng tiền', 'DP',
        'TL thực tế',
        'TL thực tế T0', 'Nợ phí LK', 'Chờ bán', 'Ngày SMS cuối', 'Thời gian SMS cuối', 'Số ngày duy trì call',
        'Số ngày rơi vào cảnh báo', 'Ngày bắt đầu call', 'Ngày bắt đầu trigger', 'Số tiền nộp thêm',
        'Số tiền cần phải bán',
        'Số tiền quá hạn cần phải xử lý', 'Số tiền dư sau bán', 'Số ngày rơi vào xử lý', 'Số tiểu khoản',
        'Tên loại hình',
        'ĐT liên lạc', 'E-mail', 'Điểm hỗ trợ', 'E-mail MG', 'ĐT MG', 'Code chi nhánh', 'Tổng GT lệnh giải chấp',
        'Tổng GT khớp bán giải chấp', 'Tỷ lệ bổ sung', 'Số ngày cộng bổ sung', 'Chạm sự kiện quyền',
        'Số ngày call cơ sở',
        'TL an toàn', 'Tổng dư nợ vay', 'Tài sản vay qui đổi', 'TS thực có tối thiểu để bảo đảm TLKQ duy trì'
    ]

    worksheet = workbook.add_worksheet('Sheet1')

    worksheet.set_column('A:A', 8.5)
    worksheet.set_column('B:B', 10)
    worksheet.set_column('C:D', 8)
    worksheet.set_column('E:E', 28)
    worksheet.set_column('F:G', 8.5)
    worksheet.set_column('H:H', 10)
    worksheet.set_column('I:M', 13)
    worksheet.set_column('N:N', 25)
    worksheet.set_column('O:P', 8)
    worksheet.set_column('Q:R', 11.5)
    worksheet.set_column('S:U', 8.5)
    worksheet.set_column('V:V', 8)
    worksheet.set_column('W:W', 10)
    worksheet.set_column('X:Z', 8)
    worksheet.set_column('AA:AB', 10)
    worksheet.set_column('AC:AE', 13)
    worksheet.set_column('AF:AV', 8)
    worksheet.set_column('AW:AY', 13)
    worksheet.set_row(0, 54)

    worksheet.write_row('A1', headers, headers_format)

    for colNum, colName in enumerate(table_loc.columns):
        if colName.lower().startswith('tl') or colName.lower().startswith('safe'):
            fmt = number_format
        elif pd.api.types.is_numeric_dtype(table[colName]):
            fmt = money_format
        elif pd.api.types.is_datetime64_dtype(table[colName]):
            fmt = date_format
        else:
            fmt = text_left_format
        worksheet.write_column(1, colNum, table_loc[colName], fmt)

    account_lst_row = table_loc.shape[0] + 2

    for colNum, colName in enumerate(account_loc.columns):
        if colName.lower().startswith('tl') or colName.lower().startswith('safe'):
            fmt = number_format
        elif pd.api.types.is_numeric_dtype(table[colName]):
            fmt = money_format
        elif pd.api.types.is_datetime64_dtype(table[colName]):
            fmt = date_format
        else:
            fmt = text_left_format
        worksheet.write_column(account_lst_row, colNum, account_loc[colName], fmt)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__ == '__main__':
        print(f"{__file__.split('/')[-1].replace('.py', '')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time() - start, 1)}s')

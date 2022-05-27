from automation.risk_management import *

"""
TK bị lệch ngày 29/04/2022
1. 022C017132	0117000598	ĐINH TIÊN HOÀNG
- Trạng thái call trên DB là No - Trạng thái call từ báo cáo bên QLRR là Yes
2. 022C018098	0203001688	NGUYỄN KIM NHUNG
- DB ngày 29/4/2022 không trả ra kết quả
3. 022C024544	0202001405	NGUYỄN THỊ CHĂM
- DB ngày 29/4/2022 không trả ra kết quả
4. 022C027172	0102002141	DƯƠNG THỊ CÁT ĐẰNG
- DB ngày 29/4/2022 không trả ra kết quả
5. 022C042269	0202003304	ĐẶNG NHÂN THỦY
- DB ngày 29/4/2022 không trả ra kết quả
6. 022C042343	0101003903	PHAN THỊ NGỌC NỮ
- DB ngày 29/4/2022 không trả ra kết quả
7. 022C087587	0101087587	VŨ KIM LIÊN
- DB ngày 29/4/2022 không trả ra kết quả
8. 022C240592	0117002925	NGUYỄN VĂN HỒ
- Trạng thái call trên DB là No - Trạng thái call từ báo cáo bên QLRR là Yes
9. 022C357999	0202002369	LÊ CHÍ CƯỜNG
- DB ngày 29/4/2022 không trả ra kết quả
10. 022C567803	0117002907	NGUYỄN ĐÌNH TƯ
- Trạng thái call trên DB là No - Trạng thái call từ báo cáo bên QLRR là Yes
11. 022C777999	0201001455	NGUYỄN XUÂN CỬ
- DB ngày 29/4/2022 không trả ra kết quả
"""


def run(  # chạy hàng ngày
    run_time=None
):
    start=time.time()
    info=get_info('daily',run_time)
    period=info['period']
    t0_date=info['end_date']
    folder_name=info['folder_name']

    # create folder
    if not os.path.isdir(join(dept_folder,folder_name,period)):
        os.mkdir((join(dept_folder,folder_name,period)))

    ###################################################
    ###################################################
    ###################################################

    table = pd.read_sql(
        f"""
        WITH [r] AS (
            SELECT 
                [date],
                [account_code],
                [broker_id],
                [sub_account]
            FROM [relationship]
            WHERE [relationship].[date] = '{t0_date}'
        ),
        [a] AS (
            SELECT
                [account_code],
                [customer_name]
            FROM [account]
        ),
        [b] AS (
            SELECT
                [broker_id],
                [broker_name]
            FROM [broker]
        )
        SELECT 
            [a].[account_code],
            [VMR0002].[TieuKhoan],
            [a].[customer_name],
            [VMR0002].[MaLoaiHinh],
            [VMR0002].[TenLoaiHinh],
            [b].[broker_name],
            [VMR0002].[SoNgayDuyTriCall],
            [VMR0002].[NgayBatDauCall],
            [VMR0002].[NgayHanCuoiCall],
            [VMR0002].[TrangThaiCall],
            [VMR0002].[LoaiCall],
            [VMR0002].[SoNgayCallVuot],
            [VMR0002].[SoTienPhaiNop],
            [VMR0002].[SoTienPhaiBan],
            [VMR0002].[SoTienDenHanVaQuaHan],
            [VMR0002].[SoTienNopThemGoc],
            [VMR0002].[ChamSuKienQuyen],
            [VMR0002].[TLThucTe],
            [VMR0002].[TLThucTeMR],
            [VMR0002].[TLThucTeTC],
            [VMR0002].[TyLeAnToan],
            [VMR0002].[ToDuNoVay],
            [VMR0002].[NoMRTCBL],
            [VMR0002].[TaiSanVayQuiDoi],
            [VMR0002].[TSThucCoToiThieuDeBaoDamTLKQDuyTri],
            [VMR0002].[ThieuHut],
            '' AS [DTLL],
            '' AS [email],
            [VMR0002].[EmailMG],
            [VMR0002].[DTMG]
        FROM [VMR0002]
        LEFT JOIN [r] ON [r].[sub_account] = [VMR0002].[TieuKhoan] AND [r].[date] = [VMR0002].[Ngay]
        LEFT JOIN [a] ON [a].[account_code] = [r].[account_code]
        LEFT JOIN [b] ON [b].[broker_id] = [r].[broker_id]
        WHERE 
            [VMR0002].[Ngay] = '{t0_date}'
            AND [VMR0002].[TenLoaiHinh] = N'Margin'
            AND [VMR0002].[TrangThaiCall] = N'Yes'
            AND [VMR0002].[LoaiCall] <> ''
            AND (
                [VMR0002].[NgayBatDauCall] IS NOT NULL
                AND [VMR0002].[NgayHanCuoiCall] IS NOT NULL
            )
        ORDER BY [account_code]
        """,
        connect_DWH_CoSo
    )

    ###################################################
    ###################################################
    ###################################################

    t0_day = t0_date[8:10]
    t0_month = int(t0_date[5:7])
    t0_month = calendar.month_name[t0_month]
    t0_year = t0_date[0:4]
    file_date = t0_month + ' ' + t0_day + ' ' + t0_year
    file_name = f'Call Margin Report on {file_date}.xlsx'
    writer = pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options':{'nan_inf_to_errors':True}}
    )
    workbook = writer.book

    ###################################################
    ###################################################
    ###################################################

    # Format
    headers_format = workbook.add_format(
        {
            'border':1,
            'bold':True,
            'align':'center',
            'valign':'top',
            'font_size':12,
            'font_name':'Calibri',
            'text_wrap':True
        }
    )
    text_left_format = workbook.add_format(
        {
            'border':1,
            'align':'left',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Calibri'
        }
    )
    stt_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Calibri'
        }
    )
    money_format = workbook.add_format(
        {
            'border':1,
            'align':'right',
            'valign':'vcenter',
            'font_size':11,
            'font_name':'Calibri',
            'num_format': '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
        }
    )
    decimal_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '#,##0.00'
        }
    )
    number_format = workbook.add_format(
        {
            'border': 1,
            'align': 'right',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format': '#,##0'
        }
    )
    date_format = workbook.add_format(
        {
            'border': 1,
            'align': 'left',
            'valign': 'vcenter',
            'font_size': 11,
            'font_name': 'Calibri',
            'num_format':'dd/mm/yyyy'
        }
    )

    ###################################################
    ###################################################
    ###################################################

    # WRITE EXCEL
    headers = [
        'No','Account No','Số tiểu khoản','Name','Mã loại hình','Tên loại hình','Broker Name','Số ngày duy trì call',
        'Call date','Call deadline','Trạng thái call','Call Type','Số ngày call vượt','Supplementary Amount',
        'Số tiền phải bán','Overdue + Due to date amount','Số tiền nộp thêm gốc','Ex-right','TL thực tế','Rtt-MR',
        'Rtt-DP','Rat','Total Outstanding','Nợ MR + TC + BL','Tài sản vay qui đổi',
        'TS thực có tối thiểu để bảo đảm TLKQ duy trì','Thiếu hụt','ĐT liên lạc','E-mail','E-mail MG','ĐT MG'
    ]

    worksheet = workbook.add_worksheet('Sheet1')
    worksheet.hide_gridlines(option=2)

    worksheet.set_column('A:A',4)
    worksheet.set_column('B:B',16)
    worksheet.set_column('D:D',27)
    worksheet.set_column('G:G',28)
    worksheet.set_column('I:J',13)
    worksheet.set_column('L:L',19)
    worksheet.set_column('N:N',22)
    worksheet.set_column('P:P',18)
    worksheet.set_column('R:R',12)
    worksheet.set_column('T:U',11)
    worksheet.set_column('V:V',8)
    worksheet.set_column('W:W',18)
    worksheet.set_column('AA:AE',0)
    worksheet.set_row(0,37)

    for col in ('CEFHKMOQSXYZ'):
        worksheet.set_column(f'{col}:{col}',0)

    worksheet.write_row('A1',headers,headers_format)
    worksheet.write_column('A2',np.arange(table.shape[0])+1,stt_format)
    for a, b in enumerate(table.columns):
        if table[f'{b}'].dtype=='object':
            fmt = text_left_format
        elif b in ['TLThucTe','TLThucTeMR','TLThucTeTC']:
            fmt = decimal_format
        elif b in ['SoNgayDuyTriCall','SoNgayCallVuot']:
            fmt = number_format
        elif table[f'{b}'].dtype=='int64' or table[f'{b}'].dtype=='float64':
            fmt = money_format
        else:
            fmt = date_format
        worksheet.write_column(1,a+1,table[f'{b}'],fmt)

    ###########################################################################
    ###########################################################################
    ###########################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')


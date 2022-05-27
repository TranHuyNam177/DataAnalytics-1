from automation.risk_management import *
from datawarehouse import *


def initiate(
    run_time=dt.datetime.now()
):
    """
    Hàm này được chạy để khởi tạo file gốc hôm nay khi chạy trước batch giữa ngày (để hôm sau đối chiếu)
    """
    info = get_info('daily',run_time)
    end_date = info['end_date']

    vmr0001 = pd.read_sql(
        f"""
            WITH [r] AS (
                SELECT
                    [relationship].[date],
                    [relationship].[sub_account],
                    [relationship].[account_code],
                    [relationship].[broker_id]
                FROM [relationship]
                WHERE [relationship].[date] = '{end_date}'
            ),
            [a] AS (
                SELECT
                    [account].[account_code],
                    [account].[customer_name]
                FROM [account]
            ),
            [b] AS (
                SELECT
                    [broker].[broker_id],
                    [broker].[broker_name]
                FROM [broker]
            ),
            [vcf51] AS (
                SELECT
                    [vcf0051].[date],
                    [vcf0051].[sub_account],
                    [vcf0051].[contract_type]
                FROM [vcf0051]
            )
            SELECT
                [VMR0001].[MaLoaiHinh],
                [vcf51].[contract_type] [TenLoaiHinh],
                [r].[account_code],
                [VMR0001].[TieuKhoan],
                [a].[customer_name],
                [b].[broker_name],
                [VMR0001].[Tien],
                [VMR0001].[TLMRThucTe],
                [VMR0001].[TLTCThucTe]
            FROM [VMR0001]
            LEFT JOIN [r] 
            ON [r].[sub_account] = [VMR0001].[TieuKhoan] AND [r].[date] = [VMR0001].[Ngay]
            LEFT JOIN [a]
            ON [a].[account_code] = [r].[account_code]
            LEFT JOIN [b]
            ON [b].[broker_id] = [r].[broker_id]
            LEFT JOIN [vcf51]
            ON [vcf51].[sub_account] = [VMR0001].[TieuKhoan] AND [vcf51].[date] = [VMR0001].[Ngay]
            WHERE [VMR0001].[Ngay] = '{end_date}'
            """,
        connect_DWH_CoSo,
        index_col='TieuKhoan'
    )
    vmr9003 = pd.read_sql(
        f"""
            SELECT
                [VMR9003].[TieuKhoan],
                [VMR9003].[DuTinhGiaiNganT0] [VMR9003]
            FROM [VMR9003]
            WHERE [VMR9003].[Ngay] = '{end_date}'
            """,
        connect_DWH_CoSo,
        index_col='TieuKhoan'
    )
    vmr0001.to_pickle(join(dirname(__file__),'pickle_file','vmr0001',f'vmr0001_{end_date.replace(".","")}.pickle'))
    vmr9003.to_pickle(join(dirname(__file__),'pickle_file','vmr9003',f'vmr9003_{end_date.replace(".","")}.pickle'))


def run(
        run_time=None
):
    start=time.time()
    info=get_info('daily', run_time)
    period=info['period']
    t0_date=info['end_date'].replace('.', '-')
    # t1_date=BDATE(t0_date,-1)
    t1_date = t0_date
    t2_date=BDATE(t0_date,-2)
    folder_name=info['folder_name']

    # create_folder
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ################################################################
    ################################################################
    ################################################################

    # Query SQL to getting data
    vmr0001_t1 = pd.read_pickle(
        join(dirname(__file__),'pickle_file','vmr0001',f'vmr0001_{t1_date.replace("-","")}.pickle')
    )
    vmr9003_t1 = pd.read_pickle(
        join(dirname(__file__),'pickle_file','vmr9003',f'vmr9003_{t1_date.replace("-","")}.pickle')
    )
    query_SQL=pd.read_sql(
        f"""
        WITH [sub_acc] AS (
            SELECT
                [sub_account].[sub_account],
                [sub_account].[account_code]
            FROM [sub_account]
        ),
        [v_t2] AS (
            SELECT
                [VMR0001].[Ngay],
                [VMR0001].[TieuKhoan],
                [VMR0001].[TLMRThucTe] [TLMRDN],
                [VMR0001].[TLTCThucTe] [TLTCDN]
            FROM [VMR0001]
            WHERE [VMR0001].[Ngay] = '{t2_date}'
            AND [VMR0001].[TenLoaiHinh] = 'Margin'
        ),
        [rln05] AS (
            SELECT
                [RLN0005].[Ngay],
                [RLN0005].[TieuKhoan],
                [RLN0005].[SoTienCapBaoLanh]
            FROM [RLN0005]
            WHERE [RLN0005].[Ngay] = '{t1_date}'
        ),
        [rln06] AS (
            SELECT
                [margin_outstanding].[date],
                [margin_outstanding].[account_code],
                ([principal_outstanding]+[interest_outstanding]+[fee_outstanding]) [sum_outstanding]
            FROM [margin_outstanding]
            WHERE [margin_outstanding].[date] = '{t1_date}'
            AND [margin_outstanding].[type] = N'Trả chậm'
        )
        SELECT
            [sub_acc].[sub_account],
            [v_t2].[TLMRDN],
            [v_t2].[TLTCDN],
            [rln05].[SoTienCapBaoLanh] [RLN0005],
            [rln06].[sum_outstanding] [RLN0006]
        FROM [sub_acc]
        LEFT JOIN [v_t2] ON [v_t2].[TieuKhoan] = [sub_acc].[sub_account]
        LEFT JOIN [rln05] ON [rln05].[TieuKhoan] = [sub_acc].[sub_account]
        LEFT JOIN [rln06] ON [rln06].[account_code] = [sub_acc].[account_code]
        """,
        connect_DWH_CoSo,
        index_col='sub_account'
    )
    table = vmr0001_t1.join(query_SQL,how='left')
    final_table = table.join(vmr9003_t1,how='left')
    final_table = final_table.reset_index().rename(columns={'index': 'sub_account'})
    final_table = final_table.sort_values('RLN0005',ascending=False,ignore_index=True)
    final_table = final_table.fillna('#N/A')

    ################################################################
    ################################################################
    ################################################################

    ### WRITE EXCEL ###
    t0_day = t0_date[8:10]
    t0_month = t0_date[5:7]
    t0_year = t0_date[0:4]
    eod = t0_day + t0_month + t0_year
    file_name=f'Checking Quota {eod}.xlsx'
    writer=pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook=writer.book

    # Set Format
    header_1_format=workbook.add_format({
        'align':'left',
        'valign':'vbottom',
        'font_size':11,
        'font_name':'Calibri'
    })
    header_2_format = workbook.add_format({
        'bold':True,
        'align':'center',
        'valign':'vcenter',
        'text_wrap': True,
        'font_size': 11,
        'font_name': 'Calibri',
        'bg_color': '#FFFF00'
    })
    text_left_format=workbook.add_format({
        'align':'left',
        'valign':'vbottom',
        'font_size':11,
        'font_name':'Calibri'
    })
    text_center_format = workbook.add_format({
        'align': 'center',
        'valign': 'vbottom',
        'font_size': 11,
        'font_name': 'Calibri'
    })
    number_format = workbook.add_format({
        'align': 'right',
        'valign': 'vbottom',
        'font_size': 11,
        'font_name': 'Calibri',
        'num_format': '#,##0.00'
    })
    money_format=workbook.add_format({
        'align':'right',
        'valign':'vbottom',
        'font_size':11,
        'font_name':'Calibri',
        'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
    })

    header_1 = [
        'Mã loại hình',
        'Tên loại hình',
        'Số TK lưu ký',
        'Số tiểu khoản',
        'Tên khách hàng',
        'Tên MG',
        'Tiền',
        'TL MR thực tế',
        'TL TC thực tế'
    ]

    header_2 = [
        'TL MR DN',
        'TL TC DN',
        'RLN0005',
        'RLN0006',
        'VMR9003',
        'Note'
    ]

    worksheet=workbook.add_worksheet('Sheet1')
    # Set Columns & Rows
    worksheet.set_column('A:B',9)
    worksheet.set_column('C:C',16)
    worksheet.set_column('D:F',8)
    worksheet.set_column('G:G',15)
    worksheet.set_column('H:K',9)
    worksheet.set_column('L:L',15)
    worksheet.set_column('M:M',9)
    worksheet.set_column('N:N',15)
    worksheet.set_column('O:O',9)
    worksheet.set_row(0, 30)

    for index, val in final_table['RLN0005'].iteritems():
        if val == '#N/A':
            worksheet.set_row(index+1,0)

    worksheet.write_row('A1',header_1,header_1_format)
    worksheet.write_row('J1',header_2,header_2_format)
    worksheet.write_column('A2',final_table['MaLoaiHinh'],text_left_format)
    worksheet.write_column('B2',final_table['TenLoaiHinh'],text_left_format)
    worksheet.write_column('C2',final_table['account_code'],text_left_format)
    worksheet.write_column('D2',final_table['sub_account'],text_left_format)
    worksheet.write_column('E2',final_table['customer_name'],text_left_format)
    worksheet.write_column('F2',final_table['broker_name'],text_left_format)
    worksheet.write_column('G2',final_table['Tien'],money_format)
    worksheet.write_column('H2',final_table['TLMRThucTe'],number_format)
    worksheet.write_column('I2',final_table['TLTCThucTe'],number_format)
    worksheet.write_column('J2',final_table['TLMRDN'],number_format)
    worksheet.write_column('K2',final_table['TLTCDN'],number_format)
    worksheet.write_column('L2',final_table['RLN0005'],money_format)
    worksheet.write_column('M2',final_table['RLN0006'],money_format)
    worksheet.write_column('N2',final_table['VMR9003'],money_format)
    worksheet.write_column('O2',['']*final_table.shape[0],text_left_format)

    ################################################################
    ################################################################
    ################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')




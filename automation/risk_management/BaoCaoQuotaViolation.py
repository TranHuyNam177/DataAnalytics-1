from automation.risk_management import *
from datawarehouse import *


def run(
        run_time=None
):
    start=time.time()
    info=get_info('daily', run_time)
    period=info['period']
    t0_date=info['end_date'].replace('.', '-')
    t1_date=BDATE(t0_date,-1)
    folder_name=info['folder_name'] 

    # create_folder
    if not os.path.isdir(join(dept_folder, folder_name, period)):
        os.mkdir((join(dept_folder, folder_name, period)))

    ################################################################
    ################################################################
    ################################################################

    # Query SQL to getting data
    table=pd.read_sql(
        f"""
        SELECT
            CASE
                WHEN [relationship].[branch_id] = '0001' THEN 'Headquarter'
                WHEN [relationship].[branch_id] = '0101' THEN 'District 3'
                WHEN [relationship].[branch_id] = '0102' THEN 'Phu My Hung'
                WHEN [relationship].[branch_id] = '0104' THEN 'District 7'
                WHEN [relationship].[branch_id] = '0105' THEN 'Tan Binh'
                WHEN [relationship].[branch_id] = '0111' THEN 'Institutional Business'
                WHEN [relationship].[branch_id] = '0113' THEN 'IB'
                WHEN [relationship].[branch_id] = '0117' THEN 'District 1'
                WHEN [relationship].[branch_id] = '0118' THEN 'AMD-03'
                WHEN [relationship].[branch_id] = '0119' THEN 'Institutional Business 02'
                WHEN [relationship].[branch_id] = '0201' THEN 'Ha Noi'
                WHEN [relationship].[branch_id] = '0202' THEN 'Thanh Xuan'
                WHEN [relationship].[branch_id] = '0203' THEN 'Cau Giay'
                WHEN [relationship].[branch_id] = '0301' THEN 'Hai Phong'
            END [branch_name],
            [account].[account_code],
            [account].[customer_name],
            [vmr0003].[guarantee_debt] [quota_violation],
            [vmr0003].[date],
            [broker].[broker_name]
        FROM
            [vmr0003]
        LEFT JOIN
            [relationship]
        ON [relationship].[sub_account] = [vmr0003].[sub_account]
        AND [relationship].[date] = [vmr0003].[date]
        LEFT JOIN
            [account]
        ON [account].[account_code] = [relationship].[account_code]
        LEFT JOIN
            [broker]
        ON [broker].[broker_id] = [relationship].[broker_id]
        WHERE
            [vmr0003].[date] = '{t1_date}'
        AND [vmr0003].[sub_account_type] <> N'Tự doanh'
        AND [vmr0003].[guarantee_debt] <> 0
        """,
        connect_DWH_CoSo
    )

    ################################################################
    ################################################################
    ################################################################

    t0_day = t0_date[8:10]
    t0_month = int(t0_date[5:7])
    t0_month = calendar.month_name[t0_month]
    t0_year = t0_date[0:4]
    eod = t0_month + ' ' + t0_day + ' ' + t0_year
    file_name=f'Quota Violation Report on {eod}.xlsx'
    writer=pd.ExcelWriter(
        join(dept_folder,folder_name,period,file_name),
        engine='xlsxwriter',
        engine_kwargs={'options': {'nan_inf_to_errors': True}}
    )
    workbook=writer.book
    title_format=workbook.add_format({
        'bold':True,
        'align':'center',
        'valign':'vbottom',
        'font_size':14,
        'font_name':'Arial'
    })
    header_format=workbook.add_format({
        'bold':True,
        'border':1,
        'align':'center',
        'valign':'vcenter',
        'font_size':10,
        'font_name':'Arial'
    })
    text_center_format=workbook.add_format({
        'border':1,
        'align':'center',
        'valign':'vcenter',
        'font_size':10,
        'font_name':'Arial'
    })
    money_format=workbook.add_format({
        'border':1,
        'align':'right',
        'valign':'vcenter',
        'font_size':10,
        'font_name':'Arial',
        'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
    })
    date_format=workbook.add_format({
        'border':1,
        'align':'center',
        'valign':'vcenter',
        'font_name':'Arial',
        'font_size':10,
        'num_format':'dd/mm/yyyy'
    })
    total_format=workbook.add_format({
        'bold':True,
        'border':1,
        'align':'center',
        'valign':'vcenter',
        'text_wrap': True,
        'font_name':'Arial',
        'font_size':10,
    })
    total_val_format=workbook.add_format({
        'bold':True,
        'border':1,
        'align':'right',
        'valign':'vbottom',
        'font_name':'Arial',
        'font_size':11,
        'num_format':'_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
    })
    title_name='QUOTA VIOLATIONS'
    now = dt.datetime.now()
    date_title=now.strftime('%d/%m/%Y')
    worksheet=workbook.add_worksheet('Sheet1')

    # Set Columns & Rows
    worksheet.set_column('A:B',13)
    worksheet.set_column('C:C',38.29)
    worksheet.set_column('D:D',22.43)
    worksheet.set_column('E:E',12.29)
    worksheet.set_column('F:F',33.86)
    worksheet.set_column('G:G',21.71)
    worksheet.merge_range('B6:D6',title_name, title_format)
    headers = ['Branch','Code','Name','Quota Violation Amount','From','Broker','Note']
    worksheet.write('C7', date_title, title_format)
    worksheet.write_row('A10',headers,header_format)
    worksheet.write_column('A11',table['branch_name'],text_center_format)
    worksheet.write_column('B11',table['account_code'],text_center_format)
    worksheet.write_column('C11',table['customer_name'],text_center_format)
    worksheet.write_column('D11',table['quota_violation'],money_format)
    worksheet.write_column('E11',table['date'],date_format)
    worksheet.write_column('F11',table['broker_name'],text_center_format)
    worksheet.write_column('G11',['']*table.shape[0],total_format)
    sum_row = table.shape[0] + 11
    worksheet.merge_range(f'A{sum_row}:C{sum_row}','Total',total_format)
    worksheet.write(
        f'D{sum_row}',
        table['quota_violation'].sum(),
        total_val_format
    )
    worksheet.merge_range(f'E{sum_row}:G{sum_row}','',total_format)

    ################################################################
    ################################################################
    ################################################################

    writer.close()
    if __name__=='__main__':
        print(f"{__file__.split('/')[-1].replace('.py','')}::: Finished")
    else:
        print(f"{__name__.split('.')[-1]} ::: Finished")
    print(f'Total Run Time ::: {np.round(time.time()-start,1)}s')




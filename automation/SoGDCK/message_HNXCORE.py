from automation.SoGDCK import *


def create_dict():
    file_path = join(dirname(__file__), 'file', 'HNXCORE_TKCT_InfoGate.xlsx')
    hnx_infogate_message = pd.read_excel(file_path, dtype={'Tag': object}).fillna('')
    hnx_infogate_message = hnx_infogate_message.drop_duplicates(ignore_index=True)
    # save InfoGate message dictionary to a pickle file
    message_dict = {str(x[0]): x[1:] for x in hnx_infogate_message.itertuples(index=False)}
    with open(join(dirname(__file__), 'file', 'message_dictionary.pickle'), 'wb') as handle:
        pickle.dump(message_dict, handle, protocol=pickle.HIGHEST_PROTOCOL)
    return message_dict


# read file data.pkl
with open(r'C:\Users\namtran\Share Folder\SoGDCK\data.pkl', 'rb') as file:
    data = pickle.load(file)


def readList(lst: list):
    message_idx = {
        '8','9','35','49','52','1','2','3','4','5','6','7','14','19','21','22','23','24','25','26','27','18'
    }
    message_roChiSo = {'8','9','35','49','52','1','2','15','55','11','12','28'}
    message_stockInfo = {
        '8','9','35','49','52','55','15','425','336','340','326','327','167','225','106','107','132','1321','133','1331',
        '134','135','260','333','332','334','31','32','137','138','139','140','387','3871','631','388','399','400',
        '109','17','230','232','233','244','255','2551','266','2661','277','310','320','321','391','392','393','3931',
        '394','3941','395','3951','3952','396','3961','3962','397','3971','398','3981','3301','541','223','1341','1351'
    }
    message_topNPrice = {'8','9','35','49','52','55','425','555','556','132','1321','133','1331'}
    message_topPriceOddLot = {
        '8','9','35','49','52','55','425','132','1321','133','1331','134','1341','135','1351','136','1361','137','1371'
    }
    message_boardInfo = {
        '8','9','35','49','52','425','426','336','340','421','422','388','399','270','250','251','252','253','17',
        '220','221','210','211','240','241','341'
    }
    message_autionMatch = {'8','9','35','49','52','33','55','31','32'}
    message_ETF_NET_VALUE = {'8','9','35','49','52','56','57','58'}
    message_ETF_TRACKING_ERROR = {'8','9','35','49','52','56','59','60','61'}
    message_derivativesInfo = {
        '8','9','35','49','52','55','15','800','425','336','340','326','327','167','801','8011','802','803','132','1321',
        '133','1331','134','135','260','333','332','31','32','137','138','804','139','140','805','387','3871','388',
        '399','400','17','255','2551','266','2661','310','320','321','391','392','393','3931','394','3941','814','815',
        '397','8141','8151','3971','816','817','398','8161','8171','3981'
    }
    if lst[-1].decode('utf-8') == '\r\n':
        lst = lst[:-1]
    new_lst = [ele.decode('utf-8').split('=')[0] for ele in lst]
    lst_to_set = set(new_lst)
    if lst_to_set.issubset(message_idx):
        table_name = 'message_index'
    elif lst_to_set.issubset(message_roChiSo):
        table_name = 'message_ro_chi_so'
    elif lst_to_set.issubset(message_stockInfo):
        table_name = 'message_stockInfo'
    elif lst_to_set.issubset(message_topNPrice):
        table_name = 'message_topNPrice'
    elif lst_to_set.issubset(message_topPriceOddLot):
        table_name = 'message_topPriceOddLot'
    elif lst_to_set.issubset(message_boardInfo):
        table_name = 'message_boardInfo'
    elif lst_to_set.issubset(message_autionMatch):
        table_name = 'message_autionMatch'
    elif lst_to_set.issubset(message_derivativesInfo):
        table_name = 'message_derivativesInfo'
    elif lst_to_set.issubset(message_ETF_TRACKING_ERROR):
        table_name = 'message_ETF_TRACKING_ERROR'
    elif lst_to_set.issubset(message_ETF_NET_VALUE):
        table_name = 'message_ETF_NET_VALUE'
    else:
        raise Exception('no table name matches data!!!')

    return table_name

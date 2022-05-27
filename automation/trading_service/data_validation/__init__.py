from automation.trading_service import *

driver = '{SQL Server}'
server = 'SRV-RPT'
database = 'RiskDb'
with open(r'C:\Users\namtran\Desktop\Passwords\DataBase\DataBase.txt') as file:
    user_id,user_password = file.readlines()
    user_id = user_id.replace('\n','')

connect = pyodbc.connect(
    f'Driver={driver};'
    f'Server={server};'
    f'Database={database};'
    f'uid={user_id};'
    f'pwd={user_password}'
)

TableNames = pd.read_sql(
    'SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES',connect
)

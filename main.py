import pandas as pd
from sqlalchemy import create_engine

# Create engine

df = pd.read_excel('D:/Work/POOF/Poof_Product_List_Excel.xlsx')
Codes = {}
for
username = 'root'
password = '13579111315szxM'
engine = create_engine(f'mysql+mysqlconnector://{username}:{password}@127.0.0.1/poof_schema')
df.to_sql(name = "product_list", con = engine, if_exists= 'append', index = False)
c = engine.connect()
import pandas as pd
from sqlalchemy import create_engine
df = pd.read_excel('D:/Work/POOF/Poof_Product_List_Excel')
username = 'root'
password = '13579111315szxM'
engine = create_engine('mysql://{username}:{password}@localhost/poof_schema')
df.to_sql(name = "product_list", con = engine, if_exists= 'append', index = False)

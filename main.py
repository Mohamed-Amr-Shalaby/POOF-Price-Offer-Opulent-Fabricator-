import pandas as pd
from sqlalchemy import create_engine

# Create engine

df = pd.read_excel('D:/Work/POOF/Poof_Product_List_Excel.xlsx')
Codes = {}
username = 'root'
password = '13579111315szxM'
engine = create_engine(f'mysql+mysqlconnector://{username}:{password}@127.0.0.1/poof_schema')
df.to_sql(name = "product_list", con = engine, if_exists= 'replace', index = False)
c = engine.connect()


def get_date():
    print("Please enter the date: ")
    date = input()
    return date


def get_product(code, name):
    pname = name
    pcode = code
    pprice = c.execute(f"SELECT Price FROM product_list WHERE code = {pcode}")
    pdescription = c.execute(f"SELECT Description FROM product_list WHERE code = {pcode}")
    pimg_dir = c.execute(f"SELECT Image_Directory FROM product_list WHERE code = {pcode}")
    product = [pcode, pname, pprice, pdescription, pimg_dir]
    return product


def get_quantity():
    print("Please enter the quantity: ")
    quantity = input()
    return quantity


def get_client_name():
    print("Please enter the client name: ")
    client = input()
    return client


from flask import Flask, render_template, request, redirect
from flask_wtf import FlaskForm
from wtforms import StringField, SubmitField, IntegerField
import pandas as pd
from sqlalchemy import create_engine
from werkzeug.security import generate_password_hash, check_password_hash

# Create engine

df = pd.read_excel('D:/Work/POOF/Poof_Product_List_Excel.xlsx')
Codes = {}
username = 'root'
password = '13579111315szxM'
engine = create_engine(f'mysql+mysqlconnector://{username}:{password}@127.0.0.1/poof_schema')
df.to_sql(name = "product_list", con = engine, if_exists= 'replace', index = False)
c = engine.connect()

# Connect Flask
app = Flask(__name__)
app.config['SECRET_KEY'] = 'mysecret'

@app.route('/', methods = ["GET", "POST"])

def index():
    if request.method == "GET":
        return render_template('signIn.html')
    elif request.method == "POST":
        name = request.form.get("Username")
        Pass = request.form.get("Password")
        if not name or not Pass:
            return render_template('Invalid_Credentials.html')
        authority = c.execute("SELECT Authority_Level FROM authorized_personnel WHERE Employee_Name = ?", name)
        if len(authority) != 1 or not check_password_hash(authority[0]["hash"], Pass):
            return render_template('Invalid_Credentials.html')
        for row in c:
            print(row)
        '''
        if authority == "Administrator" and password == Pass:
            return render_template('admin_options.html')
        elif authority == "Data_Entry" and password == Pass:
            return render_template('Create_Quotation.html')
        elif authority == "Developer" and password == Pass:
            return render_template('Developer_Options.html')
        else:
            return render_template('Invalid_Credentials.html')
        '''

@app.route("/register", methods = ["GET", "POST"])
def register():
    """TODO Register user"""
    if request.method == "POST":
        name = request.form.get("username")
        password = request.form.get("password")
        Cpassword = request.form.get("confirmation")
        hash = generate_password_hash(password, method='pbkdf2', salt_length=16)
        if not name:
            return render_template("Invalid_Credentials.html")

        # Ensure password was submitted
        elif not password:
            return render_template("Invalid_Credentials.html")

        # Query database for username
        rows = c.execute("SELECT * FROM authorized_personnel WHERE Employee_Name = ?", name)

        # Ensure username exists and password is correct
        if len(rows) != 0:
            return render_template("Invalid_Credentials.html")

        if not password == Cpassword:
            return render_template("Invalid_Credentials.html")
        c.execute("INSERT INTO authorized_personnel(Employee_Name, Password) VALUES (?, ?)", name, hash)

        # Redirect user to home page
        return redirect("/")
    else:
        return render_template("register.html")


def get_date():
    print("Please enter the date: ")
    date = input()
    return date


def get_product(code, name):
    pname = name
    pcode = code
    pprice = c.execute("SELECT Price FROM product_list WHERE code =? ", pcode)
    pdescription = c.execute(f"SELECT Description FROM product_list WHERE code = ?", pcode)
    pimg_dir = c.execute(f"SELECT Image_Directory FROM product_list WHERE code = ?", pcode)
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


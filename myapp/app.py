from flask import Flask, render_template, request, redirect
from flask_wtf import FlaskForm
from wtforms import StringField, SubmitField, IntegerField
import pandas as pd
from sqlalchemy import create_engine, text
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
@app.route('/', methods = ["GET", "POST"])

def index():
    # If the user is not logged in, redirect to login page
    if request.method == "GET":
        return render_template('signIn.html')
    elif request.method == "POST":
        # Get username and password
        name = request.form.get("Username")
        Pass = request.form.get("Password")
        # Query Database for the data of the employee
        authorityres = c.execute(text(f'SELECT * FROM authorized_personnel WHERE Employee_Name = "{name}"'))
        c.commit()
        authority = authorityres.all()
        # Ensure username and password were submitted  
        if not name or not Pass:
            print("Name or Pass is empty")
            return render_template('Invalid_Credentials.html', name = name, Pass = Pass)
        # Ensure username exists and password is correct    
        if len(authority) != 1 or not check_password_hash(authority[0][2], Pass):
            print("Wrong name or wrong Password")
            return render_template('Invalid_Credentials.html', authority = authorityres)
        # Redirect user to appropriate page depending on the authority level
        if authority[0][3] == "Administrator":
            return render_template('admin_options.html')
        elif authority[0][3] == "Data_Entry":
            return render_template('Create_Quotation.html')
        elif authority[0][3] == "Developer":
            return render_template('Developer_Options.html')
        else:
            return render_template('Invalid_Credentials.html')
        

@app.route("/register", methods = ["GET", "POST"])
def register():
    """TODO Register user"""
    if request.method == "POST":
        name = request.form.get("username")
        password = request.form.get("password")
        Cpassword = request.form.get("confirmation")
        authority = request.form.get("authority")
        hash = generate_password_hash(password, method='pbkdf2', salt_length=16)
        if not name:
            return render_template("Invalid_Credentials.html")

        # Ensure password was submitted
        elif not password:
            return render_template("Invalid_Credentials.html")

        # Query database for username
        result = c.execute(text(f'SELECT * FROM authorized_personnel WHERE Employee_Name = "{name}"'))
        rows = result.all()
        # Ensure username exists and password is correct
        if len(rows) != 0:
            return render_template("Invalid_Credentials.html")

        if password != Cpassword:
            return render_template("Invalid_Credentials.html")
        c.execute(text(f'INSERT INTO authorized_personnel(Employee_Name, Password, Authority_Level) VALUES ("{name}", "{hash}", "{authority}")'))
        c.commit()
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


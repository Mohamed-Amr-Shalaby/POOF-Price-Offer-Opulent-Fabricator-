from werkzeug.security import generate_password_hash, check_password_hash
from flask import Flask, render_template, request, redirect
from wtforms import StringField, SubmitField, IntegerField
from sqlalchemy import create_engine, text
from flask_wtf import FlaskForm
import pandas as pd

# Create engine

df = pd.read_excel("D:/Work/POOF/complete_product_list_spreadsheet.xlsx")
Codes = {}
username = "root"
password = "13579111315szxM"
engine = create_engine(
    f"mysql+mysqlconnector://{username}:{password}@127.0.0.1/poof_schema"
)
df.to_sql(name="product_list", con=engine, if_exists="replace", index=False)
c = engine.connect()

# Connect Flask
app = Flask(__name__)


@app.route("/", methods=["GET", "POST"])
def index():
    # TODO session tracking
    # If the user is not logged in, redirect to login page
    if request.method == "GET":
        return render_template("signIn.html")
    elif request.method == "POST":
        # Get username and password
        name = request.form.get("Username")
        Pass = request.form.get("Password")

        # Query Database for the data of the employee
        # TODO Protect against SQL injection
        authorityres = c.execute(
            text(f'SELECT * FROM authorized_personnel WHERE Employee_Name = "{name}"')
        )
        c.commit()
        authority = authorityres.all()

        # Ensure username and password were submitted
        if not name or not Pass:
            print("Name or Pass is empty")
            return render_template("Invalid_Credentials.html", name=name, Pass=Pass)

        # Ensure username exists and password is correct
        if len(authority) != 1 or not check_password_hash(authority[0][2], Pass):
            print("Wrong name or wrong Password")
            return render_template("Invalid_Credentials.html", authority=authorityres)

        # Redirect user to appropriate page depending on the authority level
        if authority[0][3] == "Administrator":
            return render_template("admin_options.html")
        elif authority[0][3] == "Data_Entry":
            return render_template("Create_Quotation.html")
        elif authority[0][3] == "Developer":
            return render_template("Developer_Options.html")
        else:
            return render_template("Invalid_Credentials.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    # TODO Make a change password page, hashes are great but now I don't know anyone's passwords and can't check them directly

    if request.method == "POST":
        # Get username and password and authority level

        name = request.form.get("username")
        password = request.form.get("up")
        Cpassword = request.form.get("up2")
        authority = request.form.get("authority")
        hash = generate_password_hash(password, method="pbkdf2", salt_length=16)
        print(hash)
        if not name:
            return render_template("Invalid_Credentials.html")

        # Ensure password was submitted
        elif not password:
            return render_template("Invalid_Credentials.html")

        # Query database for username
        # TODO Protect against SQL injection
        result = c.execute(
            text(f'SELECT * FROM authorized_personnel WHERE Employee_Name = "{name}"')
        )
        rows = result.all()

        # Ensure username exists and password is correct
        if len(rows) != 0:
            return render_template("Invalid_Credentials.html")

        if password != Cpassword:
            return render_template("Invalid_Credentials.html")

        # TODO Protect against SQL injection
        c.execute(
            text(
                f'''INSERT INTO authorized_personnel(Employee_Name, Password, Authority_Level) 
                VALUES ("{name}", "{hash}", "{authority}")'''
            )
        )
        c.commit()

        # Redirect user to home page
        return redirect("/")

    else:
        return render_template("register.html")


# Tasks page
@app.route("/task", methods=["GET", "POST"])
def task():
    if request.method == "GET":
        return render_template("admin_options.html")
    elif request.method == "POST":
        choice = request.form.get("task")
        print(choice)
        rows = c.execute(text(f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"))
        prods = rows.all()
        if choice == "Create_quotation":
            print(type(prods))
            return render_template("Create_Quotation.html", products = prods)
        elif choice == "Edit product prices":
            return render_template("edit_prices.html")
        elif choice == "Add new product":
            return render_template("add_product.html")
        elif choice == "Add Employee":
            return render_template("register.html")
        else:
            return render_template("Invalid_Choice.html")

# Create Quotation Page and handling Queries, autocomplete, and dynamic table row insertion
@app.route("/Create_Quotation", methods=["GET", "POST"])
def Create_Quotation():
    if request.method == "POST":
        # Get the product code and quantity from the submitted form
        code = request.form.get("product_code")
        quantity = request.form.get("quantity")
        quantity = float(quantity)
        print(code, quantity)
        rows = c.execute(text(f"SELECT * FROM product_list WHERE product_code = '{code}'"))
        prod = rows.all()
        print(f"price is {prod[0][2]}")
        # TODO redo the sql table and replace it with a list of lists in the frontend instead of a SQL table
        c.execute(text(f'INSERT INTO current_quotation VALUES ("{prod[0][0]}", "{prod[0][4]}", "{prod[0][1]}", "{prod[0][3]}", "{quantity}", "{prod[0][2]}", "{prod[0][2] * quantity}")'))
        c.commit()
        rows = c.execute(text(f"SELECT * FROM current_quotation"))
        entries = rows.all()
        rows = c.execute(text(f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"))
        prods = rows.all()
        return render_template("Create_Quotation.html", entries = entries, products = prods)



@app.route("/query", methods=["GET", "POST"])
def query():
    if request.method == "GET":
        return render_template("Create_Quotation.html")
    elif request.method == "POST":
        code = request.form.get("product_code")
        name = request.form.get("product_name")
        print(name)
        print(code)
        quantity = request.form.get("quantity")
        print(quantity)
        if (not code and not name) or not quantity:
            return render_template("insufficient_data.html")
        product_details = get_product(code, name, quantity)

        return render_template("Create_Quotation.html", product_details=product_details, quantity=quantity)

# Create Page that allows admins to change the price of products
@app.route("/price", methods = ["GET", "POST"])
def price():
    if request.method == "POST":
        choice = request.form.get("price")
        if choice == "single_product":
            return render_template("single_product.html")
        elif choice == "percentage":
            return render_template("percentage.html")
    return render_template("edit_prices.html")

# Change the price of a single product
@app.route("/single_price", methods = ["GET", "POST"])
def single_price():
    if request.method == "POST":
        code = request.form.get("code")
        price = request.form.get("price")
        c.execute(text(f'UPDATE product_list SET Price = "{price}" WHERE Product_Code = "{code}"'))
        c.commit()
        return render_template("edit_prices.html")
    return render_template("single_product.html")

# Change all product prices by a constant percentage either by increasing or decreasing
@app.route("/percentage", methods = ["GET", "POST"])
def percentage():
    if request.method == "POST":
        percentage = request.form.get("percentage")
        Type = request.form.get("change_type")
        if not percentage:
            return render_template("Invalid_Choice.html")
        if Type == "Increase":
            new_percentage = 1 + (int(percentage) / 100)
        elif Type == "Decrease":
            new_percentage = 1 - (int(percentage) / 100)
        c.execute(text(f'UPDATE product_list SET Price = Price * {new_percentage}'))
        c.commit()
        return render_template("edit_prices.html")
    return render_template("percentage.html")

# Add a new product to the database
@app.route("/add_product", methods = ["GET", "POST"])
def add_product():
    if request.method == "POST":
        code = request.form.get("code")
        name = request.form.get("name")
        price = request.form.get("price")
        description = request.form.get("description")
        image = request.form.get("image") #TODO Convert image file to directory
        c.execute(text(f'INSERT INTO product_list(Product_Code, Product_Name, Price, Description, Image_Directory) VALUES ("{code}", "{name}", "{price}", "{description}", "{image}")'))
        c.commit()
        return render_template("admin_options.html")
    return render_template("add_product.html")

def get_date():
    print("Please enter the date: ")
    date = input()
    return date


def get_product(code, name, quantity):
    pname = name
    pcode = code
    
    if not pcode:
        pcode = c.execute(text(f"SELECT Product_Code FROM product_list WHERE Product_Name = '{pname}'"))
        pcode = pcode.all()
    if not pname:
        pname = c.execute(text(f"SELECT Product_Name FROM product_list WHERE Product_Code = '{pcode}'"))
        pname = pname.all()
    pprice = c.execute(text(f"SELECT Price FROM product_list WHERE product_code = '{pcode}'"))
    pprice = pprice.all()

    pdescription = c.execute(text(f"SELECT Description FROM product_list WHERE product_code = '{pcode}'"))
    pdescription = pdescription.all()
    pimg_dir = c.execute(text(f"SELECT Image_Directory FROM product_list WHERE product_code = '{pcode}'"))
    pimg_dir = pimg_dir.all()
    product = [pimg_dir, pcode, pname, pdescription, quantity, pprice]
    return product


def get_quantity():
    print("Please enter the quantity: ")
    quantity = input()
    return quantity


def get_client_name():
    print("Please enter the client name: ")
    client = input()
    return client

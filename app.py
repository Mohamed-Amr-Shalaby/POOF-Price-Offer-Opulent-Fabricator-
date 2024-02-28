from werkzeug.security import generate_password_hash, check_password_hash
from flask import Flask, render_template, request, redirect, session
from wtforms import StringField, SubmitField, IntegerField
from sqlalchemy import create_engine, text
from flask_wtf import FlaskForm
import pandas as pd
import openpyxl
from openpyxl import Workbook
from flask_session import Session
import PIL



#  Create engine

df = pd.read_excel("D:/Work/POOF/complete_product_list_spreadsheet.xlsx")
Codes = {}
current_quotation = []
current_client = {"Date": "",
                   "Customer_Name":"",
                   "Customer_Number":"",
                   "Rep_Name":"",
                   "Rep_Number":"",} 
customer_info = []
username = "root"
password = "13579111315szxM"
engine = create_engine(
    f"mysql+mysqlconnector://{username}:{password}@127.0.0.1/poof_schema"
)
try:
    df.to_sql(name="product_list", con=engine, if_exists="fail", index=False)
except ValueError:
    pass
c = engine.connect()

# Connect Flask
app = Flask(__name__)

#ٌRead the Quotation Template content
quotation = openpyxl.load_workbook("D:/Work/POOF/Quotation_Template.xlsx")
sheet = quotation.active
Date_Cell = sheet["I5"]
Customer_Name_Cell = sheet["A9"]
Customer_Number_Cell = sheet["C9"]
Rep_Name_Cell = sheet["A12"]
Rep_Number_Cell = sheet["C12"]
Quotation_Number_Cell = sheet["G5"]
""" 
# Create a document instance
quotation_doc = Document()

# Set Font of the document
style = quotation_doc.styles["Normal"]
style.font.name = "Arial"

# Add header to the document
header_section = quotation_doc.sections[0]
header = header_section.header
header_text = header.paragraphs[0]
header_text.text = "Multimedica ScO.\nAddress: 27 Al Hayah St. Tanta Qism 2, Gharbia Governorate 31511, Egypt"

 """
# Configure Session to use filesystem
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_TYPE"] = "filesystem"
Session(app)

@app.route("/", methods=["GET", "POST"])
def index():
    # If the user is not logged in, redirect to login page
    if not session.get("name"):
        return redirect("/login")
    name = session["name"].split(" ")[0]
    return render_template("index.html", name=name)

@app.route("/login", methods=["GET", "POST"])
def login():
    # TODO session tracking
    # If the user is not logged in, redirect to login page
    if request.method == "GET":
        if session.get("name") and session.get("access_level"):
            print(session.get("name"))
            print(session.get("access_level"))
            return render_template("index.html")
        return render_template("signIn.html")
    elif request.method == "POST":
        # Get username and password
        name = request.form.get("Username")
        Pass = request.form.get("Password")
        session["name"] = name
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
        
        session["user_id"] = authority[0][0]
        session["access_level"] = authority[0][3]
        # Redirect user to appropriate page depending on the authority level
        if session["access_level"]:
            return redirect("/")
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
        #print(hash)
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

# Logout the user
@app.route("/logout")
def logout():
    # Clear the session
    session.clear()
    return redirect("/")


# Tasks page
@app.route("/task", methods=["GET", "POST"])
def task():
    if request.method == "GET":
        if session["access_level"] == 'Developer':
            return render_template("admin_options.html")
        elif session["access_level"] == 'Administrator':
            return render_template("admin_options.html")
        elif session["access_level"] == 'Data_Entry':    
            return render_template("customer_info.html")
        return render_template("Invalid_Credentials.html")
    elif request.method == "POST":
        choice = request.form.get("task")
        #print(choice)
        rows = c.execute(text(f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"))
        prods = rows.all()
        if choice == "Create_quotation":
            #print(type(prods))
            return render_template("customer_info.html")
        elif choice == "Edit product prices":
            return render_template("edit_prices.html")
        elif choice == "Add new product":
            return render_template("add_product.html")
        elif choice == "Add Employee":
            return render_template("register.html")
        else:
            return render_template("Invalid_Choice.html")


@app.route("/customer_info", methods=["GET", "POST"])
def Customer_Info():
    if request.method == "POST":
        current_client["Date"] = request.form.get("quotation_date")
        current_client["Customer_Name"] = request.form.get("customer_name")
        current_client["Customer_Number"] = request.form.get("customer_number")
        current_client["Rep_Name"] = request.form.get("rep_name")
        current_client["Rep_Number"] = request.form.get("rep_number")
        rows = c.execute(text(f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"))
        prods = rows.all()
        if not current_client["Customer_Name"] or not current_client["Date"] or not current_client["Rep_Name"] or not current_client["Rep_Number"]:
            return render_template("insufficient_data.html")
        return render_template("Create_Quotation.html", products = prods, customer_info = current_client)
    pass

@app.route("/Edit_Quotation", methods=["GET", "POST"])
def Edit_Quotation():
    if request.method == "POST":
        rows = c.execute(text(f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"))
        prods = rows.all()
        return render_template("Create_Quotation.html", products = prods, customer_info = current_client, entries = current_quotation)

# Create Quotation Page and handling Queries, autocomplete, and dynamic table row insertion
@app.route("/Create_Quotation", methods=["GET", "POST"])
def Create_Quotation():
    if request.method == "POST":
        # Get the product code and quantity from the submitted form
        code = request.form.get("product_code")
        quantity = request.form.get("quantity")
        quantity = float(quantity)
        rows = c.execute(text(f"SELECT * FROM product_list WHERE product_code = '{code}'"))
        prod = rows.all()
        # redo the sql table and replace it with a list of lists in the frontend instead of a SQL table
        if len(prod) > 0:
            current_quotation.append([prod[0][4], prod[0][0], prod[0][1], prod[0][3], quantity, prod[0][2], prod[0][2] * quantity])
        else:
            render_template("insufficient_data.html")
        """ print(prod[0][0])
        print(prod[0][1])
        print(prod[0][2])
        print(prod[0][3])
        print(prod[0][4]) """
        rows = c.execute(text(f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"))
        prods = rows.all()
        return render_template("Create_Quotation.html", entries = current_quotation, products = prods, customer_info = current_client)

#TODO Add a route that allows the user to delete a product from the quotation. This will require a new form for each table row in the HTML

@app.route("/preview", methods=["GET", "POST"])
def preview():
    #TODO Add a preview page that shows the current quotation, an option to go back and edit, and an option to submit
    return render_template("preview_quotation.html", customer_info = current_client, entries = current_quotation)


# Convert the current quotation to a dataframe, submit it to the database, and export it as an excel file
@app.route("/export", methods=["GET", "POST"])
def submit():
    if request.method == "POST":
        client_columns = ['Date', 'Customer_Name', 'Customer_Number', 'Rep_Name', 'Rep_Number']
        product_columns = ['Image', 'Product_Code', 'Product_Name', 'Description', 'Quantity', 'Price', 'Total']
        client_list = [[current_client["Date"], current_client["Customer_Name"], current_client["Customer_Number"], current_client["Rep_Name"], current_client["Rep_Number"]]]
        client_data = pd.DataFrame(client_list, columns = client_columns)
        product_data = pd.DataFrame(current_quotation, columns = product_columns)
        #Export the dataframes to an Excel file, then save the file as a pdf
        #and save the pdf to the database alongside the name of the user and the date of submission
        #Add serializtion to the quotation files
        """ with pd.ExcelWriter("D:/Work/POOF/Quotation.xlsx") as writer:
            client_data.to_excel(writer, sheet_name = "Client_Info", index = True)
            product_data.to_excel(writer, sheet_name = "Product_Info", index = True) """
        #Clear the current quotation and client info
        current_quotation.clear()
        current_client.clear()
        Number = 1
        # Add the client data to the quotation template
        Date_Cell.value = client_data["Date"][0]
        Customer_Name_Cell.value = client_data["Customer_Name"][0]
        Customer_Number_Cell.value = client_data["Customer_Number"][0]
        Rep_Name_Cell.value = client_data["Rep_Name"][0]
        Rep_Number_Cell.value = client_data["Rep_Number"][0]
        Quotation_Number_Cell.value = Number
        # Add the product data to the quotation template
        rows = sheet.iter_rows(min_row = 14, max_row = 14 + len(product_data), min_col = 1, max_col = 9)
        print("Rows: ", rows)
        for i, row in enumerate(rows):
            if i == len(product_data):
                break
            row[6].value = product_data["Quantity"][i]
            row[0].value = product_data["Product_Name"][i]
            row[2].value = product_data["Description"][i]
            row[7].value = product_data["Price"][i]
            row[8].value = product_data["Total"][i]     

        quotation.save(filename = f'Quotation_{Number}.xlsx')
        
        """ 
        #Add Title
        quotation_doc.add_heading(f"Quotation No. {Number}", 0)
        # Add Paragraphs
        p = quotation_doc.add_paragraph("Date: ")
        p.add_run(client_data["Date"][0])
        p = quotation_doc.add_paragraph("Customer Name: ")
        p.add_run(client_data["Customer_Name"][0])
        p = quotation_doc.add_paragraph("Customer Number: ")
        p.add_run(client_data["Customer_Number"][0])
        p = quotation_doc.add_paragraph("Rep Name: ")
        p.add_run(client_data["Rep_Name"][0])
        p = quotation_doc.add_paragraph("Rep Number: ")
        p.add_run(client_data["Rep_Number"][0])
        
        # Add QR Code of Quotation
        quotation_doc.add_picture("D:/Work/POOF/qr_code.png", width=Cm(3.0), height=Cm(3.0))

        # Add Table
        table = quotation_doc.add_table(rows=1, cols=7)
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = "Image"
        hdr_cells[1].text = "Product Code"
        hdr_cells[2].text = "Product Name"
        hdr_cells[3].text = "Description"
        hdr_cells[4].text = "Quantity"
        hdr_cells[5].text = "Price"
        hdr_cells[6].text = "Total"
        for i in range(len(product_data)):
            row_cells = table.add_row().cells
            print(f"Product Code is: {product_data["Product_Code"][i]}")
            print(i)
            row_cells[0].text = str(product_data["Image"][i])
            row_cells[1].text = str(product_data["Product_Code"][i])
            row_cells[2].text = str(product_data["Product_Name"][i])
            row_cells[3].text = str(product_data["Description"][i])
            row_cells[4].text = str(product_data["Quantity"][i])
            row_cells[5].text = str(product_data["Price"][i])
            row_cells[6].text = str(product_data["Total"][i])
        quotation_doc.add_page_break()
        quotation_doc.save(f'Quotation_{Number}.docx')
        """
        #TODO Export the quotation to the database and save the pdf to the database
        # Clear Dataframes
        client_data = pd.DataFrame()
        product_data = pd.DataFrame()
        return render_template("successful_submission.html"), {"Refresh": "5; url=/"}
    pass


# Create Page that allows admins to change the price of products
@app.route("/price", methods = ["GET", "POST"])
def price():
    rows = c.execute(text(f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"))
    prods = rows.all()
    if request.method == "POST":
        choice = request.form.get("price")
        if choice == "single_product":
            return render_template("single_product.html", products = prods)
        elif choice == "percentage":
            return render_template("percentage.html")
    return render_template("edit_prices.html")

# Change the price of a single product
@app.route("/single_price", methods = ["GET", "POST"])
def single_price():
    rows = c.execute(text(f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"))
    prods = rows.all()
    if request.method == "POST":
        code = request.form.get("product_code")
        price = request.form.get("price")
        c.execute(text(f'UPDATE product_list SET Price = "{price}" WHERE Product_Code = "{code}"'))
        c.commit()
        return render_template("index.html")
    return render_template("single_product.html", products = prods)

# Change all product prices by a constant percentage either by increasing or decreasing
@app.route("/percentage", methods = ["GET", "POST"])
def percentage():
    if request.method == "POST":
        percentage = request.form.get("percentage")
        Type = request.form.get("change_type")
        if not percentage:
            return render_template("Invalid_Choice.html")
        if Type == "Increase":
            new_percentage = 1 + (float(percentage) / 100.0)
        elif Type == "Decrease":
            new_percentage = 1 - (float(percentage) / 100.0)
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

""" 
# Unused Route, replaced, delete later
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
 """


"""
# Unused Route, replaced, delete later
def get_date():
    print("Please enter the date: ")
    date = input()
    return date
 """

""" 
# Unused Route, replaced, delete later
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
 """
""" 
# Unused Route, replaced, delete later
def get_quantity():
    print("Please enter the quantity: ")
    quantity = input()
    return quantity
 """
""" 
# Unused Route, replaced, delete later
def get_client_name():
    print("Please enter the client name: ")
    client = input()
    return client """
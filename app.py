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
import qrcode
from qrcode.image.styledpil import StyledPilImage
from qrcode.image.styles.moduledrawers.pil import RoundedModuleDrawer
from qrcode.image.styles.colormasks import RadialGradiantColorMask
import datetime
import os
import pythoncom
import win32com.client
import math
from dotenv import load_dotenv


# TODO: Learn what a logger is. (python logging module)


def initialize_com():
    pythoncom.CoInitialize()


#  Create engine
df = pd.read_excel("D:/Work/POOF/complete_product_list_spreadsheet_updated.xlsx")
current_quotation = []
current_client = {
    "Date": "",
    "Customer_Name": "",
    "Customer_Number": "",
    "Rep_Name": "",
    "Rep_Number": "",
}


# Load the environment variables
load_dotenv()
username = os.getenv("USER")
password = os.getenv("PASSWORD")
schema = os.getenv("SCHEMA")

quotation_dir = "static/quotations"
excel_quotation_dir = "static/editable_quotations"
if not os.path.exists(quotation_dir):
    os.mkdir(quotation_dir)


engine = create_engine(
    f"mysql+mysqlconnector://{username}:{password}@127.0.0.1/{schema}"
)
#df.to_sql(name="product_list", con=engine, index=False)


try:
    df.to_sql(name="product_list", con=engine, if_exists="fail", index=False)
except ValueError:
    pass


conn = engine.connect()

# Connect Flask
app = Flask(__name__)

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
    # Clear the current quotation and client info
    

    return render_template("index.html", name=name)


@app.route("/login", methods=["GET", "POST"])
def login():

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
        authorityres = conn.execute(
            text(f'SELECT * FROM authorized_personnel WHERE Employee_Name = "{name}"')
        )
        conn.commit()
        authority = authorityres.all()

        # Ensure username and password were submitted
        if not name or not Pass:
            print("Name or Pass is empty")
            return render_template("invalid_credentials.html", name=name, Pass=Pass)

        # Ensure username exists and password is correct
        if len(authority) != 1 or not check_password_hash(authority[0][2], Pass):
            print("Wrong name or wrong Password")
            return render_template("invalid_credentials.html", authority=authorityres)

        session["user_id"] = authority[0][0]
        session["access_level"] = authority[0][3]
        # Redirect user to appropriate page depending on the authority level
        if session["access_level"]:
            return redirect("/")
        return render_template("invalid_credentials.html")


@app.route("/register", methods=["GET", "POST"])
def register():

    if request.method == "POST":
        # Get username and password and authority level

        name = request.form.get("username")
        password = request.form.get("up")
        cpassword = request.form.get("up2")
        authority = request.form.get("authority")
        hash = generate_password_hash(password, method="pbkdf2", salt_length=16)
        # print(hash)
        if not name:
            return render_template("invalid_credentials.html")

        # Ensure password was submitted
        elif not password:
            return render_template("invalid_credentials.html")

        # Query database for username
        # TODO Protect against SQL injection
        result = conn.execute(
            text(f'SELECT * FROM authorized_personnel WHERE Employee_Name = "{name}"')
        )
        rows = result.all()

        # Ensure username exists and password is correct
        if len(rows) != 0:
            return render_template("invalid_credentials.html")

        if password != cpassword:
            return render_template("invalid_credentials.html")

        # TODO Protect against SQL injection
        conn.execute(
            text(
                f"""INSERT INTO authorized_personnel(Employee_Name, Password, Authority_Level) 
                VALUES ("{name}", "{hash}", "{authority}")"""
            )
        )
        conn.commit()

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

        match session["access_level"]:
            case "Developer":
                return render_template("admin_options.html")
            case "Administrator":
                return render_template("admin_options.html")
            case "Data_Entry":
                return render_template("customer_info.html")
            case _:
                return render_template("invalid_credentials.html")

    elif request.method == "POST":
        choice = request.form.get("task")
        # print(choice)
        match choice:
            case "create_quotation":
                return render_template("customer_info.html")
            case "Edit product prices":
                return render_template("edit_prices.html")
            case "Add new product":
                return render_template("add_product.html")
            case "Add Employee":
                return render_template("register.html")
            case _:
                return render_template("invalid_choice.html")


@app.route("/customer_info", methods=["GET", "POST"])
def Customer_Info():
    if request.method == "POST":
        current_client["Date"] = request.form.get("quotation_date")
        current_client["Customer_Name"] = request.form.get("customer_name")
        current_client["Customer_Number"] = request.form.get("customer_number")
        current_client["Rep_Name"] = request.form.get("rep_name")
        current_client["Rep_Number"] = request.form.get("rep_number")

        rows = conn.execute(
            text(
                f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"
            )
        )
        prods = rows.all()
        if (
            not current_client["Customer_Name"]
            or not current_client["Date"]
            or not current_client["Rep_Name"]
            or not current_client["Rep_Number"]
        ):
            return render_template("insufficient_data.html")
        return render_template(
            "create_quotation.html", products=prods, customer_info=current_client
        )
    pass


@app.route("/edit_quotation", methods=["GET", "POST"])
def edit_quotation():
    if request.method == "POST":
        rows = conn.execute(
            text(
                f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"
            )
        )
        prods = rows.all()
        return render_template(
            "create_quotation.html",
            products=prods,
            customer_info=current_client,
            entries=current_quotation,
        )


# TODO: FIX THE RANDOMIZED CASE IN PAGE ROUTES


# Create Quotation Page and handling Queries, autocomplete, and dynamic table row insertion
@app.route("/create_quotation", methods=["GET", "POST"])
def create_quotation():
    if request.method == "POST":
        # Get the product code and quantity from the submitted form
        code = request.form.get("product_code")
        quantity = request.form.get("quantity")
        if quantity is None:
            return render_template("insufficient_data.html")
        quantity = float(quantity)
        rows = conn.execute(
            text(f"SELECT * FROM product_list WHERE product_code = '{code}'")
        )
        prod = rows.all()
        # redo the sql table and replace it with a list of lists in the frontend instead of a SQL table

        if len(prod) == 0 or prod is None:
            return render_template("insufficient_data.html")

        # Add the product to the current quotation
        current_quotation.append(
            [
                prod[0][0], # Product Code
                prod[0][1], # Product Name
                prod[0][5], # Image
                prod[0][3], # Description
                prod[0][4], # Specs
                quantity,   # Quantity    
                prod[0][2], # Price
                prod[0][2] * quantity # Total
            ]
        )

        rows = conn.execute(
            text(
                f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"
            )
        )
        prods = rows.all()
        return render_template(
            "create_quotation.html",
            entries=current_quotation,
            products=prods,
            customer_info=current_client,
        )


# TODO Add a route that allows the user to delete a product from the quotation. This will require a new form for each table row in the HTML


@app.route("/preview", methods=["GET", "POST"])
def preview():
    # Render a preview page that shows the current quotation, an option to go back and edit, and an option to submit
    return render_template(
        "preview_quotation.html",
        customer_info=current_client,
        entries=current_quotation,
    )


def vacant_spots(sheet):
    vacants = 0
    num_rows_per_sheet = 8
    min_row = 14
    max_row = min_row + num_rows_per_sheet - 1
    min_col = 1
    max_col = 10
    rows = sheet.iter_rows(
        min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col
    )
    for row in rows:
        if row[0].value is None:
            vacants += 1
    return vacants


def add_small_product_to_page(product_specs, biggest_product, ws, vacancy, product_data):

    # The product has 8 or less lines
    # Check if the product can fit in the current page
    size = product_specs[biggest_product][1] + 1
    if size > vacancy:
        raise Exception("Not enough vacancy for product sent to add_small_product_to_page")
    for product in product_specs:
        if product == biggest_product:
            # Add the product name to the first cell of the page
            description = list(product_data.loc[product_data["Description"] == product]["Description"])[0]
            quantity = list(product_data.loc[product_data["Description"] == product]["Quantity"])[0]
            price = list(product_data.loc[product_data["Description"] == product]["Price"])[0]
            total = list(product_data.loc[product_data["Description"] == product]["Total"])[0]
            product_name = list(product_data.loc[product_data["Description"] == product]["Product_Name"])[0]
            if description == None:
                ws.cell(row=14, column=1).value = product_name
            else:
                ws.cell(row=14, column=1).value = product
            ws.cell(row=14, column=7).value = quantity
            ws.cell(row=14, column=8).value = price
            ws.cell(row=14, column=9).value = total
            ws.cell(row=14, column=10).value = product_name
            for i, spec in enumerate(product_specs[product][0]):
                ws.cell(row=15 + i + 8 - vacancy, column=1).value = spec
            break
    return True


def empty_page(ws):
    # Empty the page
    for row in ws.iter_rows(min_row=14, max_row=21, min_col=1, max_col=10):
        row[0].value = None
        row[6].value = None
        row[7].value = None
        row[8].value = None
        row[9].value = None
    return ws


def add_large_product_to_page(product_specs, biggest_product, ws, num_sheets, sheets, quotation_temp, product_data):
    # If the product has more than 8 lines, split it into multiple page
    # Add the first 8 lines of the product to the first page and remove them from the dictionary
    # Repeat the process for the rest of the pages
    current_page_number = len(sheets)  # The current page number
    # Add number of sheets needed to the sheets list
    for _ in range(num_sheets):
            ws = quotation_temp.copy_worksheet(ws)
            # Empty the page
            empty_page(ws)
            # Add the new page to the sheets list
            sheets.append(ws)

    for i in range(current_page_number, current_page_number + num_sheets, 1):
        ws = sheets[i]
        for product in product_specs:
            if product == biggest_product:
                # Add the product name to the first cell of the page
                # Only add the product name to the first page
                if i == current_page_number:
                    print(f"Product Data:\n {product_data.to_string()}")
                    description = list(product_data.loc[product_data["Description"] == product]["Description"])[0]
                    quantity = list(product_data.loc[product_data["Description"] == product]["Quantity"])[0]
                    price = list(product_data.loc[product_data["Description"] == product]["Price"])[0]
                    total = list(product_data.loc[product_data["Description"] == product]["Total"])[0]
                    product_name = list(product_data.loc[product_data["Description"] == product]["Product_Name"])[0]
                    if description == None:
                        ws.cell(row=14, column=1).value = product_data["Product_Name"][0]
                    else:
                        ws.cell(row=14, column=1).value = product
                    # Get Quantity, Price, Total, and Product Name from the product_data dictionary where product name is equal to product
                    
                    
                    ws.cell(row=14, column=7).value = quantity
                    ws.cell(row=14, column=8).value = price
                    ws.cell(row=14, column=9).value = total
                    ws.cell(row=14, column=10).value = product_name
                    # Add the first 7 lines of the specs to the page
                    for j in range(7):
                        ws.cell(row=15 + j, column=1).value = product_specs[product][0][j]
                    # Remove the first 7 lines from the product
                    product_specs[product][0] = product_specs[product][0][7:]
                    product_specs[product][1] -= 7
                # if it's not the first page, keep adding the specs to the page until the page is full or the product is finished
                else:
                    # If the product has less than 8 lines left, add them to the page
                    if product_specs[product][1] <= 8:
                        for j, spec in enumerate(product_specs[product][0]):
                            ws.cell(row=14 + j, column=1).value = spec
                        # Remove the specs from the product
                        product_specs[product][0] = []
                        product_specs[product][1] = 0
                        break
                    # If the product has more than 8 lines left, add the first 8 lines to the page
                    else:
                        for j in range(8):
                            ws.cell(row=14 + j, column=1).value = product_specs[product][0][j]
                        # Remove the first 8 lines from the product
                        product_specs[product][0] = product_specs[product][0][8:]
                        product_specs[product][1] -= 8
                        break

# Convert the current quotation to a dataframe, submit it to the database, and export it as an excel file
@app.route("/export", methods=["GET", "POST"])
def submit():
    # TODO: Break this down into atomic functions

    if request.method == "GET":
        pass
    
    client_columns = [
        "Date",
        "Customer_Name",
        "Customer_Number",
        "Rep_Name",
        "Rep_Number",
    ]
    product_columns = [

        "Product_Code",
        "Product_Name",
        "Image",
        "Description",
        "Specs",
        "Quantity",
        "Price",
        "Total",
    ]

    client_list = [
        [
            current_client["Date"],
            current_client["Customer_Name"],
            current_client["Customer_Number"],
            current_client["Rep_Name"],
            current_client["Rep_Number"],
        ]
    ]
    client_data = pd.DataFrame(client_list, columns=client_columns)
    product_data = pd.DataFrame(current_quotation, columns=product_columns)
    # Export the dataframes to an Excel file, then save the file as a pdf
    # and save the pdf to the database alongside the name of the user and the date of submission
    # Add serializtion to the quotation files
    # Clear the current quotation and client info
    current_quotation.clear()
    current_client.clear()
    no_of_quotations = conn.execute(
        text("SELECT MAX(quotation_id) FROM exported_quotations")
    ).all()[0][0]
    if not no_of_quotations:
        no_of_quotations = 0
    quotation_id = no_of_quotations + 1
    
    # Open Template file
    quotation_temp = openpyxl.load_workbook("D:/Work/POOF/Quotation_Template.xlsx")
    ws = quotation_temp.active
    sheets = [ws]
    Date_Cell = ws["I5"]
    Customer_Name_Cell = ws["A9"]
    Customer_Number_Cell = ws["C9"]
    Rep_Name_Cell = ws["A12"]
    Rep_Number_Cell = ws["C12"]
    Quotation_Number_Cell = ws["G5"]
    
    # Add the client data to the quotation template
    Date_Cell.value = client_data["Date"][0]
    Customer_Name_Cell.value = client_data["Customer_Name"][0]
    Customer_Number_Cell.value = client_data["Customer_Number"][0]
    Rep_Name_Cell.value = client_data["Rep_Name"][0]
    Rep_Number_Cell.value = client_data["Rep_Number"][0]
    Quotation_Number_Cell.value = quotation_id
    
    
    # Get the true number of rows by unpacking product data from the database
    product_specs = {}
    for code in product_data["Product_Code"]:
        rows = conn.execute(
            text(f"SELECT Description, Specs FROM product_list WHERE product_code = '{code}'")
        )
        specs = rows.all()
        # If there's no specs, set the product specs to None
        product_name = specs[0][0]
        if specs[0][1] == None:
            product_specs[product_name] = None
        # If there are specs, split them and store them in the product_specs dictionary, along with the number of specs as a list
        else:
            # product_specs["Product_Name"][0] is the product name, product_specs["Product_Name"][1] is the number of specs
            product_specs[product_name] = [specs[0][1].split("@"), len(specs[0][1].split("@"))]
    # Sort the product specs by the number of specs in descending order
    product_specs = dict(sorted(product_specs.items(), key=lambda item: item[1][1], reverse=True))
    
    while len(product_specs) > 0:

        # Check current page vacancy and save the number of vacant spots in a variable vacancy
        vacancy = vacant_spots(ws)
        # Check the biggest product in the dictionary and save its size in variable size and its name in variable largest_product
        biggest_product = list(product_specs.keys())[0]
        size = product_specs[biggest_product][1] + 1
        # If size is less than 8:

        if size <= 8:
            # If vacancy is greater than or equal to size, add the product to the current page and remove it from the dictionary
            if vacancy >= size:
                add_small_product_to_page(product_specs, biggest_product, ws, vacancy, product_data)
                del product_specs[biggest_product]
            # Elif vacancy is less than size:
            elif vacancy < size:
                found = False
                # Go through the rest of the products in the dictionary and check if any of them can fit in the current page
                for product in product_specs:
                    temp_size = product_specs[product][1] + 1
                    if temp_size <= vacancy:
                        # If a product is found, add it to the current page and remove it from the dictionary
                        add_small_product_to_page(product_specs, product, ws, vacancy, sheets, quotation_temp, product_data)
                        del product_specs[product]
                        found = True
                        break
                # If no product is found, move to the next page and add the product to the new page
                if not found:
                    ws = quotation_temp.copy_worksheet(ws)
                    sheets.append(ws)
                    # Empty the page
                    empty_page(ws)
        # Elif size is greater than 8:
        elif size > 8:
            # Create as many pages as needed to fit the product
            num_sheets = math.ceil(product_specs[biggest_product][1] / 8)
            # product_specs, biggest_product, ws, num_sheets, sheets, quotation_temp, product_data
            add_large_product_to_page(product_specs, biggest_product, ws, num_sheets, sheets, quotation_temp, product_data)
            del product_specs[biggest_product]
    # Go over the total column in each page and calculate the total for all pages
    total = 0
    for sheet in sheets:
        for row in sheet.iter_rows(min_row=14, max_row=21, min_col=1, max_col=10):
            if row[8].value is None:
                break
            print(f"Row: {row[8].value}")
            total += row[8].value
    # Add total to the last page only at cell I23
    ws["I23"].value = total
    

    

    pdf_path = submit_quotation_to_db(session["user_id"], quotation_temp, ws)
    print(f"PDF Path: {pdf_path}")

    # Clear Dataframes
    client_data = pd.DataFrame()
    product_data = pd.DataFrame()
    
    # Redirect the user to a page that allows them to download the quotation
    return redirect("/download")
    

#Create a page that the user is redirected to after submitting a quotation that allows them to download the PDF file
@app.route("/download", methods=["GET"])
def download():
    no_of_quotations = conn.execute(
        text("SELECT MAX(quotation_id) FROM exported_quotations")
    ).all()[0][0]
    if not no_of_quotations:
        no_of_quotations = 0
    quotation_id = no_of_quotations
    pdf_path = os.path.join(quotation_dir, f"{quotation_id}.pdf")
    print(f"PDF Path: {pdf_path}")
    return render_template("successful_submission.html", pdf_path=pdf_path)

# Create Page that allows admins to change the price of products
@app.route("/price", methods=["GET", "POST"])
def price():
    rows = conn.execute(
        text(
            f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"
        )
    )
    prods = rows.all()
    if request.method == "GET":
        return render_template("edit_prices.html")

    choice = request.form.get("price")
    if choice == "single_product":
        return render_template("single_product.html", products=prods)
    elif choice == "percentage":
        return render_template("percentage.html")


# Change the price of a single product
@app.route("/single_price", methods=["GET", "POST"])
def single_price():
    rows = conn.execute(
        text(
            f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"
        )
    )
    prods = rows.all()
    if request.method == "POST":
        code = request.form.get("product_code")
        price = request.form.get("price")
        conn.execute(
            text(
                f'UPDATE product_list SET Price = "{price}" WHERE Product_Code = "{code}"'
            )
        )
        conn.commit()
        return render_template("index.html")
    return render_template("single_product.html", products=prods)


# Change all product prices by a constant percentage either by increasing or decreasing
@app.route("/percentage", methods=["GET", "POST"])
def percentage():
    if request.method == "POST":
        percentage = request.form.get("percentage")
        Type = request.form.get("change_type")
        if not percentage:
            return render_template("invalid_choice.html")
        if Type == "Increase":
            new_percentage = 1 + (float(percentage) / 100.0)
        elif Type == "Decrease":
            new_percentage = 1 - (float(percentage) / 100.0)
        conn.execute(text(f"UPDATE product_list SET Price = Price * {new_percentage}"))
        conn.commit()
        return render_template("edit_prices.html")
    return render_template("percentage.html")


# Add a new product to the database
@app.route("/add_product", methods=["GET", "POST"])
def add_product():
    if request.method == "POST":
        code = request.form.get("code")
        name = request.form.get("name")
        price = request.form.get("price")
        description = request.form.get("description")
        image = request.form.get("image")
        conn.execute(
            text(
                f'INSERT INTO product_list(Product_Code, Product_Name, Price, Description, Image_Directory) VALUES ("{code}", "{name}", "{price}", "{description}", "{image}")'
            )
        )
        conn.commit()
        return render_template("admin_options.html")
    return render_template("add_product.html")

# View the Quotation
@app.route("/view", methods=["GET"])  # expection a url/view?quotation_id=1
def view_quotation():
    quotation_id = request.args.get("quotation_id")

    # Read the excel and put the values into the entries variable
    # Get the path of the quotation file
    path = os.path.join(excel_quotation_dir, f"{quotation_id}.xlsx")
    path = os.path.abspath(path)
    # Open the quotation file
    requested_quotation = openpyxl.load_workbook(path)
    sheet = requested_quotation.active
    # Assign the values of the cells to the client_data dictionary
    Date_Cell = sheet["I5"]
    Customer_Name_Cell = sheet["A9"]
    Customer_Number_Cell = sheet["C9"]
    Rep_Name_Cell = sheet["A12"]
    Rep_Number_Cell = sheet["C12"]
    Quotation_Number_Cell = sheet["G5"]

    # Create a dictionary to store the client data
    client_data = {
        "Date": None,
        "Customer_Name": None,
        "Customer_Number": None,
        "Rep_Name": None,
        "Rep_Number": None,
    }

    # Assign the values of the cells to the client_data dictionary
    client_data["Date"] = Date_Cell.value
    client_data["Customer_Name"] = Customer_Name_Cell.value
    client_data["Customer_Number"] = Customer_Number_Cell.value
    client_data["Rep_Name"] = Rep_Name_Cell.value
    client_data["Rep_Number"] = Rep_Number_Cell.value

    entries = []
    # Get the values of the cells and put them into the entries list
    rows = sheet.iter_rows(min_row=14, max_row=21, min_col=1, max_col=10)
    print("Rows: ", rows)
    for i, row in enumerate(rows):
        if not row[0].value and not row[9].value:
            break
        quantity = row[6].value

        # Get the Image_Directory from database
        description = row[0].value 
        if row[0].value == None:
            description = row[9].value
        dir = conn.execute(text(f"SELECT Image_Directory FROM product_list WHERE product_name = '{description}' OR Description = '{description}'"))
        directories = dir.all()
        conn.commit()
        print(f"Description: {description}")
        price = row[7].value
        total = row[8].value
        entries.append(
            [f"{directories[0][0]}", description, price, quantity, total]
        )

    total = 0
    vat = 0
    for entry in entries:
        total += entry[4]
    vat = total * 0.14
    total += vat
    quotation_pdf = os.path.join(quotation_dir, f"{quotation_id}.pdf")
    return render_template("view_quotation.html", entries=entries, total=total, vat=vat, quotation_pdf=quotation_pdf)


@app.route("/price_list", methods=["GET", "POST"])
def price_list():
    prods = conn.execute(text(f"SELECT product_code, product_name FROM product_list"))
    products = prods.all()
    conn.commit()

    if request.method == "POST":
        code = request.form.get("product_code")
        prods = conn.execute(
            text(
                f"SELECT product_code, product_name, price FROM product_list WHERE product_code = '{code}'"
            )
        )
        data = prods.all()
        if len(data) > 0:
            plist = [[data[0][0], data[0][1], data[0][2]]]
        conn.commit()

        return render_template("price_list.html", plist=plist, products=products)
    return render_template("price_list.html", products=products)


def submit_quotation_to_db(employee_id: int, quotations, sheet) -> bool:
    no_of_quotations = conn.execute(
        text("SELECT MAX(quotation_id) FROM exported_quotations")
    ).all()[0][0]
    if not no_of_quotations:
        no_of_quotations = 0

    quotation_id = no_of_quotations + 1
    submission_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    quotation_url = f"http://localhost:5000/view?quotation_id={quotation_id}"  # TODO: Change this to the actual URL
    quotation_file_path = os.path.join(excel_quotation_dir, f"{quotation_id}.xlsx")
    pdf_file_path = os.path.join(quotation_dir, f"{quotation_id}.pdf")
    conn.execute(
        text(
            f"INSERT INTO exported_quotations(submission_date, employee_id, quotation_url, quotation_file_path) VALUES ('{submission_date}', {employee_id}, '{quotation_url}', '{quotation_file_path}')"
        )
    )
    conn.commit()
    
    qr_code = convert_url_to_qr_code(quotation_url)
    path = f"{quotation_dir}/{quotation_id}.png"
    qr_code = qr_code.save(path, "PNG")
    print(type(qr_code))
    print(path)
    # Add QR code to the Excel

    img = openpyxl.drawing.image.Image(path)
    img.anchor = "C24"
    img.height = 100
    img.width = 100
    sheet.add_image(img)

    # Save the quotation to the file system
    quotations.save(filename=quotation_file_path)

    # Open Microsoft Excel

    quotation_file_path = os.path.abspath(quotation_file_path)
    pdf_file_path = os.path.abspath(pdf_file_path)

    initialize_com()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    sheets = excel.Workbooks.Open(quotation_file_path)
    # Get all worksheets
    work_sheets = sheets.Worksheets

    # Convert into PDF File
    # Merge all the worksheets into one PDF
    sheets.ExportAsFixedFormat(0, pdf_file_path)
    
    return pdf_file_path


def update_password(user_name: str, new_password) -> bool:
    # Make a SQL request to the DB to update the password hash for the user
    new_hash = generate_password_hash(new_password, method="pbkdf2", salt_length=16)

    query = f"UPDATE authorized_personnel SET Password = '{new_hash}' WHERE Employee_Name = '{user_name}'"
    response = conn.execute(text(query))
    conn.commit()
    return (
        response.rowcount == 1
    )  # If the row was updated, return True, else return False


# Return the image of the QR code to the calling function
def convert_url_to_qr_code(url: str, rounded_corners=True, logo_path=None) -> PIL.Image:
    qr = qrcode.QRCode(error_correction=qrcode.constants.ERROR_CORRECT_L)
    qr.add_data(url)

    if rounded_corners and logo_path:
        img = qr.make_image(
            image_factory=StyledPilImage,
            module_drawer=RoundedModuleDrawer(),
            embeded_image_path=logo_path,
        )
    elif rounded_corners:
        img = qr.make_image(
            image_factory=StyledPilImage, module_drawer=RoundedModuleDrawer()
        )
    elif logo_path:
        img = qr.make_image(image_factory=StyledPilImage, embeded_image_path=logo_path)
    else:
        img = qr.make_image()
    print(type(img))
    return img
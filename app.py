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

# TODO: Learn what a logger is. (python logging module)


def initialize_com():
    pythoncom.CoInitialize()


#  Create engine
df = pd.read_excel("D:/Work/POOF/complete_product_list_spreadsheet.xlsx")
current_quotation = []
current_client = {
    "Date": "",
    "Customer_Name": "",
    "Customer_Number": "",
    "Rep_Name": "",
    "Rep_Number": "",
}

# TODO: Use python-dotenv to store sensitive information in a .env file

username = "root"
password = "13579111315szxM"
quotation_dir = "quotations"

if not os.path.exists(quotation_dir):
    os.mkdir(quotation_dir)


engine = create_engine(
    f"mysql+mysqlconnector://{username}:{password}@127.0.0.1/poof_schema"
)
try:
    df.to_sql(name="product_list", con=engine, if_exists="fail", index=False)
except ValueError:
    pass
conn = engine.connect()

# Connect Flask
app = Flask(__name__)

# Read the Quotation Template content

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
    # Clear the current quotation and client info
    current_quotation.clear()  # TODO: IDK man.. do somthing
    current_client.clear()

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
        authorityres = conn.execute(
            text(f'SELECT * FROM authorized_personnel WHERE Employee_Name = "{name}"')
        )
        conn.commit()
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

    if request.method == "POST":
        # Get username and password and authority level

        name = request.form.get("username")
        password = request.form.get("up")
        Cpassword = request.form.get("up2")
        authority = request.form.get("authority")
        hash = generate_password_hash(password, method="pbkdf2", salt_length=16)
        # print(hash)
        if not name:
            return render_template("Invalid_Credentials.html")

        # Ensure password was submitted
        elif not password:
            return render_template("Invalid_Credentials.html")

        # Query database for username
        # TODO Protect against SQL injection
        result = conn.execute(
            text(f'SELECT * FROM authorized_personnel WHERE Employee_Name = "{name}"')
        )
        rows = result.all()

        # Ensure username exists and password is correct
        if len(rows) != 0:
            return render_template("Invalid_Credentials.html")

        if password != Cpassword:
            return render_template("Invalid_Credentials.html")

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
                return render_template("Invalid_Credentials.html")

    elif request.method == "POST":
        choice = request.form.get("task")
        # print(choice)
        match choice:
            case "Create_quotation":
                return render_template("customer_info.html")
            case "Edit product prices":
                return render_template("edit_prices.html")
            case "Add new product":
                return render_template("add_product.html")
            case "Add Employee":
                return render_template("register.html")
            case _:
                return render_template("Invalid_Choice.html")


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
            "Create_Quotation.html", products=prods, customer_info=current_client
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
            "Create_Quotation.html",
            products=prods,
            customer_info=current_client,
            entries=current_quotation,
        )


# TODO: FIX THE RANDOMIZED CASE IN PAGE ROUTES


# Create Quotation Page and handling Queries, autocomplete, and dynamic table row insertion
@app.route("/Create_Quotation", methods=["GET", "POST"])
def Create_Quotation():
    if request.method == "POST":
        # Get the product code and quantity from the submitted form
        code = request.form.get("product_code")
        quantity = request.form.get("quantity")
        quantity = float(quantity)  # TODO: validate that quantity is not None
        rows = conn.execute(
            text(f"SELECT * FROM product_list WHERE product_code = '{code}'")
        )
        prod = rows.all()
        # redo the sql table and replace it with a list of lists in the frontend instead of a SQL table

        if len(prod) == 0 or prod is None:
            return render_template("insufficient_data.html")

        ## TODO: Addd comments on wtf this does
        current_quotation.append(
            [
                prod[0][4],
                prod[0][0],
                prod[0][1],
                prod[0][3],
                quantity,
                prod[0][2],
                prod[0][2] * quantity,
            ]
        )

        rows = conn.execute(
            text(
                f"SELECT product_code, product_name FROM product_list ORDER BY product_code ASC"
            )
        )
        prods = rows.all()
        return render_template(
            "Create_Quotation.html",
            entries=current_quotation,
            products=prods,
            customer_info=current_client,
        )


# TODO Add a route that allows the user to delete a product from the quotation. This will require a new form for each table row in the HTML


@app.route("/preview", methods=["GET", "POST"])
def preview():
    # TODO Add a preview page that shows the current quotation, an option to go back and edit, and an option to submit
    return render_template(
        "preview_quotation.html",
        customer_info=current_client,
        entries=current_quotation,
    )


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
        "Image",
        "Product_Code",
        "Product_Name",
        "Description",
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
    """ with pd.ExcelWriter("D:/Work/POOF/Quotation.xlsx") as writer:
        client_data.to_excel(writer, sheet_name = "Client_Info", index = True)
        product_data.to_excel(writer, sheet_name = "Product_Info", index = True) """
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
    quotation = openpyxl.load_workbook("D:/Work/POOF/Quotation_Template.xlsx")
    sheet = quotation.active
    Date_Cell = sheet["I5"]
    Customer_Name_Cell = sheet["A9"]
    Customer_Number_Cell = sheet["C9"]
    Rep_Name_Cell = sheet["A12"]
    Rep_Number_Cell = sheet["C12"]
    Quotation_Number_Cell = sheet["G5"]
    # Add the client data to the quotation template
    Date_Cell.value = client_data["Date"][0]
    Customer_Name_Cell.value = client_data["Customer_Name"][0]
    Customer_Number_Cell.value = client_data["Customer_Number"][0]
    Rep_Name_Cell.value = client_data["Rep_Name"][0]
    Rep_Number_Cell.value = client_data["Rep_Number"][0]
    Quotation_Number_Cell.value = quotation_id
    # Add the product data to the quotation template
    rows = sheet.iter_rows(
        min_row=14, max_row=14 + len(product_data), min_col=1, max_col=9
    )
    print("Rows: ", rows)
    for i, row in enumerate(rows):
        if i == len(product_data):
            break
        row[6].value = product_data["Quantity"][i]
        row[0].value = product_data["Product_Name"][i]
        row[2].value = product_data["Description"][i]
        row[7].value = product_data["Price"][i]
        row[8].value = product_data["Total"][i]

    submit_quotation_to_db(session["user_id"], quotation, sheet)

    # quotation.save(filename=f"Quotation_{Number}.xlsx")

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
    # Clear Dataframes
    client_data = pd.DataFrame()
    product_data = pd.DataFrame()
    return render_template("successful_submission.html"), {"Refresh": "5; url=/"}


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
            return render_template("Invalid_Choice.html")
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
        image = request.form.get("image")  # TODO Convert image file to directory
        conn.execute(
            text(
                f'INSERT INTO product_list(Product_Code, Product_Name, Price, Description, Image_Directory) VALUES ("{code}", "{name}", "{price}", "{description}", "{image}")'
            )
        )
        conn.commit()
        return render_template("admin_options.html")
    return render_template("add_product.html")


################################################################################################
########################################### LUK LUK ############################################
################################################################################################


# View the Quotation
@app.route("/view", methods=["GET"])  # expection a url/view?quotation_id=1
def view_quotation():
    quotation_id = request.args.get("quotation_id")

    ## TODO: Read the excel and put the values into the entries variable
    path = os.path.join(quotation_dir, f"{quotation_id}.xlsx")
    path = os.path.abspath(path)

    requested_quotation = openpyxl.load_workbook(path)
    sheet = requested_quotation.active
    Date_Cell = sheet["I5"]
    Customer_Name_Cell = sheet["A9"]
    Customer_Number_Cell = sheet["C9"]
    Rep_Name_Cell = sheet["A12"]
    Rep_Number_Cell = sheet["C12"]
    Quotation_Number_Cell = sheet["G5"]

    client_data = {
        "Date": None,
        "Customer_Name": None,
        "Customer_Number": None,
        "Rep_Name": None,
        "Rep_Number": None,
    }

    client_data["Date"] = Date_Cell.value
    client_data["Customer_Name"] = Customer_Name_Cell.value
    client_data["Customer_Number"] = Customer_Number_Cell.value
    client_data["Rep_Name"] = Rep_Name_Cell.value
    client_data["Rep_Number"] = Rep_Number_Cell.value

    entries = []
    rows = sheet.iter_rows(min_row=14, max_row=21, min_col=1, max_col=9)
    print("Rows: ", rows)
    for i, row in enumerate(rows):
        if not row[0].value:
            break
        quantity = row[6].value
        product_name = row[0].value
        description = row[2].value
        price = row[7].value
        total = row[8].value
        entries.append(
            [f"{product_name}.png", product_name, description, price, quantity, total]
        )

    total = 0
    vat = 0
    print(f"Entries are {entries}")
    for entry in entries:
        total += entry[5]
    vat = total * 0.14
    total += vat

    return render_template("view_quotation.html", entries=entries, total=total, vat=vat)


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


def submit_quotation_to_db(employee_id: int, quotation, sheet) -> bool:
    no_of_quotations = conn.execute(
        text("SELECT MAX(quotation_id) FROM exported_quotations")
    ).all()[0][0]
    if not no_of_quotations:
        no_of_quotations = 0

    quotation_id = no_of_quotations + 1
    submission_date = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    quotation_url = f"http://localhost:5000/view?quotation_id={quotation_id}"  # TODO: Change this to the actual URL
    quotation_file_path = os.path.join(quotation_dir, f"{quotation_id}.xlsx")
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
    quotation.save(filename=quotation_file_path)

    # TODO: Convert the excel to pdf and save as quotation_dir/quotation_id.pdf
    # Open Microsoft Excel

    quotation_file_path = os.path.abspath(quotation_file_path)
    pdf_file_path = os.path.abspath(pdf_file_path)

    initialize_com()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False

    sheets = excel.Workbooks.Open(quotation_file_path)
    work_sheets = sheets.Worksheets[0]

    # Convert into PDF File
    work_sheets.ExportAsFixedFormat(0, pdf_file_path)

    """ excel = client.Dispatch("Excel.Application")    
    # Read Excel File
    sheets = excel.Workbooks.Open(quotation_file_path)
    work_sheets = sheets.Worksheets[0]

    # Convert into PDF file
    work_sheets.ExportAsFixedFormat(0, pdf_file_path) """
    return True


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

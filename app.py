from flask import Flask, render_template, request, redirect, url_for
from flask_mysqldb import MySQL
from datetime import date, timedelta

from flask import send_file
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

from openpyxl import Workbook

app = Flask(__name__)

# ---------------- MySQL Configuration ----------------
app.config["MYSQL_HOST"] = "localhost"
app.config["MYSQL_USER"] = "root"
app.config["MYSQL_PASSWORD"] = ""
app.config["MYSQL_DB"] = "pims_db"

mysql = MySQL(app)

# ---------------- Login Credentials ----------------
USERNAME = "Purvashree14"
PASSWORD = "Purvashree07"

# ---------------- Login ----------------
@app.route("/", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        username = request.form["username"]
        password = request.form["password"]

        if username == USERNAME and password == PASSWORD:
            return redirect(url_for("dashboard"))
        else:
            return render_template(
                "login.html",
                error="Invalid Username or Password"
            )

    return render_template("login.html")


# ---------------- Dashboard ----------------
@app.route("/dashboard")
def dashboard():

    cur = mysql.connection.cursor()

    # Total Drugs
    cur.execute("SELECT COUNT(*) FROM drugs")
    total_drugs = cur.fetchone()[0]

    # Total Stock
    cur.execute("SELECT SUM(stock) FROM drugs")
    total_stock = cur.fetchone()[0] or 0

    # Low Stock
    cur.execute("""
        SELECT drug_name, stock
        FROM drugs
        WHERE stock <= 10
        ORDER BY stock ASC
    """)
    low_stock = cur.fetchall()

    # Expiry Alert
    today = date.today()
    next_30_days = today + timedelta(days=30)

    cur.execute("""
        SELECT drug_name, expiry_date
        FROM drugs
        WHERE expiry_date BETWEEN %s AND %s
        ORDER BY expiry_date ASC
    """, (today, next_30_days))

    expiry_alert = cur.fetchall()

    # Total Sales
    cur.execute("SELECT COUNT(*) FROM sales")
    total_sales = cur.fetchone()[0]

    # All Drugs
    cur.execute("SELECT * FROM drugs ORDER BY id DESC")
    drugs = cur.fetchall()

    # Chart Data
    labels = []
    stock_data = []

    for drug in drugs:
        labels.append(drug[1])      # Drug Name
        stock_data.append(drug[8])  # Stock

    cur.close()

    return render_template(
        "dashboard.html",
        total_drugs=total_drugs,
        total_stock=total_stock,
        low_stock=low_stock,
        expiry_alert=expiry_alert,
        total_sales=total_sales,
        drugs=drugs,
        labels=labels,
        stock_data=stock_data,
        
    )

# ---------------- Add Drug ----------------
@app.route("/add-drug", methods=["GET", "POST"])
def add_drug():

    if request.method == "POST":

        drug_name = request.form["drug_name"]
        brand_name = request.form["brand_name"]
        pharmacological_class = request.form["pharmacological_class"]
        schedule = request.form["schedule"]
        dosage_form = request.form["dosage_form"]
        storage = request.form["storage"]
        category = request.form["category"]
        stock = request.form["stock"]
        expiry_date = request.form["expiry_date"]
        price = request.form["price"]

        cur = mysql.connection.cursor()

        cur.execute("""
        INSERT INTO drugs
        (drug_name, brand_name, pharmacological_class, schedule,
        dosage_form, storage, category, stock, expiry_date, price)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, (
            drug_name,
            brand_name,
            pharmacological_class,
            schedule,
            dosage_form,
            storage,
            category,
            stock,
            expiry_date,
            price
        ))

        mysql.connection.commit()
        cur.close()

        return redirect(url_for("dashboard"))

    return render_template("add_drug.html")


# ---------------- Edit Drug ----------------
@app.route("/edit-drug/<int:id>", methods=["GET", "POST"])
def edit_drug(id):

    cur = mysql.connection.cursor()

    if request.method == "POST":

        drug_name = request.form["drug_name"]
        brand_name = request.form["brand_name"]
        pharmacological_class = request.form["pharmacological_class"]
        schedule = request.form["schedule"]
        dosage_form = request.form["dosage_form"]
        storage = request.form["storage"]
        category = request.form["category"]
        stock = request.form["stock"]
        expiry_date = request.form["expiry_date"]
        price = request.form["price"]

        cur.execute("""
        UPDATE drugs
        SET
            drug_name=%s,
            brand_name=%s,
            pharmacological_class=%s,
            schedule=%s,
            dosage_form=%s,
            storage=%s,
            category=%s,
            stock=%s,
            expiry_date=%s,
            price=%s
        WHERE id=%s
        """, (
            drug_name,
            brand_name,
            pharmacological_class,
            schedule,
            dosage_form,
            storage,
            category,
            stock,
            expiry_date,
            price,
            id
        ))

        mysql.connection.commit()
        cur.close()

        return redirect(url_for("dashboard"))

    cur.execute("SELECT * FROM drugs WHERE id=%s", (id,))
    drug = cur.fetchone()
    cur.close()

    return render_template("edit_drug.html", drug=drug)

# ---------------- Delete Drug ----------------
@app.route("/delete-drug/<int:id>")
def delete_drug(id):

    cur = mysql.connection.cursor()

    cur.execute("DELETE FROM drugs WHERE id=%s", (id,))
    mysql.connection.commit()

    cur.close()

    return redirect(url_for("dashboard"))


@app.route("/get-price/<drug_name>")
def get_price(drug_name):

    cur = mysql.connection.cursor()

    cur.execute(
        "SELECT price, stock FROM drugs WHERE drug_name=%s",
        (drug_name,)
    )

    data = cur.fetchone()

    cur.close()

    if data:
        return {
            "price": float(data[0]),
            "stock": int(data[1])
        }

    return {
        "price": 0,
        "stock": 0
    }

@app.route("/billing", methods=["GET", "POST"])
def billing():

    cur = mysql.connection.cursor()

    if request.method == "POST":

        customer_name = request.form["customer_name"]
        drug_name = request.form["drug_name"]
        quantity = int(request.form["quantity"])

        cur.execute(
            "SELECT price, stock FROM drugs WHERE drug_name=%s",
            (drug_name,)
        )

        drug = cur.fetchone()

        if drug:

            price = float(drug[0])
            stock = int(drug[1])

            if quantity > stock:
                cur.close()
                return "Not enough stock!"

            total = price * quantity

            cur.execute("""
                INSERT INTO sales
                (customer_name, drug_name, quantity, total_amount, sale_date)
                VALUES (%s,%s,%s,%s,NOW())
            """, (customer_name, drug_name, quantity, total))

            cur.execute("""
                UPDATE drugs
                SET stock = stock-%s
                WHERE drug_name=%s
            """, (quantity, drug_name))

            cur.execute("""
            SELECT id, drug_name, brand_name, category,
stock, expiry_date, price
FROM drugs
""")


            mysql.connection.commit()

    cur.execute("SELECT * FROM drugs")
    drugs = cur.fetchall()

    cur.close()

    return render_template("billing.html", drugs=drugs)

# ---------------- Sales ----------------
@app.route("/sales")
def sales():

    cur = mysql.connection.cursor()

    cur.execute("SELECT * FROM sales ORDER BY id DESC")
    sales = cur.fetchall()

    cur.close()

    return render_template("sales.html", sales=sales)

# ---------- Add Sale ----------
@app.route("/add-sale", methods=["GET", "POST"])
def add_sale():

    cur = mysql.connection.cursor()

    cur.execute("SELECT drug_name, price, stock FROM drugs ORDER BY drug_name")
    drugs = cur.fetchall()

    if request.method == "POST":

        customer_name = request.form["customer_name"]
        drug_name = request.form["drug_name"]
        quantity = int(request.form["quantity"])

        cur.execute(
            "SELECT price, stock FROM drugs WHERE drug_name=%s",
            (drug_name,)
        )
        drug = cur.fetchone()

        if not drug:
            cur.close()
            return "Drug not found!"

        price = float(drug[0])
        stock = int(drug[1])

        if quantity > stock:
            cur.close()
            return "Not enough stock available!"

        total_amount = quantity * price

        cur.execute("""
        INSERT INTO sales
        (customer_name, drug_name, quantity, total_amount)
        VALUES (%s,%s,%s,%s)
        """, (
            customer_name,
            drug_name,
            quantity,
            total_amount
        ))

        cur.execute("""
        UPDATE drugs
        SET stock = stock - %s
        WHERE drug_name = %s
        """, (
            quantity,
            drug_name
        ))

        mysql.connection.commit()
        cur.close()

        return redirect(url_for("sales"))

    cur.close()

    return render_template("add_sale.html", drugs=drugs)

      

# ---------------- Edit Sale ----------------
@app.route("/edit-sale/<int:id>", methods=["GET", "POST"])
def edit_sale(id):

    cur = mysql.connection.cursor()

    if request.method == "POST":

        customer_name = request.form["customer_name"]
        drug_name = request.form["drug_name"]
        quantity = int(request.form["quantity"])
        total_amount = float(request.form["total_amount"])

        cur.execute("""
        UPDATE sales
        SET
            customer_name=%s,
            drug_name=%s,
            quantity=%s,
            total_amount=%s
        WHERE id=%s
        """, (
            customer_name,
            drug_name,
            quantity,
            total_amount,
            id
        ))

        mysql.connection.commit()
        cur.close()

        return redirect(url_for("sales"))

    cur.execute("SELECT * FROM sales WHERE id=%s", (id,))
    sale = cur.fetchone()

    cur.close()

    return render_template("edit_sale.html", sale=sale)


# ---------------- Delete Sale ----------------
@app.route("/delete-sale/<int:id>")
def delete_sale(id):

    cur = mysql.connection.cursor()

    cur.execute("DELETE FROM sales WHERE id=%s", (id,))
    mysql.connection.commit()

    cur.close()

    return redirect(url_for("sales"))


@app.route("/invoice/<int:id>")
def invoice(id):

    cur = mysql.connection.cursor()

    cur.execute("""
        SELECT
            customer_name,
            drug_name,
            quantity,
            total_amount,
            sale_date
        FROM sales
        WHERE id=%s
    """, (id,))

    invoice = cur.fetchone()

    cur.close()

    if not invoice:
        return "Invoice not found!"

    return render_template(
        "invoice.html",
        invoice=invoice
    )

# ---------- Reports ----------
@app.route("/reports")
def reports():

    search = request.args.get("search", "")
    from_date = request.args.get("from_date", "")
    to_date = request.args.get("to_date", "")

    cur = mysql.connection.cursor()

    query = """
    SELECT
        id,
        drug_name,
        stock,
        (stock * price) AS subtotal,
        expiry_date
    FROM drugs
    WHERE 1=1
    """

    params = []

    if search:
        query += " AND drug_name LIKE %s"
        params.append("%" + search + "%")

    if from_date:
        query += " AND expiry_date >= %s"
        params.append(from_date)

    if to_date:
        query += " AND expiry_date <= %s"
        params.append(to_date)

    query += " ORDER BY expiry_date ASC"

    cur.execute(query, tuple(params))

    reports = cur.fetchall()

    cur.close()

    return render_template(
        "reports.html",
        reports=reports,
        search=search,
        from_date=from_date,
        to_date=to_date
    )
    
@app.route("/reports/pdf")
def report_pdf():

    cur = mysql.connection.cursor()

    cur.execute("""
        SELECT drug_name, stock, price, (stock * price)
        FROM drugs
        ORDER BY drug_name
    """)

    data = cur.fetchall()
    cur.close()

    pdf_file = "PIMS_Report.pdf"

    doc = SimpleDocTemplate(pdf_file)

    table_data = [["Drug Name", "Stock", "Price", "Subtotal"]]

    for row in data:
        table_data.append(list(row))

    table = Table(table_data)

    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.green),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("GRID", (0,0), (-1,-1), 1, colors.black),
        ("BACKGROUND", (0,1), (-1,-1), colors.beige),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
    ]))

    doc.build([table])

    return send_file(pdf_file, as_attachment=True)



@app.route("/reports/excel")
def export_excel():

    wb = Workbook()
    ws = wb.active
    ws.title = "PIMS Report"

    ws.append([
        "Drug Name",
        "Quantity",
        "Price",
        "Subtotal",
        "Expiry Date"
    ])

    cur = mysql.connection.cursor()

    cur.execute("""
        SELECT
            drug_name,
            stock,
            price,
            (stock*price),
            expiry_date
        FROM drugs
        ORDER BY drug_name
    """)

    for row in cur.fetchall():
        ws.append(row)

    cur.close()

    filename = "PIMS_Report.xlsx"
    wb.save(filename)

    return send_file(
        filename,
        as_attachment=True,
        download_name="PIMS_Report.xlsx"
    )

# ---------------- Test Database ----------------
@app.route("/testdb")
def testdb():

    cur = mysql.connection.cursor()
    cur.execute("SELECT 1")
    cur.close()

    return "MySQL Connected Successfully!"


# ---------------- Run App ----------------
if __name__ == "__main__":
    app.run(debug=True)
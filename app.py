from flask import Flask, render_template, request, redirect, url_for
import pymysql

pymysql.install_as_MySQLdb()

from flask_mysqldb import MySQL
from datetime import date, timedelta

from flask import send_file
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

from openpyxl import Workbook

app = Flask(__name__)

# ---------------- MySQL Configuration ----------------
import os

app.config["MYSQL_HOST"] = "localhost"
app.config["MYSQL_USER"] = "root"
app.config["MYSQL_PASSWORD"] = ""
app.config["MYSQL_DB"] = "pims_db"
app.config["MYSQL_PORT"] = 3306

mysql = MySQL(app)

# ---------------- Login Credentials ----------------
USERNAME = "Purvashree14"
PASSWORD = "Purvashree07"

# ==========================
# Login
# ==========================
@app.route("/", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        username = request.form["username"]
        password = request.form["password"]

        if username == "Purvashree14" and password == "Purvashree07":
            return redirect(url_for("dashboard"))
        else:
            return render_template(
                "login.html",
                error="Invalid Username or Password"
            )

    return render_template("login.html")

# ==========================
# Dashboard
# ==========================

@app.route("/dashboard")
def dashboard():

    cur = mysql.connection.cursor()

    # Total Drugs
    cur.execute("SELECT COUNT(*) FROM drugs")
    total_drugs = cur.fetchone()[0]

    # Total Stock
    cur.execute("SELECT IFNULL(SUM(stock),0) FROM drugs")
    total_stock = cur.fetchone()[0]

    # Low Stock
    cur.execute("SELECT COUNT(*) FROM drugs WHERE stock <= 10")
    low_stock = cur.fetchone()[0]

    # Expiry Alert
    cur.execute("SELECT COUNT(*) FROM drugs WHERE expiry_date <= CURDATE() + INTERVAL 30 DAY")
    expiry_alert = cur.fetchone()[0]

    # Total Sales
    cur.execute("SELECT IFNULL(SUM(total_amount),0) FROM sales")
    total_sales = cur.fetchone()[0]

    # Recent Drugs
    cur.execute("""
        SELECT drug_name, stock
        FROM drugs
        ORDER BY id DESC
        LIMIT 5
    """)
    drugs = cur.fetchall()

    labels = [row[0] for row in drugs]
    stock_data = [row[1] for row in drugs]

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
        stock_data=stock_data
    )

        # ==========================
        # Add Drug
        # ==========================

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
            (
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
            )
            VALUES
            (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """,(
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

        return redirect(url_for("drugs"))

    return render_template("add_drug.html")

    # ==========================
     # Drug Management
     # ==========================

@app.route("/drugs")
def drugs():

    cur = mysql.connection.cursor()

    cur.execute("""
        SELECT
            id,
            drug_name,
            brand_name,
            category,
            stock,
            expiry_date,
            price
        FROM drugs
        ORDER BY id DESC
    """)

    drugs = cur.fetchall()

    cur.close()

    return render_template(
        "drugs.html",
        drugs=drugs
    )

 # ==========================
# Edit Drug
# ==========================

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

        return redirect(url_for("drugs"))

    cur.execute("SELECT * FROM drugs WHERE id=%s", (id,))
    drug = cur.fetchone()

    cur.close()

    return render_template("edit_drug.html", drug=drug)


  # ==========================
# Delete Drug
# ==========================

@app.route("/delete-drug/<int:id>")
def delete_drug(id):

    cur = mysql.connection.cursor()

    cur.execute(
        "DELETE FROM drugs WHERE id=%s",
        (id,)
    )

    mysql.connection.commit()

    cur.close()

    return redirect(url_for("drugs"))


  # ==========================
# Sales Management
# ==========================

@app.route("/sales")
def sales():

    cur = mysql.connection.cursor()

    cur.execute("""
        SELECT
            id,
            invoice_no,
            customer_name,
            drug_name,
            quantity,
            price,
            total_amount,
            payment_method,
            sale_date,
            status
        FROM sales
        ORDER BY id DESC
    """)

    sales = cur.fetchall()

    cur.execute("""
        SELECT drug_name, price
        FROM drugs
        ORDER BY drug_name
    """)

    drugs = cur.fetchall()

    cur.close()

    return render_template(
        "sales.html",
        sales=sales,
        drugs=drugs
    )


# ==========================
# Save Sale
# ==========================

@app.route("/save_sale", methods=["POST"])
def save_sale():

    customer_name = request.form["customer_name"]
    drug_name = request.form["drug_name"]
    quantity = int(request.form["quantity"])
    price = float(request.form["price"])
    payment_method = request.form["payment_method"]

    total_amount = quantity * price

    cur = mysql.connection.cursor()

    # Generate Invoice Number
    cur.execute("SELECT COUNT(*) FROM sales")
    count = cur.fetchone()[0] + 1
    invoice_no = f"INV-{1000 + count}"

    # Save Sale
    cur.execute("""
        INSERT INTO sales
        (
            invoice_no,
            customer_name,
            drug_name,
            quantity,
            price,
            total_amount,
            payment_method,
            sale_date,
            status
        )
        VALUES
        (%s,%s,%s,%s,%s,%s,%s,CURDATE(),'Paid')
    """, (
        invoice_no,
        customer_name,
        drug_name,
        quantity,
        price,
        total_amount,
        payment_method
    ))

    # Reduce Stock
    cur.execute("""
        UPDATE drugs
        SET stock = stock - %s
        WHERE drug_name = %s
    """, (quantity, drug_name))

    mysql.connection.commit()
    cur.close()

    return redirect(url_for("sales"))

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
        SELECT *
        FROM sales
        WHERE id=%s
    """, (id,))

    invoice = cur.fetchone()

    cur.close()

    return render_template(
        "invoice.html",
        invoice=invoice
    )

# ==========================
# Reports
# ==========================

@app.route("/reports")
def reports():

    cur = mysql.connection.cursor()

    # Total Drugs
    cur.execute("SELECT COUNT(*) FROM drugs")
    total_drugs = cur.fetchone()[0]

    # Total Sales
    try:
        cur.execute("SELECT COUNT(*) FROM sales")
        total_sales = cur.fetchone()[0]
    except:
        total_sales = 0

    # Total Revenue
    try:
        cur.execute("SELECT SUM(total_amount) FROM sales")
        revenue = cur.fetchone()[0]
        total_revenue = revenue if revenue else 0
    except:
        total_revenue = 0

    # Low Stock Count
    cur.execute("SELECT COUNT(*) FROM drugs WHERE stock <= 10")
    low_stock_count = cur.fetchone()[0]

    # Low Stock Drugs
    cur.execute("""
        SELECT
            id,
            drug_name,
            brand_name,
            category,
            stock
        FROM drugs
        WHERE stock <= 10
        ORDER BY stock ASC
    """)
    low_stock_drugs = cur.fetchall()

    # Expiry Drugs
    cur.execute("""
        SELECT
            id,
            drug_name,
            brand_name,
            category,
            stock,
            expiry_date
        FROM drugs
        ORDER BY expiry_date ASC
        LIMIT 10
    """)
    expiry_drugs = cur.fetchall()

    cur.close()

    return render_template(
        "reports.html",
        total_drugs=total_drugs,
        total_sales=total_sales,
        total_revenue=total_revenue,
        low_stock_count=low_stock_count,
        low_stock_drugs=low_stock_drugs,
        expiry_drugs=expiry_drugs
    )

# ==========================
# Settings
# ==========================

@app.route("/settings", methods=["GET", "POST"])
def settings():

    cur = mysql.connection.cursor()

    if request.method == "POST":
        full_name = request.form["full_name"]
        email = request.form["email"]
        phone = request.form["phone"]

        cur.execute("""
            UPDATE admin_profile
            SET full_name=%s,
                email=%s,
                phone=%s
            WHERE id=1
        """, (full_name, email, phone))

        mysql.connection.commit()

    cur.execute("SELECT * FROM admin_profile WHERE id=1")
    profile = cur.fetchone()

    cur.execute("SELECT * FROM pharmacy_settings WHERE id=1")
    pharmacy = cur.fetchone()

    cur.execute("SELECT * FROM notification_settings WHERE id=1")
    notification = cur.fetchone()

    cur.execute("SELECT * FROM theme_settings WHERE id=1")
    theme = cur.fetchone()

    cur.close()

    return render_template(
        "settings.html",
        profile=profile,
        pharmacy=pharmacy,
        notification=notification,
        theme=theme
    )

@app.route("/reports/pdf")
def report_pdf():

    return send_file(
        "PIMS_Report.pdf",
        as_attachment=True
    )

@app.route("/reports/excel")
def export_excel():

    return send_file(
        "PIMS_Report.xlsx",
        as_attachment=True
    )

@app.route("/testdb")
def testdb():

    cur = mysql.connection.cursor()

    cur.execute("SELECT 1")

    cur.close()

    return "MySQL Connected Successfully!"

if __name__ == "__main__":
    app.run(debug=True)
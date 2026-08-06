from flask import Flask, render_template, request, redirect, url_for, send_file
import pymysql
from datetime import date, timedelta
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from openpyxl import Workbook

app = Flask(__name__)

# ---------------- TiDB / MySQL Configuration ----------------

DB_CONFIG = {
    "host": "gateway01.ap-southeast-1.prod.aws.tidbcloud.com",
    "user": "9uePyKrdnmUXdoN.root",
    "password": "53pSEKwGkdVb7sdl",
    "database": "test",
    "port": 4000,
    "charset": "utf8mb4",
    "cursorclass": pymysql.cursors.Cursor,
    "autocommit": True
}

def get_connection():
    return pymysql.connect(**DB_CONFIG)

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

        if username == USERNAME and password == PASSWORD:
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

    conn = get_connection()
    cur = conn.cursor()

    # Total Drugs
    cur.execute("SELECT COUNT(*) FROM drugs")
    total_drugs = cur.fetchone()[0]

    # Total Stock
    cur.execute("SELECT IFNULL(SUM(total_stock),0) FROM drugs")
    total_stock = cur.fetchone()[0]

    # Low Stock
    cur.execute("SELECT COUNT(*) FROM drugs WHERE total_stock <= reorder_level")
    low_stock = cur.fetchone()[0]

    # Expiry Alert
    cur.execute("""
        SELECT COUNT(*)
        FROM drugs
        WHERE expiry_date <= CURDATE() + INTERVAL 30 DAY
    """)
    expiry_alert = cur.fetchone()[0]

    # Total Sales
    cur.execute("SELECT COUNT(*) FROM sales")
    total_sales = cur.fetchone()[0]

    # Recent Drugs Chart
    cur.execute("""
        SELECT drug_name, total_stock
        FROM drugs
        ORDER BY id DESC
        LIMIT 5
    """)

    drugs = cur.fetchall()

    labels = [row[0] for row in drugs]
    stock_data = [row[1] for row in drugs]

    cur.close()
    conn.close()

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
        total_stock = request.form["total_stock"]
        expiry_date = request.form["expiry_date"]
        reorder_level = request.form["reorder_level"]
        price = request.form["price"]

        conn = get_connection()
        cur = conn.cursor()

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
                total_stock,
                expiry_date,
                reorder_level,
                price
            )
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """,(
            drug_name,
            brand_name,
            pharmacological_class,
            schedule,
            dosage_form,
            storage,
            category,
            total_stock,
            expiry_date,
            reorder_level,
            price
        ))

        conn.commit()
        cur.close()
        conn.close()

        return redirect(url_for("drugs"))

    return render_template("add_drug.html")


# ==========================
# Drug Management
# ==========================

@app.route("/drugs")
def drugs():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT
            id,
            drug_name,
            brand_name,
            category,
            total_stock,
            expiry_date,
            price
        FROM drugs
        ORDER BY id DESC
    """)

    drugs = cur.fetchall()

    cur.close()
    conn.close()

    return render_template("drugs.html", drugs=drugs)


# ==========================
# Edit Drug
# ==========================

@app.route("/edit-drug/<int:id>", methods=["GET","POST"])
def edit_drug(id):

    conn = get_connection()
    cur = conn.cursor()

    if request.method == "POST":

        drug_name = request.form["drug_name"]
        brand_name = request.form["brand_name"]
        pharmacological_class = request.form["pharmacological_class"]
        schedule = request.form["schedule"]
        dosage_form = request.form["dosage_form"]
        storage = request.form["storage"]
        category = request.form["category"]
        total_stock = request.form["total_stock"]
        expiry_date = request.form["expiry_date"]
        reorder_level = request.form["reorder_level"]
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
                total_stock=%s,
                expiry_date=%s,
                reorder_level=%s,
                price=%s
            WHERE id=%s
        """,(
            drug_name,
            brand_name,
            pharmacological_class,
            schedule,
            dosage_form,
            storage,
            category,
            total_stock,
            expiry_date,
            reorder_level,
            price,
            id
        ))

        conn.commit()

        cur.close()
        conn.close()

        return redirect(url_for("drugs"))

    cur.execute("SELECT * FROM drugs WHERE id=%s",(id,))
    drug = cur.fetchone()

    cur.close()
    conn.close()

    return render_template("edit_drug.html", drug=drug)


# ==========================
# Delete Drug
# ==========================

@app.route("/delete-drug/<int:id>")
def delete_drug(id):

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("DELETE FROM drugs WHERE id=%s",(id,))

    conn.commit()

    cur.close()
    conn.close()

    return redirect(url_for("drugs"))

# ==========================
# Sales Management
# ==========================

@app.route("/sales")
def sales():

    conn = get_connection()
    cur = conn.cursor()

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
        SELECT id, drug_name
        FROM drugs
        ORDER BY drug_name
    """)

    drugs = cur.fetchall()

    cur.close()
    conn.close()

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

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM sales")
    count = cur.fetchone()[0] + 1

    invoice_no = f"INV-{1000 + count}"

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
    """,(
        invoice_no,
        customer_name,
        drug_name,
        quantity,
        price,
        total_amount,
        payment_method
    ))

    cur.execute("""
        UPDATE drugs
        SET total_stock = total_stock - %s
        WHERE drug_name = %s
    """,(quantity, drug_name))

    conn.commit()

    cur.close()
    conn.close()

    return redirect(url_for("sales"))


# ==========================
# Delete Sale
# ==========================

@app.route("/delete-sale/<int:id>")
def delete_sale(id):

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("DELETE FROM sales WHERE id=%s",(id,))

    conn.commit()

    cur.close()
    conn.close()

    return redirect(url_for("sales"))


# ==========================
# Invoice
# ==========================

@app.route("/invoice/<int:id>")
def invoice(id):

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("""
        SELECT *
        FROM sales
        WHERE id=%s
    """,(id,))

    invoice = cur.fetchone()

    cur.close()
    conn.close()

    return render_template(
        "invoice.html",
        invoice=invoice
    )

# ==========================
# Reports
# ==========================

@app.route("/reports")
def reports():

    conn = get_connection()
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM drugs")
    total_drugs = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM sales")
    total_sales = cur.fetchone()[0]

    cur.execute("SELECT IFNULL(SUM(total_amount),0) FROM sales")
    total_revenue = cur.fetchone()[0]

    cur.execute("""
        SELECT COUNT(*)
        FROM drugs
        WHERE total_stock <= reorder_level
    """)
    low_stock_count = cur.fetchone()[0]

    cur.execute("""
        SELECT
            id,
            drug_name,
            total_stock,
            reorder_level
        FROM drugs
        WHERE total_stock <= reorder_level
        ORDER BY total_stock ASC
    """)
    low_stock_drugs = cur.fetchall()

    cur.execute("""
        SELECT
            id,
            drug_name,
            total_stock,
            expiry_date
        FROM drugs
        ORDER BY expiry_date ASC
        LIMIT 10
    """)
    expiry_drugs = cur.fetchall()

    cur.close()
    conn.close()

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

    conn = get_connection()
    cur = conn.cursor()

    if request.method == "POST":

        full_name = request.form["full_name"]
        email = request.form["email"]
        phone = request.form["phone"]

        cur.execute("""
            UPDATE admin_profile
            SET
                full_name=%s,
                email=%s,
                phone=%s
            WHERE id=1
        """, (full_name, email, phone))

        conn.commit()

    cur.execute("SELECT * FROM admin_profile WHERE id=1")
    profile = cur.fetchone()

    cur.execute("SELECT * FROM pharmacy_settings WHERE id=1")
    pharmacy = cur.fetchone()

    cur.execute("SELECT * FROM notification_settings WHERE id=1")
    notification = cur.fetchone()

    cur.execute("SELECT * FROM theme_settings WHERE id=1")
    theme = cur.fetchone()

    cur.close()
    conn.close()

    return render_template(
        "settings.html",
        profile=profile,
        pharmacy=pharmacy,
        notification=notification,
        theme=theme
    )


# ==========================
# PDF Export
# ==========================

@app.route("/reports/pdf")
def report_pdf():

    return send_file(
        "PIMS_Report.pdf",
        as_attachment=True
    )


# ==========================
# Excel Export
# ==========================

@app.route("/reports/excel")
def export_excel():

    return send_file(
        "PIMS_Report.xlsx",
        as_attachment=True
    )


# ==========================
# Database Test
# ==========================

@app.route("/testdb")
def testdb():

    try:

        conn = get_connection()
        cur = conn.cursor()

        cur.execute("SELECT 1")

        cur.close()
        conn.close()

        return "✅ Database Connected Successfully!"

    except Exception as e:

        return f"❌ Database Error : {e}"


# ==========================
# Run Application
# ==========================

if __name__ == "__main__":
    app.run(debug=True)
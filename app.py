from flask import (
    Flask,
    render_template,
    request,
    redirect,
    url_for,
    flash,
    send_file,
    jsonify
)

from flask_mysqldb import MySQL

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)

from openpyxl import Workbook

from io import BytesIO

import os
from datetime import datetime, date, timedelta


# =========================================================
# FLASK APP
# =========================================================

app = Flask(__name__)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "pims_premium_secret_2026"
)


# =========================================================
# MYSQL CONFIGURATION
# =========================================================

app.config["MYSQL_HOST"] = os.environ.get(
    "MYSQL_HOST",
    "127.0.0.1"
)

app.config["MYSQL_PORT"] = int(
    os.environ.get(
        "MYSQL_PORT",
        3306
    )
)

app.config["MYSQL_USER"] = os.environ.get(
    "MYSQL_USER",
    "root"
)

app.config["MYSQL_PASSWORD"] = os.environ.get(
    "MYSQL_PASSWORD",
    ""
)

app.config["MYSQL_DB"] = os.environ.get(
    "MYSQL_DB",
    "pims_db"
)

app.config["MYSQL_CURSORCLASS"] = "DictCursor"

mysql = MySQL(app)


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def safe_int(value, default=0):
    try:
        if value is None or value == "":
            return default

        return int(float(value))

    except (TypeError, ValueError):
        return default


def safe_float(value, default=0.0):
    try:
        if value is None or value == "":
            return default

        return float(value)

    except (TypeError, ValueError):
        return default


def get_drug_status(stock, reorder_level):
    stock = safe_int(stock)

    reorder_level = safe_int(
        reorder_level,
        10
    )

    if stock <= 0:
        return "Out of Stock"

    elif stock <= reorder_level:
        return "Low Stock"

    return "Available"


def close_cursor(cur):
    try:
        if cur:
            cur.close()
    except Exception:
        pass


# =========================================================
# LOGIN
# =========================================================

@app.route(
    "/",
    methods=["GET", "POST"]
)
def login():

    if request.method == "POST":

        username = request.form.get(
            "username",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        ).strip()

        if (
            username == "Purvashree14"
            and
            password == "Purvashree07"
        ):
            return redirect(
                url_for("dashboard")
            )

        flash(
            "Invalid username or password!",
            "danger"
        )

    return render_template(
        "login.html"
    )


# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout")
def logout():

    return redirect(
        url_for("login")
    )

## =========================================================
# DASHBOARD
# =========================================================

@app.route("/dashboard")
def dashboard():

    cur = None

    # =====================================================
    # DEFAULT VALUES
    # =====================================================

    total_drugs = 0
    total_stock = 0
    low_stock = 0
    out_of_stock = 0
    expired = 0
    near_expiry = 0
    reorder_count = 0

    total_sales = 0
    total_sales_amount = 0.0

    drugs = []

    # Stock chart
    labels = []
    stock_data = []

    # Sales chart
    sales_labels = []
    sales_data = []

    try:

        # =================================================
        # DATABASE CONNECTION
        # =================================================

        cur = mysql.connection.cursor()


        # =================================================
        # TOTAL DRUGS
        # =================================================

        cur.execute("""
            SELECT COUNT(*) AS total
            FROM drugs
        """)

        row = cur.fetchone()

        if row:
            total_drugs = safe_int(row["total"])


        # =================================================
        # TOTAL STOCK
        # =================================================

        cur.execute("""
            SELECT COALESCE(SUM(stock), 0) AS total
            FROM drugs
        """)

        row = cur.fetchone()

        if row:
            total_stock = safe_int(row["total"])


        # =================================================
        # LOW STOCK
        # =================================================
        # Stock > 0 AND Stock <= Reorder Level
        # =================================================

        cur.execute("""
            SELECT COUNT(*) AS total
            FROM drugs
            WHERE COALESCE(stock, 0) > 0
            AND COALESCE(stock, 0)
                <= COALESCE(reorder_level, 10)
        """)

        row = cur.fetchone()

        if row:
            low_stock = safe_int(row["total"])


        # =================================================
        # OUT OF STOCK
        # =================================================

        cur.execute("""
            SELECT COUNT(*) AS total
            FROM drugs
            WHERE COALESCE(stock, 0) <= 0
        """)

        row = cur.fetchone()

        if row:
            out_of_stock = safe_int(row["total"])


        # =================================================
        # EXPIRED DRUGS
        # =================================================

        cur.execute("""
            SELECT COUNT(*) AS total
            FROM drugs
            WHERE expiry_date IS NOT NULL
            AND expiry_date < CURDATE()
        """)

        row = cur.fetchone()

        if row:
            expired = safe_int(row["total"])


        # =================================================
        # NEAR EXPIRY
        # NEXT 30 DAYS
        # =================================================

        cur.execute("""
            SELECT COUNT(*) AS total
            FROM drugs
            WHERE expiry_date IS NOT NULL
            AND expiry_date >= CURDATE()
            AND expiry_date <= DATE_ADD(
                CURDATE(),
                INTERVAL 30 DAY
            )
        """)

        row = cur.fetchone()

        if row:
            near_expiry = safe_int(row["total"])


        # =================================================
        # REORDER COUNT
        # =================================================

        cur.execute("""
            SELECT COUNT(*) AS total
            FROM drugs
            WHERE COALESCE(stock, 0)
            <= COALESCE(reorder_level, 10)
        """)

        row = cur.fetchone()

        if row:
            reorder_count = safe_int(row["total"])


        # =================================================
        # TOTAL SALES
        # =================================================

        try:

            cur.execute("""
                SELECT COUNT(*) AS total
                FROM sales
            """)

            row = cur.fetchone()

            if row:
                total_sales = safe_int(row["total"])

        except Exception as e:

            print("TOTAL SALES ERROR:", e)

            total_sales = 0


        # =================================================
        # TOTAL SALES AMOUNT
        # =================================================

        try:

            cur.execute("""
                SELECT COALESCE(
                    SUM(total_amount),
                    0
                ) AS total
                FROM sales
            """)

            row = cur.fetchone()

            if row:

                total_sales_amount = safe_float(
                    row["total"]
                )

        except Exception as e:

            print(
                "TOTAL SALES AMOUNT ERROR:",
                e
            )

            total_sales_amount = 0.0


        # =================================================
        # RECENT DRUGS
        # =================================================

        cur.execute("""
            SELECT
                id,
                drug_name,
                brand_name,
                category,
                pharmacological_class,
                schedule,
                dosage_form,
                storage,
                stock,
                reorder_level,
                expiry_date,
                price
            FROM drugs
            ORDER BY id DESC
           limit 1000
        """)

        drugs = cur.fetchall()


        # =================================================
        # DRUG STATUS
        # =================================================

        for drug in drugs:

            stock = safe_int(
                drug.get("stock", 0)
            )

            reorder_level = safe_int(
                drug.get(
                    "reorder_level",
                    10
                )
            )

            # ---------------------------------------------
            # DEFAULT REORDER LEVEL
            # ---------------------------------------------

            if reorder_level <= 0:

                reorder_level = 10


            # ---------------------------------------------
            # EXPIRY DATE
            # ---------------------------------------------

            expiry_date = drug.get(
                "expiry_date"
            )


            # ---------------------------------------------
            # DEFAULT STATUS
            # ---------------------------------------------

            status = "In Stock"


            # ---------------------------------------------
            # STOCK STATUS
            # ---------------------------------------------

            if stock <= 0:

                status = "Out of Stock"

            elif stock <= reorder_level:

                status = "Low Stock"


            # ---------------------------------------------
            # EXPIRY STATUS
            # ---------------------------------------------

            if expiry_date:

                try:

                    today = date.today()


                    if expiry_date < today:

                        status = "Expired"


                    elif expiry_date <= (
                        today + timedelta(days=30)
                    ):

                        status = "Near Expiry"


                except Exception as e:

                    print(
                        "STATUS ERROR:",
                        e
                    )


            # ---------------------------------------------
            # SAVE STATUS
            # ---------------------------------------------

            drug["status"] = status


        # =================================================
        # STOCK CHART DATA
        # =================================================

        for drug in drugs:

            drug_name = drug.get(
                "drug_name",
                ""
            )

            stock = safe_int(
                drug.get(
                    "stock",
                    0
                )
            )

            labels.append(
                str(drug_name)
            )

            stock_data.append(
                stock
            )


        # =================================================
        # SALES CHART DATA
        # =================================================
        #
        # Last 7 days sales
        #
        # IMPORTANT:
        # If sales table does not exist, chart remains empty.
        #
        # =================================================

        try:

            cur.execute("""
                SELECT
                    DATE(sale_date) AS sale_day,
                    COUNT(*) AS sale_count
                FROM sales
                WHERE sale_date >= DATE_SUB(
                    CURDATE(),
                    INTERVAL 6 DAY
                )
                GROUP BY DATE(sale_date)
                ORDER BY DATE(sale_date)
            """)

            sales_rows = cur.fetchall()


            for row in sales_rows:

                sales_day = row.get(
                    "sale_day"
                )

                sale_count = safe_int(
                    row.get(
                        "sale_count",
                        0
                    )
                )


                if sales_day:

                    sales_labels.append(
                        str(sales_day)
                    )

                    sales_data.append(
                        sale_count
                    )


        except Exception as e:

            print(
                "SALES CHART ERROR:",
                e
            )

            sales_labels = []
            sales_data = []


        # =================================================
        # FALLBACK SALES CHART
        # =================================================

        if not sales_labels:

            sales_labels = [
                "Sales"
            ]

            sales_data = [
                total_sales
            ]


        # =================================================
        # DASHBOARD TEMPLATE
        # =================================================

        return render_template(

            "dashboard.html",

            # ---------------------------------------------
            # INVENTORY CARDS
            # ---------------------------------------------

            total_drugs=total_drugs,

            total_stock=total_stock,

            low_stock=low_stock,

            out_of_stock=out_of_stock,

            expired=expired,

            near_expiry=near_expiry,

            reorder_count=reorder_count,


            # ---------------------------------------------
            # SALES
            # ---------------------------------------------

            total_sales=total_sales,

            total_sales_amount=total_sales_amount,


            # ---------------------------------------------
            # DRUG DATA
            # ---------------------------------------------

            drugs=drugs,

            recent_drugs=drugs,


            # ---------------------------------------------
            # EXPIRY ALIASES
            # ---------------------------------------------

            expired_drugs=expired,

            expiry_alert=near_expiry,


            # ---------------------------------------------
            # STOCK CHART
            # ---------------------------------------------

            labels=labels,

            stock_data=stock_data,


            # ---------------------------------------------
            # SALES CHART
            # ---------------------------------------------

            sales_labels=sales_labels,

            sales_data=sales_data

        )


    # =====================================================
    # ERROR HANDLING
    # =====================================================

    except Exception as e:

        print(
            "DASHBOARD ERROR:",
            e
        )


        flash(
            "Dashboard error: " + str(e),
            "danger"
        )


        # =================================================
        # SAFE FALLBACK PAGE
        # =================================================

        return render_template(

            "dashboard.html",

            total_drugs=0,

            total_stock=0,

            low_stock=0,

            out_of_stock=0,

            expired=0,

            near_expiry=0,

            reorder_count=0,

            total_sales=0,

            total_sales_amount=0.0,

            drugs=[],

            recent_drugs=[],

            expired_drugs=0,

            expiry_alert=0,

            labels=[],

            stock_data=[],

            sales_labels=[
                "Sales"
            ],

            sales_data=[
                0
            ]

        )


    finally:

        # =================================================
        # CLOSE CURSOR
        # =================================================

        try:

            if cur:

                cur.close()

        except Exception:

            pass


# =========================================================
# ADD DRUG
# =========================================================

@app.route(
    "/add-drug",
    methods=["GET", "POST"]
)
def add_drug():

    if request.method == "POST":

        drug_name = request.form.get(
            "drug_name",
            ""
        ).strip()

        brand_name = request.form.get(
            "brand_name",
            ""
        ).strip()

        category = request.form.get(
            "category",
            ""
        ).strip()

        pharmacological_class = request.form.get(
            "pharmacological_class",
            ""
        ).strip()

        schedule = request.form.get(
            "schedule",
            ""
        ).strip()

        dosage_form = request.form.get(
            "dosage_form",
            ""
        ).strip()

        storage = request.form.get(
            "storage",
            ""
        ).strip()

        stock = safe_int(
            request.form.get(
                "total_stock",
                request.form.get(
                    "stock",
                    0
                )
            )
        )

        reorder_level = safe_int(
            request.form.get(
                "reorder_level",
                10
            ),
            10
        )

        expiry_date = request.form.get(
            "expiry_date",
            ""
        ).strip()

        price = safe_float(
            request.form.get(
                "price",
                0
            )
        )

        if not drug_name:

            flash(
                "Drug name is required!",
                "danger"
            )

            return redirect(
                url_for("add_drug")
            )

        if stock < 0:

            flash(
                "Stock cannot be negative!",
                "danger"
            )

            return redirect(
                url_for("add_drug")
            )

        if reorder_level < 0:

            flash(
                "Reorder level cannot be negative!",
                "danger"
            )

            return redirect(
                url_for("add_drug")
            )

        if price < 0:

            flash(
                "Price cannot be negative!",
                "danger"
            )

            return redirect(
                url_for("add_drug")
            )

        status = get_drug_status(
            stock,
            reorder_level
        )

        cur = None

        try:

            cur = mysql.connection.cursor()

            cur.execute("""
                INSERT INTO drugs
                (
                    drug_name,
                    brand_name,
                    category,
                    stock,
                    reorder_level,
                    expiry_date,
                    price,
                    pharmacological_class,
                    schedule,
                    dosage_form,
                    storage,
                    status
                )
                VALUES
                (
                    %s,
                    %s,
                    %s,
                    %s,
                    %s,
                    %s,
                    %s,
                    %s,
                    %s,
                    %s,
                    %s,
                    %s
                )
            """, (

                drug_name,
                brand_name,
                category,
                stock,
                reorder_level,
                expiry_date if expiry_date else None,
                price,
                pharmacological_class,
                schedule,
                dosage_form,
                storage,
                status
            ))

            mysql.connection.commit()

            flash(
                "Drug added successfully!",
                "success"
            )

            return redirect(
                url_for("drugs")
            )

        except Exception as e:

            try:
                mysql.connection.rollback()
            except Exception:
                pass

            print(
                "ADD DRUG ERROR:",
                e
            )

            flash(
                "Unable to add drug: " + str(e),
                "danger"
            )

            return redirect(
                url_for("add_drug")
            )

        finally:

            close_cursor(cur)

    return render_template(
        "add_drug.html"
    )


# =========================================================
# DRUG MANAGEMENT
# =========================================================

@app.route("/drugs")
def drugs():

    cur = None

    try:

        cur = mysql.connection.cursor()

        cur.execute("""
            SELECT
                id,
                drug_name,
                brand_name,
                pharmacological_class,
                schedule,
                dosage_form,
                storage,
                category,
                stock,
                reorder_level,
                expiry_date,
                price,
                status
            FROM drugs
            ORDER BY id DESC
        """)

        drugs_list = cur.fetchall()

        for drug in drugs_list:

            drug["status"] = get_drug_status(
                drug.get("stock", 0),
                drug.get("reorder_level", 10)
            )

        return render_template(
            "drugs.html",
            drugs=drugs_list
        )

    except Exception as e:

        print(
            "DRUGS ERROR:",
            e
        )

        flash(
            "Unable to load drugs: " + str(e),
            "danger"
        )

        return render_template(
            "drugs.html",
            drugs=[]
        )

    finally:

        close_cursor(cur)


# =========================================================
# EDIT DRUG
# =========================================================

@app.route(
    "/edit-drug/<int:id>",
    methods=["GET", "POST"]
)
def edit_drug(id):

    cur = None

    try:

        cur = mysql.connection.cursor()

        if request.method == "POST":

            drug_name = request.form.get(
                "drug_name",
                ""
            ).strip()

            brand_name = request.form.get(
                "brand_name",
                ""
            ).strip()

            category = request.form.get(
                "category",
                ""
            ).strip()

            pharmacological_class = request.form.get(
                "pharmacological_class",
                ""
            ).strip()

            schedule = request.form.get(
                "schedule",
                ""
            ).strip()

            dosage_form = request.form.get(
                "dosage_form",
                ""
            ).strip()

            storage = request.form.get(
                "storage",
                ""
            ).strip()

            stock = safe_int(
                request.form.get(
                    "total_stock",
                    request.form.get(
                        "stock",
                        0
                    )
                )
            )

            reorder_level = safe_int(
                request.form.get(
                    "reorder_level",
                    10
                ),
                10
            )

            expiry_date = request.form.get(
                "expiry_date",
                ""
            ).strip()

            price = safe_float(
                request.form.get(
                    "price",
                    0
                )
            )

            if not drug_name:

                flash(
                    "Drug name is required!",
                    "danger"
                )

                return redirect(
                    url_for(
                        "edit_drug",
                        id=id
                    )
                )

            if stock < 0:

                flash(
                    "Stock cannot be negative!",
                    "danger"
                )

                return redirect(
                    url_for(
                        "edit_drug",
                        id=id
                    )
                )

            if reorder_level < 0:

                flash(
                    "Reorder level cannot be negative!",
                    "danger"
                )

                return redirect(
                    url_for(
                        "edit_drug",
                        id=id
                    )
                )

            if price < 0:

                flash(
                    "Price cannot be negative!",
                    "danger"
                )

                return redirect(
                    url_for(
                        "edit_drug",
                        id=id
                    )
                )

            status = get_drug_status(
                stock,
                reorder_level
            )

            cur.execute("""
                UPDATE drugs
                SET
                    drug_name=%s,
                    brand_name=%s,
                    category=%s,
                    stock=%s,
                    reorder_level=%s,
                    expiry_date=%s,
                    price=%s,
                    pharmacological_class=%s,
                    schedule=%s,
                    dosage_form=%s,
                    storage=%s,
                    status=%s
                WHERE id=%s
            """, (

                drug_name,
                brand_name,
                category,
                stock,
                reorder_level,
                expiry_date if expiry_date else None,
                price,
                pharmacological_class,
                schedule,
                dosage_form,
                storage,
                status,
                id
            ))

            mysql.connection.commit()

            flash(
                "Drug updated successfully!",
                "success"
            )

            return redirect(
                url_for("drugs")
            )

        cur.execute("""
            SELECT *
            FROM drugs
            WHERE id=%s
        """, (id,))

        drug = cur.fetchone()

        if not drug:

            flash(
                "Drug not found!",
                "danger"
            )

            return redirect(
                url_for("drugs")
            )

        return render_template(
            "edit_drug.html",
            drug=drug
        )

    except Exception as e:

        try:
            mysql.connection.rollback()
        except Exception:
            pass

        print(
            "EDIT DRUG ERROR:",
            e
        )

        flash(
            "Unable to update drug: " + str(e),
            "danger"
        )

        return redirect(
            url_for("drugs")
        )

    finally:

        close_cursor(cur)


# =========================================================
# DELETE DRUG
# =========================================================

@app.route(
    "/delete-drug/<int:id>"
)
def delete_drug(id):

    cur = None

    try:

        cur = mysql.connection.cursor()

        cur.execute("""
            DELETE FROM drugs
            WHERE id=%s
        """, (id,))

        mysql.connection.commit()

        flash(
            "Drug deleted successfully!",
            "success"
        )

    except Exception as e:

        try:
            mysql.connection.rollback()
        except Exception:
            pass

        print(
            "DELETE DRUG ERROR:",
            e
        )

        flash(
            "Unable to delete drug: " + str(e),
            "danger"
        )

    finally:

        close_cursor(cur)

    return redirect(
        url_for("drugs")
    )


# =========================================================
# SALES
# =========================================================

@app.route("/sales")
def sales():

    cur = None

    try:

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

        sales_list = cur.fetchall()

        cur.execute("""
            SELECT
                id,
                drug_name,
                stock,
                price
            FROM drugs
            WHERE stock > 0
            ORDER BY drug_name ASC
        """)

        drugs_list = cur.fetchall()

        total_sales = sum(
            safe_float(
                sale.get(
                    "total_amount",
                    0
                )
            )
            for sale in sales_list
        )

        return render_template(
            "sales.html",
            sales=sales_list,
            drugs=drugs_list,
            total_sales=total_sales
        )

    except Exception as e:

        print(
            "SALES ERROR:",
            e
        )

        flash(
            "Unable to load sales: " + str(e),
            "danger"
        )

        return render_template(
            "sales.html",
            sales=[],
            drugs=[],
            total_sales=0
        )

    finally:

        close_cursor(cur)


# =========================================================
# GET PRICE
# =========================================================

@app.route(
    "/get-price/<int:drug_id>"
)
def get_price(drug_id):

    cur = None

    try:

        cur = mysql.connection.cursor()

        cur.execute("""
            SELECT
                id,
                drug_name,
                price,
                stock
            FROM drugs
            WHERE id=%s
        """, (drug_id,))

        drug = cur.fetchone()

        if not drug:

            return jsonify({
                "error": "Drug not found"
            }), 404

        return jsonify({

            "id": drug["id"],

            "drug_name":
                drug["drug_name"],

            "price":
                safe_float(
                    drug["price"]
                ),

            "stock":
                safe_int(
                    drug["stock"]
                )
        })

    except Exception as e:

        print(
            "GET PRICE ERROR:",
            e
        )

        return jsonify({
            "error": str(e)
        }), 500

    finally:

        close_cursor(cur)


# =========================================================
# SAVE SALE
# =========================================================

@app.route(
    "/save_sale",
    methods=["POST"]
)
def save_sale():

    customer_name = request.form.get(
        "customer_name",
        ""
    ).strip()

    drug_id = safe_int(
        request.form.get(
            "drug_id",
            0
        )
    )

    quantity = safe_int(
        request.form.get(
            "quantity",
            0
        )
    )

    payment_method = request.form.get(
        "payment_method",
        "Cash"
    ).strip()

    if not drug_id:

        flash(
            "Please select a drug!",
            "danger"
        )

        return redirect(
            url_for("sales")
        )

    if quantity <= 0:

        flash(
            "Quantity must be greater than zero!",
            "danger"
        )

        return redirect(
            url_for("sales")
        )

    cur = None

    try:

        cur = mysql.connection.cursor()

        cur.execute("""
            SELECT
                id,
                drug_name,
                stock,
                price,
                reorder_level
            FROM drugs
            WHERE id=%s
            FOR UPDATE
        """, (drug_id,))

        drug = cur.fetchone()

        if not drug:

            flash(
                "Drug not found!",
                "danger"
            )

            return redirect(
                url_for("sales")
            )

        available_stock = safe_int(
            drug["stock"]
        )

        if quantity > available_stock:

            flash(
                f"Only {available_stock} units available!",
                "danger"
            )

            return redirect(
                url_for("sales")
            )

        price = safe_float(
            drug["price"]
        )

        total_amount = (
            quantity * price
        )

        cur.execute("""
            SELECT
                COALESCE(
                    MAX(id),
                    0
                ) AS last_id
            FROM sales
        """)

        row = cur.fetchone()

        last_id = safe_int(
            row["last_id"]
            if row
            else 0
        )

        invoice_no = (
            "INV-" +
            str(1001 + last_id)
        )

        cur.execute("""
            INSERT INTO sales
            (
                customer_name,
                drug_name,
                quantity,
                total_amount,
                sale_date,
                invoice_no,
                price,
                payment_method,
                status
            )
            VALUES
            (
                %s,
                %s,
                %s,
                %s,
                NOW(),
                %s,
                %s,
                %s,
                %s
            )
        """, (

            customer_name
            or "Walk-in Customer",

            drug["drug_name"],

            quantity,

            total_amount,

            invoice_no,

            price,

            payment_method
            or "Cash",

            "Completed"
        ))

        new_stock = (
            available_stock -
            quantity
        )

        reorder_level = safe_int(
            drug["reorder_level"],
            10
        )

        new_status = get_drug_status(
            new_stock,
            reorder_level
        )

        cur.execute("""
            UPDATE drugs
            SET
                stock=%s,
                status=%s
            WHERE id=%s
        """, (

            new_stock,
            new_status,
            drug_id
        ))

        mysql.connection.commit()

        flash(
            f"Sale saved successfully! Invoice: {invoice_no}",
            "success"
        )

        return redirect(
            url_for("sales")
        )

    except Exception as e:

        try:
            mysql.connection.rollback()
        except Exception:
            pass

        print(
            "SAVE SALE ERROR:",
            e
        )

        flash(
            "Unable to save sale: " + str(e),
            "danger"
        )

        return redirect(
            url_for("sales")
        )

    finally:

        close_cursor(cur)


# =========================================================
# DELETE SALE
# =========================================================

@app.route(
    "/delete-sale/<int:id>"
)
def delete_sale(id):

    cur = None

    try:

        cur = mysql.connection.cursor()

        cur.execute("""
            DELETE FROM sales
            WHERE id=%s
        """, (id,))

        mysql.connection.commit()

        flash(
            "Sale deleted successfully!",
            "success"
        )

    except Exception as e:

        try:
            mysql.connection.rollback()
        except Exception:
            pass

        print(
            "DELETE SALE ERROR:",
            e
        )

        flash(
            "Unable to delete sale: " + str(e),
            "danger"
        )

    finally:

        close_cursor(cur)

    return redirect(
        url_for("sales")
    )


# =========================================================
# REPORTS
# =========================================================

@app.route("/reports")
def reports():

    cur = None

    try:

        cur = mysql.connection.cursor()

        # TOTAL DRUGS
        cur.execute("""
            SELECT COUNT(*) AS total
            FROM drugs
        """)

        row = cur.fetchone()

        total_drugs = safe_int(
            row["total"] if row else 0
        )

        # TOTAL SALES
        total_sales = 0

        try:

            cur.execute("""
                SELECT COUNT(*) AS total
                FROM sales
            """)

            row = cur.fetchone()

            total_sales = safe_int(
                row["total"] if row else 0
            )

        except Exception as e:

            print(
                "REPORT SALES COUNT ERROR:",
                e
            )

        # TOTAL REVENUE
        total_revenue = 0.0

        try:

            cur.execute("""
                SELECT
                    COALESCE(
                        SUM(total_amount),
                        0
                    ) AS total
                FROM sales
            """)

            row = cur.fetchone()

            total_revenue = safe_float(
                row["total"] if row else 0
            )

        except Exception as e:

            print(
                "REPORT REVENUE ERROR:",
                e
            )

        # LOW STOCK
        cur.execute("""
            SELECT
                id,
                drug_name,
                stock,
                reorder_level,
                status
            FROM drugs
            WHERE stock <= COALESCE(
                reorder_level,
                10
            )
            ORDER BY stock ASC,
                     drug_name ASC
        """)

        low_stock_drugs = cur.fetchall()

        for drug in low_stock_drugs:

            drug["status"] = get_drug_status(
                drug.get("stock", 0),
                drug.get("reorder_level", 10)
            )

        # EXPIRY DRUGS
        cur.execute("""
            SELECT
                id,
                drug_name,
                stock,
                expiry_date,
                status
            FROM drugs
            WHERE expiry_date IS NOT NULL
            AND expiry_date <= DATE_ADD(
                CURDATE(),
                INTERVAL 30 DAY
            )
            ORDER BY expiry_date ASC
        """)

        expiry_drugs = cur.fetchall()

        # ALL SALES
        sales_data = []

        try:

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

            sales_data = cur.fetchall()

        except Exception as e:

            print(
                "REPORT SALES DATA ERROR:",
                e
            )

        # ALL DRUGS
        cur.execute("""
            SELECT
                id,
                drug_name,
                brand_name,
                category,
                pharmacological_class,
                schedule,
                dosage_form,
                storage,
                stock,
                reorder_level,
                expiry_date,
                price,
                status
            FROM drugs
            ORDER BY id DESC
        """)

        drugs_data = cur.fetchall()

        for drug in drugs_data:

            drug["status"] = get_drug_status(
                drug.get("stock", 0),
                drug.get("reorder_level", 10)
            )

        total_quantity = sum(
            safe_int(
                sale.get(
                    "quantity",
                    0
                )
            )
            for sale in sales_data
        )

        return render_template(

            "reports.html",

            total_drugs=total_drugs,

            total_sales=total_sales,

            total_revenue=total_revenue,

            low_stock_count=len(
                low_stock_drugs
            ),

            low_stock_drugs=low_stock_drugs,

            expiry_drugs=expiry_drugs,

            sales=sales_data,

            drugs=drugs_data,

            total_quantity=total_quantity
        )

    except Exception as e:

        print(
            "REPORTS ERROR:",
            e
        )

        flash(
            "Unable to load reports: " + str(e),
            "danger"
        )

        return render_template(

            "reports.html",

            total_drugs=0,

            total_sales=0,

            total_revenue=0,

            low_stock_count=0,

            low_stock_drugs=[],

            expiry_drugs=[],

            sales=[],

            drugs=[],

            total_quantity=0
        )

    finally:

        close_cursor(cur)


# =========================================================
# EXPORT EXCEL
# =========================================================

@app.route("/export-excel")
def export_excel():

    cur = None

    try:

        cur = mysql.connection.cursor()

        cur.execute("""
            SELECT
                id,
                drug_name,
                brand_name,
                category,
                pharmacological_class,
                schedule,
                dosage_form,
                storage,
                stock,
                reorder_level,
                expiry_date,
                price,
                status
            FROM drugs
            ORDER BY id DESC
        """)

        drugs_data = cur.fetchall()

        workbook = Workbook()

        sheet = workbook.active

        sheet.title = "Drug Inventory"

        headers = [

            "ID",
            "Drug Name",
            "Brand Name",
            "Category",
            "Pharmacological Class",
            "Schedule",
            "Dosage Form",
            "Storage",
            "Stock",
            "Reorder Level",
            "Expiry Date",
            "Price",
            "Status"

        ]

        sheet.append(headers)

        for drug in drugs_data:

            calculated_status = get_drug_status(
                drug.get("stock", 0),
                drug.get("reorder_level", 10)
            )

            sheet.append([

                drug.get("id"),

                drug.get("drug_name"),

                drug.get("brand_name"),

                drug.get("category"),

                drug.get(
                    "pharmacological_class"
                ),

                drug.get("schedule"),

                drug.get("dosage_form"),

                drug.get("storage"),

                safe_int(
                    drug.get(
                        "stock",
                        0
                    )
                ),

                safe_int(
                    drug.get(
                        "reorder_level",
                        10
                    )
                ),

                str(
                    drug.get(
                        "expiry_date"
                    )
                    or ""
                ),

                safe_float(
                    drug.get(
                        "price",
                        0
                    )
                ),

                calculated_status
            ])

        sales_sheet = workbook.create_sheet(
            "Sales Report"
        )

        sales_data = []

        try:

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

            sales_data = cur.fetchall()

        except Exception as e:

            print(
                "EXCEL SALES ERROR:",
                e
            )

        sales_headers = [

            "ID",
            "Invoice No",
            "Customer Name",
            "Drug Name",
            "Quantity",
            "Price",
            "Total Amount",
            "Payment Method",
            "Sale Date",
            "Status"

        ]

        sales_sheet.append(
            sales_headers
        )

        for sale in sales_data:

            sales_sheet.append([

                sale.get("id"),

                sale.get("invoice_no"),

                sale.get("customer_name"),

                sale.get("drug_name"),

                safe_int(
                    sale.get(
                        "quantity",
                        0
                    )
                ),

                safe_float(
                    sale.get(
                        "price",
                        0
                    )
                ),

                safe_float(
                    sale.get(
                        "total_amount",
                        0
                    )
                ),

                sale.get(
                    "payment_method"
                ),

                str(
                    sale.get(
                        "sale_date"
                    )
                    or ""
                ),

                sale.get("status")
            ])

        for ws in workbook.worksheets:

            for column_cells in ws.columns:

                max_length = 0

                column_letter = (
                    column_cells[0].column_letter
                )

                for cell in column_cells:

                    try:

                        cell_length = len(
                            str(
                                cell.value
                                or ""
                            )
                        )

                        if cell_length > max_length:

                            max_length = cell_length

                    except Exception:

                        pass

                ws.column_dimensions[
                    column_letter
                ].width = min(
                    max_length + 2,
                    35
                )

        output = BytesIO()

        workbook.save(output)

        output.seek(0)

        return send_file(

            output,

            as_attachment=True,

            download_name="PIMS_Report.xlsx",

            mimetype=(
                "application/"
                "vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            )
        )

    except Exception as e:

        print(
            "EXPORT EXCEL ERROR:",
            e
        )

        flash(
            "Unable to export Excel: " + str(e),
            "danger"
        )

        return redirect(
            url_for("reports")
        )

    finally:

        close_cursor(cur)


# =========================================================
# OLD EXCEL ROUTE
# =========================================================

@app.route("/export-drugs")
def export_drugs():

    return redirect(
        url_for("export_excel")
    )


# =========================================================
# EXPORT SALES EXCEL
# =========================================================

@app.route("/export-sales")
def export_sales():

    cur = None

    try:

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

        sales_data = cur.fetchall()

        workbook = Workbook()

        sheet = workbook.active

        sheet.title = "Sales Report"

        headers = [

            "ID",
            "Invoice No",
            "Customer Name",
            "Drug Name",
            "Quantity",
            "Price",
            "Total Amount",
            "Payment Method",
            "Sale Date",
            "Status"

        ]

        sheet.append(headers)

        for sale in sales_data:

            sheet.append([

                sale.get("id"),

                sale.get("invoice_no"),

                sale.get("customer_name"),

                sale.get("drug_name"),

                sale.get("quantity"),

                sale.get("price"),

                sale.get("total_amount"),

                sale.get("payment_method"),

                str(
                    sale.get(
                        "sale_date"
                    )
                    or ""
                ),

                sale.get("status")
            ])

        output = BytesIO()

        workbook.save(output)

        output.seek(0)

        return send_file(

            output,

            as_attachment=True,

            download_name="PIMS_Sales_Report.xlsx",

            mimetype=(
                "application/"
                "vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            )
        )

    except Exception as e:

        print(
            "EXPORT SALES ERROR:",
            e
        )

        flash(
            "Unable to export sales: " + str(e),
            "danger"
        )

        return redirect(
            url_for("reports")
        )

    finally:

        close_cursor(cur)


# =========================================================
# PDF REPORT
# =========================================================

@app.route("/report-pdf")
def report_pdf():

    cur = None

    try:

        cur = mysql.connection.cursor()

        # DRUG DATA
        cur.execute("""
            SELECT
                id,
                drug_name,
                brand_name,
                stock,
                reorder_level,
                expiry_date,
                price,
                status
            FROM drugs
            ORDER BY id DESC
        """)

        drugs_data = cur.fetchall()

        # LOW STOCK
        low_stock_drugs = [

            drug

            for drug in drugs_data

            if safe_int(
                drug.get(
                    "stock",
                    0
                )
            )
            <=
            safe_int(
                drug.get(
                    "reorder_level",
                    10
                ),
                10
            )
        ]

        # EXPIRY
        cur.execute("""
            SELECT
                id,
                drug_name,
                stock,
                expiry_date
            FROM drugs
            WHERE expiry_date IS NOT NULL
            AND expiry_date <= DATE_ADD(
                CURDATE(),
                INTERVAL 30 DAY
            )
            ORDER BY expiry_date ASC
        """)

        expiry_drugs = cur.fetchall()

        # SALES
        sales_data = []

        try:

            cur.execute("""
                SELECT
                    invoice_no,
                    customer_name,
                    drug_name,
                    quantity,
                    price,
                    total_amount,
                    payment_method,
                    sale_date
                FROM sales
                ORDER BY id DESC
            """)

            sales_data = cur.fetchall()

        except Exception as e:

            print(
                "PDF SALES QUERY ERROR:",
                e
            )

        total_revenue = sum(

            safe_float(
                sale.get(
                    "total_amount",
                    0
                )
            )

            for sale in sales_data
        )

        output = BytesIO()

        document = SimpleDocTemplate(

            output,

            pagesize=landscape(A4),

            rightMargin=8 * mm,

            leftMargin=8 * mm,

            topMargin=8 * mm,

            bottomMargin=8 * mm
        )

        styles = getSampleStyleSheet()

        story = []

        story.append(
            Paragraph(
                "PIMS Premium - Pharmacy Report",
                styles["Title"]
            )
        )

        story.append(
            Spacer(
                1,
                8
            )
        )

        story.append(
            Paragraph(
                "Generated on: "
                + datetime.now().strftime(
                    "%d-%m-%Y %H:%M"
                ),
                styles["Normal"]
            )
        )

        story.append(
            Spacer(
                1,
                8
            )
        )

        summary_data = [

            [
                "Total Drugs",
                "Total Sales",
                "Total Revenue",
                "Low Stock",
                "Expiry Alerts"
            ],

            [
                str(
                    len(drugs_data)
                ),

                str(
                    len(sales_data)
                ),

                "Rs."
                + f"{total_revenue:,.2f}",

                str(
                    len(
                        low_stock_drugs
                    )
                ),

                str(
                    len(
                        expiry_drugs
                    )
                )
            ]
        ]

        summary_table = Table(

            summary_data,

            colWidths=[
                45 * mm,
                45 * mm,
                50 * mm,
                45 * mm,
                45 * mm
            ]
        )

        summary_table.setStyle(
            TableStyle([

                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor(
                        "#20252b"
                    )
                ),

                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white
                ),

                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.grey
                ),

                (
                    "ALIGN",
                    (0, 0),
                    (-1, -1),
                    "CENTER"
                ),

                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold"
                ),

                (
                    "FONTSIZE",
                    (0, 0),
                    (-1, -1),
                    9
                )
            ])
        )

        story.append(
            summary_table
        )

        story.append(
            Spacer(
                1,
                12
            )
        )

        story.append(
            Paragraph(
                "Drug Inventory",
                styles["Heading2"]
            )
        )

        inventory_data = [[

            "ID",
            "Drug Name",
            "Brand",
            "Stock",
            "Reorder",
            "Expiry Date",
            "Price",
            "Status"

        ]]

        for drug in drugs_data:

            status = get_drug_status(

                drug.get(
                    "stock",
                    0
                ),

                drug.get(
                    "reorder_level",
                    10
                )
            )

            inventory_data.append([

                str(
                    drug.get(
                        "id",
                        ""
                    )
                ),

                str(
                    drug.get(
                        "drug_name",
                        ""
                    )
                ),

                str(
                    drug.get(
                        "brand_name",
                        ""
                    )
                ),

                str(
                    drug.get(
                        "stock",
                        0
                    )
                ),

                str(
                    drug.get(
                        "reorder_level",
                        0
                    )
                ),

                str(
                    drug.get(
                        "expiry_date",
                        ""
                    )
                    or "-"
                ),

                "Rs."
                + f"{safe_float(
                    drug.get(
                        'price',
                        0
                    )
                ):,.2f}",

                status
            ])

        inventory_table = Table(
            inventory_data,
            repeatRows=1
        )

        inventory_table.setStyle(
            TableStyle([

                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor(
                        "#20252b"
                    )
                ),

                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.white
                ),

                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.4,
                    colors.grey
                ),

                (
                    "FONTSIZE",
                    (0, 0),
                    (-1, -1),
                    7
                ),

                (
                    "FONTNAME",
                    (0, 0),
                    (-1, 0),
                    "Helvetica-Bold"
                ),

                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                )
            ])
        )

        story.append(
            inventory_table
        )

        story.append(
            Spacer(
                1,
                12
            )
        )

        if sales_data:

            story.append(
                Paragraph(
                    "Sales Report",
                    styles["Heading2"]
                )
            )

            sales_table_data = [[

                "Invoice",
                "Customer",
                "Drug",
                "Qty",
                "Price",
                "Total",
                "Payment",
                "Date"

            ]]

            for sale in sales_data:

                sales_table_data.append([

                    str(
                        sale.get(
                            "invoice_no",
                            ""
                        )
                    ),

                    str(
                        sale.get(
                            "customer_name",
                            ""
                        )
                    ),

                    str(
                        sale.get(
                            "drug_name",
                            ""
                        )
                    ),

                    str(
                        sale.get(
                            "quantity",
                            0
                        )
                    ),

                    "Rs."
                    + f"{safe_float(
                        sale.get(
                            'price',
                            0
                        )
                    ):,.2f}",

                    "Rs."
                    + f"{safe_float(
                        sale.get(
                            'total_amount',
                            0
                        )
                    ):,.2f}",

                    str(
                        sale.get(
                            "payment_method",
                            ""
                        )
                    ),

                    str(
                        sale.get(
                            "sale_date",
                            ""
                        )
                    )
                ])

            sales_table = Table(
                sales_table_data,
                repeatRows=1
            )

            sales_table.setStyle(
                TableStyle([

                    (
                        "BACKGROUND",
                        (0, 0),
                        (-1, 0),
                        colors.HexColor(
                            "#20252b"
                        )
                    ),

                    (
                        "TEXTCOLOR",
                        (0, 0),
                        (-1, 0),
                        colors.white
                    ),

                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.4,
                        colors.grey
                    ),

                    (
                        "FONTSIZE",
                        (0, 0),
                        (-1, -1),
                        7
                    ),

                    (
                        "FONTNAME",
                        (0, 0),
                        (-1, 0),
                        "Helvetica-Bold"
                    ),

                    (
                        "VALIGN",
                        (0, 0),
                        (-1, -1),
                        "MIDDLE"
                    )
                ])
            )

            story.append(
                sales_table
            )

        document.build(
            story
        )

        output.seek(0)

        return send_file(

            output,

            as_attachment=True,

            download_name="PIMS_Report.pdf",

            mimetype="application/pdf"
        )

    except Exception as e:

        print(
            "REPORT PDF ERROR:",
            e
        )

        flash(
            "Unable to generate PDF: " + str(e),
            "danger"
        )

        return redirect(
            url_for("reports")
        )

    finally:

        close_cursor(cur)


# =========================================================
# OLD PDF ROUTE
# =========================================================

@app.route("/sales-pdf")
def sales_pdf():

    return redirect(
        url_for("report_pdf")
    )


# =========================================================
# INVOICE
# =========================================================

@app.route(
    "/invoice/<int:sale_id>"
)
def invoice(sale_id):

    cur = None

    try:

        cur = mysql.connection.cursor()

        cur.execute("""
            SELECT *
            FROM sales
            WHERE id=%s
        """, (sale_id,))

        sale = cur.fetchone()

        if not sale:

            flash(
                "Invoice not found!",
                "danger"
            )

            return redirect(
                url_for("sales")
            )

        return render_template(
            "invoice.html",
            sale=sale
        )

    except Exception as e:

        print(
            "INVOICE ERROR:",
            e
        )

        flash(
            "Unable to load invoice: " + str(e),
            "danger"
        )

        return redirect(
            url_for("sales")
        )

    finally:

        close_cursor(cur)


# =========================================================
# BILLING
# =========================================================

@app.route("/billing")
def billing():

    cur = None

    try:

        cur = mysql.connection.cursor()

        cur.execute("""
            SELECT
                id,
                drug_name,
                stock,
                price
            FROM drugs
            WHERE stock > 0
            ORDER BY drug_name ASC
        """)

        drugs_list = cur.fetchall()

        return render_template(
            "billing.html",
            drugs=drugs_list
        )

    except Exception as e:

        print(
            "BILLING ERROR:",
            e
        )

        flash(
            "Unable to load billing: " + str(e),
            "danger"
        )

        return render_template(
            "billing.html",
            drugs=[]
        )

    finally:

        close_cursor(cur)


# =========================================================
# SETTINGS - GET
# =========================================================

@app.route(
    "/settings",
    methods=["GET", "POST"]
)
def settings():

    cur = None

    try:

        cur = mysql.connection.cursor()

        # =================================================
        # SAVE SETTINGS
        # =================================================

        if request.method == "POST":

            # ---------------------------------------------
            # ADMIN PROFILE
            # ---------------------------------------------

            full_name = request.form.get(
                "full_name",
                ""
            ).strip()

            email = request.form.get(
                "email",
                ""
            ).strip()

            phone = request.form.get(
                "phone",
                ""
            ).strip()

            # ---------------------------------------------
            # PHARMACY INFORMATION
            # ---------------------------------------------

            pharmacy_name = request.form.get(
                "pharmacy_name",
                ""
            ).strip()

            license_no = request.form.get(
                "license_no",
                ""
            ).strip()

            gst_no = request.form.get(
                "gst_no",
                ""
            ).strip()

            contact = request.form.get(
                "contact",
                ""
            ).strip()

            address = request.form.get(
                "address",
                ""
            ).strip()

            # ---------------------------------------------
            # NOTIFICATIONS
            # ---------------------------------------------

            low_stock_alert = (
                1
                if request.form.get(
                    "low_stock_alert"
                )
                else 0
            )

            expiry_alert = (
                1
                if request.form.get(
                    "expiry_alert"
                )
                else 0
            )

            # ---------------------------------------------
            # THEME
            # ---------------------------------------------

            theme_value = request.form.get(
                "theme",
                "Light"
            ).strip()

            if theme_value not in [
                "Light",
                "Dark"
            ]:
                theme_value = "Light"

            # =============================================
            # ADMIN PROFILE
            # =============================================

            cur.execute("""
                SELECT id
                FROM admin_profile
                ORDER BY id ASC
                LIMIT 1
            """)

            profile_row = cur.fetchone()

            if profile_row:

                cur.execute("""
                    UPDATE admin_profile
                    SET
                        full_name=%s,
                        email=%s,
                        phone=%s
                    WHERE id=%s
                """, (

                    full_name,
                    email,
                    phone,
                    profile_row["id"]
                ))

            else:

                cur.execute("""
                    INSERT INTO admin_profile
                    (
                        full_name,
                        email,
                        phone
                    )
                    VALUES
                    (
                        %s,
                        %s,
                        %s
                    )
                """, (

                    full_name,
                    email,
                    phone
                ))# =============================================
            # PHARMACY SETTINGS
            # =============================================

            try:

                cur.execute("""
                    SELECT id
                    FROM pharmacy_settings
                    ORDER BY id ASC
                    LIMIT 1
                """)

                pharmacy_row = cur.fetchone()

                if pharmacy_row:

                    cur.execute("""
                        UPDATE pharmacy_settings
                        SET
                            pharmacy_name=%s,
                            license_no=%s,
                            gst_no=%s,
                            contact=%s,
                            address=%s
                        WHERE id=%s
                    """, (

                        pharmacy_name,
                        license_no,
                        gst_no,
                        contact,
                        address,
                        pharmacy_row["id"]
                    ))

                else:

                    cur.execute("""
                        INSERT INTO pharmacy_settings
                        (
                            pharmacy_name,
                            license_no,
                            gst_no,
                            contact,
                            address
                        )
                        VALUES
                        (
                            %s,
                            %s,
                            %s,
                            %s,
                            %s
                        )
                    """, (

                        pharmacy_name,
                        license_no,
                        gst_no,
                        contact,
                        address
                    ))

            except Exception as e:

                print(
                    "PHARMACY SETTINGS ERROR:",
                    e
                )

            # =============================================
            # NOTIFICATION SETTINGS
            # =============================================

            try:

                cur.execute("""
                    SELECT id
                    FROM notification_settings
                    ORDER BY id ASC
                    LIMIT 1
                """)

                notification_row = cur.fetchone()

                if notification_row:

                    cur.execute("""
                        UPDATE notification_settings
                        SET
                            low_stock_alert=%s,
                            expiry_alert=%s
                        WHERE id=%s
                    """, (

                        low_stock_alert,
                        expiry_alert,
                        notification_row["id"]
                    ))

                else:

                    cur.execute("""
                        INSERT INTO notification_settings
                        (
                            low_stock_alert,
                            expiry_alert
                        )
                        VALUES
                        (
                            %s,
                            %s
                        )
                    """, (

                        low_stock_alert,
                        expiry_alert
                    ))

            except Exception as e:

                print(
                    "NOTIFICATION SETTINGS ERROR:",
                    e
                )

            # =============================================
            # THEME SETTINGS
            # =============================================

            try:

                cur.execute("""
                    SELECT id
                    FROM theme_settings
                    ORDER BY id ASC
                    LIMIT 1
                """)

                theme_row = cur.fetchone()

                if theme_row:

                    cur.execute("""
                        UPDATE theme_settings
                        SET
                            theme=%s
                        WHERE id=%s
                    """, (

                        theme_value,
                        theme_row["id"]
                    ))

                else:

                    cur.execute("""
                        INSERT INTO theme_settings
                        (
                            theme
                        )
                        VALUES
                        (
                            %s
                        )
                    """, (
                        theme_value,
                    ))

            except Exception as e:

                print(
                    "THEME SETTINGS ERROR:",
                    e
                )

            # =============================================
            # COMMIT
            # =============================================

            mysql.connection.commit()

            flash(
                "Settings saved successfully!",
                "success"
            )

            return redirect(
                url_for("settings")
            )

        # =================================================
        # LOAD SETTINGS
        # =================================================

        profile = None
        pharmacy = None
        notification = None
        theme = None

        # ADMIN PROFILE
        try:

            cur.execute("""
                SELECT *
                FROM admin_profile
                ORDER BY id ASC
                LIMIT 1
            """)

            profile = cur.fetchone()

        except Exception as e:

            print(
                "LOAD PROFILE ERROR:",
                e
            )

        # PHARMACY
        try:

            cur.execute("""
                SELECT *
                FROM pharmacy_settings
                ORDER BY id ASC
                LIMIT 1
            """)

            pharmacy = cur.fetchone()

        except Exception as e:

            print(
                "LOAD PHARMACY ERROR:",
                e
            )

        # NOTIFICATIONS
        try:

            cur.execute("""
                SELECT *
                FROM notification_settings
                ORDER BY id ASC
                LIMIT 1
            """)

            notification = cur.fetchone()

        except Exception as e:

            print(
                "LOAD NOTIFICATION ERROR:",
                e
            )

        # THEME
        try:

            cur.execute("""
                SELECT *
                FROM theme_settings
                ORDER BY id ASC
                LIMIT 1
            """)

            theme = cur.fetchone()

        except Exception as e:

            print(
                "LOAD THEME ERROR:",
                e
            )

        return render_template(

            "settings.html",

            profile=profile,

            pharmacy=pharmacy,

            notification=notification,

            theme=theme
        )

    except Exception as e:

        try:
            mysql.connection.rollback()
        except Exception:
            pass

        print(
            "SETTINGS ERROR:",
            e
        )

        flash(
            "Unable to save settings: " + str(e),
            "danger"
        )

        return render_template(

            "settings.html",

            profile=None,

            pharmacy=None,

            notification=None,

            theme=None
        )

    finally:

        close_cursor(cur)


# =========================================================
# HEALTH CHECK
# =========================================================

@app.route("/health")
def health():

    cur = None

    try:

        cur = mysql.connection.cursor()

        cur.execute(
            "SELECT 1 AS test"
        )

        result = cur.fetchone()

        return jsonify({

            "application":
                "PIMS Premium",

            "status":
                "OK",

            "database":
                "Connected",

            "test":
                result["test"]
                if result
                else 1

        })

    except Exception as e:

        print(
            "HEALTH ERROR:",
            e
        )

        return jsonify({

            "application":
                "PIMS Premium",

            "status":
                "ERROR",

            "database":
                "Disconnected",

            "error":
                str(e)

        }), 500

    finally:

        close_cursor(cur)


# =========================================================
# SERVER ERROR
# =========================================================

@app.errorhandler(500)
def internal_server_error(error):

    return """

    <!DOCTYPE html>

    <html>

    <head>

        <title>PIMS - Server Error</title>

        <style>

            body {
                font-family: Arial;
                background: #f5f7fb;
                text-align: center;
                padding-top: 100px;
            }

            h1 {
                color: #dc2626;
            }

            a {
                text-decoration: none;
                padding: 12px 20px;
                background: #2563eb;
                color: white;
                border-radius: 8px;
            }

        </style>

    </head>

    <body>

        <h1>PIMS Server Error</h1>

        <p>
            Something went wrong on the server.
        </p>

        <br>

        <a href="/">
            Back to Login
        </a>

    </body>

    </html>

    """, 500


# =========================================================
# RUN APPLICATION
# =========================================================

if __name__ == "__main__":

    print()
    print("=" * 60)
    print("       PIMS PREMIUM - PHARMACY INVENTORY")
    print("=" * 60)
    print()

    print(
        "Server : http://127.0.0.1:5000"
    )

    print()

    print(
        "Username : Purvashree14"
    )

    print(
        "Password : Purvashree07"
    )

    print()

    print("=" * 60)
    print()

    host = os.environ.get(
        "HOST",
        "127.0.0.1"
    )

    port = int(
        os.environ.get(
            "PORT",
            5000
        )
    )

    app.run(

        host=host,

        port=port,

        debug=True

    )
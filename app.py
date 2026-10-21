
from flask import Flask, render_template, request, redirect, session, jsonify, send_file
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime, date
from io import BytesIO
import pandas as pd
from werkzeug.utils import secure_filename

import os
from urllib.parse import quote_plus
from dotenv import load_dotenv

load_dotenv()  
app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "fg_dispatch_secret")

# ============================================================
# MSSQL DATABASE
# ============================================================
# SQL Server is the application's source of truth.
#
# Environment variables:
# MSSQL_SERVER=localhost
# MSSQL_PORT=1433
# MSSQL_DATABASE=daily_dispatch_report
# MSSQL_USERNAME=sa
# MSSQL_PASSWORD=YourPassword
# MSSQL_DRIVER=ODBC Driver 18 for SQL Server
#
# For Windows Authentication:
# MSSQL_TRUSTED_CONNECTION=yes
# and leave username/password empty.

MSSQL_SERVER = os.getenv("MSSQL_SERVER", "localhost")
MSSQL_PORT = os.getenv("MSSQL_PORT", "1433")
MSSQL_DATABASE = os.getenv("MSSQL_DATABASE", "daily_dispatch_report")
MSSQL_USERNAME = os.getenv("MSSQL_USERNAME", "")
MSSQL_PASSWORD = os.getenv("MSSQL_PASSWORD", "")
MSSQL_DRIVER = os.getenv("MSSQL_DRIVER", "ODBC Driver 18 for SQL Server")
MSSQL_TRUSTED_CONNECTION = os.getenv(
    "MSSQL_TRUSTED_CONNECTION", "no"
).lower() in ("yes", "true", "1")

if MSSQL_TRUSTED_CONNECTION:
    connection_string = (
        f"DRIVER={{{MSSQL_DRIVER}}};"
        f"SERVER={MSSQL_SERVER},{MSSQL_PORT};"
        f"DATABASE={MSSQL_DATABASE};"
        "Trusted_Connection=yes;"
        "TrustServerCertificate=yes;"
    )
else:
    connection_string = (
        f"DRIVER={{{MSSQL_DRIVER}}};"
        f"SERVER={MSSQL_SERVER},{MSSQL_PORT};"
        f"DATABASE={MSSQL_DATABASE};"
        f"UID={MSSQL_USERNAME};"
        f"PWD={MSSQL_PASSWORD};"
        "TrustServerCertificate=yes;"
    )

app.config["SQLALCHEMY_DATABASE_URI"] = (
    "mssql+pyodbc:///?odbc_connect="
    + quote_plus(connection_string)
)
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "pool_pre_ping": True,
    "pool_recycle": 1800,
    "pool_size": 10,
    "max_overflow": 20,
}

from werkzeug.utils import secure_filename

SALES_UPLOAD_FOLDER = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "sales_uploads"
)

os.makedirs(SALES_UPLOAD_FOLDER, exist_ok=True)

app.config["SALES_UPLOAD_FOLDER"] = SALES_UPLOAD_FOLDER

db = SQLAlchemy(app)


class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100))
    password = db.Column(db.String(100))
    role = db.Column(db.String(50))


class DailyDispatchReport(db.Model):
    __tablename__ = "daily_dispatch_report"

    id = db.Column(db.Integer, primary_key=True)
    dispatch_date = db.Column(db.Date, nullable=False, index=True)
    m_code = db.Column(db.String(100))
    customer_name = db.Column(db.String(255))
    part_number = db.Column(db.String(255))
    size = db.Column(db.String(100))
    dispatch_plan_qty = db.Column(db.Integer, default=0)
    sales_order_no = db.Column(db.String(255))
    invoice_number = db.Column(db.String(255))
    remarks = db.Column(db.Text)
    complete = db.Column(db.Integer, default=0)
    pending_qty = db.Column(db.Integer, default=0)
    status = db.Column(db.String(100))
    ageing = db.Column(db.Integer, default=0)
    date_of_completion = db.Column(db.Date)
    reason = db.Column(db.Text)
    edited_flag = db.Column(db.String(100), default="")


class DailyDispatchSummary(db.Model):
    __tablename__ = "daily_dispatch_summary"

    id = db.Column(db.Integer, primary_key=True)
    date = db.Column(db.Date, nullable=False, unique=True, index=True)
    total_dispatch_plan_qty = db.Column(db.Integer, default=0)
    total_dispatch_qty = db.Column(db.Integer, default=0)
    percentage = db.Column(db.String(50), default="0%")
    daily_production_qty = db.Column(db.Integer, default=0)


def parse_date(value):
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return pd.to_datetime(value).date()
    except Exception:
        return None


def clean(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def to_int(value, default=0):
    if value is None or value == "":
        return default
    try:
        if pd.isna(value):
            return default
    except Exception:
        pass
    try:
        return int(float(value))
    except Exception:
        return default


def date_str(value):
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    return str(value).split(" ")[0]


def record_dict(r):
    return {
        "id": r.id,
        # Keep excel_row as an alias so existing JS does not need to change.
        "excel_row": r.id,
        "dispatch_date": date_str(r.dispatch_date),
        "m_code": r.m_code or "",
        "customer": r.customer_name or "",
        "part_number": r.part_number or "",
        "size": r.size or "",
        "dispatch_qty": r.dispatch_plan_qty or 0,
        "sales_order_no": r.sales_order_no or "",
        "invoice_no": r.invoice_number or "",
        "remarks": r.remarks or "",
        "complete": r.complete or 0,
        "pending": r.pending_qty or 0,
        "status": r.status or "",
        "ageing": r.ageing or 0,
        "completion_date": date_str(r.date_of_completion),
        "reason": r.reason or "",
        "edited": r.edited_flag or "",
    }


def get_record(record_id):
    try:
        return db.session.get(DailyDispatchReport, int(record_id))
    except Exception:
        return None


# ==========================
# IMPORT SALES EXCEL
# ==========================
# Excel is accepted only as an input upload. The uploaded report is not stored.
@app.route("/import_sales_excel", methods=["POST"])
def import_sales_excel():
    try:
        dispatch_date = parse_date(request.form.get("dispatch_date"))
        if not dispatch_date:
            return "Please select a valid Dispatch Plan Date."

        if "excel_file" not in request.files:
            return "No file selected"

        file = request.files["excel_file"]
        if file.filename == "":
            return "No file selected"

        original_filename = secure_filename(file.filename)

        if not original_filename:
            return "Invalid file name"

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

        archived_filename = (
            f"{dispatch_date.strftime('%Y-%m-%d')}_"
            f"{timestamp}_"
            f"{original_filename}"
        )

        archived_path = os.path.join(
            app.config["SALES_UPLOAD_FOLDER"],
            archived_filename
        )

        # Save original Excel for record/audit
        file.save(archived_path)

        # Read the archived copy and import it into SQL Server
        df = pd.read_excel(archived_path)

        for _, row in df.iterrows():
            qty = to_int(row.get("Dispatch Qty", 0))
            db.session.add(DailyDispatchReport(
                dispatch_date=dispatch_date,
                m_code=clean(row.get("M Code")),
                customer_name=clean(row.get("Customer")),
                part_number=clean(row.get("Part Number")),
                size=clean(row.get("Size")),
                dispatch_plan_qty=qty,
                sales_order_no=clean(row.get("Sales Order No")),
                invoice_number=clean(row.get("Invoice No")),
                remarks=clean(row.get("Remarks")),
                complete=0,
                pending_qty=qty,
                status="",
                ageing=0,
                date_of_completion=None,
                reason="",
                edited_flag=""
            ))

        db.session.commit()
        return redirect("/sales")
    except Exception as e:
        db.session.rollback()
        return f"Import Error: {str(e)}"


# ==========================
# LOGIN
# ==========================
@app.route("/admin")
def admin():
    if session.get("role") != "admin":
        return redirect("/")
    return render_template("admin.html")


@app.route("/", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = User.query.filter_by(
            username=request.form["username"],
            password=request.form["password"]
        ).first()

        if user:
            session["username"] = user.username
            session["role"] = user.role
            return redirect({
                "sales": "/sales",
                "fg": "/fg",
                "admin": "/admin"
            }.get(user.role, "/"))

        return render_template("login.html", error="Invalid Login")

    return render_template("login.html")


# ==========================
# SALES
# ==========================
@app.route("/sales")
def sales():
    if session.get("role") not in ["sales", "admin"]:
        return redirect("/")
    return render_template("sales.html", records=[], role=session["role"])


@app.route("/load_sales_date", methods=["POST"])
def load_sales_date():
    try:
        d = parse_date(request.form.get("dispatch_date"))
        if not d:
            return jsonify([])

        rows = DailyDispatchReport.query.filter_by(
            dispatch_date=d
        ).order_by(DailyDispatchReport.id).all()

        return jsonify([record_dict(r) for r in rows])
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/save_sales_bulk", methods=["POST"])
def save_sales_bulk():
    try:
        data = request.get_json() or {}
        d = parse_date(data.get("dispatch_date"))
        rows = data.get("rows", [])

        if not d:
            return jsonify({"success": False, "error": "Invalid dispatch date"})
        if not rows:
            return jsonify({"success": False, "error": "No rows received"})

        for item in rows:
            qty = to_int(item.get("dispatch_qty"))
            db.session.add(DailyDispatchReport(
                dispatch_date=d,
                m_code=clean(item.get("m_code")),
                customer_name=clean(item.get("customer")),
                part_number=clean(item.get("part_number")),
                size=clean(item.get("size")),
                dispatch_plan_qty=qty,
                sales_order_no=clean(item.get("sales_order_no")),
                invoice_number=clean(item.get("invoice_no")),
                remarks=clean(item.get("remarks")),
                complete=0,
                pending_qty=qty,
                status="",
                ageing=0,
                date_of_completion=None,
                reason="",
                edited_flag=""
            ))

        db.session.commit()
        return jsonify({"success": True})
    except Exception as e:
        db.session.rollback()
        return jsonify({"success": False, "error": str(e)})


# ==========================
# FG
# ==========================
@app.route("/fg")
def fg():
    if session.get("role") not in ["fg", "admin"]:
        return redirect("/")
    return render_template("fg.html", role=session["role"])


@app.route("/load_date", methods=["POST"])
def load_date():
    try:
        d = parse_date(request.form.get("dispatch_date"))
        if not d:
            return jsonify([])

        rows = DailyDispatchReport.query.filter(
            DailyDispatchReport.dispatch_date == d,
            DailyDispatchReport.complete < DailyDispatchReport.dispatch_plan_qty
        ).order_by(DailyDispatchReport.id).all()

        return jsonify([record_dict(r) for r in rows])
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/update_fg_bulk", methods=["POST"])
def update_fg_bulk():
    try:
        rows = (request.get_json() or {}).get("rows", [])
        if not rows:
            return jsonify({"success": False, "error": "No rows selected"})

        for item in rows:
            r = get_record(item.get("id") or item.get("excel_row"))
            if not r:
                continue

            r.complete = to_int(item.get("complete"))
            r.pending_qty = to_int(item.get("pending"))
            r.status = clean(item.get("status"))
            r.ageing = to_int(item.get("ageing"))
            r.date_of_completion = parse_date(item.get("completion_date"))
            r.reason = clean(item.get("reason"))

        db.session.commit()
        return jsonify({"success": True})
    except Exception as e:
        db.session.rollback()
        return jsonify({"success": False, "error": str(e)})


@app.route("/update_sales_row", methods=["POST"])
def update_sales_row():
    try:
        data = request.get_json() or {}
        r = get_record(data.get("id") or data.get("excel_row"))

        if not r:
            return jsonify({"success": False, "error": "Record not found"}), 404

        old = {
            "date": date_str(r.dispatch_date),
            "m_code": r.m_code or "",
            "customer": r.customer_name or "",
            "part_number": r.part_number or "",
            "size": r.size or "",
            "dispatch_qty": str(r.dispatch_plan_qty or 0),
            "sales_order_no": r.sales_order_no or "",
            "invoice_no": r.invoice_number or "",
            "remarks": r.remarks or ""
        }

        new_date = parse_date(data.get("dispatch_date"))
        new_values = {
            "date": date_str(new_date),
            "m_code": clean(data.get("m_code")),
            "customer": clean(data.get("customer")),
            "part_number": clean(data.get("part_number")),
            "size": clean(data.get("size")),
            "dispatch_qty": str(to_int(data.get("dispatch_qty"))),
            "sales_order_no": clean(data.get("sales_order_no")),
            "invoice_no": clean(data.get("invoice_no")),
            "remarks": clean(data.get("remarks"))
        }

        edited = [k for k in old if old[k] != new_values[k]]

        r.dispatch_date = new_date
        r.m_code = new_values["m_code"]
        r.customer_name = new_values["customer"]
        r.part_number = new_values["part_number"]
        r.size = new_values["size"]
        r.dispatch_plan_qty = int(new_values["dispatch_qty"])
        r.sales_order_no = new_values["sales_order_no"]
        r.invoice_number = new_values["invoice_no"]
        r.remarks = new_values["remarks"]
        r.pending_qty = max(r.dispatch_plan_qty - (r.complete or 0), 0)
        r.edited_flag = ",".join(edited)

        db.session.commit()
        return jsonify({"success": True})
    except Exception as e:
        db.session.rollback()
        return jsonify({"success": False, "error": str(e)})


# ==========================
# REPORT
# ==========================
@app.route("/report")
def report():
    return render_template("report.html")


@app.route("/load_report", methods=["POST"])
def load_report():
    try:
        start = parse_date(request.form.get("from_date"))
        end = parse_date(request.form.get("to_date"))
        if not start or not end:
            return jsonify([])

        rows = DailyDispatchReport.query.filter(
            DailyDispatchReport.dispatch_date >= start,
            DailyDispatchReport.dispatch_date <= end
        ).order_by(DailyDispatchReport.dispatch_date, DailyDispatchReport.id).all()

        return jsonify([{
            "dispatch_date": date_str(r.dispatch_date),
            "m_code": r.m_code or "",
            "customer": r.customer_name or "",
            "part_number": r.part_number or "",
            "size": r.size or "",
            "dispatch_qty": r.dispatch_plan_qty or 0,
            "complete": r.complete or 0,
            "pending": r.pending_qty or 0,
            "status": r.status or "",
            "ageing": r.ageing or 0,
            "completion_date": date_str(r.date_of_completion),
            "reason": r.reason or ""
        } for r in rows])
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ==========================
# EXCEL DOWNLOADS
# ==========================
# These generate a temporary Excel export from DB data.
# Excel is NOT used for persistence.
def make_excel(data, columns, sheet_name, filename):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(data, columns=columns).to_excel(
            writer, index=False, sheet_name=sheet_name
        )
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/download_report")
def download_report():
    try:
        start = parse_date(request.args.get("from_date"))
        end = parse_date(request.args.get("to_date"))
        if not start or not end:
            return "Invalid date range", 400

        rows = DailyDispatchReport.query.filter(
            DailyDispatchReport.dispatch_date >= start,
            DailyDispatchReport.dispatch_date <= end
        ).order_by(DailyDispatchReport.dispatch_date, DailyDispatchReport.id).all()

        data = [[
            r.dispatch_date, r.m_code, r.customer_name, r.part_number, r.size,
            r.dispatch_plan_qty, r.complete, r.pending_qty, r.status, r.ageing,
            r.date_of_completion, r.reason
        ] for r in rows]

        columns = [
            "Date", "M Code", "Customer", "Part Number", "Size",
            "Dispatch Qty", "Complete", "Pending", "Status", "Ageing",
            "Completion Date", "Reason"
        ]

        return make_excel(
            data, columns, "FG Report", "Dispatch_Report.xlsx"
        )
    except Exception as e:
        return f"Download Error: {str(e)}", 500


@app.route("/download_fg")
def download_fg():
    try:
        d = parse_date(request.args.get("dispatch_date"))
        if not d:
            return "Invalid dispatch date", 400

        rows = DailyDispatchReport.query.filter_by(
            dispatch_date=d
        ).order_by(DailyDispatchReport.id).all()

        data = [[
            r.dispatch_date, r.m_code, r.customer_name, r.part_number, r.size,
            r.dispatch_plan_qty, r.sales_order_no, r.invoice_number, r.remarks,
            r.complete, r.pending_qty, r.status, r.ageing,
            r.date_of_completion, r.reason
        ] for r in rows]

        columns = [
            "Date", "M Code", "Customer", "Part Number", "Size",
            "Dispatch Qty", "Sales Order No.", "Invoice No", "Remarks",
            "Complete", "Pending", "Status", "Ageing",
            "Completion Date", "Reason"
        ]

        return make_excel(
            data, columns, "FG Report", f"FG_Report_{d}.xlsx"
        )
    except Exception as e:
        return f"Download Error: {str(e)}", 500


# ==========================
# SUMMARY
# ==========================
@app.route("/summary_report")
def summary_report():
    return render_template("summary_report.html")


@app.route("/load_summary", methods=["POST"])
def load_summary():
    try:
        month = request.form.get("month", "").strip()
        if not month:
            return jsonify([])

        rows = DailyDispatchReport.query.order_by(
            DailyDispatchReport.dispatch_date
        ).all()

        summary = {}

        for r in rows:
            ds = date_str(r.dispatch_date)
            if not ds.startswith(month):
                continue

            if ds not in summary:
                summary[ds] = {
                    "date": ds,
                    "plan_qty": 0,
                    "dispatch_qty": 0,
                    "daily_production_qty": 0
                }

            summary[ds]["plan_qty"] += r.dispatch_plan_qty or 0
            summary[ds]["dispatch_qty"] += r.complete or 0

        saved = DailyDispatchSummary.query.all()
        for s in saved:
            ds = date_str(s.date)
            if ds in summary:
                summary[ds]["daily_production_qty"] = s.daily_production_qty or 0

        result = []
        for item in summary.values():
            plan = item["plan_qty"]
            dispatch = item["dispatch_qty"]
            item["percentage"] = round((dispatch / plan) * 100) if plan else 0
            result.append(item)

        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/save_summary", methods=["POST"])
def save_summary():
    try:
        data = request.get_json() or []

        for item in data:
            d = parse_date(item.get("date"))
            if not d:
                continue

            percentage = item.get("percentage", 0)
            if isinstance(percentage, str):
                percentage = percentage.replace("%", "").strip()
            try:
                percentage = f"{float(percentage):g}%"
            except Exception:
                percentage = "0%"

            s = DailyDispatchSummary.query.filter_by(date=d).first()

            if not s:
                s = DailyDispatchSummary(date=d)
                db.session.add(s)

            s.total_dispatch_plan_qty = to_int(item.get("plan_qty"))
            s.total_dispatch_qty = to_int(item.get("dispatch_qty"))
            s.percentage = percentage
            s.daily_production_qty = to_int(item.get("daily_production_qty"))

        db.session.commit()
        return jsonify({"success": True})
    except Exception as e:
        db.session.rollback()
        return jsonify({"success": False, "error": str(e)})




# ============================================================
# ADMIN-ONLY DISPATCH DATA MANAGEMENT API
# These endpoints are used by templates/admin_data.html.
# ============================================================
def require_admin_api():
    if session.get("role") != "admin":
        return jsonify({"success": False, "error": "Admin access required."}), 403
    return None


def admin_report_dict(row):
    return {
        "id": row.id,
        "dispatch_date": date_str(row.dispatch_date),
        "m_code": row.m_code or "",
        "customer_name": row.customer_name or "",
        "part_number": row.part_number or "",
        "size": row.size or "",
        "dispatch_plan_qty": row.dispatch_plan_qty or 0,
        "sales_order_no": row.sales_order_no or "",
        "invoice_number": row.invoice_number or "",
        "remarks": row.remarks or "",
        "complete": row.complete or 0,
        "pending_qty": row.pending_qty or 0,
        "status": row.status or "",
        "ageing": row.ageing or 0,
        "date_of_completion": date_str(row.date_of_completion),
        "reason": row.reason or "",
        "edited_flag": row.edited_flag or "",
    }


def admin_summary_dict(row):
    return {
        "id": row.id,
        "date": date_str(row.date),
        "total_dispatch_plan_qty": row.total_dispatch_plan_qty or 0,
        "total_dispatch_qty": row.total_dispatch_qty or 0,
        "percentage": row.percentage or "0%",
        "daily_production_qty": row.daily_production_qty or 0,
    }


@app.route("/admin/data")
@app.route("/admin/data-management")
def admin_data_management():
    if session.get("role") != "admin":
        return redirect("/")
    return render_template("admin_data.html")


@app.route("/admin/data/report", methods=["GET"])
def admin_data_report_list():
    denied = require_admin_api()
    if denied:
        return denied
    try:
        rows = DailyDispatchReport.query.order_by(
            DailyDispatchReport.dispatch_date, DailyDispatchReport.id
        ).all()
        return jsonify([admin_report_dict(row) for row in rows])
    except Exception as exc:
        app.logger.exception("Admin report list failed")
        return jsonify({"success": False, "error": "Could not load dispatch report records."}), 500


@app.route("/admin/data/summary", methods=["GET"])
def admin_data_summary_list():
    denied = require_admin_api()
    if denied:
        return denied
    try:
        rows = DailyDispatchSummary.query.order_by(
            DailyDispatchSummary.date, DailyDispatchSummary.id
        ).all()
        return jsonify([admin_summary_dict(row) for row in rows])
    except Exception as exc:
        app.logger.exception("Admin summary list failed")
        return jsonify({"success": False, "error": "Could not load dispatch summary records."}), 500


@app.route("/admin/data/report/<int:record_id>", methods=["PUT", "DELETE"])
def admin_data_report_record(record_id):
    denied = require_admin_api()
    if denied:
        return denied
    row = db.session.get(DailyDispatchReport, record_id)
    if row is None:
        return jsonify({"success": False, "error": "Dispatch report record not found."}), 404

    if request.method == "DELETE":
        try:
            db.session.delete(row)
            db.session.commit()
            return jsonify({"success": True, "id": record_id})
        except Exception:
            db.session.rollback()
            app.logger.exception("Admin report delete failed")
            return jsonify({"success": False, "error": "Could not delete dispatch report record."}), 500

    data = request.get_json(silent=True) or {}
    try:
        allowed_text = {
            "m_code": 100, "customer_name": 255, "part_number": 255, "size": 100,
            "sales_order_no": 255, "invoice_number": 255, "status": 100, "edited_flag": 100
        }
        for field, max_length in allowed_text.items():
            if field in data:
                value = data[field]
                value = "" if value is None else str(value).strip()
                if len(value) > max_length:
                    return jsonify({"success": False, "error": f"{field} exceeds {max_length} characters."}), 400
                setattr(row, field, value or None)

        for field in ("remarks", "reason"):
            if field in data:
                setattr(row, field, "" if data[field] is None else str(data[field]))

        for field in ("dispatch_plan_qty", "complete", "pending_qty", "ageing"):
            if field in data:
                value = data[field]
                if value in (None, ""):
                    value = 0
                try:
                    value = int(float(value))
                except (TypeError, ValueError):
                    return jsonify({"success": False, "error": f"{field} must be a whole number."}), 400
                setattr(row, field, value)

        if "dispatch_date" in data:
            parsed = parse_date(data["dispatch_date"])
            if not parsed:
                return jsonify({"success": False, "error": "Dispatch Date is required and must be valid."}), 400
            row.dispatch_date = parsed

        if "date_of_completion" in data:
            raw = data["date_of_completion"]
            parsed = parse_date(raw) if raw else None
            if raw and not parsed:
                return jsonify({"success": False, "error": "Date of Completion must be a valid date."}), 400
            row.date_of_completion = parsed

        db.session.commit()
        return jsonify({"success": True, "record": admin_report_dict(row)})
    except Exception:
        db.session.rollback()
        app.logger.exception("Admin report update failed")
        return jsonify({"success": False, "error": "Could not save dispatch report changes."}), 500


@app.route("/admin/data/summary/<int:record_id>", methods=["PUT", "DELETE"])
def admin_data_summary_record(record_id):
    denied = require_admin_api()
    if denied:
        return denied
    row = db.session.get(DailyDispatchSummary, record_id)
    if row is None:
        return jsonify({"success": False, "error": "Dispatch summary record not found."}), 404

    if request.method == "DELETE":
        try:
            db.session.delete(row)
            db.session.commit()
            return jsonify({"success": True, "id": record_id})
        except Exception:
            db.session.rollback()
            app.logger.exception("Admin summary delete failed")
            return jsonify({"success": False, "error": "Could not delete dispatch summary record."}), 500

    data = request.get_json(silent=True) or {}
    try:
        if "date" in data:
            parsed = parse_date(data["date"])
            if not parsed:
                return jsonify({"success": False, "error": "Summary Date is required and must be valid."}), 400
            duplicate = DailyDispatchSummary.query.filter(
                DailyDispatchSummary.date == parsed,
                DailyDispatchSummary.id != record_id
            ).first()
            if duplicate:
                return jsonify({"success": False, "error": "A summary record already exists for this date."}), 409
            row.date = parsed

        for field in ("total_dispatch_plan_qty", "total_dispatch_qty", "daily_production_qty"):
            if field in data:
                value = data[field]
                if value in (None, ""):
                    value = 0
                try:
                    value = int(float(value))
                except (TypeError, ValueError):
                    return jsonify({"success": False, "error": f"{field} must be a whole number."}), 400
                setattr(row, field, value)

        if "percentage" in data:
            percentage = str(data["percentage"] or "0").replace("%", "").strip()
            try:
                row.percentage = f"{float(percentage):g}%"
            except ValueError:
                return jsonify({"success": False, "error": "Percentage must be numeric."}), 400

        db.session.commit()
        return jsonify({"success": True, "record": admin_summary_dict(row)})
    except Exception:
        db.session.rollback()
        app.logger.exception("Admin summary update failed")
        return jsonify({"success": False, "error": "Could not save dispatch summary changes."}), 500


@app.route("/logout")
def logout():
    session.clear()
    return redirect("/")


# ==========================
# FIRST RUN
# ==========================
if __name__ == "__main__":
    with app.app_context():
        try:
            # Verify SQL Server before starting the application.
            with db.engine.connect() as connection:
                connection.exec_driver_sql("SELECT 1")

            # Create the application tables if they do not exist.
            db.create_all()

            if not User.query.filter_by(username="admin").first():
                db.session.add_all([
                    User(username="admin", password="admin123", role="admin"),
                    User(username="sales", password="sales123", role="sales"),
                    User(username="fg", password="fg123", role="fg")
                ])
                db.session.commit()

            print("SQL Server connection successful.")
            print(f"Database: {MSSQL_DATABASE}")

        except Exception as e:
            print("\nDATABASE CONNECTION ERROR")
            print("------------------------")
            print(str(e))
            print("\nCheck MSSQL_SERVER, MSSQL_DATABASE, credentials,")
            print("ODBC Driver 18 for SQL Server, and SQL Server availability.")
            raise

    app.run(debug=True)
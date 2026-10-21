
from flask_sqlalchemy import SQLAlchemy


def init_models(db: SQLAlchemy):
    """Register all application models on the supplied SQLAlchemy object."""

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

    return User, DailyDispatchReport, DailyDispatchSummary
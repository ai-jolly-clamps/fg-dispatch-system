import pandas as pd
from sqlalchemy import create_engine

# =========================
# DATABASE CONNECTION
# =========================

server = r"PEW-SERVER"
database = "fg_dispatch_db"

connection_string = (
    "mssql+pyodbc://@"
    + server
    + "/"
    + database
    + "?driver=ODBC+Driver+18+for+SQL+Server"
    "&trusted_connection=yes"
    "&TrustServerCertificate=yes"
)

engine = create_engine(connection_string)

# =========================
# EXCEL FILE
# =========================

file_path = r"E:/vaidehi/FG_Dispatch_System_sandip/Daily Dispatch adherence report.xlsx"

# Header is on Excel row 5 -> pandas header=4
df = pd.read_excel(
    file_path,
    sheet_name="FG Report",
    header=4
)

# =========================
# RENAME EXCEL COLUMNS
# =========================

column_mapping = {
    "dispatch date": "dispatch_date",
    "M Code": "m_code",
    "Customer Name": "customer_name",
    "Part Number": "part_number",
    "Size": "size",
    "Dispatch Plan Qty.": "dispatch_plan_qty",
    "Sales Order No": "sales_order_no",
    "Invoice Number": "invoice_number",
    "Remarks": "remarks",
    "Complete": "complete",
    "Pending Qty": "pending_qty",
    "Status": "status",
    "Ageing": "ageing",
    "Date of Completion": "date_of_completion",
    "Reason": "reason",
    "Edited Flag": "edited_flag"
}

df = df.rename(columns=column_mapping)

# =========================
# REMOVE COMPLETELY EMPTY ROWS
# =========================

# df = df.dropna(
#     how="all",
#     subset=list(column_mapping.keys())
# )

# =========================
# KEEP ONLY SQL TABLE COLUMNS
# =========================

sql_columns = [
    "dispatch_date",
    "m_code",
    "customer_name",
    "part_number",
    "size",
    "dispatch_plan_qty",
    "sales_order_no",
    "invoice_number",
    "remarks",
    "complete",
    "pending_qty",
    "status",
    "ageing",
    "date_of_completion",
    "reason",
    "edited_flag"
]

df = df[sql_columns]

# =========================
# DATE COLUMNS
# =========================

df["dispatch_date"] = pd.to_datetime(
    df["dispatch_date"],
    errors="coerce"
)

df["date_of_completion"] = pd.to_datetime(
    df["date_of_completion"],
    errors="coerce"
)

# =========================
# APPEND TO SQL TABLE
# =========================

df.to_sql(
    "daily_dispatch_report",
    con=engine,
    if_exists="append",
    index=False,
    chunksize=500,
    method="multi"
)

print(f"{len(df)} rows successfully appended.")
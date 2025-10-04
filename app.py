# -*- coding: utf-8 -*-
import os, io, json, uuid, socket, logging
from datetime import datetime
from functools import wraps
from urllib.parse import quote

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, jsonify, send_file
)
from werkzeug.security import generate_password_hash, check_password_hash
from dotenv import load_dotenv
from flask_wtf import CSRFProtect
import pyodbc
import pandas as pd

load_dotenv()
app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY', 'change-me')
app.config.update({
    'SESSION_COOKIE_HTTPONLY': True,
    'SESSION_COOKIE_SAMESITE': 'Lax'
})

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler('app.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)
csrf = CSRFProtect(app)

# ------------------------
# Environment / Settings
# ------------------------
DB_SERVER   = os.getenv('DB_SERVER', 'localhost')
DB_NAME     = os.getenv('DB_NAME', 'EquipmentDB')  # אפשר לשנות ל- GINEGAR-IT בקובץ .env
DB_TRUSTED  = os.getenv('DB_TRUSTED', '1') == '1'
DB_USERNAME = os.getenv('DB_USERNAME')
DB_PASSWORD = os.getenv('DB_PASSWORD')
DB_DRIVER   = os.getenv('DB_DRIVER', 'ODBC Driver 17 for SQL Server')

ZEBRA_IP    = os.getenv('ZEBRA_PRINTER_IP', '')
ZEBRA_PORT  = int(os.getenv('ZEBRA_PRINTER_PORT', '9100'))
ZEBRA_SEND  = os.getenv('ZEBRA_SEND_TO_PRINTER', '0').lower() in ('1','true','yes')
COMPANY_NAME= os.getenv('COMPANY_NAME', 'Ginegar')
LABEL_WIDTH_MM  = int(os.getenv('LABEL_WIDTH_MM', '50'))
LABEL_HEIGHT_MM = int(os.getenv('LABEL_HEIGHT_MM', '30'))

# ------------------------
# DB helpers
# ------------------------
def build_conn_str() -> str:
    driver = (DB_DRIVER or '').strip() or 'ODBC Driver 17 for SQL Server'
    base = f"Driver={{{driver}}};Server={DB_SERVER};Database={DB_NAME};TrustServerCertificate=yes;"
    if DB_TRUSTED:
        return base + 'Trusted_Connection=yes;'
    return base + f'UID={DB_USERNAME};PWD={DB_PASSWORD};'

def get_db():
    return pyodbc.connect(build_conn_str())

@app.after_request
def add_no_cache_headers(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response

# ------------------------
# Bootstrap schema (idempotent)
# ------------------------
def bootstrap_db():
    with get_db() as conn:
        cur = conn.cursor()
        # equipment
        cur.execute("""
IF NOT EXISTS (SELECT * FROM sys.objects WHERE name='equipment' AND type='U')
BEGIN
    CREATE TABLE dbo.equipment (
        id INT IDENTITY(1,1) PRIMARY KEY,
        name NVARCHAR(100),
        vendor NVARCHAR(100),
        warranty_expiry DATE,
        status NVARCHAR(50),
        barcode NVARCHAR(100),
        history NVARCHAR(MAX)
    )
END
""")
        # users
        cur.execute("""
IF NOT EXISTS (SELECT * FROM sys.objects WHERE name='users' AND type='U')
BEGIN
    CREATE TABLE dbo.users (
        id INT IDENTITY(1,1) PRIMARY KEY,
        username NVARCHAR(50) UNIQUE NOT NULL,
        password_hash NVARCHAR(255) NOT NULL,
        email NVARCHAR(100),
        role NVARCHAR(20) NOT NULL DEFAULT 'User',
        status NVARCHAR(20) DEFAULT 'Active'
    )
END
""")
        # created_at on users (hotfix)
        cur.execute("IF COL_LENGTH('dbo.users','created_at') IS NULL ALTER TABLE dbo.users ADD created_at DATETIME NULL DEFAULT GETDATE()")
        # employees
        cur.execute("""
IF NOT EXISTS (SELECT * FROM sys.objects WHERE name='employees' AND type='U')
BEGIN
    CREATE TABLE dbo.employees (
        id INT IDENTITY(1,1) PRIMARY KEY,
        first_name NVARCHAR(100) NOT NULL,
        last_name NVARCHAR(100) NOT NULL,
        emp_no NVARCHAR(50) UNIQUE NOT NULL,
        random_id NVARCHAR(50) NOT NULL,
        station_no NVARCHAR(50) NULL,
        created_at DATETIME DEFAULT GETDATE()
    )
END
""")
        # stations
        cur.execute("""
IF NOT EXISTS (SELECT * FROM sys.objects WHERE name='stations' AND type='U')
BEGIN
    CREATE TABLE dbo.stations (
        id INT IDENTITY(1,1) PRIMARY KEY,
        station_no NVARCHAR(50) UNIQUE NOT NULL,
        display_name NVARCHAR(100) NULL,
        created_at DATETIME DEFAULT GETDATE()
    )
END
""")
        # add-on columns for equipment
        cur.execute("IF COL_LENGTH('dbo.equipment','assigned_to') IS NULL ALTER TABLE dbo.equipment ADD assigned_to NVARCHAR(100) NULL")
        cur.execute("IF COL_LENGTH('dbo.equipment','station') IS NULL ALTER TABLE dbo.equipment ADD station NVARCHAR(50) NULL")
        cur.execute("IF COL_LENGTH('dbo.equipment','placement') IS NULL ALTER TABLE dbo.equipment ADD placement NVARCHAR(20) NULL")
        cur.execute("IF COL_LENGTH('dbo.equipment','sold_to') IS NULL ALTER TABLE dbo.equipment ADD sold_to NVARCHAR(100) NULL")
        # link users -> employees
        cur.execute("IF COL_LENGTH('dbo.users','employee_id') IS NULL ALTER TABLE dbo.users ADD employee_id INT NULL")
        cur.execute("""
IF NOT EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name='FK_users_employees')
AND COL_LENGTH('dbo.users','employee_id') IS NOT NULL
BEGIN
    ALTER TABLE dbo.users WITH NOCHECK
    ADD CONSTRAINT FK_users_employees FOREIGN KEY (employee_id) REFERENCES dbo.employees(id)
END
""")
        # audit table
        cur.execute("""
IF NOT EXISTS (SELECT * FROM sys.objects WHERE name='equipment_audit' AND type='U')
BEGIN
    CREATE TABLE dbo.equipment_audit (
        audit_id BIGINT IDENTITY(1,1) PRIMARY KEY,
        equipment_id INT NULL,
        action NVARCHAR(20) NOT NULL,
        changed_by NVARCHAR(100) NULL,
        changed_at DATETIME NOT NULL DEFAULT GETDATE(),
        before_json NVARCHAR(MAX) NULL,
        after_json NVARCHAR(MAX) NULL
    )
END
""")
        # default admin user
        cur.execute("SELECT id FROM dbo.users WHERE username='admin'")
        if not cur.fetchone():
            cur.execute(
                "INSERT INTO dbo.users (username, password_hash, email, role, status) VALUES (?,?,?,?,?)",
                ('admin', generate_password_hash('admin'), 'admin@example.com', 'Admin', 'Active')
            )
        conn.commit()

try:
    bootstrap_db()
except Exception as e:
    logger.error('Bootstrap DB failed: %s', e)

# ------------------------
# Global context
# ------------------------
@app.context_processor
def inject_globals():
    return {
        'role': session.get('role'),
        'current_username': session.get('username')
    }

# ------------------------
# Auth decorators
# ------------------------
def login_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return fn(*args, **kwargs)
    return wrapper

def admin_required(fn):
    @wraps(fn)
    def wrapper(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if session.get('role') != 'Admin':
            flash('אין לך הרשאה לבצע פעולה זו', 'warning')
            return redirect(url_for('dashboard'))
        return fn(*args, **kwargs)
    return wrapper

# ------------------------
# Auth routes
# ------------------------
@app.route('/', methods=['GET', 'POST'])
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = (request.form.get('username') or '').strip()
        password = request.form.get('password') or ''
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("SELECT id, password_hash, role, status FROM dbo.users WHERE username=?", (username,))
            row = cur.fetchone()
            if row and row[3] == 'Active' and check_password_hash(row[1], password):
                session['user_id'] = row[0]
                session['username'] = username
                session['role'] = row[2]
                return redirect(url_for('dashboard'))
        flash('שם משתמש או סיסמה שגויים, או שהמשתמש אינו פעיל', 'danger')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    session.clear()
    flash('התנתקת בהצלחה', 'success')
    return redirect(url_for('login'))


# ------------------------
# Reports Page
# ------------------------
@app.route('/reports')
@login_required
def reports():
    return render_template('reports.html', title='דוחות')
# ------------------------
# Dashboard
# ------------------------
@app.route('/dashboard')
@login_required
def dashboard():
    role = session.get('role', 'User')  # או איך שאתה שומר את התפקיד

    with get_db() as conn:
        cur = conn.cursor()

        # פריטי ציוד כלליים
        cur.execute('SELECT COUNT(*) FROM dbo.equipment')
        eq_count = cur.fetchone()[0] or 0

        # משתמשים פעילים
        cur.execute("SELECT COUNT(*) FROM dbo.users WHERE status='Active'")
        active_users = cur.fetchone()[0] or 0

        # מחשבים בשימוש (לא במחסן)
        cur.execute("""
            SELECT COUNT(*)
FROM dbo.equipment e
JOIN dbo.equipment_types t ON e.equipment_type_id = t.id
WHERE t.name LIKE N'%מחשב%'
  AND LTRIM(RTRIM(e.placement)) = N'מחסן'
  AND LTRIM(RTRIM(e.status)) IN (N'חדש', N'תקין');

        """)
        computers_in_use = cur.fetchone()[0] or 0

        # מדפסות בשימוש (לא במחסן)
        cur.execute("""
            SELECT COUNT(*)
FROM dbo.equipment e
JOIN dbo.equipment_types t ON e.equipment_type_id = t.id
WHERE t.name LIKE N'%מדפסת%'
  AND LTRIM(RTRIM(e.placement)) = N'מחסן'
  AND LTRIM(RTRIM(e.status)) IN (N'חדש', N'תקין');

        """)
        printers_in_use = cur.fetchone()[0] or 0

        # מסכים במחסן
        cur.execute("""
            SELECT COUNT(*)
FROM dbo.equipment e
JOIN dbo.equipment_types t ON e.equipment_type_id = t.id
WHERE t.name LIKE N'%מסך%'
  AND LTRIM(RTRIM(e.placement)) = N'מחסן'
  AND LTRIM(RTRIM(e.status)) IN (N'חדש', N'תקין');

        """)
        screens_in_stock = cur.fetchone()[0] or 0

        # טונרים קריטיים (אחרון במלאי לפי סוג מדפסת)
        cur.execute("""
    SELECT t.printer_type, COUNT(*) AS in_stock_count
    FROM dbo.toners t
    WHERE t.status = 'InStock'
    GROUP BY t.printer_type
    HAVING COUNT(*) <= 1
""")
        critical_toners = cur.fetchall()
        last_toner_count = len(critical_toners)

        # תוכנה לחידוש פחות מ-3 חודשים
        cur.execute("""
            SELECT COUNT(*)
            FROM dbo.software
            WHERE DATEDIFF(MONTH, GETDATE(), renewal_next) < 3
        """)
        software_expiring = cur.fetchone()[0] or 0

    return render_template(
        'dashboard.html',
        role=role,
        eq_count=eq_count,
        active_users=active_users,
        computers_in_use=computers_in_use,
        printers_in_use=printers_in_use,
        screens_in_stock=screens_in_stock,
        last_toner_count=last_toner_count,
        software_expiring=software_expiring
    )




# ------------------------
# Legacy users list (optional UI)
# ------------------------
@app.route('/users')
@admin_required
def users():
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute('SELECT id, username, email, role, status FROM dbo.users ORDER BY id DESC')
        rows = cur.fetchall()
        data = [{'id': r[0], 'username': r[1], 'email': r[2], 'role': r[3], 'status': r[4]} for r in rows]
    return render_template('users.html', users=data)

@app.route('/users/add', methods=['GET', 'POST'])
@admin_required
def add_user():
    if request.method == 'POST':
        username = (request.form.get('username') or '').strip()
        password = request.form.get('password') or ''
        email = (request.form.get('email') or '').strip()
        role = (request.form.get('role') or 'User').strip()
        if not username or not password:
            flash('שם משתמש וסיסמה הינם שדות חובה', 'warning')
            return render_template('add_user.html')
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute('SELECT 1 FROM dbo.users WHERE username=?', (username,))
            if cur.fetchone():
                flash('שם המשתמש כבר קיים', 'danger')
                return render_template('add_user.html')
            cur.execute('INSERT INTO dbo.users (username, password_hash, email, role, status) VALUES (?, ?, ?, ?, "Active")',
                        (username, generate_password_hash(password), email, role))
            conn.commit()
        flash('המשתמש נוסף בהצלחה', 'success')
        return redirect(url_for('users'))
    return render_template('add_user.html')

# ------------------------
# Equipment pages
# ------------------------
@app.route('/equipment')
@login_required
def equipment_list():
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT e.id, e.name, e.vendor, e.warranty_expiry, e.status, e.barcode,
                   e.assigned_to, e.station, e.placement, e.sold_to,
                   e.equipment_type_id, et.name as equipment_type
            FROM dbo.equipment e
            LEFT JOIN dbo.equipment_types et
              ON e.equipment_type_id = et.id
            ORDER BY e.id DESC
        """)
        rows = cur.fetchall()

        sdate = lambda d: d.strftime('%Y-%m-%d') if d else ''
        items = [{
            'id': r[0],
            'name': r[1],
            'vendor': r[2],
            'warranty_expiry': sdate(r[3]),
            'status': r[4],
            'barcode': r[5],
            'assigned_to': r[6],
            'station': r[7],
            'placement': r[8],
            'sold_to': r[9],
            'equipment_type_id': r[10],
            'equipment_type': r[11] or ""   # כאן יהיה השם של סוג הציוד
        } for r in rows]

    return render_template('main.html', items=items)

@app.route('/equipment_list')
@login_required
def equipment_filtered_list():
    equipment_type = request.args.get('type')
    placement = request.args.get('placement')
    status = request.args.get('status')  # 'חדש,תקין'
    statuses = status.split(',') if status else []

    query = """
        SELECT e.id, e.name, e.status, e.placement, t.name AS equipment_type
        FROM dbo.equipment e
        JOIN dbo.equipment_types t ON e.equipment_type_id = t.id
        WHERE 1=1
    """
    params = []

    if equipment_type:
        query += " AND t.name LIKE ?"
        params.append(f"%{equipment_type}%")
    if placement:
        query += " AND LTRIM(RTRIM(e.placement)) = ?"
        params.append(placement)
    if statuses:
        query += " AND LTRIM(RTRIM(e.status)) IN ({})".format(','.join(['?']*len(statuses)))
        params.extend(statuses)

    with get_db() as conn:
        cur = conn.cursor()
        cur.execute(query, params)
        equipment = cur.fetchall()

    return render_template('equipment_list.html', equipment=equipment)




@app.route('/equipment/add', methods=['GET', 'POST'])
@login_required
def add_equipment():
    if request.method == 'POST':
        # קריאה וטיוב הקלטים
        name = (request.form.get('name') or '').strip()
        vendor = (request.form.get('vendor') or '').strip()
        equipment_type = request.form.get('equipment_type')  # <--- כאן השדה החדש
        warranty_expiry_raw = request.form.get('warranty_expiry')
        status_ = request.form.get('status') or 'Available'
        history = (request.form.get('history') or '').strip()
        assigned_to = request.form.get('assigned_to')
        station = request.form.get('station')
        placement = (request.form.get('placement') or '').strip()
        sold_to = (request.form.get('sold_to') or '').strip()

        # המרת תאריך אחריות ל-MMYY
        if warranty_expiry_raw:
            try:
                warranty_expiry = datetime.strptime(warranty_expiry_raw, "%Y-%m-%d")
                warranty_mm_yy = warranty_expiry.strftime("%m%y")
            except:
                warranty_expiry = None
                warranty_mm_yy = "0000"
        else:
            warranty_expiry = None
            warranty_mm_yy = "0000"

        # יצירת barcode אוטומטי אם לא הוזן
        barcode = request.form.get('barcode')
        if not barcode:
            import random
            vendor_code = (vendor[:3].upper() if vendor else "XXX")
            unique_id = f"{random.randint(1000, 9999)}"
            barcode = f"{vendor_code}-{warranty_mm_yy}-{unique_id}"

        try:
            # חיבור למסד הנתונים
            with get_db() as conn:
                cur = conn.cursor()
                # בדיקה אם ה-barcode כבר קיים
                cur.execute("SELECT id FROM dbo.equipment WHERE barcode = ?", (barcode,))
                if cur.fetchone():
                    flash("This barcode already exists in the system!", "danger")
                    return redirect(url_for('add_equipment'))

                # הוספת הציוד למסד כולל סוג ציוד
                cur.execute("""
                    INSERT INTO dbo.equipment
                    (name, vendor, equipment_type_id, warranty_expiry, status, barcode, history, assigned_to, station, placement, sold_to)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (name, vendor, equipment_type, warranty_expiry, status_, barcode, history, assigned_to, station, placement, sold_to))
                conn.commit()

            flash(f"Equipment added successfully! Barcode: {barcode}", "success")
            return redirect(url_for('equipment_list'))

        except pyodbc.IntegrityError:
            flash("Database error: Could not add equipment. Check barcode uniqueness.", "danger")
            return redirect(url_for('add_equipment'))

    # GET request – להביא את רשימת סוגי הציוד
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, name FROM dbo.equipment_types ORDER BY name")
        equipment_types = [{"id": row.id, "name": row.name} for row in cur.fetchall()]
        
        cur.execute("SELECT id, name FROM dbo.vendors ORDER BY name")
        vendors = [{"id": row[0], "name": row[1]} for row in cur.fetchall()]

    return render_template('add_equipment.html', equipment_types=equipment_types, vendors=vendors)








@app.route('/equipment/edit/<int:eq_id>', methods=['GET', 'POST'])
@login_required
def edit_equipment(eq_id: int):
    with get_db() as conn:
        cur = conn.cursor()

        if request.method == 'POST':
            # --- Load current values (before) ---
            cur.execute("""
                SELECT id, name, vendor, warranty_expiry, status, barcode, history,
                       assigned_to, station, placement, sold_to, equipment_type_id
                FROM dbo.equipment
                WHERE id=?
            """, (eq_id,))
            rb = cur.fetchone()
            before = None
            if rb:
                before = {
                    'id': rb[0], 'name': rb[1], 'vendor': rb[2],
                    'warranty_expiry': rb[3].strftime('%Y-%m-%d') if rb[3] else None,
                    'status': rb[4], 'barcode': rb[5], 'history': rb[6],
                    'assigned_to': rb[7], 'station': rb[8], 'placement': rb[9],
                    'sold_to': rb[10], 'equipment_type_id': rb[11]
                }

            # --- New values from form ---
            name = (request.form.get('name') or '').strip()
            vendor = (request.form.get('vendor') or '').strip()
            warranty_expiry = request.form.get('warranty_expiry') or None
            status_ = (request.form.get('status') or '').strip()
            barcode = (request.form.get('barcode') or '').strip()
            history = (request.form.get('history') or '').strip()
            assigned_to = (request.form.get('assigned_to') or '').strip() or None
            station = (request.form.get('station') or '').strip() or None
            placement = (request.form.get('placement') or '').strip() or None
            sold_to = (request.form.get('sold_to') or '').strip() or None
            equipment_type_id = request.form.get('equipment_type') or None

            # --- Update equipment ---
            cur.execute("""
                UPDATE dbo.equipment
                SET name=?, vendor=?, warranty_expiry=?, status=?, barcode=?, history=?,
                    assigned_to=?, station=?, placement=?, sold_to=?, equipment_type_id=?
                WHERE id=?
            """, (name, vendor, warranty_expiry, status_, barcode, history,
                  assigned_to, station, placement, sold_to, equipment_type_id, eq_id))

            # --- After values for audit ---
            after = {
                'id': eq_id, 'name': name, 'vendor': vendor,
                'warranty_expiry': warranty_expiry, 'status': status_, 'barcode': barcode,
                'history': history, 'assigned_to': assigned_to, 'station': station,
                'placement': placement, 'sold_to': sold_to, 'equipment_type_id': equipment_type_id
            }

            # --- Insert into audit log ---
            cur2 = conn.cursor()
            cur2.execute("""
                INSERT INTO dbo.equipment_audit 
                    (equipment_id, action, changed_by, before_json, after_json)
                VALUES (?, 'UPDATE', ?, ?, ?)
            """, (eq_id, session.get('username'),
                  json.dumps(before or {}, default=str),
                  json.dumps(after, default=str)))

            conn.commit()
            flash('הפריט עודכן', 'success')
            return redirect(url_for('equipment_list'))

        # --- GET: Load equipment ---
        cur.execute("""
            SELECT id, name, vendor, warranty_expiry, status, barcode, history,
                   assigned_to, station, placement, sold_to, equipment_type_id
            FROM dbo.equipment
            WHERE id=?
        """, (eq_id,))
        r = cur.fetchone()
        if not r:
            flash('פריט הציוד לא נמצא', 'warning')
            return redirect(url_for('equipment_list'))

        item = {
            'id': r[0], 'name': r[1], 'vendor': r[2],
            'warranty_expiry': r[3].strftime('%Y-%m-%d') if r[3] else '',
            'status': r[4], 'barcode': r[5], 'history': r[6],
            'assigned_to': r[7], 'station': r[8], 'placement': r[9],
            'sold_to': r[10], 'equipment_type_id': r[11]
        }

        # --- Load equipment types ---
        cur.execute("SELECT id, name FROM dbo.equipment_types ORDER BY name")
        equipment_types = [{"id": row[0], "name": row[1]} for row in cur.fetchall()]
        
        cur.execute("SELECT id, name FROM dbo.vendors ORDER BY name")
        vendors = [{"id": row[0], "name": row[1]} for row in cur.fetchall()]


    return render_template('edit_equipment.html', item=item, equipment_types=equipment_types, vendors=vendors)


# ------------------------
# Equipment API & Export
# ------------------------
@app.route('/api/equipment')
@login_required
def api_equipment():
    q = (request.args.get('q') or '').strip()
    status = (request.args.get('status') or '').strip()
    vendor = (request.args.get('vendor') or '').strip()
    sort = (request.args.get('sort') or 'id').lower()
    order = 'DESC' if (request.args.get('order') or 'desc').lower() == 'desc' else 'ASC'
    allowed = {'id','name','vendor','warranty_expiry','status','barcode'}
    if sort not in allowed:
        sort = 'id'
    where, params = [], []
    if q:
        where.append('(name LIKE ? OR vendor LIKE ? OR barcode LIKE ?)')
        like = f"%{q}%"; params += [like, like, like]
    if status:
        where.append('status = ?'); params.append(status)
    if vendor:
        where.append('vendor = ?'); params.append(vendor)
    where_sql = (' WHERE ' + ' AND '.join(where)) if where else ''
    sql = f"""
SELECT id, name, vendor, warranty_expiry, status, barcode,
       assigned_to, station, placement, sold_to
FROM dbo.equipment
{where_sql}
ORDER BY {sort} {order}
"""
    with get_db() as conn:
        cur = conn.cursor(); cur.execute(sql, params); rows = cur.fetchall()
        sdate = lambda d: d.strftime('%Y-%m-%d') if d else ''
        data = [{
            'id': r[0], 'name': r[1], 'vendor': r[2], 'warranty_expiry': sdate(r[3]),
            'status': r[4], 'barcode': r[5], 'assigned_to': r[6], 'station': r[7],
            'placement': r[8], 'sold_to': r[9]
        } for r in rows]
    return jsonify(data)

@app.route('/export/csv')
@login_required
def export_csv():
    data = api_equipment().get_json()
    buf = io.StringIO()
    pd.DataFrame(data).to_csv(buf, index=False)
    mem = io.BytesIO(buf.getvalue().encode('utf-8-sig'))
    return send_file(mem, mimetype='text/csv', as_attachment=True,
                     download_name=f"equipment_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")

@app.route('/export/xlsx')
@login_required
def export_xlsx():
    data = api_equipment().get_json()
    mem = io.BytesIO()
    with pd.ExcelWriter(mem, engine='openpyxl') as writer:
        pd.DataFrame(data).to_excel(writer, index=False, sheet_name='Equipment')
    mem.seek(0)
    return send_file(mem, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,
                     download_name=f"equipment_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

# ------------------------
# Quick search APIs (datalist)
# ------------------------
@app.route('/api/employees')
@login_required
def api_employees_search():
    q = (request.args.get('q') or '').strip()
    with get_db() as conn:
        cur = conn.cursor()
        if q:
            like = f"%{q}%"
            cur.execute("SELECT TOP 50 id, first_name, last_name, emp_no, station_no FROM dbo.employees WHERE first_name LIKE ? OR last_name LIKE ? OR emp_no LIKE ? ORDER BY id DESC",
                        (like, like, like))
        else:
            cur.execute("SELECT TOP 50 id, first_name, last_name, emp_no, station_no FROM dbo.employees ORDER BY id DESC")
        rows = cur.fetchall()
    data = [{'id': r[0], 'first_name': r[1], 'last_name': r[2], 'emp_no': r[3], 'station_no': r[4]} for r in rows]
    return jsonify(data)

@app.route('/api/stations')
@login_required
def api_stations_search():
    q = (request.args.get('q') or '').strip()
    with get_db() as conn:
        cur = conn.cursor()
        if q:
            like = f"%{q}%"; cur.execute("SELECT TOP 20 id, station_no, display_name FROM dbo.stations WHERE station_no LIKE ? OR display_name LIKE ? ORDER BY id DESC", (like, like))
        else:
            cur.execute("SELECT TOP 20 id, station_no, display_name FROM dbo.stations ORDER BY id DESC")
        rows = cur.fetchall()
    data = [{'id': r[0], 'station_no': r[1], 'display_name': r[2]} for r in rows]
    return jsonify(data)

# ------------------------
# Zebra label (optional send), simple ZPL
# ------------------------
def build_zpl(item: dict) -> str:
    mm_to_dots = lambda mm: int(mm * 8)  # ~203dpi
    width = mm_to_dots(LABEL_WIDTH_MM)
    height = mm_to_dots(LABEL_HEIGHT_MM)
    name = item.get('name', '')
    vendor = item.get('vendor', '')
    barcode = item.get('barcode', '') or str(item.get('id', ''))
    return f"""
^XA
^PW{width}
^LL{height}
^CI28
^CF0,28
^FO20,20^FD{COMPANY_NAME}^FS
^CF0,24
^FO20,60^FD{name}^FS
^FO20,95^FD{vendor}^FS
^BY2,2,60
^FO20,130^BCN,60,Y,N,N
^FD{barcode}^FS
^XZ
"""

@app.route('/equipment/<int:eq_id>/print_label')
@login_required
def print_label(eq_id: int):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, name, vendor, barcode FROM dbo.equipment WHERE id=?", (eq_id,))
        r = cur.fetchone()
        if not r:
            flash('פריט לא נמצא', 'warning')
            return redirect(url_for('equipment_list'))
        item = {'id': r[0], 'name': r[1], 'vendor': r[2], 'barcode': r[3]}
        zpl = build_zpl(item)
        if ZEBRA_SEND and ZEBRA_IP:
            try:
                with socket.create_connection((ZEBRA_IP, ZEBRA_PORT), timeout=5) as s:
                    s.sendall(zpl.encode('utf-8'))
                flash('התווית נשלחה להדפסה', 'success')
                return redirect(url_for('equipment_list'))
            except Exception as e:
                flash(f'שגיאת הדפסה: {e}', 'danger')
                return redirect(url_for('equipment_list'))
        mem = io.BytesIO(zpl.encode('utf-8'))
        return send_file(mem, mimetype='text/plain', as_attachment=True, download_name=f'label_{eq_id}.zpl')

# ------------------------
# SYSTEM – Admin page + APIs (DataTables)
# ------------------------
@app.route('/system')
@admin_required
def system():
    counts = {'users': 0, 'employees': 0, 'stations': 0}
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute('SELECT COUNT(*) FROM dbo.users'); counts['users'] = cur.fetchone()[0]
        cur.execute('SELECT COUNT(*) FROM dbo.employees'); counts['employees'] = cur.fetchone()[0]
        cur.execute('SELECT COUNT(*) FROM dbo.stations'); counts['stations'] = cur.fetchone()[0]
    return render_template('system.html', counts=counts)

@app.route('/system/api/counts')
@admin_required
def system_counts():
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute('SELECT COUNT(*) FROM dbo.users'); u = cur.fetchone()[0]
        cur.execute('SELECT COUNT(*) FROM dbo.employees'); e = cur.fetchone()[0]
        cur.execute('SELECT COUNT(*) FROM dbo.stations'); s = cur.fetchone()[0]
    return jsonify({'users': u, 'employees': e, 'stations': s})

# ---- Users API (JSON for DataTables) ----
@app.route('/system/api/users')
@admin_required
def api_users_list():
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
SELECT u.id, u.username, u.email, u.role, u.status, u.created_at,
       u.employee_id, e.emp_no, e.first_name, e.last_name
FROM dbo.users u
LEFT JOIN dbo.employees e ON e.id = u.employee_id
ORDER BY u.id DESC
""")
        rows = cur.fetchall()
    data = [{
        'id': r[0], 'username': r[1], 'email': r[2], 'role': r[3], 'status': r[4],
        'created_at': r[5].strftime('%Y-%m-%d %H:%M') if r[5] else '',
        'employee_id': r[6], 'emp_no': r[7], 'emp_first_name': r[8], 'emp_last_name': r[9]
    } for r in rows]
    return jsonify({'data': data})

@app.route('/system/api/users', methods=['POST'])
@admin_required
@csrf.exempt
def api_users_create():
    p = request.get_json(force=True, silent=True) or {}
    username = (p.get('username') or '').strip()
    password = p.get('password') or ''
    email = (p.get('email') or '').strip() or None
    role = (p.get('role') or 'User').strip()
    status = (p.get('status') or 'Active').strip()
    emp_id = p.get('employee_id')
    if not username or not password:
        return jsonify({'ok': False, 'error': 'שם משתמש וסיסמה הם חובה'}), 400
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute('SELECT 1 FROM dbo.users WHERE username=?', (username,))
        if cur.fetchone():
            return jsonify({'ok': False, 'error': 'שם המשתמש כבר קיים'}), 400
        cur.execute('INSERT INTO dbo.users (username, password_hash, email, role, status, employee_id) VALUES (?,?,?,?,?,?)',
                    (username, generate_password_hash(password), email, role, status, int(emp_id) if emp_id else None))
        conn.commit()
    return jsonify({'ok': True})

@app.route('/system/api/users/<int:user_id>', methods=['PUT'])
@admin_required
@csrf.exempt
def api_users_update(user_id: int):
    p = request.get_json(force=True, silent=True) or {}
    email = (p.get('email') or '').strip() or None
    role = (p.get('role') or '').strip() or None
    status = (p.get('status') or '').strip() or None
    new_pwd = p.get('password') or None
    emp_id = p.get('employee_id')
    with get_db() as conn:
        cur = conn.cursor()
        if new_pwd:
            cur.execute('UPDATE dbo.users SET email=?, role=?, status=?, password_hash=?, employee_id=? WHERE id=?',
                        (email, role, status, generate_password_hash(new_pwd), int(emp_id) if emp_id else None, user_id))
        else:
            cur.execute('UPDATE dbo.users SET email=?, role=?, status=?, employee_id=? WHERE id=?',
                        (email, role, status, int(emp_id) if emp_id else None, user_id))
        conn.commit()
    return jsonify({'ok': True})

@app.route('/system/api/users/<int:user_id>', methods=['DELETE'])
@admin_required
@csrf.exempt
def api_users_delete(user_id: int):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute('SELECT username FROM dbo.users WHERE id=?', (user_id,))
        r = cur.fetchone()
        if not r:
            return jsonify({'ok': False, 'error': 'המשתמש לא נמצא'}), 404
        if r[0] == 'admin':
            return jsonify({'ok': False, 'error': 'לא ניתן למחוק את משתמש העל (admin)'}), 400
        cur.execute('DELETE FROM dbo.users WHERE id=?', (user_id,))
        conn.commit()
    return jsonify({'ok': True})

# ---- Employees API (DataTables) ----
@app.route('/system/api/employees')
@admin_required
def api_employees_list_full():
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute('SELECT id, first_name, last_name, emp_no, station_no FROM dbo.employees ORDER BY id DESC')
        rows = cur.fetchall()
    data = [{'id': r[0], 'first_name': r[1], 'last_name': r[2], 'emp_no': r[3], 'station_no': r[4]} for r in rows]
    return jsonify({'data': data})

@app.route('/system/api/employees', methods=['POST'])
@admin_required
@csrf.exempt
def api_employees_create():
    p = request.get_json(force=True, silent=True) or {}
    first = (p.get('first_name') or '').strip()
    last  = (p.get('last_name') or '').strip()
    empno = (p.get('emp_no') or '').strip()
    stno  = (p.get('station_no') or '').strip() or None
    if not first or not last or not empno:
        return jsonify({'ok': False, 'error': 'שם פרטי, שם משפחה ומס׳ עובד הם חובה'}), 400
    rnd = uuid.uuid4().hex[:12]
    with get_db() as conn:
        cur = conn.cursor()
        try:
            cur.execute('INSERT INTO dbo.employees (first_name, last_name, emp_no, random_id, station_no) VALUES (?,?,?,?,?)',
                        (first, last, empno, rnd, stno))
            conn.commit()
        except Exception:
            return jsonify({'ok': False, 'error': 'מס׳ עובד כבר קיים או שגיאה בנתונים'}), 400
    return jsonify({'ok': True})

@app.route('/system/api/employees/<int:eid>', methods=['PUT'])
@admin_required
@csrf.exempt
def api_employees_update(eid: int):
    p = request.get_json(force=True, silent=True) or {}
    first = (p.get('first_name') or '').strip()
    last  = (p.get('last_name') or '').strip()
    empno = (p.get('emp_no') or '').strip()
    stno  = (p.get('station_no') or '').strip() or None
    if not first or not last or not empno:
        return jsonify({'ok': False, 'error': 'שם פרטי, שם משפחה ומס׳ עובד הם חובה'}), 400
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute('UPDATE dbo.employees SET first_name=?, last_name=?, emp_no=?, station_no=? WHERE id=?',
                    (first, last, empno, stno, eid))
        conn.commit()
    return jsonify({'ok': True})

@app.route('/system/api/employees/<int:eid>', methods=['DELETE'])
@admin_required
@csrf.exempt
def api_employees_delete(eid: int):
    with get_db() as conn:
        cur = conn.cursor(); cur.execute('DELETE FROM dbo.employees WHERE id=?', (eid,)); conn.commit()
    return jsonify({'ok': True})

# ---- Stations API (DataTables) ----
@app.route('/system/api/stations')
@admin_required
def api_stations_list_full():
    with get_db() as conn:
        cur = conn.cursor(); cur.execute('SELECT id, station_no, display_name FROM dbo.stations ORDER BY id DESC')
        rows = cur.fetchall()
    data = [{'id': r[0], 'station_no': r[1], 'display_name': r[2]} for r in rows]
    return jsonify({'data': data})

@app.route('/system/api/stations', methods=['POST'])
@admin_required
@csrf.exempt
def api_stations_create():
    p = request.get_json(force=True, silent=True) or {}
    stno = (p.get('station_no') or '').strip()
    dname= (p.get('display_name') or '').strip() or None
    if not stno:
        return jsonify({'ok': False, 'error': 'מס׳ עמדה הוא שדה חובה'}), 400
    with get_db() as conn:
        cur = conn.cursor()
        try:
            cur.execute('INSERT INTO dbo.stations (station_no, display_name) VALUES (?,?)', (stno, dname))
            conn.commit()
        except Exception:
            return jsonify({'ok': False, 'error': 'מס׳ עמדה כבר קיים או שגיאה בנתונים'}), 400
    return jsonify({'ok': True})

@app.route('/system/api/stations/<int:sid>', methods=['PUT'])
@admin_required
@csrf.exempt
def api_stations_update(sid: int):
    p = request.get_json(force=True, silent=True) or {}
    stno = (p.get('station_no') or '').strip()
    dname= (p.get('display_name') or '').strip() or None
    if not stno:
        return jsonify({'ok': False, 'error': 'מס׳ עמדה הוא שדה חובה'}), 400
    with get_db() as conn:
        cur = conn.cursor(); cur.execute('UPDATE dbo.stations SET station_no=?, display_name=? WHERE id=?', (stno, dname, sid)); conn.commit()
    return jsonify({'ok': True})

@app.route('/system/api/stations/<int:sid>', methods=['DELETE'])
@admin_required
@csrf.exempt
def api_stations_delete(sid: int):
    with get_db() as conn:
        cur = conn.cursor(); cur.execute('DELETE FROM dbo.stations WHERE id=?', (sid,)); conn.commit()
    return jsonify({'ok': True})

# ------------------------
# Error handlers
# ------------------------
@app.errorhandler(403)
def forbidden(e):
    return render_template('error.html', title='403', message='אין הרשאה'), 403

@app.errorhandler(404)
def not_found(e):
    return render_template('error.html', title='404', message='הדף לא נמצא'), 404

@app.errorhandler(500)
def server_error(e):
    return render_template('error.html', title='500', message='שגיאת שרת'), 500

# ------------------------
# Main
# ------------------------
# ------------------------
# Vendors pages
# ------------------------
@app.route('/vendors')
@login_required
def vendors_list():
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT id, name, contact_person, phone, email, created_at
            FROM dbo.vendors
            ORDER BY id DESC
        """)
        rows = cur.fetchall()
        items = [{
            'id': r[0],
            'name': r[1],
            'contact_person': r[2] or '',
            'phone': r[3] or '',
            'email': r[4] or '',
            'created_at': r[5].strftime('%Y-%m-%d %H:%M') if r[5] else ''
        } for r in rows]
    return render_template('vendors_list.html', items=items)

@app.route('/vendors/add', methods=['GET', 'POST'])
@login_required
def add_vendor():
    if request.method == 'POST':
        name = (request.form.get('name') or '').strip()
        contact_person = (request.form.get('contact_person') or '').strip()
        phone = (request.form.get('phone') or '').strip()
        email = (request.form.get('email') or '').strip()
        if not name:
            flash('שם הספק/ספק הוא חובה', 'warning')
            return render_template('add_vendor.html')
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("INSERT INTO dbo.vendors (name, contact_person, phone, email) VALUES (?, ?, ?, ?)",
                        (name, contact_person, phone, email))
            conn.commit()
        flash('הספק נוסף בהצלחה', 'success')
        return redirect(url_for('vendors_list'))
    return render_template('add_vendor.html')

@app.route('/vendors/edit/<int:vid>', methods=['GET', 'POST'])
@login_required
def edit_vendor(vid: int):
    with get_db() as conn:
        cur = conn.cursor()
        if request.method == 'POST':
            name = (request.form.get('name') or '').strip()
            contact_person = (request.form.get('contact_person') or '').strip()
            phone = (request.form.get('phone') or '').strip()
            email = (request.form.get('email') or '').strip()
            if not name:
                flash('שם הספק/ספק הוא חובה', 'warning')
                return redirect(url_for('edit_vendor', vid=vid))
            cur.execute("""
                UPDATE dbo.vendors
                SET name=?, contact_person=?, phone=?, email=?
                WHERE id=?
            """, (name, contact_person, phone, email, vid))
            conn.commit()
            flash('הספק עודכן בהצלחה', 'success')
            return redirect(url_for('vendors_list'))
        # GET
        cur.execute("SELECT id, name, contact_person, phone, email FROM dbo.vendors WHERE id=?", (vid,))
        r = cur.fetchone()
        if not r:
            flash('הספק לא נמצא', 'warning')
            return redirect(url_for('vendors_list'))
        item = {'id': r[0], 'name': r[1], 'contact_person': r[2], 'phone': r[3], 'email': r[4]}
    return render_template('edit_vendor.html', item=item)



# ------------------------
# Vendors API (CRUD JSON, CSRF exempt)
# ------------------------
@app.route('/system/api/vendors', methods=['GET', 'POST'])
@csrf.exempt
@login_required
def api_vendors():
    from flask import request, jsonify
    with get_db() as conn:
        cur = conn.cursor()
        if request.method == 'GET':
            cur.execute("SELECT id, name, contact_person, phone, email, created_at FROM dbo.vendors ORDER BY id DESC")
            rows = cur.fetchall()
            data = [{
                'id': r[0],
                'name': r[1],
                'contact_person': r[2] or '',
                'phone': r[3] or '',
                'email': r[4] or '',
                'created_at': r[5].strftime('%Y-%m-%d %H:%M') if r[5] else ''
            } for r in rows]
            return jsonify({'data': data})

        # POST – הוספת ספק חדש
        if request.method == 'POST':
            body = request.get_json() or {}
            name = (body.get('name') or '').strip()
            contact_person = (body.get('contact_person') or '').strip()
            phone = (body.get('phone') or '').strip()
            email = (body.get('email') or '').strip()
            if not name:
                return jsonify({'error': 'שם הספק חובה'}), 400
            cur.execute(
                "INSERT INTO dbo.vendors (name, contact_person, phone, email) VALUES (?, ?, ?, ?)",
                (name, contact_person, phone, email)
            )
            conn.commit()
            return jsonify({'success': True})

@app.route('/system/api/vendors/<int:vid>', methods=['DELETE'])
@csrf.exempt
@login_required
def api_delete_vendor(vid):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM dbo.vendors WHERE id=?", (vid,))
        conn.commit()
    return jsonify({'success': True})
@app.route('/system/api/vendors/<int:vid>', methods=['PUT'])
@csrf.exempt
@login_required
def api_update_vendor(vid):
    from flask import request, jsonify
    body = request.get_json() or {}
    name = (body.get('name') or '').strip()
    contact_person = (body.get('contact_person') or '').strip()
    phone = (body.get('phone') or '').strip()
    email = (body.get('email') or '').strip()

    if not name:
        return jsonify({'error': 'שם הספק חובה'}), 400

    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
            UPDATE dbo.vendors
            SET name=?, contact_person=?, phone=?, email=?
            WHERE id=?
        """, (name, contact_person, phone, email, vid))
        if cur.rowcount == 0:
            return jsonify({'error': 'ספק לא נמצא'}), 404
        conn.commit()

    return jsonify({'success': True})

# ------------------------
# Equipment Types pages
# ------------------------
@app.route('/equipment/types')
@login_required
def equipment_types_list():
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT id, name, created_at
            FROM dbo.equipment_types
            ORDER BY id DESC
        """)
        rows = cur.fetchall()
        items = [{
            'id': r[0],
            'name': r[1],
            'created_at': r[2].strftime('%Y-%m-%d %H:%M') if r[2] else ''
        } for r in rows]
    return render_template('equipment_types_list.html', items=items)


@app.route('/equipment/types/add', methods=['GET', 'POST'])
@login_required
def add_equipment_type():
    if request.method == 'POST':
        name = (request.form.get('name') or '').strip()
        if not name:
            flash('שם סוג הציוד הוא חובה', 'warning')
            return render_template('add_equipment_type.html')
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("INSERT INTO dbo.equipment_types (name) VALUES (?)", (name,))
            conn.commit()
        flash('סוג הציוד נוסף בהצלחה', 'success')
        return redirect(url_for('equipment_types_list'))
    return render_template('add_equipment_type.html')


@app.route('/equipment/types/edit/<int:tid>', methods=['GET', 'POST'])
@login_required
def edit_equipment_type(tid: int):
    with get_db() as conn:
        cur = conn.cursor()
        if request.method == 'POST':
            name = (request.form.get('name') or '').strip()
            if not name:
                flash('שם סוג הציוד הוא חובה', 'warning')
                return redirect(url_for('edit_equipment_type', tid=tid))
            cur.execute("UPDATE dbo.equipment_types SET name=? WHERE id=?", (name, tid))
            conn.commit()
            flash('סוג הציוד עודכן בהצלחה', 'success')
            return redirect(url_for('equipment_types_list'))
        # GET
        cur.execute("SELECT id, name FROM dbo.equipment_types WHERE id=?", (tid,))
        r = cur.fetchone()
        if not r:
            flash('סוג הציוד לא נמצא', 'warning')
            return redirect(url_for('equipment_types_list'))
        item = {'id': r[0], 'name': r[1]}
    return render_template('edit_equipment_type.html', item=item)


# ------------------------
# Equipment Types API (CRUD JSON, CSRF exempt)
# ------------------------
@app.route('/system/api/equipment_types', methods=['GET', 'POST'])
@csrf.exempt
@login_required
def api_equipment_types():
    from flask import request, jsonify
    with get_db() as conn:
        cur = conn.cursor()
        if request.method == 'GET':
            cur.execute("SELECT id, name, created_at FROM dbo.equipment_types ORDER BY id DESC")
            rows = cur.fetchall()
            data = [{
                'id': r[0],
                'name': r[1],
                'created_at': r[2].strftime('%Y-%m-%d %H:%M') if r[2] else ''
            } for r in rows]
            return jsonify({'data': data})

        if request.method == 'POST':
            body = request.get_json() or {}
            name = (body.get('name') or '').strip()
            if not name:
                return jsonify({'error': 'שם סוג הציוד חובה'}), 400
            cur.execute("INSERT INTO dbo.equipment_types (name) VALUES (?)", (name,))
            conn.commit()
            return jsonify({'success': True})


@app.route('/system/api/equipment_types/<int:tid>', methods=['DELETE'])
@csrf.exempt
@login_required
def api_delete_equipment_type(tid):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM dbo.equipment_types WHERE id=?", (tid,))
        conn.commit()
    return jsonify({'success': True})

# ------------------------
# Toners pages
# ------------------------
@app.route('/toners')
@login_required
def toners_list():
    with get_db() as conn:
        cur = conn.cursor()
        # אם t.assigned_printer הוא INT:
        cur.execute("""
            SELECT
                t.id,
                t.serial_number,
                t.model,
                t.printer_type,         -- אם אין עמודה כזו אצלך, ראה הערות למטה
                t.vendor_id,               -- נמפה ל-vendor_name בטבלה
                t.status,
                t.assigned_printer,
                e.name AS printer_name  -- שם המדפסת שהטונר הוקצה אליה
            FROM dbo.toners t
            LEFT JOIN dbo.equipment e
              ON e.id = t.assigned_printer
            ORDER BY t.id DESC
        """)
        rows = cur.fetchall()
        toners = [{
            'id': r[0],
            'serial_number': r[1],
            'model': r[2],
            'printer_type': r[3],      # אם אין — שים None או '-'
            'vendor_name': r[4],       # הטמפלייט משתמש vendor_name
            'status': r[5],
            'assigned_printer': r[6],
            'printer_name': r[7],      # הטמפלייט מציג את זה
        } for r in rows]

    return render_template('toners_list.html', toners=toners)


@app.route('/toners/add', methods=['GET', 'POST'])
@login_required
def add_toner():
    with get_db() as conn:
        cur = conn.cursor()
        # רשימת ספקים
        cur.execute("SELECT id, name FROM dbo.vendors ORDER BY name")
        vendors = cur.fetchall()
        # רשימת מדפסות (לדוגמה equipment_type_id=1)
        cur.execute("SELECT id, model FROM dbo.printer_models WHERE is_active=1 ORDER BY model")
        printer_models = cur.fetchall()

    if request.method == 'POST':
        serial_number = (request.form.get('serial_number') or '').strip()
        model = (request.form.get('model') or '').strip()
        printer_type = (request.form.get('printer_type') or '').strip()  # הערך מה־Dropdown
        vendor_id = request.form.get('vendor_id') or None
        status = 'InStock'  # ברירת מחדל
        assigned_printer = None  # בהוספה ראשונית אין printer מוקצה

        # בדיקות חובה
        if not model:
            flash("דגם חובה", "danger")
            return redirect(url_for('add_toner'))

        if not printer_type:
            flash("חובה לבחור מדפסת מתאימה", "danger")
            return redirect(url_for('add_toner'))

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO dbo.toners (serial_number, model, printer_type, vendor_id, status, assigned_printer)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (serial_number, model, printer_type, vendor_id, status, assigned_printer))
            conn.commit()

        flash("טונר נוסף בהצלחה", "success")
        return redirect(url_for('toners_list'))

    return render_template('add_toners.html', vendors=vendors, printer_models=printer_models)


@app.route('/toners/edit/<int:tid>', methods=['GET', 'POST'])
@login_required
def edit_toner(tid):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("SELECT id, serial_number, model, printer_type, vendor_id, status, assigned_printer FROM dbo.toners WHERE id=?", (tid,))
        toner = cur.fetchone()
        if not toner:
            flash("טונר לא נמצא", "warning")
            return redirect(url_for('toners_list'))

        # רשימת ספקים
        cur.execute("SELECT id, name FROM dbo.vendors ORDER BY name")
        vendors = cur.fetchall()
        # רשימת מדפסות – פתרון 1
        cur.execute("""
            SELECT e.id, e.name
            FROM dbo.equipment e
            JOIN dbo.equipment_types t ON e.equipment_type_id = t.id
            WHERE t.name = ?
            ORDER BY e.name
        """, ('Printer',))
        printers = cur.fetchall()

    if request.method == 'POST':
        serial_number = (request.form.get('serial_number') or '').strip()
        model = (request.form.get('model') or '').strip()
        printer_type = (request.form.get('printer_type') or '').strip()
        vendor_id = request.form.get('vendor_id') or None
        status = request.form.get('status') or 'InStock'
        assigned_printer = request.form.get('assigned_printer') or None

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                UPDATE dbo.toners
                SET serial_number=?, model=?, printer_type=?, vendor_id=?, status=?, assigned_printer=?
                WHERE id=?
            """, (serial_number, model, printer_type, vendor_id, status, assigned_printer, tid))
            conn.commit()
        flash("הטונר עודכן בהצלחה", "success")
        return redirect(url_for('toners_list'))

    item = {
        'id': toner[0],
        'serial_number': toner[1],
        'model': toner[2],
        'printer_type': toner[3],
        'vendor_id': toner[4],
        'status': toner[5],
        'assigned_printer': toner[6]
    }
    return render_template('edit_toner.html', item=item, vendors=vendors, printers=printers)



@app.route('/toners/delete/<int:tid>', methods=['POST'])
@login_required
def delete_toner(tid):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM dbo.toners WHERE id=?", (tid,))
        conn.commit()
    flash("הטונר נמחק", "success")
    return redirect(url_for('toners_list'))


# ------------------------
# Toners API (CRUD JSON, CSRF exempt)
# ------------------------
@app.route('/system/api/toners', methods=['GET', 'POST'])
@csrf.exempt
@login_required
def api_toners():
    from flask import request, jsonify
    with get_db() as conn:
        cur = conn.cursor()

        if request.method == 'GET':
            critical = request.args.get('critical')
            printer_type = request.args.get('printer_type')

            # מצב קריטי - דגמים עם טונר אחד או פחות במלאי
            if critical == "1":
                cur.execute("""
                    SELECT t.printer_type, COUNT(*) as available_count
                    FROM dbo.toners t
                    WHERE t.status = 'InStock'
                    GROUP BY t.printer_type
                    HAVING COUNT(*) <= 1
                    ORDER BY t.printer_type
                """)
                rows = cur.fetchall()
                data = [{'printer_type': r[0], 'available_count': r[1]} for r in rows]
                return jsonify({'data': data})

            # מצב רגיל או סינון לפי סוג מדפסת
            sql = """
                SELECT t.id, t.serial_number, t.model, t.printer_type, v.name, e.name, t.status, t.created_at
                FROM dbo.toners t
                LEFT JOIN dbo.vendors v ON t.vendor_id = v.id
                LEFT JOIN dbo.equipment e ON t.assigned_printer = e.id
            """
            params = []
            if printer_type:
                sql += " WHERE t.printer_type = ?"
                params.append(printer_type)

            sql += " ORDER BY t.id DESC"
            cur.execute(sql, params)
            rows = cur.fetchall()
            data = [{
                'id': r[0],
                'serial_number': r[1],
                'model': r[2],
                'printer_type': r[3],
                'vendor': r[4] or '',
                'assigned_printer': r[5] or '',
                'status': r[6],
                'created_at': r[7].strftime('%Y-%m-%d %H:%M') if r[7] else ''
            } for r in rows]
            return jsonify({'data': data})

        if request.method == 'POST':
            body = request.get_json() or {}
            serial_number = (body.get('serial_number') or '').strip()
            model = (body.get('model') or '').strip()
            printer_type = (body.get('printer_type') or '').strip()
            vendor_id = body.get('vendor_id')
            status = (body.get('status') or 'InStock').strip()
            assigned_printer = body.get('assigned_printer')

            if not serial_number:
                return jsonify({'error': 'מספר סידורי חובה'}), 400

            cur.execute("""
                INSERT INTO dbo.toners (serial_number, model, printer_type, vendor_id, status, assigned_printer)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (serial_number, model, printer_type, vendor_id, status, assigned_printer))
            conn.commit()
            return jsonify({'success': True})

@app.route('/toners/critical')
@login_required
def toners_critical():
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                SELECT t.id, t.serial_number, t.model, t.printer_type, v.name AS vendor, e.name AS assigned_printer, t.status, t.created_at
                FROM dbo.toners t
                LEFT JOIN dbo.vendors v ON t.vendor_id = v.id
                LEFT JOIN dbo.equipment e ON t.assigned_printer = e.id
                WHERE t.status='InStock'
                AND t.printer_type IN (
                    SELECT printer_type
                    FROM dbo.toners
                    WHERE status='InStock'
                    GROUP BY printer_type
                    HAVING COUNT(*) <= 1
                )
                ORDER BY t.printer_type
            """)
            toners = [dict(zip([col[0] for col in cur.description], r)) for r in cur.fetchall()]
        return render_template('toners_critical.html', toners=toners)
    except Exception as e:
        return f"Error: {e}"


@app.route('/system/api/toners/<int:tid>', methods=['DELETE'])
@csrf.exempt
@login_required
def api_delete_toner(tid):
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("DELETE FROM dbo.toners WHERE id=?", (tid,))
        conn.commit()
    return jsonify({'success': True})

@app.route('/toners/assign/<int:tid>', methods=['GET', 'POST'])
@login_required
def assign_toner(tid):
    with get_db() as conn:
        cur = conn.cursor()
        # בדיקה אם הטונר קיים
        cur.execute("SELECT id, status FROM dbo.toners WHERE id=?", (tid,))
        toner = cur.fetchone()
        if not toner:
            flash("טונר לא נמצא", "warning")
            return redirect(url_for('toners_list'))

        if request.method == 'POST':
            printer_id = request.form.get('printer_id')
            cur.execute("UPDATE dbo.toners SET assigned_printer=?, status='Assigned' WHERE id=?", (printer_id, tid))
            conn.commit()
            flash("טונר הוקצה בהצלחה", "success")
            return redirect(url_for('toners_list'))

        # רשימת מדפסות לבחירה
        cur.execute("SELECT id, name FROM dbo.equipment WHERE equipment_type_id=2")
        printers = cur.fetchall()

    return render_template('assign_toner.html', toner=toner, printers=printers)
    
    
    
    
    

# --------------------------------------------
# Warehouse (matrix + filters) + Export + Print
# Fixed v2: ASCII-safe filename (no Unicode via isalnum), RFC 5987 filename*
# --------------------------------------------
from collections import OrderedDict
from flask import request, redirect, url_for, render_template, Response
from datetime import datetime
from urllib.parse import quote
import io, csv

# ---- Buckets per user rules ----
BUCKETS = [
    "אצל עובד",
    "מחסן חדש",
    "מחסן תקין",
    "מחסן תקול",
    "מחסן למחזור",
    "מיחזור",
    "נמכר",
]

BUCKET_COLORS = {
    "אצל עובד":     "bg-info text-white",
    "מחסן חדש":     "bg-primary text-white",
    "מחסן תקין":    "bg-success text-white",
    "מחסן תקול":    "bg-danger text-white",
    "מחסן למחזור":  "bg-warning text-dark",
    "מיחזור":       "bg-secondary text-white",
    "נמכר":         "bg-dark text-white",
}

WAREHOUSE_PLACE = "(e.placement = N'מחסן')"  # מחסן בלבד

BUCKETS_SQL = OrderedDict([
    ("אצל עובד",   "(e.placement = N'אצל עובד')"),
    ("מיחזור",     "(e.placement = N'מיחזור')"),
    ("נמכר",       "(e.placement = N'נמכר')"),

    ("מחסן חדש",     f"({WAREHOUSE_PLACE} AND e.status = N'חדש')"),
    ("מחסן תקין",    f"({WAREHOUSE_PLACE} AND e.status = N'תקין')"),
    ("מחסן תקול",    f"({WAREHOUSE_PLACE} AND e.status = N'תקול')"),
    ("מחסן למחזור",  f"({WAREHOUSE_PLACE} AND e.status = N'למחזור')"),
])

ALLOWED_COLUMNS = {
    "status": "e.status",
    "placement": "e.placement",
    "vendor": "e.vendor",
    "equipment_type_id": "e.equipment_type_id",
}

# ---- Helper: strict ASCII filename + RFC 5987 filename* ----
def _ascii_safe(name: str) -> str:
    # Keep only ASCII letters/digits and ._-; replace others with '_', collapse spaces to '_'
    out = []
    for ch in name:
        if ch.isascii() and (ch.isalnum() or ch in ('-', '_', '.', ' ')):
            out.append(ch)
        else:
            out.append('_')
    safe = ''.join(out).strip().replace(' ', '_')
    return safe or 'download'

def set_download_headers(resp, base_name_no_ascii: str):
    ts = datetime.now().strftime('%Y%m%d_%H%M')
    safe_ascii = _ascii_safe(base_name_no_ascii)
    filename_ascii = f"{safe_ascii}_{ts}.csv"
    filename_utf8 = f"{base_name_no_ascii}_{ts}.csv"
    filename_star = "UTF-8''" + quote(filename_utf8, safe='')
    resp.headers['Content-Disposition'] = f'attachment; filename="{filename_ascii}"; filename*={filename_star}'
    resp.headers['Content-Type'] = 'text/csv; charset=utf-8'
    return resp

# -------------- /warehouse (matrix) --------------
@app.route('/warehouse')
@login_required
def warehouse():
    with get_db() as conn:
        cur = conn.cursor()
        bucket_selects = []
        for bucket_name in BUCKETS:  # סדר קבוע לפי BUCKETS
            cond = BUCKETS_SQL[bucket_name]
            alias = f"bucket_{abs(hash(bucket_name))}"
            bucket_selects.append(f"SUM(CASE WHEN {cond} THEN 1 ELSE 0 END) AS [{alias}]")
        select_clause = ",\n                ".join(bucket_selects)
        query = f"""
            SELECT
                et.id AS type_id,
                et.name AS type_name,
                {select_clause}
            FROM equipment e
            LEFT JOIN equipment_types et ON e.equipment_type_id = et.id
            GROUP BY et.id, et.name
            ORDER BY et.name
        """
        cur.execute(query)
        rows = cur.fetchall()

        types_matrix = []
        for r in rows:
            type_id = r[0]
            type_name = r[1] if r[1] else "ללא סוג"
            counts = {bucket: (r[i] or 0) for i, bucket in enumerate(BUCKETS, start=2)}
            types_matrix.append({"type_id": type_id, "type_name": type_name, "counts": counts})

    return render_template('warehouse.html', buckets=BUCKETS, types_matrix=types_matrix, bucket_colors=BUCKET_COLORS)

# -------------- /warehouse/export (CSV of matrix) --------------
@app.route('/warehouse/export')
@login_required
def warehouse_export():
    with get_db() as conn:
        cur = conn.cursor()
        bucket_selects = []
        for bucket_name in BUCKETS:
            cond = BUCKETS_SQL[bucket_name]
            alias = f"bucket_{abs(hash(bucket_name))}"
            bucket_selects.append(f"SUM(CASE WHEN {cond} THEN 1 ELSE 0 END) AS [{alias}]")
        select_clause = ",\n                ".join(bucket_selects)
        query = f"""
            SELECT
                et.id AS type_id,
                et.name AS type_name,
                {select_clause}
            FROM equipment e
            LEFT JOIN equipment_types et ON e.equipment_type_id = et.id
            GROUP BY et.id, et.name
            ORDER BY et.name
        """
        cur.execute(query)
        rows = cur.fetchall()

    output = io.StringIO(newline='')
    writer = csv.writer(output)
    writer.writerow(['סוג ציוד'] + BUCKETS)
    for r in rows:
        type_name = r[1] if r[1] else 'ללא סוג'
        data = [(r[i] or 0) for i, _ in enumerate(BUCKETS, start=2)]
        writer.writerow([type_name] + data)

    data = output.getvalue()
    bom = '\ufeff'
    resp = Response((bom + data).encode('utf-8'), mimetype='text/csv')
    return set_download_headers(resp, 'warehouse_matrix')

# -------------- /warehouse/filtered --------------
@app.route('/warehouse/filtered')
@login_required
def warehouse_filtered():
    type_id   = request.args.get('type_id')  # יכול להיות 'NULL'
    bucket    = request.args.get('bucket')
    filter_by = request.args.get('filter_by')
    value     = request.args.get('value')

    where_clauses, params = [], []

    if type_id:
        if str(type_id).upper() == 'NULL':
            where_clauses.append("e.equipment_type_id IS NULL")
        else:
            where_clauses.append("e.equipment_type_id = ?")
            params.append(type_id)

    if bucket:
        cond = BUCKETS_SQL.get(bucket)
        if cond:
            where_clauses.append(f"({cond})")

    if filter_by and value and not bucket:
        col = ALLOWED_COLUMNS.get(filter_by)
        if col:
            where_clauses.append(f"{col} = ?")
            params.append(value)

    base_query = """
        SELECT e.id, e.name, e.vendor, e.status, e.placement, et.name
        FROM equipment e
        LEFT JOIN equipment_types et ON e.equipment_type_id = et.id
    """
    if where_clauses:
        base_query += " WHERE " + " AND ".join(where_clauses)

    with get_db() as conn:
        cur = conn.cursor()
        cur.execute(base_query, params)
        items = cur.fetchall()

        title_parts = []
        if type_id:
            if str(type_id).upper() == 'NULL':
                type_name = "ללא סוג"
            else:
                cur.execute("SELECT name FROM equipment_types WHERE id = ?", (type_id,))
                r = cur.fetchone()
                type_name = r[0] if r and r[0] else f"Type #{type_id}"
            title_parts.append(f"סוג ציוד: {type_name}")
        if bucket:
            title_parts.append(f"פילטר: {bucket}")
        elif filter_by and value:
            title_parts.append(f"{filter_by} = {value}")
        title = " | ".join(title_parts) if title_parts else "פריטים מסוננים"

        export_url = url_for('warehouse_filtered_export', **request.args.to_dict(flat=True))

    return render_template('warehouse_filtered.html', items=items, title=title, export_url=export_url)

# -------------- /warehouse/filtered/export (CSV of filtered items) --------------
@app.route('/warehouse/filtered/export')
@login_required
def warehouse_filtered_export():
    type_id   = request.args.get('type_id')
    bucket    = request.args.get('bucket')
    filter_by = request.args.get('filter_by')
    value     = request.args.get('value')

    where_clauses, params = [], []

    if type_id:
        if str(type_id).upper() == 'NULL':
            where_clauses.append("e.equipment_type_id IS NULL")
        else:
            where_clauses.append("e.equipment_type_id = ?")
            params.append(type_id)

    if bucket:
        cond = BUCKETS_SQL.get(bucket)
        if cond:
            where_clauses.append(f"({cond})")

    if filter_by and value and not bucket:
        col = ALLOWED_COLUMNS.get(filter_by)
        if col:
            where_clauses.append(f"{col} = ?")
            params.append(value)

    base_query = """
        SELECT e.id, e.name, e.vendor, e.status, e.placement, et.name
        FROM equipment e
        LEFT JOIN equipment_types et ON e.equipment_type_id = et.id
    """
    if where_clauses:
        base_query += " WHERE " + " AND ".join(where_clauses)

    with get_db() as conn:
        cur = conn.cursor()
        cur.execute(base_query, params)
        rows = cur.fetchall()

    output = io.StringIO(newline='')
    writer = csv.writer(output)
    writer.writerow(['ID','שם','ספק','סטטוס','מיקום','סוג ציוד'])
    for r in rows:
        writer.writerow([r[0], r[1], r[2], r[3], r[4], r[5]])

    data = output.getvalue()
    bom = '\ufeff'
    resp = Response((bom + data).encode('utf-8'), mimetype='text/csv')

    base = 'warehouse_filtered'
    if bucket:
        base += '_' + bucket

    return set_download_headers(resp, base)

# -------------- legacy path redirect --------------
@app.route('/warehouse/filtered/<filter_by>/<value>')
@login_required
def warehouse_filtered_legacy(filter_by, value):
    return redirect(url_for('warehouse_filtered', filter_by=filter_by, value=value))



# ================== END: WAREHOUSE PATCH ==================

# ============================
# SYSTEM API: Printer Models (CRUD)
# ============================

@app.route('/system/api/printer_models', methods=['GET'])
@login_required
def system_api_printer_models_list():
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute("""
            SELECT id, brand, model, is_active
            FROM dbo.printer_models
            ORDER BY brand, model
        """)
        rows = cur.fetchall()
    data = [{"id": r[0], "brand": r[1], "model": r[2], "is_active": bool(r[3])} for r in rows]
    return jsonify({"data": data}), 200


@app.route('/system/api/printer_models', methods=['POST'])
@login_required
def system_api_printer_models_create():
    payload = request.get_json(force=True) or {}
    brand = (payload.get('brand') or '').strip()
    model = (payload.get('model') or '').strip()
    is_active = 1 if payload.get('is_active', True) else 0

    if not brand or not model:
        return jsonify({"error": "חובה למלא יצרן ודגם"}), 400

    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO dbo.printer_models (brand, model, is_active)
                OUTPUT INSERTED.id
                VALUES (?, ?, ?)
            """, (brand, model, is_active))
            new_id = cur.fetchone()[0]
            conn.commit()
        return jsonify({"id": int(new_id)}), 201
    except Exception as ex:
        msg = str(ex)
        if 'UQ_printer_models_brand_model' in msg or 'UNIQUE' in msg:
            return jsonify({"error": "הדגם כבר קיים עבור היצרן הזה"}), 409
        return jsonify({"error": "שגיאה ביצירה", "details": msg}), 500


@app.route('/system/api/printer_models/<int:item_id>', methods=['PUT'])
@login_required
def system_api_printer_models_update(item_id):
    payload = request.get_json(force=True) or {}
    brand = (payload.get('brand') or '').strip()
    model = (payload.get('model') or '').strip()
    is_active = 1 if payload.get('is_active', True) else 0

    if not brand or not model:
        return jsonify({"error": "חובה למלא יצרן ודגם"}), 400

    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                UPDATE dbo.printer_models
                SET brand = ?, model = ?, is_active = ?
                WHERE id = ?
            """, (brand, model, is_active, item_id))
            conn.commit()
        return jsonify({"ok": True}), 200
    except Exception as ex:
        msg = str(ex)
        if 'UQ_printer_models_brand_model' in msg or 'UNIQUE' in msg:
            return jsonify({"error": "הדגם כבר קיים עבור היצרן הזה"}), 409
        return jsonify({"error": "שגיאה בעדכון", "details": msg}), 500


@app.route('/system/api/printer_models/<int:item_id>', methods=['DELETE'])
@login_required
def system_api_printer_models_delete(item_id):
    try:
        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("DELETE FROM dbo.printer_models WHERE id = ?", (item_id,))
            conn.commit()
        return jsonify({"ok": True}), 200
    except Exception as ex:
        return jsonify({"error": "שגיאה במחיקה", "details": str(ex)}), 500


# ❗ פטור מ־CSRF לשלושת הפעולות שמשנות נתונים
csrf.exempt(system_api_printer_models_create)
csrf.exempt(system_api_printer_models_update)
csrf.exempt(system_api_printer_models_delete)



# רשימת תוכנות
@app.route('/software/list')
@login_required
def software_list():
    expiring = request.args.get('expiring')

    with get_db() as conn:
        cur = conn.cursor()
        query = "SELECT * FROM dbo.software"
        params = []

        if expiring == "1":
            query += " WHERE DATEDIFF(MONTH, GETDATE(), renewal_next) < 3"

        cur.execute(query, params)
        rows = cur.fetchall()

    return render_template("software_list.html", software=rows)


# הוספת תוכנה
@app.route('/software/add', methods=['GET','POST'])
def add_software():
    if request.method == 'POST':
        name = request.form['name']
        description = request.form.get('description')
        category = request.form.get('category')
        purchase_date = request.form.get('purchase_date')
        subsidiaries = request.form.getlist('subsidiaries')
        renewal_last = request.form.get('renewal_last')
        renewal_next = request.form.get('renewal_next')
        last_cost = request.form.get('last_cost')

        subsidiaries_str = ','.join(subsidiaries)
        updated_at = datetime.now()

        with get_db() as conn:
            cur = conn.cursor()
            cur.execute("""
                INSERT INTO software
                (name, description, category, purchase_date, subsidiaries, renewal_last, renewal_next, last_cost, created_at, updated_at)
                VALUES (?,?,?,?,?,?,?,?,?,?)
            """, name, description, category, purchase_date, subsidiaries_str, renewal_last, renewal_next, last_cost, datetime.now(), updated_at)
            conn.commit()
        flash("תוכנה נשמרה בהצלחה", "success")
        return redirect(url_for('software_list'))

    return render_template('add_software.html')

# עריכת תוכנה
@app.route('/software/edit/<int:id>', methods=['GET','POST'])
def edit_software(id):
    with get_db() as conn:
        cur = conn.cursor()
        
        if request.method == 'POST':
            # ערכים חדשים מהטופס
            name = request.form['name']
            description = request.form.get('description')
            category = request.form.get('category')
            purchase_date = request.form.get('purchase_date')
            subsidiaries = ','.join(request.form.getlist('subsidiaries'))
            renewal_last = request.form.get('renewal_last')
            renewal_next = request.form.get('renewal_next')
            last_cost = request.form.get('last_cost')
            updated_at = datetime.now()

            # קח את השורה הקיימת לפני העדכון
            cur.execute("SELECT * FROM software WHERE id=?", (id,))
            old_row = cur.fetchone()
            old = dict(zip([col[0] for col in cur.description], old_row))

            # ערכים חדשים
            new = {
                'name': name,
                'description': description,
                'category': category,
                'purchase_date': purchase_date,
                'subsidiaries': subsidiaries,
                'renewal_last': renewal_last,
                'renewal_next': renewal_next,
                'last_cost': last_cost
            }

            # שמירה בהיסטוריה עבור כל שדה ששונה
            for key in new:
                if str(old.get(key) or '') != str(new[key] or ''):
                    cur.execute("""
                        INSERT INTO software_history
                        (software_id, changed_at, change_description, field_name, old_value, new_value, changed_by)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (id, datetime.now(), f"שינוי בשדה {key}", key, old.get(key), new[key], "admin"))

            # עדכון תוכנה
            cur.execute("""
                UPDATE software SET
                    name=?, description=?, category=?, purchase_date=?, subsidiaries=?,
                    renewal_last=?, renewal_next=?, last_cost=?, updated_at=?
                WHERE id=?
            """, (name, description, category, purchase_date, subsidiaries,
                  renewal_last, renewal_next, last_cost, updated_at, id))

            conn.commit()
            flash("תוכנה עודכנה בהצלחה", "success")
            return redirect(url_for('software_list'))

        # שליפת תוכנה
        cur.execute("SELECT * FROM software WHERE id=?", (id,))
        software_row = cur.fetchone()
        software = dict(zip([col[0] for col in cur.description], software_row))

        # היסטוריה
        cur.execute("SELECT * FROM software_history WHERE software_id=? ORDER BY changed_at DESC", (id,))
        history = [dict(zip([col[0] for col in cur.description], row)) for row in cur.fetchall()]

    return render_template('edit_software.html', software=software, history=history)


# API להחזרת תוכנות
@app.route('/api/software')
def api_software():
    q = request.args.get('q', '')
    expiring = request.args.get('expiring', '')  # <–– נוסיף את הפרמטר
    
    with get_db() as conn:
        cur = conn.cursor()
        sql = """SELECT id, name, description, category, purchase_date,
                        subsidiaries, renewal_last, renewal_next, last_cost
                 FROM software"""
        params = []
        where_clauses = []

        if q:
            where_clauses.append("(name LIKE ? OR category LIKE ? OR subsidiaries LIKE ?)")
            qlike = f"%{q}%"
            params.extend([qlike, qlike, qlike])

        if expiring == "3m":
            # תוכנות שתאריך החידוש הבא עד 3 חודשים קדימה
            where_clauses.append("renewal_next <= DATE('now', '+3 months')")

        if where_clauses:
            sql += " WHERE " + " AND ".join(where_clauses)

        sql += " ORDER BY id DESC"
        cur.execute(sql, params)
        rows = [dict(zip([col[0] for col in cur.description], r)) for r in cur.fetchall()]
    return jsonify(rows)



# API להיסטוריה
@app.route('/api/software/history/<int:software_id>')
def api_software_history(software_id):
    try:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("""
            SELECT changed_at, field_name, old_value, new_value, changed_by, change_description
            FROM software_history
            WHERE software_id = ?
            ORDER BY changed_at DESC
        """, (software_id,))  # חייב להיות tuple עם פסיק
        rows = cur.fetchall()

        history = []
        for row in rows:
            history.append({
                'changed_at': row[0].strftime('%Y-%m-%d %H:%M:%S') if row[0] else '',
                'field_name': row[1] or '',
                'old_value': row[2] or '',
                'new_value': row[3] or '',
                'user_name': row[4] or '',
                'change_description': row[5] or '',
            })
        return jsonify(history)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        conn.close()

    return jsonify(history)

# דף מערכת (ייבוא CSV)
@app.route('/system/import')
@csrf.exempt
@login_required
def system_import():
    return render_template('system.html')

# Preview CSV (תצוגה מקדימה)
@app.route('/system/api/import_csv_preview', methods=['POST'])
@csrf.exempt
@login_required
def import_csv_preview():
    if 'file' not in request.files:
        return jsonify({"error": "לא נבחר קובץ"}), 400

    target = request.form.get("target")
    if not target:
        return jsonify({"error": "חסר target"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "שם הקובץ ריק"}), 400

    try:
        stream = io.StringIO(file.stream.read().decode("utf-8"))
        reader = csv.DictReader(stream)
        data = [row for row in reader]
        columns = reader.fieldnames

        # בדיקת עמודות חובה
        required = []
        if target == "employees":
            required = ["first_name", "last_name", "emp_no"]
        elif target == "stations":
            required = ["station_no", "display_name"]
        elif target == "equipment":
            required = ["name", "vendor", "barcode"]

        for col in required:
            if col not in columns:
                return jsonify({"error": f"עמודת חובה חסרה: {col}"}), 400

        return jsonify({"columns": columns, "data": data})

    except Exception as e:
        return jsonify({"error": str(e)}), 400



# Commit CSV - שמירה לפי target (employees, stations, equipment)
@app.route('/system/api/import_csv_commit', methods=['POST'])
@csrf.exempt
@login_required
def import_csv_commit():
    import csv, io
    target = request.form.get("target")
    file = request.files.get("file")

    if not target:
        return jsonify({"error": "חסר target"}), 400
    if not file:
        return jsonify({"error": "חסר קובץ"}), 400

    try:
        stream = io.StringIO(file.stream.read().decode("utf-8"))
        reader = csv.DictReader(stream)
        imported = 0

        with get_db() as conn:
            cur = conn.cursor()

            if target == "employees":
                for row in reader:
                    cur.execute("""
                        INSERT INTO employees (first_name, last_name, email)
                        VALUES (?, ?, ?)
                    """, (row["first_name"], row["last_name"], row["email"]))
                    imported += 1

            elif target == "stations":
                for row in reader:
                    cur.execute("""
                        INSERT INTO stations (station_code, location)
                        VALUES (?, ?)
                    """, (row["station_code"], row["location"]))
                    imported += 1

            elif target == "equipment":
                for row in reader:
                    cur.execute("""
                        INSERT INTO equipment (name, vendor, barcode)
                        VALUES (?, ?, ?)
                    """, (row["name"], row["vendor"], row["barcode"]))
                    imported += 1

            else:
                return jsonify({"error": "target לא חוקי"}), 400

            conn.commit()

        return jsonify({"success": True, "imported": imported})

    except Exception as e:
        return jsonify({"error": f"שגיאה בעת הייבוא: {str(e)}"}), 500


    













# ===================== Reports Add-on (Advanced) =====================
from flask import jsonify, send_file
import pandas as pd
import io


def _where_and_params_from_args(args):
    where = []
    params = []
    if args.get('equipment_type_id'):
        where.append("e.equipment_type_id = ?")
        params.append(args.get('equipment_type_id'))
    if args.get('status'):
        where.append("LTRIM(RTRIM(e.status)) = ?")
        params.append(args.get('status'))
    if args.get('placement'):
        where.append("LTRIM(RTRIM(e.placement)) = ?")
        params.append(args.get('placement'))
    if args.get('vendor'):
        where.append("e.vendor LIKE ?")
        params.append(f"%{args.get('vendor')}%")
    if args.get('assigned_to'):
        where.append("e.assigned_to LIKE ?")
        params.append(f"%{args.get('assigned_to')}%")
    if args.get('station'):
        where.append("LTRIM(RTRIM(e.station)) LIKE ?")
        params.append(f"%{args.get('station')}%")
    return where, params


@app.route('/reports/api/wide')
@login_required
def reports_api_wide():
    where, params = _where_and_params_from_args(request.args)
    where_sql = (" WHERE " + " AND ".join(where)) if where else ""
    sql = f"""
    SELECT e.id, e.name, e.vendor, et.name AS equipment_type,
           e.barcode, e.status, e.placement, e.assigned_to, e.station,
           e.warranty_expiry, MAX(a.changed_at) AS last_updated
    FROM dbo.equipment e
    LEFT JOIN dbo.equipment_types et ON e.equipment_type_id = et.id
    LEFT JOIN dbo.equipment_audit a ON a.equipment_id = e.id
    {where_sql}
    GROUP BY e.id, e.name, e.vendor, et.name, e.barcode, e.status,
             e.placement, e.assigned_to, e.station, e.warranty_expiry
    ORDER BY e.id DESC
    """
    with get_db() as conn:
        cur = conn.cursor(); cur.execute(sql, params)
        rows = cur.fetchall()
    def sdate(d):
        try: return d.strftime('%Y-%m-%d')
        except: return ''
    data = [[r[0], r[1], r[2], r[3] or '', r[4] or '', r[5] or '', r[6] or '',
             r[7] or '', r[8] or '', sdate(r[9]),
             r[10].strftime('%Y-%m-%d %H:%M') if r[10] else ''] for r in rows]
    return jsonify(data)


ADV_ALLOWED = {
    "vendor": "e.vendor",
    "status": "e.status",
    "placement": "e.placement",
    "equipment_type_id": "e.equipment_type_id",
    "assigned_to": "e.assigned_to",
    "station": "e.station",
    "warranty_expiry": "e.warranty_expiry",
}

@app.route('/reports/api/advanced')
@login_required
def reports_api_advanced():
    field = request.args.get('field', '')
    op    = request.args.get('op', 'eq')
    val   = request.args.get('value', '')
    val2  = request.args.get('value2', '')
    col   = ADV_ALLOWED.get(field)
    if not col:
        return jsonify([])

    where_sql = ""
    params = []
    if op == 'eq':
        where_sql = f"WHERE {col} = ?"; params = [val]
    elif op == 'neq':
        where_sql = f"WHERE {col} <> ?"; params = [val]
    elif op == 'contains':
        where_sql = f"WHERE {col} LIKE ?"; params = [f"%{val}%"]
    elif op == 'in':
        items = [x.strip() for x in val.split(',') if x.strip()]
        if not items: return jsonify([])
        where_sql = "WHERE " + col + " IN (" + ",".join(["?"]*len(items)) + ")"
        params = items
    elif op == 'between' and field == 'warranty_expiry':
        where_sql = f"WHERE {col} BETWEEN ? AND ?"; params = [val, val2]

    sql = f"""
    SELECT e.id, e.name, e.vendor, et.name AS equipment_type,
           e.barcode, e.status, e.placement, e.assigned_to, e.station, e.warranty_expiry
    FROM dbo.equipment e
    LEFT JOIN dbo.equipment_types et ON e.equipment_type_id = et.id
    {where_sql}
    ORDER BY e.id DESC
    """
    with get_db() as conn:
        cur = conn.cursor(); cur.execute(sql, params)
        rows = cur.fetchall()
    def sdate(d):
        try: return d.strftime('%Y-%m-%d')
        except: return ''
    data = [[r[0], r[1], r[2], r[3] or '', r[4] or '', r[5] or '', r[6] or '',
             r[7] or '', r[8] or '', sdate(r[9])] for r in rows]
    return jsonify(data)


@app.route('/reports/api/history')
@login_required
def reports_api_history():
    employee = (request.args.get('employee') or '').strip()
    station  = (request.args.get('station') or '').strip()
    dt_from  = request.args.get('from')
    dt_to    = request.args.get('to')

    where = ["1=1"]
    params = []
    if employee:
        where.append("(JSON_VALUE(ea.before_json,'$.assigned_to') LIKE ? OR JSON_VALUE(ea.after_json ,'$.assigned_to') LIKE ?)")
        like_emp = f"%{employee}%"
        params += [like_emp, like_emp]
    if station:
        where.append("(JSON_VALUE(ea.before_json,'$.station') LIKE ? OR JSON_VALUE(ea.after_json ,'$.station') LIKE ?)")
        like_st = f"%{station}%"
        params += [like_st, like_st]
    if dt_from:
        where.append("ea.changed_at >= ?"); params.append(dt_from + " 00:00:00")
    if dt_to:
        where.append("ea.changed_at <= ?"); params.append(dt_to + " 23:59:59")

    where_sql = " WHERE " + " AND ".join(where)
    sql = f"""
SELECT ea.equipment_id, e.name AS equipment_name, ea.action, ea.changed_by, ea.changed_at,
       JSON_VALUE(ea.before_json,'$.assigned_to') AS before_assigned_to,
       JSON_VALUE(ea.after_json ,'$.assigned_to') AS after_assigned_to,
       JSON_VALUE(ea.before_json,'$.station')     AS before_station,
       JSON_VALUE(ea.after_json ,'$.station')     AS after_station
FROM dbo.equipment_audit ea
LEFT JOIN dbo.equipment e ON e.id = ea.equipment_id
{where_sql}
ORDER BY ea.changed_at DESC
"""

    with get_db() as conn:
        cur = conn.cursor(); cur.execute(sql, params); rows = cur.fetchall()
    data = [[r[0], r[1], r[2], r[3] or '', r[4].strftime('%Y-%m-%d %H:%M') if r[4] else '',
         r[5] or '', r[6] or '', r[7] or '', r[8] or ''] for r in rows]

    return jsonify(data)


# ---------- Export helpers ----------

def _df_response_csv(df: pd.DataFrame, fname: str):
    out = io.StringIO()
    df.to_csv(out, index=False)
    mem = io.BytesIO(out.getvalue().encode('utf-8-sig'))
    return send_file(mem, mimetype='text/csv', as_attachment=True, download_name=f"{fname}.csv")


def _df_response_xlsx(df: pd.DataFrame, fname: str):
    mem = io.BytesIO()
    with pd.ExcelWriter(mem, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Report')
    mem.seek(0)
    return send_file(mem, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f"{fname}.xlsx")


@app.route('/reports/export/wide/<fmt>')
@login_required
def export_wide(fmt: str):
    where, params = _where_and_params_from_args(request.args)
    where_sql = (" WHERE " + " AND ".join(where)) if where else ""
    sql = f"""
    SELECT e.id, e.name, e.vendor, et.name AS equipment_type, e.barcode,
           e.status, e.placement, e.assigned_to, e.station, e.warranty_expiry
    FROM dbo.equipment e
    LEFT JOIN dbo.equipment_types et ON e.equipment_type_id = et.id
    {where_sql}
    ORDER BY e.id DESC
    """
    with get_db() as conn:
        cur = conn.cursor(); cur.execute(sql, params)
        cols = [c[0] for c in cur.description]
        rows = cur.fetchall()
    df = pd.DataFrame.from_records(rows, columns=cols)
    fname = f"wide_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    return _df_response_csv(df, fname) if fmt == 'csv' else _df_response_xlsx(df, fname)


@app.route('/reports/export/advanced/<fmt>')
@login_required
def export_advanced(fmt: str):
    with app.test_request_context(query_string=request.query_string):
        data = reports_api_advanced().get_json()
    cols = ["id","name","vendor","equipment_type","barcode","status","placement","assigned_to","station","warranty_expiry"]
    df = pd.DataFrame(data, columns=cols)
    fname = f"advanced_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    return _df_response_csv(df, fname) if fmt == 'csv' else _df_response_xlsx(df, fname)


@app.route('/reports/export/history/<fmt>')
@login_required
def export_history(fmt: str):
    with app.test_request_context(query_string=request.query_string):
        data = reports_api_history().get_json()
    cols = ["equipment_id","action","changed_by","changed_at",
            "before_assigned_to","after_assigned_to","before_station","after_station"]
    df = pd.DataFrame(data, columns=cols)
    fname = f"history_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    return _df_response_csv(df, fname) if fmt == 'csv' else _df_response_xlsx(df, fname)


@app.route('/reports/export/all/<fmt>')
@login_required
def export_all(fmt: str):
    sql = """
    SELECT 
      e.id, e.name, e.vendor, e.warranty_expiry, e.status, e.barcode, e.history,
      e.assigned_to, e.station, e.placement, e.sold_to,
      e.equipment_type_id, et.name AS equipment_type
    FROM dbo.equipment e
    LEFT JOIN dbo.equipment_types et ON e.equipment_type_id = et.id
    ORDER BY e.id ASC
    """
    with get_db() as conn:
        cur = conn.cursor()
        cur.execute(sql)
        cols = [c[0] for c in cur.description]
        rows = cur.fetchall()
    df = pd.DataFrame.from_records(rows, columns=cols)

    if fmt.lower() == 'csv':
        out = io.StringIO(); df.to_csv(out, index=False)
        mem = io.BytesIO(out.getvalue().encode('utf-8-sig'))
        return send_file(mem, mimetype='text/csv', as_attachment=True,
                         download_name=f"equipment_all_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")

    mem = io.BytesIO()
    with pd.ExcelWriter(mem, engine='openpyxl') as w:
        df.to_excel(w, index=False, sheet_name='All Equipment')
    mem.seek(0)
    return send_file(mem,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f"equipment_all_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
# =================== End of Reports Add-on ===================


# ---------- Toners Reports ----------

# ---------- Toners Reports Wide ----------
@app.route('/reports/api/toners/wide')
@login_required
def toners_wide():
    assigned = request.args.get('assigned_printer')
    status   = request.args.get('status')

    where = []
    params = []
    if assigned:
        where.append("t.assigned_printer LIKE ?")
        params.append(f"%{assigned}%")
    if status:
        where.append("t.status = ?")
        params.append(status)
    where_sql = "WHERE " + " AND ".join(where) if where else ""

    sql = f"""
    SELECT t.id, t.serial_number, t.model, t.printer_type, t.vendor_id,
           t.assigned_printer, t.status, t.created_at,
           MAX(a.changed_at) AS last_updated
    FROM dbo.toners t
    LEFT JOIN dbo.equipment_audit a ON a.equipment_id = t.id
    {where_sql}
    GROUP BY t.id, t.serial_number, t.model, t.printer_type, t.vendor_id,
             t.assigned_printer, t.status, t.created_at
    ORDER BY t.id DESC
    """
    with get_db() as conn:
        cur = conn.cursor(); cur.execute(sql, params); rows = cur.fetchall()
    def sdate(d): return d.strftime('%Y-%m-%d') if d else ''
    data = [[r[0], r[1], r[2] or '', r[3] or '', r[4] or '', r[5] or '', r[6] or '',
             sdate(r[7]), r[8].strftime('%Y-%m-%d %H:%M') if r[8] else ''] for r in rows]
    return jsonify(data)


# ---------- Toners Reports History ----------
@app.route('/reports/api/toners/history')
@login_required
def toners_history():
    assigned = request.args.get('assigned_printer')
    status   = request.args.get('status')
    dt_from  = request.args.get('from')
    dt_to    = request.args.get('to')

    where = ["t.id IS NOT NULL"]
    params = []

    if assigned:
        where.append("(JSON_VALUE(ea.before_json,'$.assigned_printer') LIKE ? OR JSON_VALUE(ea.after_json,'$.assigned_printer') LIKE ?)")
        like_assigned = f"%{assigned}%"
        params += [like_assigned, like_assigned]
    if status:
        where.append("(JSON_VALUE(ea.before_json,'$.status') = ? OR JSON_VALUE(ea.after_json,'$.status') = ?)")
        params += [status, status]
    if dt_from:
        where.append("ea.changed_at >= ?"); params.append(dt_from + " 00:00:00")
    if dt_to:
        where.append("ea.changed_at <= ?"); params.append(dt_to + " 23:59:59")

    where_sql = "WHERE " + " AND ".join(where)

    sql = f"""
    SELECT ea.equipment_id, t.serial_number, t.model, ea.action, ea.changed_by, ea.changed_at,
           JSON_VALUE(ea.before_json,'$.assigned_printer') AS before_assigned,
           JSON_VALUE(ea.after_json,'$.assigned_printer') AS after_assigned,
           JSON_VALUE(ea.before_json,'$.status') AS before_status,
           JSON_VALUE(ea.after_json,'$.status') AS after_status
    FROM dbo.equipment_audit ea
    LEFT JOIN dbo.toners t ON t.id = ea.equipment_id
    {where_sql}
    ORDER BY ea.changed_at DESC
    """
    with get_db() as conn:
        cur = conn.cursor(); cur.execute(sql, params); rows = cur.fetchall()

    data = [[r[0], r[1] or '', r[2] or '', r[3], r[4] or '',
             r[5].strftime('%Y-%m-%d %H:%M') if r[5] else '',
             r[6] or '', r[7] or '', r[8] or '', r[9] or ''] for r in rows]
    return jsonify(data)


# ---------- Export Helpers ----------
def _df_response_csv(df: pd.DataFrame, fname: str):
    out = io.StringIO(); df.to_csv(out, index=False)
    mem = io.BytesIO(out.getvalue().encode('utf-8-sig'))
    return send_file(mem, mimetype='text/csv', as_attachment=True, download_name=f"{fname}.csv")

def _df_response_xlsx(df: pd.DataFrame, fname: str):
    mem = io.BytesIO()
    with pd.ExcelWriter(mem, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='Report')
    mem.seek(0)
    return send_file(mem, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f"{fname}.xlsx")


@app.route('/reports/export/toners/wide/<fmt>')
@login_required
def export_toners_wide(fmt: str):
    data = toners_wide().get_json()
    cols = ["ID","Serial Number","Model","Printer Type","Vendor ID",
            "Assigned Printer","Status","Created At","Last Updated"]
    df = pd.DataFrame(data, columns=cols)
    fname = f"toners_wide_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    return _df_response_csv(df, fname) if fmt=='csv' else _df_response_xlsx(df, fname)


@app.route('/reports/export/toners/history/<fmt>')
@login_required
def export_toners_history(fmt: str):
    data = toners_history().get_json()
    cols = ["ID","Serial Number","Model","Action","Changed By","Changed At",
            "Assigned (Before)","Assigned (After)","Status (Before)","Status (After)"]
    df = pd.DataFrame(data, columns=cols)
    fname = f"toners_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    return _df_response_csv(df, fname) if fmt=='csv' else _df_response_xlsx(df, fname)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)


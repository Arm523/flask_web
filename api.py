from flask import Blueprint, jsonify, request
from utils import get_db_connection, get_now
from datetime import date, datetime


api = Blueprint('api', __name__)

today = get_now(mocked=True).date()

# =============================
# 1) จำนวนสถานะห้อง (Pie chart)
# =============================
@api.route("/api/status_pie")
def status_pie():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute("""
        SELECT
            s.status_id,
            s.name_status,
            COUNT(u.unit_id) AS count
        FROM s_unit s
        LEFT JOIN unit u ON u.status_id = s.status_id
        GROUP BY s.status_id, s.name_status
        ORDER BY s.status_id
    """)
    rows = cursor.fetchall()

    summary = {
        "empty": 0,        # 1
        "occupied": 0,     # 2
        "contract": 0,     # 3
        "maintenance": 0,  # 4
        "waiting": 0,      # 5
        "checkout": 0      # 6
    }

    for r in rows:
        if r["status_id"] == 1:
            summary["empty"] = r["count"]
        elif r["status_id"] == 2:
            summary["occupied"] = r["count"]
        elif r["status_id"] == 3:
            summary["contract"] = r["count"]
        elif r["status_id"] == 4:
            summary["maintenance"] = r["count"]
        elif r["status_id"] == 5:
            summary["waiting"] = r["count"]
        elif r["status_id"] == 6:
            summary["checkout"] = r["count"]

    return jsonify({
        "labels": [r["name_status"] for r in rows],
        "counts": [r["count"] for r in rows],
        "summary": summary
    })

@api.route('/api/finance_summary')
def finance_bar_data():
    selected_year = request.args.get('year', str(today.year))
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    query = """
        SELECT MONTH(transaction_date) as m_num, type, SUM(amount) as total 
        FROM transactions 
        WHERE YEAR(transaction_date) = %s
        GROUP BY m_num, type
    """
    cursor.execute(query, (selected_year,))
    res_bar = cursor.fetchall()

    income_list = [0] * 12
    expense_list = [0] * 12
    

    for row in res_bar:
        m_idx = int(row['m_num']) - 1 
        if row['type'] == 'income':
            income_list[m_idx] = float(row['total'] or 0)
        elif row['type'] == 'expense':
            expense_list[m_idx] = float(row['total'] or 0)

    profit_loss_list = []
    for i in range(12):
        profit_loss_list.append(income_list[i] - expense_list[i])

    cursor.close()
    conn.close()
    return jsonify({"income_list": income_list, "expense_list": expense_list, "profit_loss_list": profit_loss_list})

# 2. ข้อมูลวงกลมสรุปการจ่ายเงิน (ฝั่งขวาบน)
@api.route('/api/payment_summary')
def payment_summary():
    today = datetime.now()
    year = request.args.get('year')
    month = request.args.get('month')

    # ดักจับกรณี JavaScript ส่งค่าว่าง หรือค่า 'undefined' มา
    if not year or year == 'undefined': year = str(today.year)
    if not month or month == 'undefined': month = str(today.month)
    
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    # 1. คำนวณ % การจ่ายเงิน (ใช้ DECIMAL เพื่อความแม่นยำใน SQL)
    query_stats = """
        SELECT 
            SUM(CASE WHEN status = 'paid' THEN total_amount ELSE 0 END) as paid_sum, 
            SUM(total_amount) as total_sum 
        FROM invoices 
        WHERE status != 'cancelled'
          AND YEAR(billing_period_start) = %s 
          AND MONTH(billing_period_start) = %s
    """
    cursor.execute(query_stats, (year, month))
    res = cursor.fetchone()
    
    # ดึงค่าออกมาและจัดการกับ None ให้เป็น 0.0
    paid_amt = float(res['paid_sum'] or 0)
    total_amt = float(res['total_sum'] or 0)

    # คำนวณเปอร์เซ็นต์
    if total_amt > 0:
        pay_percent = int(round((paid_amt / total_amt) * 100))
    else:
        # (แนะนำ 0 เพื่อให้กราฟว่าง แต่ถ้าอยากให้ขึ้น 'จ่ายครบ' อาจใช้ 100)
        pay_percent = 0 

    # 2. ดึงรายชื่อห้องที่ค้างชำระ (เพิ่ม status 'unpaid' หรืออื่นๆ ให้ครบ)
    query_unpaid = """
        SELECT u.name as room_, 
               COALESCE(CONCAT(t.fname, ' ', t.lname), CONCAT(i.guest_fname, ' ', i.guest_lname)) as tenant_name,
               i.total_amount as amount
        FROM invoices i
        JOIN unit u ON i.unit_id = u.unit_id
        LEFT JOIN tenants t ON i.tenant_id = t.tenant_id
        WHERE i.status NOT IN ('paid', 'cancelled') 
          AND YEAR(i.billing_period_start) = %s 
          AND MONTH(i.billing_period_start) = %s
        ORDER BY u.name ASC
    """
    cursor.execute(query_unpaid, (year, month))
    unpaid_rooms = cursor.fetchall()
    
    cursor.close()
    conn.close()

    return jsonify({
        "paid_percent": pay_percent, 
        "paid_amount": paid_amt, 
        "total_amount": total_amt,
        "unpaid_rooms": unpaid_rooms
    })

# 3. ดึงรายการธุรกรรมล่าสุด (Recent Transactions)
@api.route('/api/recent_transactions')
def get_recent_transactions():
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)
    
    # ดึงข้อมูลจากตารางเดียว (transactions)
    query = """
        SELECT 
            transaction_date, 
            type, 
            category, 
            COALESCE(note, '') as description, 
            amount
        FROM transactions
        ORDER BY transaction_date DESC
        LIMIT 15
    """
    try:
        cursor.execute(query)
        rows = cursor.fetchall()
        
        # จัดรูปแบบวันที่ให้ JS เข้าใจง่าย
        for row in rows:
            if row['transaction_date']:
                row['transaction_date'] = row['transaction_date'].strftime('%d/%m/%Y')
        
        return jsonify(rows)
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        cursor.close()
        conn.close()

@api.route('/api/finance_data')
def get_finance_data():
    year = request.args.get('year')
    month = request.args.get('month')
    
    conn = get_db_connection()
    cursor = conn.cursor(dictionary=True)

    # กรองตามปีเสมอ
    query = "SELECT type, SUM(amount) as total FROM transactions WHERE YEAR(transaction_date) = %s"
    params = [year]

    # ถ้าไม่ได้เลือก "ทุกเดือน" ให้กรองเดือนเพิ่ม
    if month != 'all':
        query += " AND MONTH(transaction_date) = %s"
        params.append(month)

    query += " GROUP BY type"
    
    cursor.execute(query, params)
    rows = cursor.fetchall()
    
    # ดึงค่าออกมา (เช็คตัวสะกด 'income' และ 'expense' ให้ตรงกับใน DB)
    income = next((float(r['total']) for r in rows if r['type'] == 'income'), 0.0)
    expense = next((float(r['total']) for r in rows if r['type'] == 'expense'), 0.0)
    profit = income - expense
    if income > 0:
        margin = ((income - expense) / income) * 100
    else:
        margin = 0.0

    return jsonify({
        "total_income": income,
        "total_expense": expense,
        "profit": profit,
        "margin": round(margin, 2)
    })


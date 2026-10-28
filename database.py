"""
Database layer for Procurement & Daily Purchasing Specification System.
Stores 365-day purchase plans, daily ingredient selections, benchmark catalog cache, and bill records.
"""

import sqlite3
import datetime
import os
from benchmark import load_benchmark_prices_from_excel, DEFAULT_EXCEL_PATH

USE_SUPABASE = os.environ.get("USE_SUPABASE", "false").lower() in ("true", "1", "yes")

try:
    import supabase_db
except ImportError:
    supabase_db = None

DB_PATH = os.path.join(os.path.dirname(__file__), "procurement.db")

def get_connection(db_path=DB_PATH):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn

def init_db(db_path=DB_PATH):
    """Initialize SQLite database tables."""
    conn = get_connection(db_path)
    cur = conn.cursor()

    # Benchmark prices cache
    cur.execute("""
    CREATE TABLE IF NOT EXISTS benchmark_items (
        code TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        category TEXT NOT NULL,
        unit TEXT NOT NULL,
        price REAL NOT NULL,
        supplier TEXT,
        min_qty REAL DEFAULT 1.0,
        shelf_life INTEGER DEFAULT 7,
        notes TEXT,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # Daily metadata for 365 days
    cur.execute("""
    CREATE TABLE IF NOT EXISTS daily_metadata (
        day_number INTEGER PRIMARY KEY,
        date TEXT NOT NULL UNIQUE,
        day_of_week TEXT NOT NULL,
        theme TEXT,
        notes TEXT,
        status TEXT DEFAULT 'Planned',
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    );
    """)

    # Daily purchase items (specifying what ingredients to buy each day & cost)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS daily_purchases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        day_number INTEGER NOT NULL,
        item_code TEXT NOT NULL,
        item_name TEXT NOT NULL,
        category TEXT NOT NULL,
        unit TEXT NOT NULL,
        quantity REAL NOT NULL,
        benchmark_price REAL NOT NULL,
        subtotal_cost REAL NOT NULL,
        notes TEXT,
        FOREIGN KEY (day_number) REFERENCES daily_metadata (day_number),
        FOREIGN KEY (item_code) REFERENCES benchmark_items (code)
    );
    """)

    # Index for fast lookup by day
    cur.execute("CREATE INDEX IF NOT EXISTS idx_purchases_day ON daily_purchases (day_number);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_purchases_code ON daily_purchases (item_code);")

    # Printed / Issued Bills record
    cur.execute("""
    CREATE TABLE IF NOT EXISTS bills (
        bill_number TEXT PRIMARY KEY,
        day_number INTEGER,
        start_date TEXT NOT NULL,
        end_date TEXT NOT NULL,
        bill_type TEXT DEFAULT 'Daily Purchase Bill',
        department TEXT DEFAULT 'Kitchen & Culinary Operations',
        vendor_reference TEXT,
        total_items INTEGER NOT NULL,
        total_quantity REAL NOT NULL,
        subtotal REAL NOT NULL,
        tax_rate REAL DEFAULT 0.0,
        tax_amount REAL DEFAULT 0.0,
        total_amount REAL NOT NULL,
        status TEXT DEFAULT 'Generated',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        notes TEXT
    );
    """)

    # Bill & PO configuration settings (ผู้ขาย, ส่วนราชการ, ผู้สั่งซื้อ, ผู้รับใบสั่งซื้อ)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS bill_settings (
        key TEXT PRIMARY KEY,
        value TEXT
    );
    """)

    conn.commit()
    conn.close()

def init_empty_365_days(db_path=DB_PATH):
    """
    Ensure the 365 days calendar framework exists in daily_metadata
    with empty menu themes and notes. Never populates mock purchases.
    """
    conn = get_connection(db_path)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM daily_metadata")
    cnt = cur.fetchone()[0]
    if cnt < 365:
        start_date = datetime.date(2026, 1, 1)
        thai_days = ["วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์", "วันอาทิตย์"]
        for i in range(365):
            day_num = i + 1
            dt = start_date + datetime.timedelta(days=i)
            date_str = dt.strftime("%Y-%m-%d")
            weekday_name = thai_days[dt.weekday()]
            cur.execute("""
                INSERT OR IGNORE INTO daily_metadata (day_number, date, day_of_week, theme, notes, status)
                VALUES (?, ?, ?, '', '', 'Active')
            """, (day_num, date_str, weekday_name))
        conn.commit()
    conn.close()

def sync_benchmark_from_excel(excel_path=DEFAULT_EXCEL_PATH, db_path=DB_PATH):
    """Sync fixed benchmark prices from Excel into SQLite."""
    items = load_benchmark_prices_from_excel(excel_path)
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        return supabase_db.sb_replace_benchmark_items(items)
    conn = get_connection(db_path)
    cur = conn.cursor()

    cur.execute("DELETE FROM benchmark_items;")
    for it in items:
        cur.execute("""
        INSERT INTO benchmark_items (code, name, category, unit, price, supplier, min_qty, shelf_life, notes, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        """, (
            it["code"], it["name"], it["category"], it["unit"], it["price"],
            it["supplier"], it["min_qty"], it["shelf_life"], it["notes"]
        ))

    # Also update any existing daily purchases to reflect updated benchmark prices
    cur.execute("""
        UPDATE daily_purchases
        SET benchmark_price = (SELECT price FROM benchmark_items WHERE benchmark_items.code = daily_purchases.item_code),
            subtotal_cost = ROUND(quantity * (SELECT price FROM benchmark_items WHERE benchmark_items.code = daily_purchases.item_code), 2)
        WHERE item_code IN (SELECT code FROM benchmark_items);
    """)

    conn.commit()
    conn.close()
    return len(items)

def get_all_benchmark_items(db_path=DB_PATH):
    """Retrieve all benchmark items ordered by category and name."""
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            return supabase_db.sb_get_all_benchmark_items()
        except Exception as e:
            print(f"[Supabase Fallback] get_all_benchmark_items: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()
    cur.execute("SELECT * FROM benchmark_items ORDER BY category ASC, name ASC")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows

def get_benchmark_item(code, db_path=DB_PATH):
    """Retrieve single benchmark item by item code."""
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            return supabase_db.sb_get_benchmark_item(code)
        except Exception as e:
            print(f"[Supabase Fallback] get_benchmark_item: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()
    cur.execute("SELECT * FROM benchmark_items WHERE code = ?", (code,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None

def get_daily_purchase_spec(day_number, db_path=DB_PATH):
    """Retrieve purchasing specification for a single day."""
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            res = supabase_db.sb_get_daily_purchase_spec(day_number)
            if res is not None:
                return res
        except Exception as e:
            print(f"[Supabase Fallback] get_daily_purchase_spec: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()

    cur.execute("SELECT * FROM daily_metadata WHERE day_number = ?", (day_number,))
    meta_row = cur.fetchone()
    if not meta_row:
        conn.close()
        return None

    meta = dict(meta_row)

    cur.execute("""
        SELECT id, day_number, item_code, item_name, category, unit, quantity, benchmark_price, subtotal_cost, notes
        FROM daily_purchases
        WHERE day_number = ?
        ORDER BY category ASC, item_name ASC
    """, (day_number,))
    items = [dict(r) for r in cur.fetchall()]

    total_cost = round(sum(it["subtotal_cost"] for it in items), 2)
    total_qty = round(sum(it["quantity"] for it in items), 2)
    item_count = len(items)

    # Category breakdown
    cat_summary = {}
    for it in items:
        c = it["category"]
        cat_summary[c] = round(cat_summary.get(c, 0.0) + it["subtotal_cost"], 2)

    conn.close()
    return {
        "metadata": meta,
        "items": items,
        "total_cost": total_cost,
        "total_quantity": total_qty,
        "item_count": item_count,
        "category_summary": cat_summary
    }

def update_daily_purchases(day_number, items_data, theme=None, notes=None, db_path=DB_PATH):
    """
    Replace the purchase specification for a given day.
    items_data: list of dicts with keys:
      item_code (required)
      quantity (required)
      notes (optional)
    The system looks up current benchmark price and calculates subtotal automatically.
    """
    sb_result = None
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            sb_result = supabase_db.sb_update_daily_purchases(day_number, items_data, theme=theme, notes=notes)
        except Exception as e:
            print(f"[Supabase Fallback] update_daily_purchases: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()

    # Update metadata if theme or notes provided
    if theme is not None or notes is not None:
        updates = []
        vals = []
        if theme is not None:
            updates.append("theme = ?")
            vals.append(theme)
        if notes is not None:
            updates.append("notes = ?")
            vals.append(notes)
        updates.append("updated_at = CURRENT_TIMESTAMP")
        vals.append(day_number)
        cur.execute(f"UPDATE daily_metadata SET {', '.join(updates)} WHERE day_number = ?", tuple(vals))

    # Delete existing purchase items for this day
    cur.execute("DELETE FROM daily_purchases WHERE day_number = ?", (day_number,))

    # Insert new items
    for it in items_data:
        code = it["item_code"]
        qty = float(it.get("quantity", 1.0))
        item_note = it.get("notes", "")

        # Fetch benchmark price and metadata
        cur.execute("SELECT name, category, unit, price FROM benchmark_items WHERE code = ?", (code,))
        b_item = cur.fetchone()
        if b_item:
            name = b_item["name"]
            cat = b_item["category"]
            unit = b_item["unit"]
            price = b_item["price"]
        else:
            name = it.get("item_name", "Custom Ingredient")
            cat = it.get("category", "General")
            unit = it.get("unit", "unit")
            price = float(it.get("benchmark_price", 0.0))

        subtotal = round(qty * price, 2)

        cur.execute("""
            INSERT INTO daily_purchases (day_number, item_code, item_name, category, unit, quantity, benchmark_price, subtotal_cost, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (day_number, code, name, cat, unit, qty, price, subtotal, item_note))

    conn.commit()
    conn.close()
    return sb_result if sb_result else get_daily_purchase_spec(day_number, db_path)

def get_365_days_summary(db_path=DB_PATH):
    """
    Get 365-day high-level summary:
    - Array of daily total cost, item count, date, day_of_week, theme for each day
    - Overall annual total cost, average daily spend, top 10 cost ingredients, monthly totals
    """
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            return supabase_db.sb_get_365_days_summary()
        except Exception as e:
            print(f"[Supabase Fallback] get_365_days_summary: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()

    cur.execute("""
        SELECT 
            m.day_number,
            m.date,
            m.day_of_week,
            m.theme,
            COALESCE(COUNT(p.id), 0) AS item_count,
            COALESCE(ROUND(SUM(p.quantity), 2), 0) AS total_quantity,
            COALESCE(ROUND(SUM(p.subtotal_cost), 2), 0) AS daily_total
        FROM daily_metadata m
        LEFT JOIN daily_purchases p ON m.day_number = p.day_number
        GROUP BY m.day_number
        ORDER BY m.day_number ASC
    """)
    daily_stats = [dict(r) for r in cur.fetchall()]

    annual_total = round(sum(d["daily_total"] for d in daily_stats), 2)
    avg_daily = round(annual_total / len(daily_stats), 2) if daily_stats else 0.0
    active_days = [d for d in daily_stats if d["daily_total"] > 0]
    max_day = max(active_days, key=lambda d: d["daily_total"]) if active_days else None
    min_day = min(active_days, key=lambda d: d["daily_total"]) if active_days else None

    # Top ingredients by annual spend
    cur.execute("""
        SELECT 
            item_code,
            item_name,
            category,
            unit,
            benchmark_price,
            ROUND(SUM(quantity), 2) AS annual_quantity,
            ROUND(SUM(subtotal_cost), 2) AS annual_spend,
            COUNT(DISTINCT day_number) AS days_purchased
        FROM daily_purchases
        GROUP BY item_code
        ORDER BY annual_spend DESC
        LIMIT 10
    """)
    top_ingredients = [dict(r) for r in cur.fetchall()]

    # Monthly breakdown
    cur.execute("""
        SELECT 
            SUBSTR(m.date, 1, 7) AS month,
            COUNT(DISTINCT m.day_number) AS days_count,
            COALESCE(ROUND(SUM(p.subtotal_cost), 2), 0) AS month_total,
            COALESCE(ROUND(SUM(p.quantity), 2), 0) AS month_qty
        FROM daily_metadata m
        LEFT JOIN daily_purchases p ON m.day_number = p.day_number
        GROUP BY SUBSTR(m.date, 1, 7)
        ORDER BY month ASC
    """)
    monthly_totals = [dict(r) for r in cur.fetchall()]

    # Category annual breakdown
    cur.execute("""
        SELECT 
            category,
            ROUND(SUM(subtotal_cost), 2) AS total_spend,
            ROUND(SUM(quantity), 2) AS total_qty
        FROM daily_purchases
        GROUP BY category
        ORDER BY total_spend DESC
    """)
    category_totals = [dict(r) for r in cur.fetchall()]

    conn.close()
    return {
        "daily_stats": daily_stats,
        "annual_total": annual_total,
        "average_daily": avg_daily,
        "max_day": max_day,
        "min_day": min_day,
        "top_ingredients": top_ingredients,
        "monthly_totals": monthly_totals,
        "category_totals": category_totals
    }

from benchmark import save_benchmark_prices_to_excel

def update_benchmark_item(code, data, excel_path=DEFAULT_EXCEL_PATH, db_path=DB_PATH):
    """Update benchmark item and automatically re-sync to master Excel file."""
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            supabase_db.sb_update_benchmark_item(code, data, excel_path)
        except Exception as e:
            print(f"[Supabase Fallback] update_benchmark_item: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()

    name = data["name"].strip()
    category = data.get("category", "ทั่วไป").strip()
    unit = data.get("unit", "กก.").strip()
    price = round(float(data["price"]), 2)
    supplier = data.get("supplier", "ร้านทัศนีย์ สุรเกียรติ").strip()
    min_qty = float(data.get("min_qty", 1.0))
    shelf_life = int(float(data.get("shelf_life", 7)))
    notes = data.get("notes", "").strip()

    cur.execute("""
        UPDATE benchmark_items
        SET name = ?, category = ?, unit = ?, price = ?, supplier = ?, min_qty = ?, shelf_life = ?, notes = ?, updated_at = CURRENT_TIMESTAMP
        WHERE code = ?
    """, (name, category, unit, price, supplier, min_qty, shelf_life, notes, code))

    # Recalculate daily purchases that use this item code
    cur.execute("""
        UPDATE daily_purchases
        SET item_name = ?, category = ?, unit = ?, benchmark_price = ?,
            subtotal_cost = ROUND(quantity * ?, 2)
        WHERE item_code = ?
    """, (name, category, unit, price, price, code))

    conn.commit()

    # Sync all items to benchmark_prices.xlsx
    cur.execute("SELECT * FROM benchmark_items ORDER BY code ASC")
    all_items = [dict(r) for r in cur.fetchall()]
    conn.close()

    save_benchmark_prices_to_excel(all_items, excel_path)
    return get_benchmark_item(code, db_path)

def create_benchmark_item(data, excel_path=DEFAULT_EXCEL_PATH, db_path=DB_PATH):
    """Create new benchmark item and sync to master Excel file."""
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            supabase_db.sb_create_benchmark_item(data, excel_path)
        except Exception as e:
            print(f"[Supabase Fallback] create_benchmark_item: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()

    code = data.get("code", "").strip()
    if not code:
        cur.execute("SELECT COUNT(*) as cnt FROM benchmark_items")
        cnt = cur.fetchone()["cnt"] + 1
        code = f"TH-{cnt:03d}"

    name = data["name"].strip()
    category = data.get("category", "ทั่วไป").strip()
    unit = data.get("unit", "กก.").strip()
    price = round(float(data.get("price", 0.0)), 2)
    supplier = data.get("supplier", "ร้านทัศนีย์ สุรเกียรติ").strip()
    min_qty = float(data.get("min_qty", 1.0))
    shelf_life = int(float(data.get("shelf_life", 7)))
    notes = data.get("notes", "").strip()

    cur.execute("""
        INSERT OR REPLACE INTO benchmark_items (code, name, category, unit, price, supplier, min_qty, shelf_life, notes, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
    """, (code, name, category, unit, price, supplier, min_qty, shelf_life, notes))

    conn.commit()

    cur.execute("SELECT * FROM benchmark_items ORDER BY code ASC")
    all_items = [dict(r) for r in cur.fetchall()]
    conn.close()

    save_benchmark_prices_to_excel(all_items, excel_path)
    return get_benchmark_item(code, db_path)

def delete_benchmark_item(code, excel_path=DEFAULT_EXCEL_PATH, db_path=DB_PATH):
    """Delete benchmark item from SQLite and re-sync master Excel file."""
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            supabase_db.sb_delete_benchmark_item(code, excel_path)
        except Exception as e:
            print(f"[Supabase Fallback] delete_benchmark_item: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()

    cur.execute("DELETE FROM benchmark_items WHERE code = ?", (code,))
    conn.commit()

    # Re-sync Excel file
    cur.execute("SELECT * FROM benchmark_items ORDER BY code ASC")
    all_items = [dict(r) for r in cur.fetchall()]
    conn.close()

    save_benchmark_prices_to_excel(all_items, excel_path)
    return True

def clear_benchmark_catalog(excel_path=DEFAULT_EXCEL_PATH, db_path=DB_PATH):
    """Clear the benchmark catalog from its active database and Excel cache."""
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        supabase_db.sb_clear_benchmark_items()

    conn = get_connection(db_path)
    cur = conn.cursor()
    cur.execute("DELETE FROM benchmark_items")
    conn.commit()
    conn.close()

    save_benchmark_prices_to_excel([], excel_path)
    return True

def delete_daily_purchase_item(purchase_id, db_path=DB_PATH):
    """Delete a single purchase item from a day's schedule."""
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            supabase_db.sb_delete_daily_purchase_item(purchase_id)
        except Exception as e:
            print(f"[Supabase Fallback] delete_daily_purchase_item: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()

    cur.execute("SELECT day_number FROM daily_purchases WHERE id = ?", (purchase_id,))
    row = cur.fetchone()
    if not row:
        conn.close()
        return None

    day_num = row["day_number"]
    cur.execute("DELETE FROM daily_purchases WHERE id = ?", (purchase_id,))
    conn.commit()
    conn.close()

    return get_daily_purchase_spec(day_num, db_path)

# ==============================================================
# การตั้งค่าข้อมูลใบสั่งซื้อทางการ (ผู้ขาย, ส่วนราชการ, ผู้ลงนาม)
# ==============================================================

DEFAULT_BILL_SETTINGS = {
    # ผู้ขาย
    "vendor_name": "นางสาวทัศนีย์ สุรเกียรติ",
    "vendor_address_line1": "หมู่บ้านปากตารอ 67/22 หมู่ที่ 11",
    "vendor_address_line2": "ต.บานา อ.เมือง จ.ปัตตานี 94000",
    "vendor_tel": "085-1523003",
    "vendor_tax_id": "1949900022220",
    "vendor_bank_acc": "907-0-34496-3",
    "vendor_bank_name": "น.ส.ทัศนีย์ สุรเกียรติ",
    "vendor_bank_branch": "ธนาคาร กรุงไทย จำกัด สาขาปัตตานี",
    # ส่วนราชการ
    "dept_name": "สถานสงเคราะห์เด็กชายนราธิวาส",
    "dept_agency": "กรมกิจการเด็กและเยาวชน",
    "dept_address_line1": "11 หมู่ 8 ตำบลลำภู อำเภอเมืองนราธิวาส",
    "dept_address_line2": "จังหวัดนราธิวาส",
    "dept_tel": "073-532054",
    # ข้อมูลใบสั่งซื้อ
    "fiscal_year": "2569",
    "bill_year": "2570",
    # ลายมือชื่อ หน้า 2
    "buyer_name": "นางสาวรานีซะ หะมิ",
    "buyer_position": "หัวหน้าเจ้าหน้าที่",
    "receiver_name": "นางสาวทัศนีย์ สุรเกียรติ",
    "receiver_position": "เจ้าของร้าน",
    # เงื่อนไขการสั่งซื้อ (หน้า 2)
    "cond_delivery_days": "1",
    "cond_deadline_text": "",
    "cond_location": "",
    "cond_warranty": "-",
    "cond_penalty_rate": "0.2",
    "cond_penalty_min": "100.00",
    "cond_item_6": "ส่วนราชการสงวนสิทธิ์ที่จะไม่รับมอบถ้าปรากฏว่าสินค้านั้นมีลักษณะไม่ตรงตามรายการที่ระบุไว้ในใบสั่งซื้อ กรณีนี้ผู้รับจ้างจะต้องดำเนินการเปลี่ยนใหม่ให้ถูกต้องตามใบสั่งซื้อทุกประการ",
    "cond_item_7": "ผู้กรณีงานจ้าง ผู้รับจะต้องไม่เอางานทั้งหมดหรือแต่บางส่วนแห่งสัญญานี้ไปจ้างช่วงอีกทอดหนึ่ง เว้นแต่การจ้างช่วงงานแต่บางส่วนที่ได้รับอนุญาตเป็นหนังสือจากผู้ว่าจ้างแล้ว การที่ผู้ว่าจ้างได้อนุญาตให้จ้างช่วงงานแต่บางส่วนดังกล่าวนั้น ไม่เป็นเหตุให้ผู้รับจ้างหลุดพ้นจากความรับผิดหรือพันธะหน้าที่ตามสัญญานี้ และผู้รับจ้างจะยังคงต้องรับผิดในความผิดกรณีผู้รับจ้างไปจ้างช่วงงานแต่บางส่วนโดยฝ่าฝืนความในวรรคหนึ่ง ผู้รับจ้างจะต้องชำระค่าปรับให้แก่ผู้ว่าจ้างเป็นจำนวนในอัตราร้อยละ 10 (สิบ) ของวงเงินของงานที่จ้างช่วงตามสัญญา ทั้งนี้ ไม่ตัดสิทธิผู้ว่าจ้างในการบอกเลิกสัญญา",
    "cond_item_8": "การประเมินผลการปฏิบัติงานของผู้ประกอบการ หน่วยงานของรัฐสามารถนำผลการปฏิบัติงานแล้วเสร็จตามสัญญาหรือข้อตกลงของคู่สัญญาเพื่อนำมาประเมินผลการปฏิบัติงานของผู้ประกอบการ"
}

def get_bill_settings(db_path=DB_PATH):
    """Retrieve bill settings merged with defaults."""
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            return supabase_db.sb_get_bill_settings()
        except Exception as e:
            print(f"[Supabase Fallback] get_bill_settings: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS bill_settings (key TEXT PRIMARY KEY, value TEXT)")
    cur.execute("SELECT key, value FROM bill_settings")
    rows = cur.fetchall()
    conn.close()

    settings = dict(DEFAULT_BILL_SETTINGS)
    for r in rows:
        settings[r["key"]] = r["value"]
    return settings

def update_bill_settings(settings_dict, db_path=DB_PATH):
    """Update bill settings in SQLite."""
    sb_result = None
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            sb_result = supabase_db.sb_update_bill_settings(settings_dict)
        except Exception as e:
            print(f"[Supabase Fallback] update_bill_settings: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS bill_settings (key TEXT PRIMARY KEY, value TEXT)")
    for k, v in settings_dict.items():
        if k in DEFAULT_BILL_SETTINGS or k.startswith("bill_") or k.startswith("order_") or k.startswith("deadline_") or k.startswith("cond_"):
            cur.execute("""
                INSERT INTO bill_settings (key, value) VALUES (?, ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
            """, (k, str(v).strip()))
    conn.commit()
    conn.close()
    return sb_result if sb_result else get_bill_settings(db_path)

def clear_month_purchases(month_str, clear_theme=True, db_path=DB_PATH):
    """
    Clear all daily purchases for days within a specific month (e.g. '2026-01').
    Returns number of deleted purchase records.
    """
    sb_deleted = None
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            sb_deleted = supabase_db.sb_clear_month_purchases(month_str, clear_theme=clear_theme)
        except Exception as e:
            print(f"[Supabase Fallback] clear_month_purchases: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()

    cur.execute("""
        DELETE FROM daily_purchases
        WHERE day_number IN (
            SELECT day_number FROM daily_metadata WHERE date LIKE ?
        )
    """, (f"{month_str}%",))
    deleted_items = cur.rowcount

    if clear_theme:
        cur.execute("""
            UPDATE daily_metadata
            SET theme = '', notes = ''
            WHERE date LIKE ?
        """, (f"{month_str}%",))

    conn.commit()
    conn.close()
    return sb_deleted if sb_deleted is not None else deleted_items

def clear_day_purchases(day_number, clear_theme=False, db_path=DB_PATH):
    """
    Clear all purchase items for a specific day.
    Returns number of deleted purchase records.
    """
    sb_deleted = None
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            sb_deleted = supabase_db.sb_clear_day_purchases(day_number, clear_theme=clear_theme)
        except Exception as e:
            print(f"[Supabase Fallback] clear_day_purchases: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM daily_purchases WHERE day_number = ?", (day_number,))
    deleted = cur.fetchone()[0]

    cur.execute("DELETE FROM daily_purchases WHERE day_number = ?", (day_number,))
    if clear_theme:
        cur.execute("UPDATE daily_metadata SET theme = '', notes = '', updated_at = CURRENT_TIMESTAMP WHERE day_number = ?", (day_number,))

    conn.commit()
    conn.close()
    return sb_deleted if sb_deleted is not None else deleted

def reset_all_data_365(db_path=DB_PATH):
    """
    Reset all 365 days data:
    - Delete all records in daily_purchases
    - Clear theme and notes in daily_metadata for all 365 days
    - Ensure framework of 365 days exists
    """
    sb_result = None
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            sb_result = supabase_db.sb_reset_all_data_365()
        except Exception as e:
            print(f"[Supabase Fallback] reset_all_data_365: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()
    cur.execute("DELETE FROM daily_purchases")
    deleted_purchases = cur.rowcount
    cur.execute("UPDATE daily_metadata SET theme = '', notes = '', updated_at = CURRENT_TIMESTAMP")
    updated_days = cur.rowcount
    conn.commit()
    conn.close()
    init_empty_365_days(db_path)
    return {
        "deleted_purchases": deleted_purchases,
        "updated_days": updated_days,
        "supabase": sb_result
    }

def reset_all_data_365(db_path=DB_PATH):
    """
    Reset all 365 days data:
    - Delete all records in daily_purchases
    - Clear theme and notes in daily_metadata for all 365 days
    - Ensure framework of 365 days exists
    """
    sb_result = None
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            sb_result = supabase_db.sb_reset_all_data_365()
        except Exception as e:
            print(f"[Supabase Fallback] reset_all_data_365: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()
    cur.execute("DELETE FROM daily_purchases")
    deleted_purchases = cur.rowcount
    cur.execute("UPDATE daily_metadata SET theme = '', notes = '', updated_at = CURRENT_TIMESTAMP")
    updated_days = cur.rowcount
    conn.commit()
    conn.close()
    init_empty_365_days(db_path)
    return {
        "deleted_purchases": deleted_purchases,
        "updated_days": updated_days,
        "supabase": sb_result
    }

def update_daily_metadata(day_number, theme=None, notes=None, db_path=DB_PATH):
    """
    Update theme (menu name) and/or notes for a specific day without touching its items.
    """
    sb_result = None
    if USE_SUPABASE and supabase_db and supabase_db.is_supabase_configured():
        try:
            sb_result = supabase_db.sb_update_daily_metadata(day_number, theme=theme, notes=notes)
        except Exception as e:
            print(f"[Supabase Fallback] update_daily_metadata: {e}")

    conn = get_connection(db_path)
    cur = conn.cursor()
    updates = []
    vals = []
    if theme is not None:
        updates.append("theme = ?")
        vals.append(str(theme).strip())
    if notes is not None:
        updates.append("notes = ?")
        vals.append(str(notes).strip())
    if updates:
        updates.append("updated_at = CURRENT_TIMESTAMP")
        vals.append(day_number)
        cur.execute(f"UPDATE daily_metadata SET {', '.join(updates)} WHERE day_number = ?", tuple(vals))
        conn.commit()
    conn.close()
    return sb_result if sb_result else get_daily_purchase_spec(day_number, db_path)

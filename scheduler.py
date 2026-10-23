"""
365-Day Purchasing Scheduler with Authentic Thai Institutional Menus.
Coordinates daily ingredient specifications with benchmark prices in Excel,
including meal plans (เช้า, เที่ยง, เย็น) for government and welfare institutions.
"""

import datetime
from database import get_connection, sync_benchmark_from_excel, DB_PATH
from benchmark import load_benchmark_prices_from_excel, DEFAULT_EXCEL_PATH

THAI_MONTHS = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

THAI_DAYS = ["วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์", "วันอาทิตย์"]

def to_thai_date_str(date_obj, include_day_name=True):
    b_year = date_obj.year + 543
    month_name = THAI_MONTHS[date_obj.month]
    if include_day_name:
        day_name = THAI_DAYS[date_obj.weekday()]
        return f"{day_name} ที่ {date_obj.day} {month_name} {b_year}"
    return f"{date_obj.day} {month_name} {b_year}"

# Curated daily meal plans matching welfare school / institutional standards
THAI_DAILY_MENUS = {
    0: { # จันทร์
        "theme": "ข้าวมันไก่ & ต้มจืดฟัก",
        "breakfast": ("ข้าวต้มหมูสับเห็ดหอม", "นมถั่วเหลืองร้อน"),
        "lunch": "ข้าวมันไก่ต้ม พร้อมน้ำจิ้มและน้ำซุป",
        "dinner": ("แกงส้มปลากะพงผักรวม", "ไข่เจียว แตงโม"),
        "items": [
            {"code": "TH-001", "qty": 4.0, "note": "สำหรับข้าวมันไก่"},
            {"code": "TH-022", "qty": 2.0, "note": "สำหรับข้าวต้มหมูสับ"},
            {"code": "TH-020", "qty": 15.0, "note": "หุงข้าว 3 มื้อ"},
            {"code": "TH-023", "qty": 3.0, "note": "แกงส้มปลา"},
            {"code": "TH-027", "qty": 3.0, "note": "ผักรวมแกงส้ม"},
            {"code": "TH-017", "qty": 30.0, "note": "ไข่เจียวมื้อเย็น"},
            {"code": "TH-036", "qty": 6.0, "note": "ผลไม้แตงโม"},
            {"code": "TH-018", "qty": 1.0, "note": "น้ำมันปรุงอาหาร"},
            {"code": "TH-032", "qty": 1.0, "note": "น้ำปลาปรุงรส"}
        ]
    },
    1: { # อังคาร
        "theme": "ผัดกะเพราไก่ & แกงเลียงกุ้งสด",
        "breakfast": ("ขนมปังปิ้ง แยมสับปะรด", "โอวัลตินร้อน"),
        "lunch": "ผัดกะเพราไก่สับ ไข่ดาว",
        "dinner": ("แกงเลียงกุ้งสดยอดผัก", "ไก่ทอดกระเทียม กล้วยน้ำว้า"),
        "items": [
            {"code": "TH-001", "qty": 5.0, "note": "ผัดกะเพราและไก่ทอด"},
            {"code": "TH-024", "qty": 2.5, "note": "แกงเลียงกุ้งสด"},
            {"code": "TH-016", "qty": 4.0, "note": "ยอดผักแกงเลียง"},
            {"code": "TH-017", "qty": 35.0, "note": "ไข่ดาวมื้อเที่ยง"},
            {"code": "TH-038", "qty": 3.0, "note": "ขนมปังมื้อเช้า"},
            {"code": "TH-039", "qty": 1.0, "note": "แยมทาขนมปัง"},
            {"code": "TH-007", "qty": 1.0, "note": "ผงมอลต์เครื่องดื่ม"},
            {"code": "TH-008", "qty": 1.0, "note": "นมข้นหวาน"},
            {"code": "TH-030", "qty": 1.0, "note": "กระเทียมเจียว"},
            {"code": "TH-035", "qty": 3.0, "note": "กล้วยน้ำว้า"}
        ]
    },
    2: { # พุธ
        "theme": "พะแนงหมู & ต้มจืดเต้าหู้หมูสับ",
        "breakfast": ("ข้าวผัดไข่ใส่ต้นหอม", "ไมโลร้อน"),
        "lunch": "ก๋วยเตี๋ยวไก่ตุ๋นมะระ",
        "dinner": ("พะแนงหมูยอดมะพร้าว", "ต้มจืดเต้าหู้หมูสับ สับปะรด"),
        "items": [
            {"code": "TH-021", "qty": 4.0, "note": "พะแนงหมู"},
            {"code": "TH-022", "qty": 2.0, "note": "หมูสับแกงจืด"},
            {"code": "TH-040", "qty": 6.0, "note": "เต้าหู้ไข่"},
            {"code": "TH-004", "qty": 2.0, "note": "ต้มน้ำซุปก๋วยเตี๋ยว"},
            {"code": "TH-001", "qty": 3.0, "note": "เนื้อไก่ก๋วยเตี๋ยว"},
            {"code": "TH-015", "qty": 2.0, "note": "กะทิสดทำพะแนง"},
            {"code": "TH-005", "qty": 2.0, "note": "ต้นหอมโรยหน้า"},
            {"code": "TH-037", "qty": 3.0, "note": "สับปะรดผลไม้"},
            {"code": "TH-018", "qty": 1.0, "note": "น้ำมันพืช"}
        ]
    },
    3: { # พฤหัสบดี
        "theme": "แกงเขียวหวานไก่ & ไข่พะโล้",
        "breakfast": ("ข้าวต้มปลาโรยคื่นช่าย", "ช็อกโกแลตร้อน"),
        "lunch": "ไข่พะโล้หมูสามชั้น เต้าหู้ขาว",
        "dinner": ("แกงเขียวหวานไก่ยอดมะเขือ", "ปลาทอดขมิ้น แตงกวาผัดไข่"),
        "items": [
            {"code": "TH-001", "qty": 4.0, "note": "แกงเขียวหวานไก่"},
            {"code": "TH-023", "qty": 3.0, "note": "เนื้อปลาข้าวต้มและปลาทอด"},
            {"code": "TH-017", "qty": 36.0, "note": "ไข่พะโล้ 36 ฟอง"},
            {"code": "TH-021", "qty": 3.0, "note": "หมูพะโล้"},
            {"code": "TH-015", "qty": 2.0, "note": "กะทิแกงเขียวหวาน"},
            {"code": "TH-029", "qty": 3.0, "note": "แตงกวาผัดไข่"},
            {"code": "TH-009", "qty": 1.0, "note": "น้ำตาลปรุงพะโล้"},
            {"code": "TH-033", "qty": 1.0, "note": "ซีอิ๊วขาวปรุงรส"}
        ]
    },
    4: { # ศุกร์
        "theme": "ผัดซีอิ๊วหมู & ต้มยำปลากะพงน้ำใส",
        "breakfast": ("แซนวิชไก่เชียงไข่ดาว", "นมสดหวาน"),
        "lunch": "ผัดซีอิ๊วเส้นใหญ่หมูหมัก ผักคะน้า",
        "dinner": ("ต้มยำปลากะพงน้ำใสเห็ดฟาง", "ไข่เจียวหมูสับ มะละกอสุก"),
        "items": [
            {"code": "TH-021", "qty": 3.0, "note": "หมูหมักผัดซีอิ๊ว"},
            {"code": "TH-022", "qty": 2.0, "note": "หมูสับไข่เจียว"},
            {"code": "TH-023", "qty": 3.5, "note": "ปลากะพงต้มยำ"},
            {"code": "TH-003", "qty": 4.0, "note": "คะน้าผัดซีอิ๊ว"},
            {"code": "TH-010", "qty": 2.0, "note": "ไก่เชียงแซนวิช"},
            {"code": "TH-017", "qty": 30.0, "note": "ไข่ไก่"},
            {"code": "TH-013", "qty": 1.5, "note": "มะนาวต้มยำ"},
            {"code": "TH-019", "qty": 4.0, "note": "มะละกอสุก"},
            {"code": "TH-038", "qty": 2.0, "note": "ขนมปังแซนวิช"}
        ]
    },
    5: { # เสาร์ (เหมือนในเอกสารตัวอย่างใบสั่งซื้อ 24/2570 วันที่ 24 ตุลาคม 2569)
        "theme": "ข้าวผัดซอสแดง & แกงคั่วหอยใบชะพลู",
        "breakfast": ("ข้าวผัดซอสแดง แกงจืด", "ไมโลร้อน"),
        "lunch": "ยำไก่เชียง",
        "dinner": ("แกงคั่วหอยกับใบชะพลู", "ไข่เจียว มะละกอสุก"),
        "items": [
            {"code": "TH-001", "qty": 3.0, "note": "อกไก่ (หั่น)"},
            {"code": "TH-002", "qty": 1.0, "note": "ซอสแดง"},
            {"code": "TH-003", "qty": 1.0, "note": "ผักคะน้า"},
            {"code": "TH-004", "qty": 1.0, "note": "โครงไก่"},
            {"code": "TH-005", "qty": 2.0, "note": "ต้นหอม"},
            {"code": "TH-006", "qty": 1.0, "note": "ผักกวางตุ้ง"},
            {"code": "TH-007", "qty": 1.0, "note": "ผงช็อกโกแลตมอลต์"},
            {"code": "TH-008", "qty": 1.0, "note": "นมข้นหวาน"},
            {"code": "TH-009", "qty": 1.0, "note": "น้ำตาลทราย"},
            {"code": "TH-010", "qty": 3.0, "note": "ไก่เชียง"},
            {"code": "TH-011", "qty": 0.5, "note": "หอมหัวใหญ่"},
            {"code": "TH-012", "qty": 0.5, "note": "มะเขือเทศ"},
            {"code": "TH-013", "qty": 0.5, "note": "มะนาว"},
            {"code": "TH-014", "qty": 4.0, "note": "หอยแครง"},
            {"code": "TH-015", "qty": 1.0, "note": "น้ำกะทิ"},
            {"code": "TH-016", "qty": 3.0, "note": "ยอดผักต่างๆ"},
            {"code": "TH-017", "qty": 36.0, "note": "ไข่ไก่ เบอร์ 2"},
            {"code": "TH-018", "qty": 1.0, "note": "น้ำมันพืช"},
            {"code": "TH-019", "qty": 4.0, "note": "มะละกอสุก"}
        ]
    },
    6: { # อาทิตย์
        "theme": "ก๋วยเตี๋ยวคั่วไก่ & แกงส้มผักกาดขาว",
        "breakfast": ("โจ๊กหมูสับใส่ไข่ลวก", "นมสดร้อน"),
        "lunch": "ก๋วยเตี๋ยวคั่วไก่กรอบ ผักกาดหอม",
        "dinner": ("แกงส้มปลาผักกาดขาว", "หมูแดดเดียว แตงโมหวาน"),
        "items": [
            {"code": "TH-001", "qty": 4.0, "note": "คั่วไก่"},
            {"code": "TH-022", "qty": 2.5, "note": "หมูสับโจ๊ก"},
            {"code": "TH-021", "qty": 3.0, "note": "หมูแดดเดียว"},
            {"code": "TH-023", "qty": 3.0, "note": "ปลาแกงส้ม"},
            {"code": "TH-017", "qty": 30.0, "note": "ไข่ลวกและคั่วไก่"},
            {"code": "TH-027", "qty": 3.0, "note": "ผักกาดขาว"},
            {"code": "TH-036", "qty": 5.0, "note": "แตงโมหวาน"},
            {"code": "TH-018", "qty": 1.0, "note": "น้ำมันพืช"}
        ]
    }
}

PROCUREMENT_TEMPLATES = {
    "saturday_special": {
        "name": "ชุดวันเสาร์ (ตามตัวอย่างใบสั่งซื้อ)",
        "description": "ข้าวผัดซอสแดง แกงจืด ยำไก่เชียง แกงคั่วหอยใบชะพลู ไข่เจียว มะละกอสุก",
        "theme": "ข้าวผัดซอสแดง & แกงคั่วหอยใบชะพลู",
        "items": THAI_DAILY_MENUS[5]["items"]
    },
    "monday_chicken_rice": {
        "name": "ชุดวันจันทร์ - ข้าวมันไก่ & ต้มจืด",
        "description": "ข้าวมันไก่ แกงส้มปลากะพง ไข่เจียว แตงโม",
        "theme": "ข้าวมันไก่ & ต้มจืดฟัก",
        "items": THAI_DAILY_MENUS[0]["items"]
    },
    "tuesday_basil": {
        "name": "ชุดวันอังคาร - ผัดกะเพรา & แกงเลียงกุ้ง",
        "description": "ผัดกะเพราไก่สับ แกงเลียงกุ้งสด ไก่ทอดกระเทียม",
        "theme": "ผัดกะเพราไก่ & แกงเลียงกุ้งสด",
        "items": THAI_DAILY_MENUS[1]["items"]
    },
    "wednesday_panang": {
        "name": "ชุดวันพุธ - พะแนงหมู & ก๋วยเตี๋ยวไก่",
        "description": "ก๋วยเตี๋ยวไก่ พะแนงหมู ต้มจืดเต้าหู้หมูสับ",
        "theme": "พะแนงหมู & ต้มจืดเต้าหู้หมูสับ",
        "items": THAI_DAILY_MENUS[2]["items"]
    },
    "thursday_green_curry": {
        "name": "ชุดวันพฤหัสบดี - แกงเขียวหวาน & ไข่พะโล้",
        "description": "ไข่พะโล้ แกงเขียวหวานไก่ ปลาทอดขมิ้น",
        "theme": "แกงเขียวหวานไก่ & ไข่พะโล้",
        "items": THAI_DAILY_MENUS[3]["items"]
    },
    "friday_seafood": {
        "name": "ชุดวันศุกร์ - ผัดซีอิ๊ว & ต้มยำปลากะพง",
        "description": "ผัดซีอิ๊วหมู ต้มยำปลากะพงน้ำใส ไข่เจียว",
        "theme": "ผัดซีอิ๊วหมู & ต้มยำปลากะพงน้ำใส",
        "items": THAI_DAILY_MENUS[4]["items"]
    }
}

def generate_365_days_schedule(start_date=None, excel_path=DEFAULT_EXCEL_PATH, db_path=DB_PATH):
    """
    Generate the complete 365-day purchasing schedule based on authentic institutional menus.
    """
    if start_date is None:
        start_date = datetime.date(2026, 1, 1)
    elif isinstance(start_date, str):
        start_date = datetime.datetime.strptime(start_date, "%Y-%m-%d").date()

    sync_benchmark_from_excel(excel_path, db_path)
    benchmark_items = {it["code"]: it for it in load_benchmark_prices_from_excel(excel_path)}

    conn = get_connection(db_path)
    cur = conn.cursor()

    cur.execute("DELETE FROM daily_purchases;")
    cur.execute("DELETE FROM daily_metadata;")

    for day_num in range(1, 366):
        current_date = start_date + datetime.timedelta(days=day_num - 1)
        weekday_idx = current_date.weekday()
        weekday_name = THAI_DAYS[weekday_idx]
        date_str = current_date.strftime("%Y-%m-%d")

        menu_plan = THAI_DAILY_MENUS[weekday_idx]
        theme = menu_plan["theme"]

        b_main, b_drink = menu_plan["breakfast"]
        l_main = menu_plan["lunch"]
        d_main, d_sub = menu_plan["dinner"]
        
        meal_notes_formatted = (
            f"มื้อเช้า  {b_main}\n"
            f"        {b_drink}\n"
            f"มื้อเที่ยง {l_main}\n"
            f"มื้อเย็น  {d_main}\n"
            f"        {d_sub}"
        )

        cur.execute("""
            INSERT INTO daily_metadata (day_number, date, day_of_week, theme, notes, status)
            VALUES (?, ?, ?, ?, ?, 'Active')
        """, (day_num, date_str, weekday_name, theme, meal_notes_formatted))

        for it in menu_plan["items"]:
            code = it["code"]
            qty = float(it["qty"])
            note = it.get("note", "")

            b_info = benchmark_items.get(code)
            if b_info:
                name = b_info["name"]
                cat = b_info["category"]
                unit = b_info["unit"]
                price = b_info["price"]
            else:
                name = f"รายการ {code}"
                cat = "ทั่วไป"
                unit = "กก."
                price = 50.0

            subtotal = round(qty * price, 2)

            cur.execute("""
                INSERT INTO daily_purchases (day_number, item_code, item_name, category, unit, quantity, benchmark_price, subtotal_cost, notes)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (day_num, code, name, cat, unit, qty, price, subtotal, note))

    conn.commit()
    conn.close()
    return 365

def apply_template_to_days(template_key, day_numbers, replace=False, db_path=DB_PATH):
    """Apply a selected template to specific days."""
    if template_key not in PROCUREMENT_TEMPLATES:
        raise ValueError(f"Unknown template: {template_key}")

    template = PROCUREMENT_TEMPLATES[template_key]
    conn = get_connection(db_path)
    cur = conn.cursor()

    cur.execute("SELECT code, name, category, unit, price FROM benchmark_items")
    benchmarks = {r["code"]: dict(r) for r in cur.fetchall()}

    for day_num in day_numbers:
        if replace:
            cur.execute("DELETE FROM daily_purchases WHERE day_number = ?", (day_num,))
            cur.execute("UPDATE daily_metadata SET theme = ?, updated_at = CURRENT_TIMESTAMP WHERE day_number = ?", 
                        (template["theme"], day_num))

        for it in template["items"]:
            code = it["code"]
            qty = float(it["qty"])
            note = it.get("note", "")

            b_info = benchmarks.get(code, {
                "name": code,
                "category": "ทั่วไป",
                "unit": "กก.",
                "price": 0.0
            })

            cur.execute("SELECT id, quantity FROM daily_purchases WHERE day_number = ? AND item_code = ?", (day_num, code))
            existing = cur.fetchone()

            if existing:
                new_qty = round(existing["quantity"] + qty, 2)
                subtotal = round(new_qty * b_info["price"], 2)
                cur.execute("""
                    UPDATE daily_purchases 
                    SET quantity = ?, subtotal_cost = ?, notes = notes || ' | ' || ?
                    WHERE id = ?
                """, (new_qty, subtotal, note, existing["id"]))
            else:
                subtotal = round(qty * b_info["price"], 2)
                cur.execute("""
                    INSERT INTO daily_purchases (day_number, item_code, item_name, category, unit, quantity, benchmark_price, subtotal_cost, notes)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (day_num, code, b_info["name"], b_info["category"], b_info["unit"], qty, b_info["price"], subtotal, note))

    conn.commit()
    conn.close()
    return len(day_numbers)

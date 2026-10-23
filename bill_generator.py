"""
Printable Purchase Order / Bill Generator (ใบสั่งซื้อ).
Matches the exact Thai government / welfare home purchase order format (ราชการ).
Uses Sarabun / TH Sarabun New font for printing, exact official tables, conditions, and signatures.
"""

import datetime
from database import get_connection, get_daily_purchase_spec, get_bill_settings, DB_PATH
from xlsx_engine import write_xlsx

THAI_MONTHS = [
    "", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

THAI_DAYS = ["วันจันทร์", "วันอังคาร", "วันพุธ", "วันพฤหัสบดี", "วันศุกร์", "วันเสาร์", "วันอาทิตย์"]

def to_thai_date(dt, include_day_name=True):
    b_year = dt.year + 543
    m_name = THAI_MONTHS[dt.month]
    if include_day_name:
        d_name = THAI_DAYS[dt.weekday()]
        return f"{d_name} ที่ {dt.day} {m_name} {b_year}"
    return f"{dt.day} {m_name} {b_year}"

def baht_text(num):
    """Convert number to Thai Baht text format, e.g. 2433.50 -> (สองพันสี่ร้อยสามสิบสามบาทห้าสิบสตางค์)"""
    num_str = f"{num:.2f}"
    baht, satang = num_str.split(".")
    
    digits = ["", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"]
    units = ["", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน", "ล้าน"]
    
    def convert_group(n_str):
        res = ""
        length = len(n_str)
        for i, ch in enumerate(n_str):
            d = int(ch)
            pos = length - i - 1
            if d == 0:
                continue
            if pos == 0 and d == 1 and length > 1:
                res += "เอ็ด"
            elif pos == 1 and d == 2:
                res += "ยี่" + units[pos]
            elif pos == 1 and d == 1:
                res += units[pos]
            else:
                res += digits[d] + units[pos]
        return res
    
    baht_val = int(baht)
    if baht_val == 0:
        b_txt = "ศูนย์บาท"
    else:
        if len(baht) > 6:
            millions = baht[:-6]
            remainder = baht[-6:]
            b_txt = convert_group(millions) + "ล้าน" + convert_group(remainder) + "บาท"
        else:
            b_txt = convert_group(baht) + "บาท"
            
    satang_val = int(satang)
    if satang_val == 0:
        s_txt = "ถ้วน"
    else:
        s_txt = convert_group(satang) + "สตางค์"
        
    return f"({b_txt}{s_txt})"

def _extract_conditions(settings, default_deadline, dept_info, day_num_key=None):
    cond_deadline = ""
    if day_num_key:
        cond_deadline = settings.get(f"cond_deadline_day_{day_num_key}", "")
    if not cond_deadline:
        cond_deadline = settings.get("cond_deadline_text", "") or default_deadline

    default_location = f"{dept_info.get('dept_name', '')} {dept_info.get('dept_agency', '')} {dept_info.get('dept_address_line1', '')} {dept_info.get('dept_address_line2', '')}".strip()
    cond_location = settings.get("cond_location") or default_location

    return {
        "cond_delivery_days": settings.get("cond_delivery_days", "1"),
        "cond_deadline_text": cond_deadline,
        "cond_location": cond_location,
        "cond_warranty": settings.get("cond_warranty", "-"),
        "cond_penalty_rate": settings.get("cond_penalty_rate", "0.2"),
        "cond_penalty_min": settings.get("cond_penalty_min", "100.00"),
        "cond_item_6": settings.get("cond_item_6", "ส่วนราชการสงวนสิทธิ์ที่จะไม่รับมอบถ้าปรากฏว่าสินค้านั้นมีลักษณะไม่ตรงตามรายการที่ระบุไว้ในใบสั่งซื้อ กรณีนี้ผู้รับจ้างจะต้องดำเนินการเปลี่ยนใหม่ให้ถูกต้องตามใบสั่งซื้อทุกประการ"),
        "cond_item_7": settings.get("cond_item_7", "ผู้กรณีงานจ้าง ผู้รับจะต้องไม่เอางานทั้งหมดหรือแต่บางส่วนแห่งสัญญานี้ไปจ้างช่วงอีกทอดหนึ่ง เว้นแต่การจ้างช่วงงานแต่บางส่วนที่ได้รับอนุญาตเป็นหนังสือจากผู้ว่าจ้างแล้ว การที่ผู้ว่าจ้างได้อนุญาตให้จ้างช่วงงานแต่บางส่วนดังกล่าวนั้น ไม่เป็นเหตุให้ผู้รับจ้างหลุดพ้นจากความรับผิดหรือพันธะหน้าที่ตามสัญญานี้ และผู้รับจ้างจะยังคงต้องรับผิดในความผิดกรณีผู้รับจ้างไปจ้างช่วงงานแต่บางส่วนโดยฝ่าฝืนความในวรรคหนึ่ง ผู้รับจ้างจะต้องชำระค่าปรับให้แก่ผู้ว่าจ้างเป็นจำนวนในอัตราร้อยละ 10 (สิบ) ของวงเงินของงานที่จ้างช่วงตามสัญญา ทั้งนี้ ไม่ตัดสิทธิผู้ว่าจ้างในการบอกเลิกสัญญา"),
        "cond_item_8": settings.get("cond_item_8", "การประเมินผลการปฏิบัติงานของผู้ประกอบการ หน่วยงานของรัฐสามารถนำผลการปฏิบัติงานแล้วเสร็จตามสัญญาหรือข้อตกลงของคู่สัญญาเพื่อนำมาประเมินผลการปฏิบัติงานของผู้ประกอบการ")
    }

def get_bill_data_for_day(day_number, db_path=DB_PATH):
    """Retrieve full bill data for a specific day."""
    spec = get_daily_purchase_spec(day_number, db_path)
    if not spec:
        return None

    settings = get_bill_settings(db_path)
    meta = spec["metadata"]
    items = spec["items"]

    # Parse delivery date
    service_date = datetime.datetime.strptime(meta["date"], "%Y-%m-%d").date()
    # Order date is 3 days prior or current
    order_date = service_date - datetime.timedelta(days=3)
    deadline_date = order_date + datetime.timedelta(days=1)
    
    bill_year = settings.get("bill_year", "2570")
    default_bill_no = f"{day_number}/{bill_year}"
    bill_no = settings.get(f"bill_number_day_{day_number}", default_bill_no)

    default_order_date_thai = to_thai_date(order_date, include_day_name=False)
    order_date_thai = settings.get(f"order_date_day_{day_number}", default_order_date_thai)

    default_deadline_date_thai = to_thai_date(deadline_date, include_day_name=False)
    deadline_date_thai = settings.get(f"deadline_date_day_{day_number}", default_deadline_date_thai)

    subtotal = spec["total_cost"]
    baht_wording = baht_text(subtotal)

    res = {
        "bill_number": bill_no,
        "day_number": day_number,
        "fiscal_year": settings.get("fiscal_year", "2569"),
        "service_date": service_date,
        "service_date_thai": to_thai_date(service_date, include_day_name=True),
        "order_date_thai": order_date_thai,
        "deadline_date_thai": deadline_date_thai,
        "theme": meta.get("theme", ""),
        "meal_notes": meta.get("notes", ""),
        "items": items,
        "total_items": spec["item_count"],
        "total_quantity": spec["total_quantity"],
        "grand_total": subtotal,
        "baht_text": baht_wording,
        # Supplier details (ผู้ขาย)
        "vendor_name": settings.get("vendor_name", "นางสาวทัศนีย์ สุรเกียรติ"),
        "vendor_address_line1": settings.get("vendor_address_line1", "หมู่บ้านปากตารอ 67/22 หมู่ที่ 11"),
        "vendor_address_line2": settings.get("vendor_address_line2", "ต.บานา อ.เมือง จ.ปัตตานี 94000"),
        "vendor_tel": settings.get("vendor_tel", "085-1523003"),
        "vendor_tax_id": settings.get("vendor_tax_id", "1949900022220"),
        "vendor_bank_acc": settings.get("vendor_bank_acc", "907-0-34496-3"),
        "vendor_bank_name": settings.get("vendor_bank_name", "น.ส.ทัศนีย์ สุรเกียรติ"),
        "vendor_bank_branch": settings.get("vendor_bank_branch", "ธนาคาร กรุงไทย จำกัด สาขาปัตตานี"),
        # Government Department details (ส่วนราชการ)
        "dept_name": settings.get("dept_name", "สถานสงเคราะห์เด็กชายนราธิวาส"),
        "dept_agency": settings.get("dept_agency", "กรมกิจการเด็กและเยาวชน"),
        "dept_address_line1": settings.get("dept_address_line1", "11 หมู่ 8 ตำบลลำภู อำเภอเมืองนราธิวาส"),
        "dept_address_line2": settings.get("dept_address_line2", "จังหวัดนราธิวาส"),
        "dept_tel": settings.get("dept_tel", "073-532054"),
        # Signatures (ผู้สั่งซื้อ & ผู้รับใบสั่งซื้อ)
        "buyer_name": settings.get("buyer_name", "นางสาวรานีซะ หะมิ"),
        "buyer_position": settings.get("buyer_position", "หัวหน้าเจ้าหน้าที่"),
        "receiver_name": settings.get("receiver_name", "นางสาวทัศนีย์ สุรเกียรติ"),
        "receiver_position": settings.get("receiver_position", "เจ้าของร้าน")
    }
    res.update(_extract_conditions(settings, deadline_date_thai, res, day_num_key=day_number))
    return res

def get_bill_data_for_month(month_str, db_path=DB_PATH):
    """
    Retrieve consolidated bill data for a calendar month.
    month_str format: 'YYYY-MM' e.g. '2026-01'
    Queries live daily_purchases data filtered by daily_metadata.date.
    """
    conn = get_connection(db_path)
    cur = conn.cursor()

    # ดึง metadata ของเดือนนั้น (วันเริ่ม-สิ้นสุด, วันแรก-สุดท้าย)
    cur.execute("""
        SELECT MIN(day_number) as start_day, MAX(day_number) as end_day,
               MIN(date) as min_date, MAX(date) as max_date,
               COUNT(DISTINCT day_number) as day_count
        FROM daily_metadata
        WHERE SUBSTR(date, 1, 7) = ?
    """, (month_str,))
    row = cur.fetchone()
    if not row or row[0] is None:
        conn.close()
        return None

    range_meta = dict(row)
    start_day = range_meta["start_day"]
    end_day   = range_meta["end_day"]

    # ดึงรายการจัดซื้อจาก daily_purchases จริงๆ (live data)
    cur.execute("""
        SELECT
            p.item_code, p.item_name, p.category, p.unit, p.benchmark_price,
            ROUND(SUM(p.quantity), 2)      AS quantity,
            ROUND(SUM(p.subtotal_cost), 2) AS subtotal_cost,
            GROUP_CONCAT(DISTINCT p.notes) AS notes
        FROM daily_purchases p
        INNER JOIN daily_metadata m ON p.day_number = m.day_number
        WHERE SUBSTR(m.date, 1, 7) = ?
        GROUP BY p.item_code
        ORDER BY p.item_code ASC
    """, (month_str,))
    items = [dict(r) for r in cur.fetchall()]
    conn.close()

    subtotal     = round(sum(it["subtotal_cost"] for it in items), 2)
    baht_wording = baht_text(subtotal)

    start_dt = datetime.datetime.strptime(range_meta["min_date"], "%Y-%m-%d").date()
    end_dt   = datetime.datetime.strptime(range_meta["max_date"], "%Y-%m-%d").date()

    settings  = get_bill_settings(db_path)
    bill_year = settings.get("bill_year", "2570")

    # ชื่อเดือนภาษาไทย + พ.ศ.
    month_name = THAI_MONTHS[start_dt.month]
    b_year     = start_dt.year + 543
    month_label = f"{month_name} {b_year}"

    res = {
        "bill_number":        f"เดือน{month_label}/{bill_year}",
        "day_number":         f"{start_day}-{end_day}",
        "fiscal_year":        settings.get("fiscal_year", "2569"),
        "service_date":       start_dt,
        "service_date_thai":  f"{to_thai_date(start_dt, False)} ถึง {to_thai_date(end_dt, False)}",
        "order_date_thai":    to_thai_date(start_dt - datetime.timedelta(days=2), False),
        "deadline_date_thai": to_thai_date(start_dt, False),
        "theme":              f"ใบสั่งซื้อรายเดือน{month_label} ({range_meta['day_count']} วัน วัน {start_day}–{end_day})",
        "meal_notes":         f"รายการจัดซื้อรวมเดือน{month_label} (วัน {start_day} ถึง {end_day})",
        "items":              items,
        "total_items":        len(items),
        "total_quantity":     round(sum(it["quantity"] for it in items), 2),
        "grand_total":        subtotal,
        "baht_text":          baht_wording,
        "month_str":          month_str,
        "month_label":        month_label,
        # Supplier details
        "vendor_name":         settings.get("vendor_name", "นางสาวทัศนีย์ สุรเกียรติ"),
        "vendor_address_line1":settings.get("vendor_address_line1", "หมู่บ้านปากตารอ 67/22 หมู่ที่ 11"),
        "vendor_address_line2":settings.get("vendor_address_line2", "ต.บานา อ.เมือง จ.ปัตตานี 94000"),
        "vendor_tel":          settings.get("vendor_tel", "085-1523003"),
        "vendor_tax_id":       settings.get("vendor_tax_id", "1949900022220"),
        "vendor_bank_acc":     settings.get("vendor_bank_acc", "907-0-34496-3"),
        "vendor_bank_name":    settings.get("vendor_bank_name", "น.ส.ทัศนีย์ สุรเกียรติ"),
        "vendor_bank_branch":  settings.get("vendor_bank_branch", "ธนาคาร กรุงไทย จำกัด สาขาปัตตานี"),
        # Government Department
        "dept_name":           settings.get("dept_name", "สถานสงเคราะห์เด็กชายนราธิวาส"),
        "dept_agency":         settings.get("dept_agency", "กรมกิจการเด็กและเยาวชน"),
        "dept_address_line1":  settings.get("dept_address_line1", "11 หมู่ 8 ตำบลลำภู อำเภอเมืองนราธิวาส"),
        "dept_address_line2":  settings.get("dept_address_line2", "จังหวัดนราธิวาส"),
        "dept_tel":            settings.get("dept_tel", "073-532054"),
        # Signatures
        "buyer_name":          settings.get("buyer_name", "นางสาวรานีซะ หะมิ"),
        "buyer_position":      settings.get("buyer_position", "หัวหน้าเจ้าหน้าที่"),
        "receiver_name":       settings.get("receiver_name", "นางสาวทัศนีย์ สุรเกียรติ"),
        "receiver_position":   settings.get("receiver_position", "เจ้าของร้าน"),
    }
    res.update(_extract_conditions(settings, to_thai_date(start_dt, False), res))
    return res


def get_bill_data_for_range(start_day, end_day, db_path=DB_PATH):
    """Retrieve consolidated bill data for a date range (live data from daily_purchases)."""
    conn = get_connection(db_path)
    cur = conn.cursor()

    cur.execute("""
        SELECT MIN(date) as min_date, MAX(date) as max_date, COUNT(DISTINCT day_number) as day_count
        FROM daily_metadata WHERE day_number BETWEEN ? AND ?
    """, (start_day, end_day))
    range_meta = dict(cur.fetchone())

    cur.execute("""
        SELECT 
            item_code, item_name, category, unit, benchmark_price,
            ROUND(SUM(quantity), 2) AS quantity,
            ROUND(SUM(subtotal_cost), 2) AS subtotal_cost,
            GROUP_CONCAT(DISTINCT notes) AS notes
        FROM daily_purchases
        WHERE day_number BETWEEN ? AND ?
        GROUP BY item_code
        ORDER BY item_code ASC
    """, (start_day, end_day))
    items = [dict(r) for r in cur.fetchall()]

    subtotal = round(sum(it["subtotal_cost"] for it in items), 2)
    baht_wording = baht_text(subtotal)

    start_dt = datetime.datetime.strptime(range_meta["min_date"], "%Y-%m-%d").date()
    end_dt = datetime.datetime.strptime(range_meta["max_date"], "%Y-%m-%d").date()

    conn.close()
    settings = get_bill_settings(db_path)
    bill_year = settings.get("bill_year", "2570")

    res = {
        "bill_number": f"{start_day}-{end_day}/{bill_year}",
        "day_number": f"{start_day}-{end_day}",
        "fiscal_year": settings.get("fiscal_year", "2569"),
        "service_date": start_dt,
        "service_date_thai": f"{to_thai_date(start_dt, True)} ถึง {to_thai_date(end_dt, True)}",
        "order_date_thai": to_thai_date(start_dt - datetime.timedelta(days=2), False),
        "deadline_date_thai": to_thai_date(start_dt, False),
        "theme": f"รายการรวม {range_meta['day_count']} วัน",
        "meal_notes": f"รายการจัดซื้อรวม วันที่ {start_day} ถึง {end_day}",
        "items": items,
        "total_items": len(items),
        "total_quantity": round(sum(it["quantity"] for it in items), 2),
        "grand_total": subtotal,
        "baht_text": baht_wording,
        # Supplier details (ผู้ขาย)
        "vendor_name": settings.get("vendor_name", "นางสาวทัศนีย์ สุรเกียรติ"),
        "vendor_address_line1": settings.get("vendor_address_line1", "หมู่บ้านปากตารอ 67/22 หมู่ที่ 11"),
        "vendor_address_line2": settings.get("vendor_address_line2", "ต.บานา อ.เมือง จ.ปัตตานี 94000"),
        "vendor_tel": settings.get("vendor_tel", "085-1523003"),
        "vendor_tax_id": settings.get("vendor_tax_id", "1949900022220"),
        "vendor_bank_acc": settings.get("vendor_bank_acc", "907-0-34496-3"),
        "vendor_bank_name": settings.get("vendor_bank_name", "น.ส.ทัศนีย์ สุรเกียรติ"),
        "vendor_bank_branch": settings.get("vendor_bank_branch", "ธนาคาร กรุงไทย จำกัด สาขาปัตตานี"),
        # Government Department details (ส่วนราชการ)
        "dept_name": settings.get("dept_name", "สถานสงเคราะห์เด็กชายนราธิวาส"),
        "dept_agency": settings.get("dept_agency", "กรมกิจการเด็กและเยาวชน"),
        "dept_address_line1": settings.get("dept_address_line1", "11 หมู่ 8 ตำบลลำภู อำเภอเมืองนราธิวาส"),
        "dept_address_line2": settings.get("dept_address_line2", "จังหวัดนราธิวาส"),
        "dept_tel": settings.get("dept_tel", "073-532054"),
        # Signatures (ผู้สั่งซื้อ & ผู้รับใบสั่งซื้อ)
        "buyer_name": settings.get("buyer_name", "นางสาวรานีซะ หะมิ"),
        "buyer_position": settings.get("buyer_position", "หัวหน้าเจ้าหน้าที่"),
        "receiver_name": settings.get("receiver_name", "นางสาวทัศนีย์ สุรเกียรติ"),
        "receiver_position": settings.get("receiver_position", "เจ้าของร้าน")
    }
    res.update(_extract_conditions(settings, to_thai_date(start_dt, False), res))
    return res

def render_printable_bill_html(bill_data):
    """
    Render exact matching HTML for the Thai government purchase order (ใบสั่งซื้อ).
    Uses 'Sarabun' / 'TH Sarabun New' font for printing.
    """
    items = bill_data["items"]

    month_str = bill_data.get("month_str")
    day_num_str = str(bill_data.get("day_number", ""))
    if month_str:
        export_url = f"/api/bill/month/export/xlsx?month={month_str}"
    elif "-" in day_num_str:
        parts = day_num_str.split("-")
        export_url = f"/api/bill/range/export/xlsx?start={parts[0]}&end={parts[1]}"
    else:
        export_url = f"/api/bill/export/xlsx?day={day_num_str}"
    
    meal_lines = [line.strip() for line in bill_data.get("meal_notes", "").split("\n") if line.strip()]

    # --- Pagination logic ---
    # Header/intro/footer on page 1 takes ~12 rows of space
    # Continuation pages have full table height
    ITEMS_PER_PAGE_1 = 15   # max rows on first page
    ITEMS_PER_CONT   = 22   # max rows on continuation pages

    def make_rows(chunk, start_idx=0):
        rows = []
        for i, it in enumerate(chunk):
            idx = start_idx + i
            no_str   = str(idx + 1)
            name_str = it["item_name"]
            qty_str  = f"{it['quantity']:.2f}"
            unit_str = it["unit"]
            price_str    = f"{it['benchmark_price']:.2f}"
            subtotal_str = f"{it['subtotal_cost']:.2f}"
            rows.append(f"""
        <tr>
          <td class="col-no">{no_str}</td>
          <td class="col-name">{name_str}</td>
          <td class="col-qty">{qty_str}</td>
          <td class="col-unit">{unit_str}</td>
          <td class="col-price">{price_str}</td>
          <td class="col-subtotal">{subtotal_str}</td>
          <td class="col-remark"></td>
        </tr>""")
        return "".join(rows)

    def make_table_header():
        return """<table class="bill-table">
    <thead>
      <tr>
        <th rowspan="2" class="col-no">ลำดับที่</th>
        <th rowspan="2" class="col-name">รายการ</th>
        <th colspan="2" style="width: 90px;">จำนวน</th>
        <th rowspan="2" class="col-price">ราคาต่อ<br>หน่วย</th>
        <th rowspan="2" class="col-subtotal">จำนวนเงิน<br>(บาท)</th>
        <th rowspan="2" class="col-remark">หมายเหตุ</th>
      </tr>
      <tr>
        <th style="width: 45px; font-size: 11pt;">จำนวน</th>
        <th style="width: 45px; font-size: 11pt;">หน่วย</th>
      </tr>
    </thead>"""

    # Split items into pages
    if not items:
        page_chunks = [[]]
    elif len(items) <= ITEMS_PER_PAGE_1:
        page_chunks = [items]
    else:
        page_chunks = [items[:ITEMS_PER_PAGE_1]]
        rem = items[ITEMS_PER_PAGE_1:]
        for i in range(0, len(rem), ITEMS_PER_CONT):
            page_chunks.append(rem[i:i + ITEMS_PER_CONT])

    # Build page 1 HTML block
    page1_rows_html = make_rows(page_chunks[0], start_idx=0)
    is_last_on_page1 = (len(page_chunks) == 1)
    foot_row = f"""
      <tr class="table-foot">
        <td colspan="2" style="text-align: center; font-weight: bold;">จำนวนเงิน</td>
        <td colspan="2" style="text-align: center; font-weight: bold; border-left: 1px solid #000; border-right: 1px solid #000;">
          {bill_data['baht_text']}
        </td>
        <td style="border-right: 1px solid #000;"></td>
        <td style="text-align: right; font-weight: bold;">{bill_data['grand_total']:,.2f}</td>
        <td></td>
      </tr>""" if is_last_on_page1 else ""

    page1_html = f"""<!-- PAGE 1: ใบสั่งซื้อ -->
<div class="page">
  <div class="header-title">ใบสั่งซื้อ</div>

  <div class="info-grid">
    <div class="info-left">
      <div><strong>ผู้ขาย</strong> {bill_data['vendor_name']}</div>
      <div>ที่อยู่ {bill_data['vendor_address_line1']}</div>
      <div>{bill_data['vendor_address_line2']}</div>
      <div>โทรศัพท์ {bill_data['vendor_tel']}</div>
      <div>เลขประจำตัวผู้เสียภาษี {bill_data['vendor_tax_id']}</div>
      <div>เลขที่บัญชีเงินฝากธนาคาร {bill_data['vendor_bank_acc']}</div>
      <div>ชื่อบัญชี {bill_data['vendor_bank_name']}</div>
      <div>ธนาคาร {bill_data['vendor_bank_branch']}</div>
    </div>
    <div class="info-right">
      <div><strong>ใบสั่งซื้อเลขที่</strong> <span id="bill-num-display" contenteditable="true" title="คลิกเพื่อแก้ไขเลขที่ใบสั่งซื้อได้โดยตรง" style="border-bottom: 1px dashed #2563eb; outline: none; padding: 0 4px; cursor: text;">{bill_data['bill_number']}</span></div>
      <div><strong>วันที่</strong> <span id="order-date-display" contenteditable="true" title="คลิกเพื่อแก้ไขวันที่สั่งซื้อได้โดยตรง" style="border-bottom: 1px dashed #2563eb; outline: none; padding: 0 4px; cursor: text;">{bill_data['order_date_thai']}</span></div>
      <div><strong>ส่วนราชการ</strong> {bill_data['dept_name']}</div>
      <div>{bill_data['dept_agency']}</div>
      <div>ที่อยู่ {bill_data['dept_address_line1']}</div>
      <div>{bill_data['dept_address_line2']}</div>
      <div>โทรศัพท์ {bill_data['dept_tel']}</div>
    </div>
  </div>

  <div class="intro-text">
    ตามที่ท่านได้เสนอราคาวัสดุอาหารและเครื่องอุปโภค บริโภค ปีงบประมาณ {bill_data['fiscal_year']} ให้นำวัสดุส่งมอบให้{bill_data['dept_name']} ตามรายการดังต่อไปนี้
  </div>

  {make_table_header()}
    <tbody>
      {page1_rows_html}
      {foot_row}
    </tbody>
  </table>

  <div class="foot-continue">/อยู่ภายใต้...</div>
</div>"""

    # Build continuation pages (if any items spill over)
    cont_pages_html = ""
    total_so_far = len(page_chunks[0])
    for pi, chunk in enumerate(page_chunks[1:], start=2):
        is_last = (pi == len(page_chunks))
        cont_rows = make_rows(chunk, start_idx=total_so_far)
        total_so_far += len(chunk)
        cont_foot = f"""
      <tr class="table-foot">
        <td colspan="2" style="text-align: center; font-weight: bold;">จำนวนเงิน</td>
        <td colspan="2" style="text-align: center; font-weight: bold; border-left: 1px solid #000; border-right: 1px solid #000;">
          {bill_data['baht_text']}
        </td>
        <td style="border-right: 1px solid #000;"></td>
        <td style="text-align: right; font-weight: bold;">{bill_data['grand_total']:,.2f}</td>
        <td></td>
      </tr>""" if is_last else ""
        cont_pages_html += f"""
<!-- ต่อหน้า {pi} -->
<div class="page page-break">
  <div class="page-num">- {pi} - (ต่อ)</div>
  <div style="font-size: 13pt; margin-bottom: 6px; color: #555;">ใบสั่งซื้อเลขที่ {bill_data['bill_number']} (ต่อ)</div>
  {make_table_header()}
    <tbody>
      {cont_rows}
      {cont_foot}
    </tbody>
  </table>
  <div class="foot-continue">/อยู่ภายใต้...</div>
</div>"""


    html = f"""<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<title>ใบสั่งซื้อเลขที่ {bill_data['bill_number']}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Sarabun:ital,wght@0,300;0,400;0,500;0,600;0,700;0,800;1,400;1,700&display=swap" rel="stylesheet">
<style>
  * {{
    box-sizing: border-box;
    margin: 0;
    padding: 0;
  }}
  body {{
    font-family: 'TH Sarabun New', 'Sarabun', 'THSarabunNew', sans-serif;
    font-size: 15pt;
    line-height: 1.25;
    color: #000000;
    background: #e5e7eb;
    padding: 20px 0;
  }}
  .no-print-bar {{
    width: 95%;
    max-width: 1300px;
    margin: 0 auto 18px auto;
    padding: 10px 18px;
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    box-shadow: 0 4px 14px rgba(15, 23, 42, 0.08);
    display: flex;
    justify-content: space-between;
    align-items: center;
    flex-wrap: wrap;
    gap: 12px;
    box-sizing: border-box;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", sans-serif;
  }}
  .bar-title {{
    font-size: 14.5px;
    font-weight: 700;
    color: #1e3a8a;
    display: flex;
    align-items: center;
    gap: 8px;
    flex-shrink: 0;
  }}
  .bar-actions {{
    display: flex;
    gap: 8px;
    align-items: center;
    flex-wrap: wrap;
  }}
  .btn {{
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 7px 14px;
    border-radius: 8px;
    font-size: 13.5px;
    font-weight: 600;
    cursor: pointer;
    border: 1px solid transparent;
    text-decoration: none;
    white-space: nowrap;
    transition: all 0.15s ease;
    font-family: inherit;
    box-sizing: border-box;
  }}
  .btn:active {{
    transform: scale(0.98);
  }}
  .ic {{
    width: 15px;
    height: 15px;
    stroke: currentColor;
    stroke-width: 2;
    stroke-linecap: round;
    stroke-linejoin: round;
    fill: none;
    vertical-align: middle;
    display: inline-block;
    flex-shrink: 0;
  }}
  .ic-lg {{
    width: 17px;
    height: 17px;
  }}
  .btn-print {{
    background: #1e40af;
    color: #ffffff;
    border-color: #1e40af;
    box-shadow: 0 2px 4px rgba(30, 64, 175, 0.25);
  }}
  .btn-print:hover {{
    background: #1e3a8a;
    border-color: #1e3a8a;
  }}
  .btn-excel {{
    background: #ffffff;
    color: #047857;
    border-color: #cbd5e1;
  }}
  .btn-excel:hover {{
    background: #f0fdf4;
    border-color: #86efac;
    color: #065f46;
  }}
  .btn-back {{
    background: #ffffff;
    color: #475569;
    border-color: #cbd5e1;
  }}
  .btn-back:hover {{
    background: #f1f5f9;
    color: #1e293b;
    border-color: #94a3b8;
  }}
  .btn-edit {{
    background: #0284c7;
    color: #ffffff;
    border-color: #0284c7;
    box-shadow: 0 2px 4px rgba(2, 132, 199, 0.2);
  }}
  .btn-edit:hover {{
    background: #0369a1;
    border-color: #0369a1;
  }}
  .bill-modal-overlay {{
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(0, 0, 0, 0.6);
    display: none;
    align-items: center;
    justify-content: center;
    z-index: 99999;
    backdrop-filter: blur(2px);
    font-family: 'Kanit', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  }}
  .bill-modal-overlay.visible {{
    display: flex;
  }}
  .bill-modal {{
    background: #ffffff;
    width: 780px;
    max-width: 95vw;
    max-height: 90vh;
    border-radius: 12px;
    box-shadow: 0 10px 25px rgba(0,0,0,0.25);
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }}
  .bill-modal-header {{
    padding: 14px 22px;
    background: #1e3a8a;
    color: #ffffff;
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-weight: 600;
    font-size: 16px;
  }}
  .bill-modal-close {{
    background: none;
    border: none;
    color: #ffffff;
    font-size: 24px;
    cursor: pointer;
    line-height: 1;
  }}
  .bill-modal-body {{
    padding: 18px 24px;
    overflow-y: auto;
    font-size: 13.5px;
    color: #1f2937;
  }}
  .bill-modal-footer {{
    padding: 12px 24px;
    background: #f9fafb;
    border-top: 1px solid #e5e7eb;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }}
  .form-sec-title {{
    font-size: 14px;
    font-weight: 700;
    color: #1e3a8a;
    margin: 14px 0 8px 0;
    padding-bottom: 4px;
    border-bottom: 2px solid #e2e8f0;
    display: flex;
    align-items: center;
    gap: 6px;
  }}
  .form-sec-title:first-child {{
    margin-top: 0;
  }}
  .field-row {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 12px;
    margin-bottom: 8px;
  }}
  .field-group {{
    margin-bottom: 6px;
  }}
  .field-group label {{
    display: block;
    font-size: 12.5px;
    font-weight: 600;
    color: #4b5563;
    margin-bottom: 3px;
  }}
  .field-group input {{
    width: 100%;
    box-sizing: border-box;
    padding: 7px 10px;
    border: 1px solid #d1d5db;
    border-radius: 6px;
    font-size: 13.5px;
    outline: none;
    font-family: inherit;
  }}
  .field-group input:focus {{
    border-color: #2563eb;
    box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.15);
  }}
  .field-group textarea {{
    width: 100%;
    box-sizing: border-box;
    padding: 7px 10px;
    border: 1px solid #d1d5db;
    border-radius: 6px;
    font-size: 13px;
    outline: none;
    font-family: inherit;
    resize: vertical;
    line-height: 1.4;
  }}
  .field-group textarea:focus {{
    border-color: #2563eb;
    box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.15);
  }}

  /* Modern Beautiful Popup Modal */
  .sw-popup-backdrop {{
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(15, 23, 42, 0.65);
    backdrop-filter: blur(5px);
    -webkit-backdrop-filter: blur(5px);
    z-index: 100000;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 16px;
    opacity: 0;
    visibility: hidden;
    transition: opacity 0.2s ease, visibility 0.2s ease;
  }}
  .sw-popup-backdrop.active {{
    opacity: 1;
    visibility: visible;
  }}
  .sw-popup-box {{
    background: #ffffff;
    border-radius: 20px;
    box-shadow: 0 25px 50px -12px rgba(15, 23, 42, 0.25), 0 0 0 1px rgba(226, 232, 240, 0.9);
    width: 100%;
    max-width: 420px;
    padding: 28px 24px 22px 24px;
    transform: scale(0.9) translateY(12px);
    transition: transform 0.25s cubic-bezier(0.16, 1, 0.3, 1), opacity 0.2s ease;
    text-align: center;
    box-sizing: border-box;
    font-family: 'Kanit', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  }}
  .sw-popup-backdrop.active .sw-popup-box {{
    transform: scale(1) translateY(0);
  }}
  .sw-popup-icon {{
    width: 64px;
    height: 64px;
    border-radius: 50%;
    margin: 0 auto 16px auto;
    display: flex;
    align-items: center;
    justify-content: center;
    position: relative;
    animation: swPopBounce 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
  }}
  @keyframes swPopBounce {{
    0% {{ transform: scale(0.5); opacity: 0; }}
    100% {{ transform: scale(1); opacity: 1; }}
  }}
  .sw-popup-icon.success {{ background: #dcfce7; color: #16a34a; }}
  .sw-popup-icon.error   {{ background: #fee2e2; color: #dc2626; }}
  .sw-popup-icon.warning {{ background: #fef3c7; color: #d97706; }}
  .sw-popup-icon.info    {{ background: #dbeafe; color: #2563eb; }}
  .sw-popup-icon svg {{
    width: 34px;
    height: 34px;
    stroke: currentColor;
    stroke-width: 2.2;
    fill: none;
    stroke-linecap: round;
    stroke-linejoin: round;
  }}
  .sw-popup-title {{
    font-size: 19px;
    font-weight: 700;
    color: #0f172a;
    margin-bottom: 8px;
    line-height: 1.35;
  }}
  .sw-popup-desc {{
    font-size: 14.5px;
    color: #475569;
    line-height: 1.6;
    margin-bottom: 22px;
    white-space: pre-line;
    word-break: break-word;
  }}
  .sw-popup-btn-row {{
    display: flex;
    gap: 10px;
    justify-content: center;
  }}
  .sw-popup-btn {{
    flex: 1;
    padding: 10px 18px;
    border-radius: 10px;
    font-size: 14.5px;
    font-weight: 600;
    cursor: pointer;
    border: none;
    transition: all 0.15s ease;
    font-family: inherit;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    gap: 6px;
    outline: none;
  }}
  .sw-popup-btn:active {{
    transform: scale(0.98);
  }}
  .sw-popup-btn-primary {{
    background: #2563eb;
    color: #ffffff;
    box-shadow: 0 4px 10px rgba(37, 99, 235, 0.25);
  }}
  .sw-popup-btn-primary:hover {{
    background: #1d4ed8;
    box-shadow: 0 6px 14px rgba(37, 99, 235, 0.35);
  }}
  .sw-popup-btn-danger {{
    background: #dc2626;
    color: #ffffff;
    box-shadow: 0 4px 10px rgba(220, 38, 38, 0.25);
  }}
  .sw-popup-btn-danger:hover {{
    background: #b91c1c;
    box-shadow: 0 6px 14px rgba(220, 38, 38, 0.35);
  }}
  .sw-popup-btn-secondary {{
    background: #f1f5f9;
    color: #475569;
    border: 1px solid #cbd5e1;
  }}
  .sw-popup-btn-secondary:hover {{
    background: #e2e8f0;
    color: #1e293b;
  }}
  .page {{
    width: 210mm;
    min-height: 297mm;
    margin: 0 auto 20px auto;
    background: #ffffff;
    box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    padding: 15mm 15mm 15mm 15mm;
    position: relative;
  }}
  .header-title {{
    text-align: center;
    font-size: 22pt;
    font-weight: bold;
    margin-bottom: 12px;
  }}
  .info-grid {{
    display: flex;
    justify-content: space-between;
    margin-bottom: 10px;
    font-size: 14pt;
    line-height: 1.35;
  }}
  .info-left {{
    width: 48%;
  }}
  .info-right {{
    width: 50%;
  }}
  .intro-text {{
    font-size: 14pt;
    line-height: 1.4;
    text-indent: 2.5cm;
    margin: 10px 0 8px 0;
    text-align: justify;
  }}
  
  /* Official Government Table Styling - Clean Aesthetic (No Underlines within table) */
  table.bill-table {{
    width: 100%;
    border-collapse: collapse;
    margin-top: 4px;
    font-size: 13.5pt;
  }}
  table.bill-table th {{
    border: 1px solid #000000;
    padding: 3px 4px;
    vertical-align: middle;
    font-weight: bold;
    text-align: center;
    background: #ffffff;
  }}
  table.bill-table tbody td {{
    border-top: none !important;
    border-bottom: none !important;
    border-left: 1px solid #000000;
    border-right: 1px solid #000000;
    padding: 4px 6px;
    vertical-align: middle;
  }}
  .col-no {{ width: 50px; text-align: center; }}
  .col-name {{ text-align: left; padding-left: 6px !important; }}
  .col-qty {{ width: 45px; text-align: right; border-right: none !important; }}
  .col-unit {{ width: 45px; text-align: center; border-left: none !important; }}
  .col-price {{ width: 68px; text-align: right; }}
  .col-subtotal {{ width: 85px; text-align: right; }}
  .col-remark {{ width: 175px; text-align: left; font-size: 12pt; padding-left: 6px !important; }}
  
  table.bill-table tr.table-foot td {{
    border-top: 1px solid #000000 !important;
    border-bottom: 1px solid #000000 !important;
    border-left: 1px solid #000000;
    border-right: 1px solid #000000;
    font-weight: bold;
    vertical-align: middle;
    padding: 6px 4px;
  }}
  .foot-continue {{
    position: absolute;
    bottom: 12mm;
    right: 15mm;
    font-size: 14pt;
  }}
  
  /* Page 2 Styling */
  .page-num {{
    text-align: center;
    font-size: 15pt;
    margin-bottom: 15px;
  }}
  .conditions-title {{
    font-size: 15pt;
    font-weight: normal;
    margin-bottom: 6px;
  }}
  .condition-list {{
    list-style: none;
    font-size: 14pt;
    line-height: 1.35;
  }}
  .condition-list li {{
    margin-bottom: 6px;
    text-align: justify;
    padding-left: 1.8em;
    text-indent: -1.8em;
  }}
  .cond-edit-span {{
    outline: none;
    cursor: text;
    border-radius: 3px;
    padding: 0 2px;
    transition: background 0.15s, box-shadow 0.15s;
    border-bottom: 1px dashed rgba(37, 99, 235, 0.4);
  }}
  .cond-edit-span:hover {{
    background: #eff6ff;
    box-shadow: 0 0 0 1px #93c5fd;
    border-bottom-color: #2563eb;
  }}
  .cond-edit-span:focus {{
    background: #ffffff;
    box-shadow: 0 0 0 2px #2563eb;
    border-bottom-color: transparent;
  }}
  .signatures-wrap {{
    margin-top: 40px;
    width: 100%;
  }}
  .sig-table {{
    width: 100%;
    border: none !important;
  }}
  .sig-table td {{
    border: none !important;
    padding: 0;
  }}
  
  /* Print Optimization */
  @media print {{
    body {{
      background: #ffffff;
      padding: 0;
    }}
    .no-print-bar, .no-print {{
      display: none !important;
    }}
    .cond-edit-span {{
      border-bottom: none !important;
      background: transparent !important;
      box-shadow: none !important;
    }}
    .page {{
      box-shadow: none;
      margin: 0;
      width: 100%;
      min-height: auto;
      padding: 10mm 15mm;
    }}
    .page-break {{
      page-break-before: always;
      break-before: page;
    }}
    @page {{
      size: A4 portrait;
      margin: 0;
    }}
    table.bill-table tbody td {{
      border-top: none !important;
      border-bottom: none !important;
    }}
    .col-remark * {{
      display: none !important;
    }}
    #bill-num-display, #order-date-display {{
      border-bottom: none !important;
    }}
  }}
</style>
</head>
<body>

<!-- SVG Icons Sprite -->
<svg xmlns="http://www.w3.org/2000/svg" style="display:none">
  <symbol id="ic-edit" viewBox="0 0 24 24"><path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/><path d="m15 5 4 4"/></symbol>
  <symbol id="ic-printer" viewBox="0 0 24 24"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></symbol>
  <symbol id="ic-chart" viewBox="0 0 24 24"><rect x="18" y="3" width="4" height="18" rx="1"/><rect x="10" y="8" width="4" height="13" rx="1"/><rect x="2" y="13" width="4" height="8" rx="1"/></symbol>
  <symbol id="ic-arrow-left" viewBox="0 0 24 24"><path d="M19 12H5M12 19l-7-7 7-7"/></symbol>
  <symbol id="ic-save" viewBox="0 0 24 24"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></symbol>
  <symbol id="ic-building" viewBox="0 0 24 24"><rect width="16" height="20" x="4" y="2" rx="2" ry="2"/><path d="M9 22v-4h6v4M8 6h.01M16 6h.01M12 6h.01M12 10h.01M12 14h.01M16 10h.01M16 14h.01M8 10h.01M8 14h.01"/></symbol>
  <symbol id="ic-landmark" viewBox="0 0 24 24"><line x1="3" x2="21" y1="22" y2="22"/><line x1="6" x2="6" y1="18"/><line x1="10" x2="10" y1="18"/><line x1="14" x2="14" y1="18"/><line x1="18" x2="18" y1="18"/><polygon points="12 2 20 7 4 7"/></symbol>
  <symbol id="ic-pen-tool" viewBox="0 0 24 24"><path d="M12 19l7-7 3 3-7 7-3-3z"/><path d="M18 13l-1.5-7.5L2 2l3.5 14.5L13 18l5-5z"/><path d="M2 2l7.586 7.586"/><circle cx="11" cy="11" r="2"/></symbol>
  <symbol id="ic-x" viewBox="0 0 24 24"><path d="M18 6 6 18M6 6l12 12"/></symbol>
</svg>

<div class="no-print-bar">
  <div class="bar-title">
    <svg class="ic" style="color: #2563eb;"><use href="#ic-printer"/></svg>
    <span>ใบสั่งซื้อทางการ (ราชการ) — {bill_data['bill_number']}</span>
  </div>
  <div class="bar-actions">
    <button onclick="openBillSettingsModal()" class="btn btn-edit">
      <svg class="ic"><use href="#ic-edit"/></svg> แก้ไขข้อมูลใบสั่งซื้อ / ผู้ลงนาม
    </button>
    <button onclick="window.print()" class="btn btn-print">
      <svg class="ic"><use href="#ic-printer"/></svg> พิมพ์ใบสั่งซื้อ (Print / Save PDF)
    </button>
    <a href="{export_url}" class="btn btn-excel">
      <svg class="ic"><use href="#ic-chart"/></svg> Export Excel
    </a>
    <button onclick="window.close(); if(history.length>1)history.back();" class="btn btn-back">
      <svg class="ic"><use href="#ic-arrow-left"/></svg> กลับสู่ระบบ
    </button>
  </div>
</div>

{page1_html}
{cont_pages_html}

<!-- PAGE 2 (เงื่อนไขและลายมือชื่อ): Always last page -->
<div class="page page-break">
  <div class="page-num">- {len(page_chunks) + 1} -</div>

  <div class="conditions-title" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
    <span>อยู่ภายใต้เงื่อนไขต่อไปนี้</span>
    <button onclick="openBillSettingsModal('sec-conditions')" class="btn btn-edit no-print" style="padding: 3px 10px; font-size: 13px;">
      <svg class="ic"><use href="#ic-edit"/></svg> แก้ไขเงื่อนไข
    </button>
  </div>

  <ul class="condition-list">
    <li>1. กำหนดส่งมอบภายใน <strong><span id="cond-span-delivery-days" class="cond-edit-span" contenteditable="true" title="คลิกเพื่อแก้ไขจำนวนวัน">{bill_data['cond_delivery_days']}</span></strong> วัน นับถัดจากวันที่ผู้ขายได้รับใบสั่งซื้อ</li>
    <li>2. ครบกำหนดส่งมอบวันที่ <span id="cond-span-deadline" class="cond-edit-span" contenteditable="true" title="คลิกเพื่อแก้ไขวันที่ครบกำหนดส่งมอบ" style="color: red; font-weight: bold;">{bill_data['cond_deadline_text']}</span></li>
    <li>3. สถานที่ส่งมอบ <span id="cond-span-location" class="cond-edit-span" contenteditable="true" title="คลิกเพื่อแก้ไขสถานที่">{bill_data['cond_location']}</span></li>
    <li>4. ระยะเวลารับประกัน <span id="cond-span-warranty" class="cond-edit-span" contenteditable="true" title="คลิกเพื่อแก้ไขระยะเวลารับประกัน">{bill_data['cond_warranty']}</span></li>
    <li>5. สงวนสิทธิ์ค่าปรับกรณีส่งมอบเกินกำหนด โดยคิดค่าปรับเป็นรายวันในอัตราร้อยละ <span id="cond-span-rate" class="cond-edit-span" contenteditable="true" title="คลิกเพื่อแก้ไขอัตราค่าปรับ">{bill_data['cond_penalty_rate']}</span> ของสิ่งของที่ยังไม่ได้รับมอบ แต่ต้องไม่ต่ำกว่าวันละ <span id="cond-span-min" class="cond-edit-span" contenteditable="true" title="คลิกเพื่อแก้ไขค่าปรับขั้นต่ำ">{bill_data['cond_penalty_min']}</span> บาท</li>
    <li>6. <span id="cond-span-6" class="cond-edit-span" contenteditable="true" title="คลิกเพื่อแก้ไขข้อความ">{bill_data['cond_item_6']}</span></li>
    <li>7. <span id="cond-span-7" class="cond-edit-span" contenteditable="true" title="คลิกเพื่อแก้ไขข้อความ">{bill_data['cond_item_7']}</span></li>
    <li>8. <span id="cond-span-8" class="cond-edit-span" contenteditable="true" title="คลิกเพื่อแก้ไขข้อความ">{bill_data['cond_item_8']}</span></li>
  </ul>

  <div class="signatures-wrap">
    <table class="sig-table">
      <tr>
        <td style="width: 48%;"></td>
        <td style="width: 52%; text-align: left;">
          <div style="font-size: 14pt;">(ลงชื่อ)...................................................... ผู้สั่งซื้อ</div>
          <div style="text-align: center; width: 230px; font-size: 14pt; margin-top: 5px;">({bill_data['buyer_name']})</div>
          <div style="text-align: center; width: 230px; font-size: 14pt; margin-top: 2px;">{bill_data['buyer_position']}</div>
        </td>
      </tr>
      <tr>
        <td style="height: 40px;"></td>
        <td></td>
      </tr>
      <tr>
        <td></td>
        <td style="text-align: left;">
          <div style="font-size: 14pt;">(ลงชื่อ)...................................................... ผู้รับใบสั่งซื้อ</div>
          <div style="text-align: center; width: 230px; font-size: 14pt; margin-top: 5px;">({bill_data['receiver_name']})</div>
          <div style="text-align: center; width: 230px; font-size: 14pt; margin-top: 2px;">{bill_data['receiver_position']}</div>
        </td>
      </tr>
    </table>
  </div>
</div>

<!-- ป๊อปอัปแก้ไขข้อมูลใบสั่งซื้อ / ส่วนราชการ / ผู้ลงนาม -->
<div class="bill-modal-overlay" id="bill-settings-modal">
  <div class="bill-modal">
    <div class="bill-modal-header">
      <div style="display: flex; align-items: center; gap: 8px;">
        <svg class="ic ic-lg"><use href="#ic-edit"/></svg> แก้ไขข้อมูลใบสั่งซื้อ / ส่วนราชการ / ผู้ลงนาม
      </div>
      <button class="bill-modal-close" onclick="closeBillSettingsModal()">&times;</button>
    </div>
    <div class="bill-modal-body">
      <!-- ส่วนที่ 1: ข้อมูลผู้ขาย (ซ้ายบน) -->
      <div class="form-sec-title">
        <svg class="ic" style="color: #2563eb;"><use href="#ic-building"/></svg> ข้อมูลผู้ขาย (ด้านซ้ายบนของใบสั่งซื้อ)
      </div>
      <div class="field-row">
        <div class="field-group">
          <label>ชื่อผู้ขาย / ร้านค้า</label>
          <input type="text" id="cfg-vendor_name" value="{bill_data['vendor_name']}">
        </div>
        <div class="field-group">
          <label>เบอร์โทรศัพท์ผู้ขาย</label>
          <input type="text" id="cfg-vendor_tel" value="{bill_data['vendor_tel']}">
        </div>
      </div>
      <div class="field-row">
        <div class="field-group">
          <label>ที่อยู่ผู้ขาย บรรทัด 1</label>
          <input type="text" id="cfg-vendor_address_line1" value="{bill_data['vendor_address_line1']}">
        </div>
        <div class="field-group">
          <label>ที่อยู่ผู้ขาย บรรทัด 2</label>
          <input type="text" id="cfg-vendor_address_line2" value="{bill_data['vendor_address_line2']}">
        </div>
      </div>
      <div class="field-row">
        <div class="field-group">
          <label>เลขประจำตัวผู้เสียภาษี</label>
          <input type="text" id="cfg-vendor_tax_id" value="{bill_data['vendor_tax_id']}">
        </div>
        <div class="field-group">
          <label>เลขที่บัญชีเงินฝากธนาคาร</label>
          <input type="text" id="cfg-vendor_bank_acc" value="{bill_data['vendor_bank_acc']}">
        </div>
      </div>
      <div class="field-row">
        <div class="field-group">
          <label>ชื่อบัญชีเงินฝาก</label>
          <input type="text" id="cfg-vendor_bank_name" value="{bill_data['vendor_bank_name']}">
        </div>
        <div class="field-group">
          <label>ธนาคาร และ สาขา</label>
          <input type="text" id="cfg-vendor_bank_branch" value="{bill_data['vendor_bank_branch']}">
        </div>
      </div>

      <!-- ส่วนที่ 2: ข้อมูลส่วนราชการ & ใบสั่งซื้อ (ขวาบน) -->
      <div class="form-sec-title">
        <svg class="ic" style="color: #2563eb;"><use href="#ic-landmark"/></svg> ข้อมูลส่วนราชการและใบสั่งซื้อ (ด้านขวาบน)
      </div>
      <div class="field-row" style="background: #eff6ff; padding: 10px; border-radius: 8px; border: 1px solid #bfdbfe; margin-bottom: 10px;">
        <div class="field-group">
          <label style="color: #1e40af; font-weight: bold;">ใบสั่งซื้อเลขที่ (สำหรับวันนี้)</label>
          <input type="text" id="cfg-bill_number" value="{bill_data['bill_number']}" style="border: 2px solid #3b82f6; font-weight: bold;">
        </div>
        <div class="field-group">
          <label style="color: #1e40af; font-weight: bold;">วันที่สั่งซื้อ (สำหรับวันนี้ เช่น {bill_data['order_date_thai']})</label>
          <input type="text" id="cfg-order_date_thai" value="{bill_data['order_date_thai']}" style="border: 2px solid #3b82f6; font-weight: bold;">
        </div>
      </div>
      <div class="field-row">
        <div class="field-group">
          <label>ชื่อส่วนราชการ</label>
          <input type="text" id="cfg-dept_name" value="{bill_data['dept_name']}">
        </div>
        <div class="field-group">
          <label>สังกัด / กรม</label>
          <input type="text" id="cfg-dept_agency" value="{bill_data['dept_agency']}">
        </div>
      </div>
      <div class="field-row">
        <div class="field-group">
          <label>ที่อยู่ส่วนราชการ บรรทัด 1</label>
          <input type="text" id="cfg-dept_address_line1" value="{bill_data['dept_address_line1']}">
        </div>
        <div class="field-group">
          <label>ที่อยู่ส่วนราชการ บรรทัด 2</label>
          <input type="text" id="cfg-dept_address_line2" value="{bill_data['dept_address_line2']}">
        </div>
      </div>
      <div class="field-row">
        <div class="field-group">
          <label>เบอร์โทรศัพท์ส่วนราชการ</label>
          <input type="text" id="cfg-dept_tel" value="{bill_data['dept_tel']}">
        </div>
        <div class="field-group">
          <label>ปีงบประมาณ (เช่น 2569)</label>
          <input type="text" id="cfg-fiscal_year" value="{bill_data['fiscal_year']}">
        </div>
      </div>

      <!-- ส่วนที่ 3: ผู้ลงนาม (หน้า 2) -->
      <div class="form-sec-title">
        <svg class="ic" style="color: #2563eb;"><use href="#ic-pen-tool"/></svg> ผู้ลงนามท้ายใบสั่งซื้อ (หน้าที่ 2)
      </div>
      <div class="field-row">
        <div class="field-group" style="background: #eff6ff; padding: 10px; border-radius: 8px;">
          <label style="color: #1e40af; font-weight: bold; margin-bottom: 6px;">ผู้สั่งซื้อ (Person placing the order)</label>
          <div style="margin-bottom: 6px;">
            <label style="font-size: 11.5px;">ชื่อ-นามสกุล ผู้สั่งซื้อ</label>
            <input type="text" id="cfg-buyer_name" value="{bill_data['buyer_name']}">
          </div>
          <div>
            <label style="font-size: 11.5px;">ตำแหน่ง ผู้สั่งซื้อ</label>
            <input type="text" id="cfg-buyer_position" value="{bill_data['buyer_position']}">
          </div>
        </div>
        <div class="field-group" style="background: #f0fdf4; padding: 10px; border-radius: 8px;">
          <label style="color: #166534; font-weight: bold; margin-bottom: 6px;">ผู้รับใบสั่งซื้อ (Recipient of the order)</label>
          <div style="margin-bottom: 6px;">
            <label style="font-size: 11.5px;">ชื่อ-นามสกุล ผู้รับใบสั่งซื้อ</label>
            <input type="text" id="cfg-receiver_name" value="{bill_data['receiver_name']}">
          </div>
          <div>
            <label style="font-size: 11.5px;">ตำแหน่ง ผู้รับใบสั่งซื้อ</label>
            <input type="text" id="cfg-receiver_position" value="{bill_data['receiver_position']}">
          </div>
        </div>
      </div>

      <!-- ส่วนที่ 4: เงื่อนไขการสั่งซื้อ (หน้า 2) -->
      <div class="form-sec-title" id="sec-conditions">
        <svg class="ic" style="color: #2563eb;"><use href="#ic-edit"/></svg> เงื่อนไขการสั่งซื้อ (หน้าที่ 2)
      </div>
      <div class="field-row">
        <div class="field-group">
          <label>1. กำหนดส่งมอบภายใน (วัน)</label>
          <input type="text" id="cfg-cond_delivery_days" value="{bill_data['cond_delivery_days']}">
        </div>
        <div class="field-group">
          <label>2. ครบกำหนดส่งมอบวันที่</label>
          <input type="text" id="cfg-cond_deadline_text" value="{bill_data['cond_deadline_text']}">
        </div>
      </div>
      <div class="field-row">
        <div class="field-group">
          <label>3. สถานที่ส่งมอบ</label>
          <input type="text" id="cfg-cond_location" value="{bill_data['cond_location']}">
        </div>
        <div class="field-group">
          <label>4. ระยะเวลารับประกัน</label>
          <input type="text" id="cfg-cond_warranty" value="{bill_data['cond_warranty']}">
        </div>
      </div>
      <div class="field-row">
        <div class="field-group">
          <label>5. ค่าปรับรายวัน (ร้อยละ)</label>
          <input type="text" id="cfg-cond_penalty_rate" value="{bill_data['cond_penalty_rate']}">
        </div>
        <div class="field-group">
          <label>5. ค่าปรับขั้นต่ำ (บาท/วัน)</label>
          <input type="text" id="cfg-cond_penalty_min" value="{bill_data['cond_penalty_min']}">
        </div>
      </div>
      <div class="field-group">
        <label>6. เงื่อนไขข้อ 6 (การไม่รับมอบ/เปลี่ยนสินค้า)</label>
        <textarea id="cfg-cond_item_6" rows="2">{bill_data['cond_item_6']}</textarea>
      </div>
      <div class="field-group">
        <label>7. เงื่อนไขข้อ 7 (การจ้างช่วง/ค่าปรับจ้างช่วง)</label>
        <textarea id="cfg-cond_item_7" rows="3">{bill_data['cond_item_7']}</textarea>
      </div>
      <div class="field-group">
        <label>8. เงื่อนไขข้อ 8 (การประเมินผลการปฏิบัติงาน)</label>
        <textarea id="cfg-cond_item_8" rows="2">{bill_data['cond_item_8']}</textarea>
      </div>
    </div>
    <div class="bill-modal-footer">
      <div style="font-size: 12px; color: #6b7280;">
        * ข้อมูลที่แก้ไขจะถูกบันทึกและแสดงผลในใบสั่งซื้อรวมถึงไฟล์ส่งออก Excel
      </div>
      <div style="display: flex; gap: 8px;">
        <button class="btn btn-back" onclick="closeBillSettingsModal()">ยกเลิก</button>
        <button class="btn btn-print" onclick="saveBillSettings()">
          <svg class="ic"><use href="#ic-save"/></svg> บันทึกข้อมูล
        </button>
      </div>
    </div>
  </div>
</div>

<script>
function openBillSettingsModal(secId) {{
  const modal = document.getElementById('bill-settings-modal');
  modal.classList.add('visible');
  if (secId) {{
    setTimeout(() => {{
      const target = document.getElementById(secId);
      if (target) {{
        target.scrollIntoView({{ behavior: 'smooth', block: 'start' }});
      }}
    }}, 120);
  }}
}}
function closeBillSettingsModal() {{
  document.getElementById('bill-settings-modal').classList.remove('visible');
}}
async function saveBillSettings() {{
  const dayNum = "{bill_data['day_number']}";
  const billNum = document.getElementById('cfg-bill_number').value.trim();
  const orderDateThai = document.getElementById('cfg-order_date_thai').value.trim();
  const condDeadline = document.getElementById('cfg-cond_deadline_text') ? document.getElementById('cfg-cond_deadline_text').value.trim() : '';

  const payload = {{
    [`bill_number_day_${{dayNum}}`]: billNum,
    [`order_date_day_${{dayNum}}`]: orderDateThai,
    [`cond_deadline_day_${{dayNum}}`]: condDeadline,
    cond_deadline_text: condDeadline,
    cond_delivery_days: document.getElementById('cfg-cond_delivery_days') ? document.getElementById('cfg-cond_delivery_days').value.trim() : '',
    cond_location: document.getElementById('cfg-cond_location') ? document.getElementById('cfg-cond_location').value.trim() : '',
    cond_warranty: document.getElementById('cfg-cond_warranty') ? document.getElementById('cfg-cond_warranty').value.trim() : '',
    cond_penalty_rate: document.getElementById('cfg-cond_penalty_rate') ? document.getElementById('cfg-cond_penalty_rate').value.trim() : '',
    cond_penalty_min: document.getElementById('cfg-cond_penalty_min') ? document.getElementById('cfg-cond_penalty_min').value.trim() : '',
    cond_item_6: document.getElementById('cfg-cond_item_6') ? document.getElementById('cfg-cond_item_6').value.trim() : '',
    cond_item_7: document.getElementById('cfg-cond_item_7') ? document.getElementById('cfg-cond_item_7').value.trim() : '',
    cond_item_8: document.getElementById('cfg-cond_item_8') ? document.getElementById('cfg-cond_item_8').value.trim() : '',
    vendor_name: document.getElementById('cfg-vendor_name').value.trim(),
    vendor_tel: document.getElementById('cfg-vendor_tel').value.trim(),
    vendor_address_line1: document.getElementById('cfg-vendor_address_line1').value.trim(),
    vendor_address_line2: document.getElementById('cfg-vendor_address_line2').value.trim(),
    vendor_tax_id: document.getElementById('cfg-vendor_tax_id').value.trim(),
    vendor_bank_acc: document.getElementById('cfg-vendor_bank_acc').value.trim(),
    vendor_bank_name: document.getElementById('cfg-vendor_bank_name').value.trim(),
    vendor_bank_branch: document.getElementById('cfg-vendor_bank_branch').value.trim(),
    dept_name: document.getElementById('cfg-dept_name').value.trim(),
    dept_agency: document.getElementById('cfg-dept_agency').value.trim(),
    dept_address_line1: document.getElementById('cfg-dept_address_line1').value.trim(),
    dept_address_line2: document.getElementById('cfg-dept_address_line2').value.trim(),
    dept_tel: document.getElementById('cfg-dept_tel').value.trim(),
    fiscal_year: document.getElementById('cfg-fiscal_year').value.trim(),
    buyer_name: document.getElementById('cfg-buyer_name').value.trim(),
    buyer_position: document.getElementById('cfg-buyer_position').value.trim(),
    receiver_name: document.getElementById('cfg-receiver_name').value.trim(),
    receiver_position: document.getElementById('cfg-receiver_position').value.trim()
  }};

  try {{
    const res = await fetch('/api/bill/settings', {{
      method: 'POST',
      headers: {{ 'Content-Type': 'application/json' }},
      body: JSON.stringify(payload)
    }});
    const data = await res.json();
    if (data.success) {{
      await showPopupAlert({{
        title: 'บันทึกข้อมูลสำเร็จ',
        message: 'บันทึกข้อมูลใบสั่งซื้อ ส่วนราชการ เงื่อนไข และผู้ลงนามเรียบร้อยแล้ว',
        type: 'success',
        btnText: 'ตกลง'
      }});
      window.location.reload();
    }} else {{
      await showPopupAlert({{
        title: 'เกิดข้อผิดพลาด',
        message: data.error || 'ไม่สามารถบันทึกข้อมูลได้',
        type: 'error',
        btnText: 'ปิด'
      }});
    }}
  }} catch (e) {{
    await showPopupAlert({{
      title: 'ข้อผิดพลาดการเชื่อมต่อ',
      message: String(e),
      type: 'error',
      btnText: 'ปิด'
    }});
  }}
}}

// Modern Beautiful Popup Alert & Confirm System
function showPopupAlert({{ title = 'แจ้งเตือน', message = '', type = 'info', btnText = 'ตกลง' }}) {{
  return new Promise((resolve) => {{
    let container = document.getElementById('sw-popup-container');
    if (!container) {{
      container = document.createElement('div');
      container.id = 'sw-popup-container';
      document.body.appendChild(container);
    }}

    const iconSvg = type === 'success'
      ? '<polyline points="20 6 9 17 4 12"/>'
      : type === 'error'
      ? '<circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>'
      : type === 'warning'
      ? '<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>'
      : '<circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/>';

    container.innerHTML = `
      <div class="sw-popup-backdrop" id="sw-popup-overlay">
        <div class="sw-popup-box">
          <div class="sw-popup-icon ${{type}}">
            <svg viewBox="0 0 24 24">${{iconSvg}}</svg>
          </div>
          <div class="sw-popup-title">${{title}}</div>
          <div class="sw-popup-desc">${{message}}</div>
          <div class="sw-popup-btn-row">
            <button class="sw-popup-btn sw-popup-btn-primary" id="sw-popup-ok-btn">${{btnText}}</button>
          </div>
        </div>
      </div>
    `;

    const overlay = document.getElementById('sw-popup-overlay');
    requestAnimationFrame(() => {{
      if (overlay) overlay.classList.add('active');
    }});

    let finished = false;
    const finish = () => {{
      if (finished) return;
      finished = true;
      document.removeEventListener('keydown', keyHandler);
      if (overlay) overlay.classList.remove('active');
      setTimeout(() => {{ container.innerHTML = ''; resolve(true); }}, 200);
    }};

    const keyHandler = (e) => {{
      if (e.key === 'Escape' || e.key === 'Enter') finish();
    }};
    document.addEventListener('keydown', keyHandler);

    setTimeout(() => {{
      const btn = document.getElementById('sw-popup-ok-btn');
      if (btn) {{
        btn.focus();
        btn.onclick = finish;
      }}
      if (overlay) {{
        overlay.onclick = (e) => {{
          if (e.target === overlay) finish();
        }};
      }}
    }}, 20);
  }});
}}

function showPopupConfirm({{ title = 'ยืนยันการทำรายการ', message = '', type = 'warning', confirmText = 'ยืนยัน', cancelText = 'ยกเลิก', isDanger = true }}) {{
  return new Promise((resolve) => {{
    let container = document.getElementById('sw-popup-container');
    if (!container) {{
      container = document.createElement('div');
      container.id = 'sw-popup-container';
      document.body.appendChild(container);
    }}

    const iconSvg = isDanger || type === 'error'
      ? '<circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>'
      : '<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>';

    const btnClass = isDanger ? 'sw-popup-btn-danger' : 'sw-popup-btn-primary';

    container.innerHTML = `
      <div class="sw-popup-backdrop" id="sw-popup-overlay">
        <div class="sw-popup-box">
          <div class="sw-popup-icon ${{isDanger ? 'error' : type}}">
            <svg viewBox="0 0 24 24">${{iconSvg}}</svg>
          </div>
          <div class="sw-popup-title">${{title}}</div>
          <div class="sw-popup-desc">${{message}}</div>
          <div class="sw-popup-btn-row">
            <button class="sw-popup-btn sw-popup-btn-secondary" id="sw-popup-cancel-btn">${{cancelText}}</button>
            <button class="sw-popup-btn ${{btnClass}}" id="sw-popup-confirm-btn">${{confirmText}}</button>
          </div>
        </div>
      </div>
    `;

    const overlay = document.getElementById('sw-popup-overlay');
    requestAnimationFrame(() => {{
      if (overlay) overlay.classList.add('active');
    }});

    let finished = false;
    const finish = (result) => {{
      if (finished) return;
      finished = true;
      document.removeEventListener('keydown', keyHandler);
      if (overlay) overlay.classList.remove('active');
      setTimeout(() => {{ container.innerHTML = ''; resolve(result); }}, 200);
    }};

    const keyHandler = (e) => {{
      if (e.key === 'Escape') finish(false);
      else if (e.key === 'Enter') finish(true);
    }};
    document.addEventListener('keydown', keyHandler);

    setTimeout(() => {{
      const confBtn = document.getElementById('sw-popup-confirm-btn');
      const cancBtn = document.getElementById('sw-popup-cancel-btn');
      if (confBtn) {{
        confBtn.focus();
        confBtn.onclick = () => finish(true);
      }}
      if (cancBtn) cancBtn.onclick = () => finish(false);
      if (overlay) {{
        overlay.onclick = (e) => {{
          if (e.target === overlay) finish(false);
        }};
      }}
    }}, 20);
  }});
}}

// Override window.alert & window.confirm for full system consistency
window.alert = function(msg) {{
  const isErr = typeof msg === 'string' && (msg.includes('ผิดพลาด') || msg.includes('ไม่พบ') || msg.includes('กรุณา'));
  return showPopupAlert({{
    title: isErr ? 'แจ้งเตือน' : 'สำเร็จ',
    message: String(msg),
    type: isErr ? 'warning' : 'success'
  }});
}};

window.confirm = function(msg) {{
  return showPopupConfirm({{
    title: 'ยืนยันการทำรายการ',
    message: String(msg),
    type: 'warning',
    confirmText: 'ยืนยัน',
    cancelText: 'ยกเลิก',
    isDanger: true
  }});
}};

// Auto-save inline edits on the page
const currentDayNum = "{bill_data['day_number']}";
const billNumEl = document.getElementById('bill-num-display');
const orderDateEl = document.getElementById('order-date-display');

if (billNumEl) {{
  billNumEl.addEventListener('blur', async function() {{
    const val = this.innerText.trim();
    const cfgEl = document.getElementById('cfg-bill_number');
    if (cfgEl) cfgEl.value = val;
    await fetch('/api/bill/settings', {{
      method: 'POST',
      headers: {{ 'Content-Type': 'application/json' }},
      body: JSON.stringify({{ [`bill_number_day_${{currentDayNum}}`]: val }})
    }});
  }});
}}

if (orderDateEl) {{
  orderDateEl.addEventListener('blur', async function() {{
    const val = this.innerText.trim();
    const cfgEl = document.getElementById('cfg-order_date_thai');
    if (cfgEl) cfgEl.value = val;
    await fetch('/api/bill/settings', {{
      method: 'POST',
      headers: {{ 'Content-Type': 'application/json' }},
      body: JSON.stringify({{ [`order_date_day_${{currentDayNum}}`]: val }})
    }});
  }});
}}

// Auto-save for condition editable spans
const conditionSpanMap = [
  {{ id: 'cond-span-delivery-days', key: 'cond_delivery_days', inputId: 'cfg-cond_delivery_days' }},
  {{ id: 'cond-span-deadline', key: 'cond_deadline_text', dayKey: `cond_deadline_day_${{currentDayNum}}`, inputId: 'cfg-cond_deadline_text' }},
  {{ id: 'cond-span-location', key: 'cond_location', inputId: 'cfg-cond_location' }},
  {{ id: 'cond-span-warranty', key: 'cond_warranty', inputId: 'cfg-cond_warranty' }},
  {{ id: 'cond-span-rate', key: 'cond_penalty_rate', inputId: 'cfg-cond_penalty_rate' }},
  {{ id: 'cond-span-min', key: 'cond_penalty_min', inputId: 'cfg-cond_penalty_min' }},
  {{ id: 'cond-span-6', key: 'cond_item_6', inputId: 'cfg-cond_item_6' }},
  {{ id: 'cond-span-7', key: 'cond_item_7', inputId: 'cfg-cond_item_7' }},
  {{ id: 'cond-span-8', key: 'cond_item_8', inputId: 'cfg-cond_item_8' }}
];

conditionSpanMap.forEach(function(item) {{
  const spanEl = document.getElementById(item.id);
  if (!spanEl) return;
  spanEl.addEventListener('blur', async function() {{
    const val = this.innerText.trim();
    const inputEl = document.getElementById(item.inputId);
    if (inputEl) inputEl.value = val;
    const body = {{ [item.key]: val }};
    if (item.dayKey) body[item.dayKey] = val;
    try {{
      await fetch('/api/bill/settings', {{
        method: 'POST',
        headers: {{ 'Content-Type': 'application/json' }},
        body: JSON.stringify(body)
      }});
    }} catch (e) {{
      console.error('Failed to auto-save condition span:', e);
    }}
  }});
}});
</script>

</body>
</html>"""
    return html

def export_bill_to_xlsx(bill_data, output_path):
    """Export the official purchase bill directly to an Excel .xlsx workbook."""
    headers = ["ลำดับที่", "รหัส", "รายการ", "หมวดหมู่", "จำนวน", "หน่วย", "ราคาต่อหน่วย (฿)", "จำนวนเงิน (฿)", "หมายเหตุ"]
    rows = [headers]

    for idx, it in enumerate(bill_data["items"], start=1):
        rows.append([
            idx,
            it["item_code"],
            it["item_name"],
            it["category"],
            float(it["quantity"]),
            it["unit"],
            float(it["benchmark_price"]),
            float(it["subtotal_cost"]),
            it.get("notes", "")
        ])

    rows.append([])
    rows.append(["", "", "", "", "", "", "รวมทั้งสิ้น:", float(bill_data["grand_total"])])
    rows.append(["", "", "", "", "", "", "จำนวนเงินตัวอักษร:", bill_data["baht_text"]])

    summary_rows = [
        ["หัวข้อ", "รายละเอียด"],
        ["เลขที่ใบสั่งซื้อ", bill_data["bill_number"]],
        ["วันที่สั่งซื้อ", bill_data["order_date_thai"]],
        ["กำหนดส่งมอบ", bill_data["deadline_date_thai"]],
        ["วันให้บริการ", bill_data["service_date_thai"]],
        ["ส่วนราชการ", bill_data["dept_name"]],
        ["ผู้ขาย", bill_data["vendor_name"]],
        ["ผู้สั่งซื้อ", f"{bill_data['buyer_name']} ({bill_data['buyer_position']})"],
        ["ผู้รับใบสั่งซื้อ", f"{bill_data['receiver_name']} ({bill_data['receiver_position']})"],
        ["ยอดรวมทั้งสิ้น (฿)", bill_data["grand_total"]],
        ["จำนวนเงินตัวอักษร", bill_data["baht_text"]]
    ]

    write_xlsx(output_path, {
        "รายการจัดซื้อ": rows,
        "ข้อมูลใบสั่งซื้อ": summary_rows
    })
    return output_path

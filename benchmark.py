"""
Benchmark Prices Manager.
Stores fixed benchmark prices in Excel, supporting Thai government procurement standards.
"""

import os
from xlsx_engine import write_xlsx, read_xlsx

DEFAULT_EXCEL_PATH = os.path.join(os.path.dirname(__file__), "benchmark_prices.xlsx")

# Comprehensive catalog including authentic Thai institutional ingredients from government PO
INITIAL_BENCHMARK_ITEMS = [
    # รายการหลักจากเอกสารราชการ (ใบสั่งซื้อสถานสงเคราะห์เด็กชายนราธิวาส)
    {"code": "TH-001", "name": "อกไก่ (หั่น)", "category": "เนื้อสัตว์", "unit": "กก.", "price": 85.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 3, "notes": "สด สะอาด ลอกหนังหั่นชิ้น"},
    {"code": "TH-002", "name": "ซอสแดง", "category": "เครื่องปรุงรส", "unit": "ขวด", "price": 20.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 180, "notes": "สำหรับข้าวผัดซอสแดง"},
    {"code": "TH-003", "name": "ผักคะน้า", "category": "ผักสด", "unit": "กก.", "price": 42.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 4, "notes": "คะน้าต้นสด ไม่ช้ำ"},
    {"code": "TH-004", "name": "โครงไก่", "category": "เนื้อสัตว์", "unit": "กก.", "price": 47.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 3, "notes": "สำหรับต้มน้ำซุปแกงจืด"},
    {"code": "TH-005", "name": "ต้นหอม", "category": "ผักสด", "unit": "ขีด", "price": 11.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 5, "notes": "ต้นหอมสดคัดสวย"},
    {"code": "TH-006", "name": "ผักกวางตุ้ง", "category": "ผักสด", "unit": "กก.", "price": 36.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 4, "notes": "กวางตุ้งฮ่องเต้/ดอก"},
    {"code": "TH-007", "name": "ผงช็อกโกแลตมอลต์", "category": "เครื่องดื่ม", "unit": "ถุง", "price": 124.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 365, "notes": "ขนาด 1,000 กรัม สำหรับมื้อเช้า"},
    {"code": "TH-008", "name": "นมข้นหวาน", "category": "ผลิตภัณฑ์นม", "unit": "กระป๋อง", "price": 26.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 180, "notes": "สูตรเข้มข้น"},
    {"code": "TH-009", "name": "น้ำตาลทราย", "category": "เครื่องปรุงรส", "unit": "กก.", "price": 28.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 365, "notes": "น้ำตาลทรายขาวบริสุทธิ์"},
    {"code": "TH-010", "name": "ไก่เชียง", "category": "เนื้อสัตว์แปรรูป", "unit": "กก.", "price": 162.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 90, "notes": "ฮาลาล เกรดพรีเมียม"},
    {"code": "TH-011", "name": "หอมหัวใหญ่", "category": "ผักสด", "unit": "กก.", "price": 36.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 0.5, "shelf_life": 21, "notes": "หัวใหญ่แน่น แห้ง"},
    {"code": "TH-012", "name": "มะเขือเทศ", "category": "ผักสด", "unit": "กก.", "price": 52.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 0.5, "shelf_life": 7, "notes": "มะเขือเทศทวยผลสด"},
    {"code": "TH-013", "name": "มะนาว", "category": "ผักสด", "unit": "กก.", "price": 59.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 0.5, "shelf_life": 14, "notes": "มะนาวแป้นน้ำเยอะ"},
    {"code": "TH-014", "name": "หอยแครง", "category": "อาหารทะเล", "unit": "กก.", "price": 190.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 2, "notes": "สด สำหรับแกงคั่ว"},
    {"code": "TH-015", "name": "น้ำกะทิ", "category": "เครื่องปรุงรส", "unit": "กก.", "price": 85.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 2, "notes": "กะทิสดคั้นแท้ 100%"},
    {"code": "TH-016", "name": "ยอดผักต่างๆ (ใบชะพลู/ผักชี)", "category": "ผักสด", "unit": "มัด", "price": 9.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 4, "notes": "ใบชะพลูสด มัดใหญ่"},
    {"code": "TH-017", "name": "ไข่ไก่ เบอร์ 2", "category": "ไข่", "unit": "ฟอง", "price": 4.50, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 10.0, "shelf_life": 25, "notes": "ไข่ไก่สด ฟองละ 4.50 บาท"},
    {"code": "TH-018", "name": "น้ำมันพืช", "category": "เครื่องปรุงรส", "unit": "ขวด", "price": 56.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 365, "notes": "น้ำมันปาล์ม/ถั่วเหลือง 1 ลิตร"},
    {"code": "TH-019", "name": "มะละกอสุก", "category": "ผลไม้", "unit": "กก.", "price": 46.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 5, "notes": "ผลสุกหวาน ผลไม้หลังอาหาร"},

    # วัตถุดิบเพิ่มเติมสำหรับหมุนเวียน 365 วัน
    {"code": "TH-020", "name": "ข้าวสารหอมมะลิ", "category": "ข้าวและแป้ง", "unit": "กก.", "price": 38.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 15.0, "shelf_life": 365, "notes": "ข้าวเสาไห้/หอมมะลิแท้"},
    {"code": "TH-021", "name": "เนื้อหมูสะโพก", "category": "เนื้อสัตว์", "unit": "กก.", "price": 145.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 2.0, "shelf_life": 3, "notes": "สด ตัดแต่งไม่ติดมัน"},
    {"code": "TH-022", "name": "หมูบด", "category": "เนื้อสัตว์", "unit": "กก.", "price": 135.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 2.0, "shelf_life": 3, "notes": "หมูบดเนื้อล้วน"},
    {"code": "TH-023", "name": "ปลากะพงขาว", "category": "อาหารทะเล", "unit": "กก.", "price": 180.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 2.0, "shelf_life": 2, "notes": "ปลาสด ขอดเกล็ดควักไส้"},
    {"code": "TH-024", "name": "กุ้งขาวแว่น", "category": "อาหารทะเล", "unit": "กก.", "price": 220.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 2.0, "shelf_life": 2, "notes": "กุ้งสด ขนาด 50 ตัว/กก."},
    {"code": "TH-025", "name": "ปลาหมึกกล้วย", "category": "อาหารทะเล", "unit": "กก.", "price": 210.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 2, "notes": "ปลาหมึกสด"},
    {"code": "TH-026", "name": "ฟักเขียว", "category": "ผักสด", "unit": "กก.", "price": 25.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 2.0, "shelf_life": 14, "notes": "ฟักแก่ สำหรับต้มซุป"},
    {"code": "TH-027", "name": "กะหล่ำปลี", "category": "ผักสด", "unit": "กก.", "price": 32.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 2.0, "shelf_life": 10, "notes": "หัวแน่น สด"},
    {"code": "TH-028", "name": "แครอท", "category": "ผักสด", "unit": "กก.", "price": 38.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 14, "notes": "แครอทหัวสวย"},
    {"code": "TH-029", "name": "แตงกวา", "category": "ผักสด", "unit": "กก.", "price": 28.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 2.0, "shelf_life": 6, "notes": "ผลสด กรอบ"},
    {"code": "TH-030", "name": "กระเทียมกลีบ", "category": "ผักสด", "unit": "กก.", "price": 85.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 0.5, "shelf_life": 60, "notes": "กระเทียมไทยคัดเม็ด"},
    {"code": "TH-031", "name": "พริกขี้หนูสวน", "category": "ผักสด", "unit": "ขีด", "price": 15.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 7, "notes": "พริกสดเผ็ดหอม"},
    {"code": "TH-032", "name": "น้ำปลาแท้", "category": "เครื่องปรุงรส", "unit": "ขวด", "price": 35.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 365, "notes": "ขวด 700 มล."},
    {"code": "TH-033", "name": "ซีอิ๊วขาว", "category": "เครื่องปรุงรส", "unit": "ขวด", "price": 42.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 365, "notes": "สูตร 1"},
    {"code": "TH-034", "name": "ซอสหอยนางรม", "category": "เครื่องปรุงรส", "unit": "ขวด", "price": 48.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 365, "notes": "ขวดใหญ่ 800 กรัม"},
    {"code": "TH-035", "name": "กล้วยน้ำว้า", "category": "ผลไม้", "unit": "หวี", "price": 30.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 2.0, "shelf_life": 5, "notes": "หวีละ 14-16 ลูก ห่ามสุก"},
    {"code": "TH-036", "name": "แตงโม", "category": "ผลไม้", "unit": "กก.", "price": 28.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 5.0, "shelf_life": 7, "notes": "ผลสด หวานฉ่ำ"},
    {"code": "TH-037", "name": "สับปะรด", "category": "ผลไม้", "unit": "ลูก", "price": 35.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 2.0, "shelf_life": 6, "notes": "ลูกใหญ่ หวานฉ่ำ"},
    {"code": "TH-038", "name": "ขนมปังแถว", "category": "ข้าวและแป้ง", "unit": "แถว", "price": 42.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 2.0, "shelf_life": 4, "notes": "ขนมปังปอนด์ตัดหนา"},
    {"code": "TH-039", "name": "แยมสับปะรด/ส้ม", "category": "ของหวาน", "unit": "กระปุก", "price": 45.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 1.0, "shelf_life": 180, "notes": "สำหรับทาขนมปังมื้อเช้า"},
    {"code": "TH-040", "name": "เต้าหู้ไข่ไก่", "category": "ผลิตภัณฑ์ถั่วเหลือง", "unit": "หลอด", "price": 12.00, "supplier": "ร้านทัศนีย์ สุรเกียรติ", "min_qty": 4.0, "shelf_life": 30, "notes": "สำหรับแกงจืดเต้าหู้หมูสับ"}
]

BENCHMARK_HEADERS = [
    "Item Code",
    "Item Name",
    "Category",
    "Unit",
    "Benchmark Price (฿)",
    "Preferred Supplier",
    "Min Order Qty",
    "Shelf Life (Days)",
    "Notes"
]

def init_benchmark_excel(filepath=DEFAULT_EXCEL_PATH, overwrite=False):
    """Create the benchmark prices Excel file with initial data if it doesn't already exist."""
    if os.path.exists(filepath) and not overwrite:
        return filepath

    rows = [BENCHMARK_HEADERS]
    for item in INITIAL_BENCHMARK_ITEMS:
        rows.append([
            item["code"],
            item["name"],
            item["category"],
            item["unit"],
            item["price"],
            item["supplier"],
            item["min_qty"],
            item["shelf_life"],
            item["notes"]
        ])

    write_xlsx(filepath, {"Benchmark Prices": rows})
    return filepath

def load_benchmark_prices_from_excel(filepath=DEFAULT_EXCEL_PATH):
    """
    Read benchmark prices from the Excel file and return a list of dictionaries.
    """
    if not os.path.exists(filepath):
        init_benchmark_excel(filepath)

    raw_rows = read_xlsx(filepath)
    if not raw_rows or len(raw_rows) < 2:
        return []

    headers = [str(c).strip().lower() if c is not None else "" for c in raw_rows[0]]
    
    col_map = {}
    for idx, h in enumerate(headers):
        if "code" in h:
            col_map["code"] = idx
        elif "name" in h:
            col_map["name"] = idx
        elif "category" in h:
            col_map["category"] = idx
        elif "unit" in h:
            col_map["unit"] = idx
        elif "price" in h or "cost" in h or "benchmark" in h or "บาท" in h or "฿" in h:
            col_map["price"] = idx
        elif "supplier" in h or "vendor" in h:
            col_map["supplier"] = idx
        elif "min" in h or "moq" in h:
            col_map["min_qty"] = idx
        elif "life" in h or "shelf" in h:
            col_map["shelf_life"] = idx
        elif "note" in h or "desc" in h:
            col_map["notes"] = idx

    items = []
    for row_idx, r in enumerate(raw_rows[1:], start=2):
        if not r or all(cell is None or cell == "" for cell in r):
            continue

        def get_val(key, default=""):
            idx = col_map.get(key)
            if idx is not None and idx < len(r) and r[idx] is not None:
                return r[idx]
            return default

        code = str(get_val("code", f"ITEM-{row_idx:03d}")).strip()
        name = str(get_val("name", "Unnamed Item")).strip()
        category = str(get_val("category", "ทั่วไป")).strip()
        unit = str(get_val("unit", "กก.")).strip()
        
        try:
            raw_p = get_val("price", 0.0)
            price = float(raw_p) if raw_p is not None else 0.0
        except (ValueError, TypeError):
            price = 0.0

        supplier = str(get_val("supplier", "ร้านทัศนีย์ สุรเกียรติ")).strip()

        try:
            raw_min = get_val("min_qty", 1.0)
            min_qty = float(raw_min) if raw_min is not None else 1.0
        except (ValueError, TypeError):
            min_qty = 1.0

        try:
            raw_life = get_val("shelf_life", 7)
            shelf_life = int(float(raw_life)) if raw_life is not None else 7
        except (ValueError, TypeError):
            shelf_life = 7

        notes = str(get_val("notes", "")).strip()

        items.append({
            "code": code,
            "name": name,
            "category": category,
            "unit": unit,
            "price": round(price, 2),
            "supplier": supplier,
            "min_qty": min_qty,
            "shelf_life": shelf_life,
            "notes": notes
        })

    return items

def save_benchmark_prices_to_excel(items, filepath=DEFAULT_EXCEL_PATH):
    """Write an updated list of benchmark items back to the Excel file."""
    rows = [BENCHMARK_HEADERS]
    for item in items:
        rows.append([
            item.get("code", ""),
            item.get("name", ""),
            item.get("category", "ทั่วไป"),
            item.get("unit", "กก."),
            round(float(item.get("price", 0.0)), 2),
            item.get("supplier", "ร้านทัศนีย์ สุรเกียรติ"),
            float(item.get("min_qty", 1.0)),
            int(item.get("shelf_life", 7)),
            item.get("notes", "")
        ])

    write_xlsx(filepath, {"Benchmark Prices": rows})
    return filepath

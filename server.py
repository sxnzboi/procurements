"""
HTTP Web Server for the Daily Procurement & Bill System (ภาษาไทย & ฟอนต์ Kanit).
Pure Python (http.server) — ไม่ต้องติดตั้ง dependency ภายนอก
ให้บริการ API endpoints และหน้า Dashboard ภาษาไทยเต็มรูปแบบ พร้อมระบบแก้ไขราคากลาง Excel
"""

import json
import os
import re
import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

from database import (
    init_db, init_empty_365_days, sync_benchmark_from_excel, get_all_benchmark_items,
    get_benchmark_item, update_benchmark_item, create_benchmark_item,
    delete_benchmark_item, delete_daily_purchase_item,
    get_daily_purchase_spec, update_daily_purchases,
    get_365_days_summary, get_bill_settings, update_bill_settings,
    clear_month_purchases, clear_day_purchases, update_daily_metadata, reset_all_data_365, get_connection, DB_PATH
)
from benchmark import init_benchmark_excel, DEFAULT_EXCEL_PATH
from scheduler import generate_365_days_schedule, apply_template_to_days, PROCUREMENT_TEMPLATES
from bill_generator import (
    get_bill_data_for_day, get_bill_data_for_range, get_bill_data_for_month,
    render_printable_bill_html, export_bill_to_xlsx
)

PORT = int(os.environ.get("PORT", 8080))

class ProcurementHandler(BaseHTTPRequestHandler):
    """ตัวจัดการ Request สำหรับ API และหน้าเว็บทั้งหมด"""

    def log_message(self, format, *args):
        pass

    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False, default=str).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.end_headers()
        self.wfile.write(body)

    def _send_html(self, html, status=200):
        body = html.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        self.end_headers()
        self.wfile.write(body)

    def _send_file(self, filepath, content_type, filename=None):
        if not os.path.exists(filepath):
            self._send_json({"error": "ไม่พบไฟล์"}, 404)
            return
        with open(filepath, "rb") as f:
            data = f.read()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        if filename:
            import urllib.parse
            encoded_fn = urllib.parse.quote(filename)
            self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{encoded_fn}")
        self.end_headers()
        self.wfile.write(data)

    def _read_body(self):
        length = int(self.headers.get("Content-Length", 0))
        if length > 0:
            raw = self.rfile.read(length)
            content_type = self.headers.get("Content-Type", "")
            if "multipart/form-data" in content_type:
                return {"_multipart_raw": raw, "_content_type": content_type}
            return json.loads(raw.decode("utf-8"))
        return {}

    def _parse_multipart(self, raw_bytes, content_type):
        """Parse multipart/form-data and return dict of {field_name: bytes}"""
        import re
        boundary_match = re.search(r'boundary=(?:"([^"]+)"|([^\s;]+))', content_type)
        if not boundary_match:
            return {}
        boundary = (boundary_match.group(1) or boundary_match.group(2)).encode()
        parts = {}
        delimiter = b'--' + boundary
        raw_parts = raw_bytes.split(delimiter)
        for part in raw_parts[1:]:
            if part.strip() in (b'', b'--', b'--\r\n'):
                continue
            if b'\r\n\r\n' not in part:
                continue
            header_bytes, _, body_bytes = part.partition(b'\r\n\r\n')
            if body_bytes.endswith(b'\r\n'):
                body_bytes = body_bytes[:-2]
            elif body_bytes.endswith(b'\n'):
                body_bytes = body_bytes[:-1]
            headers_text = header_bytes.decode('utf-8', errors='replace')
            name_match = re.search(r'name="([^"]+)"', headers_text)
            if name_match:
                field_name = name_match.group(1)
                parts[field_name] = body_bytes
        return parts

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path.rstrip("/") or "/"
        qs = parse_qs(parsed.query)

        # หน้าเว็บหลัก
        if path == "/":
            self._serve_dashboard()
        elif path == "/bill":
            self._serve_bill_page(qs)

        # API แคตตาล็อกราคากลาง
        elif path == "/api/benchmark":
            items = get_all_benchmark_items()
            self._send_json({"items": items, "count": len(items)})

        # API ดาวน์โหลดไฟล์ Excel ราคากลางต้นฉบับ
        elif path == "/api/benchmark/download":
            self._send_file(DEFAULT_EXCEL_PATH,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "benchmark_prices.xlsx")

        elif path.startswith("/api/benchmark/"):
            code = path.split("/")[-1]
            item = get_benchmark_item(code)
            if item:
                self._send_json(item)
            else:
                self._send_json({"error": "ไม่พบรายการวัตถุดิบ"}, 404)

        # API รายการจัดซื้อประจำวัน
        elif path.startswith("/api/day/") and path.split("/")[-1].isdigit():
            try:
                day_num = int(path.split("/")[-1])
                spec = get_daily_purchase_spec(day_num)
                if spec:
                    self._send_json(spec)
                else:
                    self._send_json({"error": f"ไม่พบข้อมูลวันที่ {day_num}"}, 404)
            except ValueError:
                self._send_json({"error": "หมายเลขวันไม่ถูกต้อง"}, 400)

        # API สรุปภาพรวม 365 วัน
        elif path == "/api/summary":
            summary = get_365_days_summary()
            self._send_json(summary)

        # API ชุดเทมเพลตเมนู
        elif path == "/api/templates":
            tpls = {}
            for k, v in PROCUREMENT_TEMPLATES.items():
                tpls[k] = {
                    "name": v["name"],
                    "description": v["description"],
                    "theme": v["theme"],
                    "item_count": len(v["items"])
                }
            self._send_json({"templates": tpls})

        # API พิมพ์ใบสั่งซื้อ HTML
        elif path == "/api/bill/html":
            day = int(qs.get("day", [1])[0])
            bill = get_bill_data_for_day(day)
            if bill:
                html = render_printable_bill_html(bill)
                self._send_html(html)
            else:
                self._send_html("<h1>ไม่พบข้อมูลวันที่ระบุ</h1>", 404)

        elif path == "/api/bill/range/html":
            start = int(qs.get("start", [1])[0])
            end = int(qs.get("end", [7])[0])
            bill = get_bill_data_for_range(start, end)
            html = render_printable_bill_html(bill)
            self._send_html(html)

        # API ส่งออกใบสั่งซื้อ Excel (.xlsx)
        elif path == "/api/bill/export/xlsx":
            day = int(qs.get("day", [1])[0])
            bill = get_bill_data_for_day(day)
            if bill:
                out = f"/tmp/bill_day_{day}.xlsx"
                export_bill_to_xlsx(bill, out)
                self._send_file(out, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                f"ใบสั่งซื้อ_วันที่_{day}.xlsx")
            else:
                self._send_json({"error": "ไม่พบข้อมูลวันที่ระบุ"}, 404)

        elif path == "/api/bill/range/export/xlsx":
            start = int(qs.get("start", [1])[0])
            end = int(qs.get("end", [7])[0])
            bill = get_bill_data_for_range(start, end)
            out = f"/tmp/bill_range_{start}_{end}.xlsx"
            export_bill_to_xlsx(bill, out)
            self._send_file(out, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            f"ใบสั่งซื้อ_วันที่_{start}_ถึง_{end}.xlsx")

        # ใบสั่งซื้อรายเดือน (HTML + XLSX)
        elif path == "/api/bill/month/html":
            month = qs.get("month", ["2026-01"])[0]
            bill = get_bill_data_for_month(month)
            if bill:
                html = render_printable_bill_html(bill)
                self._send_html(html)
            else:
                self._send_html("<h2>ไม่พบข้อมูลในเดือนที่เลือก หรือยังไม่มีรายการจัดซื้อในเดือนนี้</h2>", 404)

        elif path == "/api/bill/month/export/xlsx":
            month = qs.get("month", ["2026-01"])[0]
            bill = get_bill_data_for_month(month)
            if bill:
                safe_month = month.replace("-", "_")
                out = f"/tmp/bill_month_{safe_month}.xlsx"
                export_bill_to_xlsx(bill, out)
                self._send_file(out, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                f"ใบสั่งซื้อ_เดือน_{month}.xlsx")
            else:
                self._send_json({"error": "ไม่พบข้อมูลในเดือนที่เลือก"}, 404)

        # API การตั้งค่าข้อมูลหัวใบสั่งซื้อและผู้ลงนาม
        elif path == "/api/bill/settings":
            settings = get_bill_settings()
            self._send_json(settings)

        else:
            self._send_json({"error": "ไม่พบหน้านี้"}, 404)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path.rstrip("/")

        content_type = self.headers.get("Content-Type", "")
        length = int(self.headers.get("Content-Length", 0))

        # Handle multipart file upload separately (not json)
        if "multipart/form-data" in content_type and path == "/api/benchmark/upload":
            raw = self.rfile.read(length) if length > 0 else b""
            parts = self._parse_multipart(raw, content_type)
            file_bytes = parts.get("file", b"")

            if not file_bytes:
                self._send_json({"error": "ไม่พบไฟล์ที่อัปโหลด"}, 400)
                return
            try:
                import tempfile, shutil
                with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                    tmp.write(file_bytes)
                    tmp_path = tmp.name
                # Validate — try to load items from the uploaded file
                from benchmark import load_benchmark_prices_from_excel, save_benchmark_prices_to_excel
                items = load_benchmark_prices_from_excel(tmp_path)
                if not items:
                    os.unlink(tmp_path)
                    self._send_json({"error": "ไฟล์ Excel ไม่มีข้อมูลราคากลาง หรือรูปแบบไม่ถูกต้อง"}, 400)
                    return
                # Replace the main benchmark Excel file
                shutil.copy2(tmp_path, DEFAULT_EXCEL_PATH)
                os.unlink(tmp_path)
                # Re-sync DB from new Excel
                count = sync_benchmark_from_excel()
                self._send_json({"success": True, "items_loaded": len(items), "items_synced": count})
            except Exception as e:
                self._send_json({"error": f"เกิดข้อผิดพลาดในการอ่านไฟล์: {str(e)}"}, 400)
            return

        body = {}
        if length > 0 and "application/json" in content_type:
            try:
                body = json.loads(self.rfile.read(length).decode("utf-8"))
            except Exception:
                body = {}
        elif length > 0 and "multipart" not in content_type:
            try:
                body = json.loads(self.rfile.read(length).decode("utf-8"))
            except Exception:
                body = {}

        # แก้ไขข้อมูลวัตถุดิบราคากลาง (บันทึกและซิงค์ลง Excel อัตโนมัติ)
        if path == "/api/benchmark/update":
            code = body.get("code")
            if not code:
                self._send_json({"error": "กรุณาระบุรหัสวัตถุดิบ (code)"}, 400)
                return
            try:
                updated = update_benchmark_item(code, body)
                self._send_json({"success": True, "item": updated})
            except Exception as e:
                self._send_json({"error": str(e)}, 400)

        # เพิ่มวัตถุดิบใหม่เข้าราคากลาง (บันทึกและซิงค์ลง Excel)
        elif path == "/api/benchmark/create":
            try:
                created = create_benchmark_item(body)
                self._send_json({"success": True, "item": created})
            except Exception as e:
                self._send_json({"error": str(e)}, 400)

        # ลบรายการวัตถุดิบออกจากแคตตาล็อกราคากลาง (ลบและซิงค์ลง Excel อัตโนมัติ)
        elif path == "/api/benchmark/delete":
            code = body.get("code")
            if not code:
                self._send_json({"error": "กรุณาระบุรหัสวัตถุดิบ (code)"}, 400)
                return
            try:
                delete_benchmark_item(code)
                self._send_json({"success": True, "code": code})
            except Exception as e:
                self._send_json({"error": str(e)}, 400)

        # ลบรายการวัตถุดิบจัดซื้อประจำวัน (ลบรายบรรทัด)
        elif path == "/api/day/item/delete":
            purchase_id = body.get("id")
            if not purchase_id:
                self._send_json({"error": "กรุณาระบุ id ของรายการที่ต้องการลบ"}, 400)
                return
            try:
                updated_day = delete_daily_purchase_item(int(purchase_id))
                self._send_json({"success": True, "day": updated_day})
            except Exception as e:
                self._send_json({"error": str(e)}, 400)

        # บันทึกการตั้งค่าข้อมูลใบสั่งซื้อ (ผู้ขาย, ส่วนราชการ, ผู้สั่งซื้อ, ผู้รับใบสั่งซื้อ)
        elif path == "/api/bill/settings":
            try:
                updated = update_bill_settings(body)
                self._send_json({"success": True, "settings": updated})
            except Exception as e:
                self._send_json({"error": str(e)}, 400)

        # ปรับปรุงรายการจัดซื้อของวันใดวันหนึ่ง
        elif path.startswith("/api/day/") and path.split("/")[-1].isdigit():
            try:
                day_num = int(path.split("/")[-1])
                items_data = body.get("items", [])
                theme = body.get("theme")
                notes = body.get("notes")
                result = update_daily_purchases(day_num, items_data, theme=theme, notes=notes)
                self._send_json({"success": True, "day": result})
            except Exception as e:
                self._send_json({"error": str(e)}, 400)

        # ใช้งานเทมเพลตกับวันที่เลือก
        elif path == "/api/template/apply":
            tpl_key = body.get("template_key")
            day_numbers = body.get("day_numbers", [])
            replace = body.get("replace", False)
            if not tpl_key or not day_numbers:
                self._send_json({"error": "กรุณาระบุ template_key และ day_numbers"}, 400)
                return
            try:
                count = apply_template_to_days(tpl_key, day_numbers, replace=replace)
                self._send_json({"success": True, "days_updated": count})
            except ValueError as e:
                self._send_json({"error": str(e)}, 400)

        # รีเซ็ตข้อมูลทั้งหมด 365 วัน (สำหรับเริ่มปีใหม่)
        elif path == "/api/reset/all":
            try:
                res = reset_all_data_365()
                self._send_json({"success": True, "result": res})
            except Exception as e:
                self._send_json({"error": str(e)}, 500)

        # สร้างตาราง 365 วันใหม่ทั้งหมด
        elif path == "/api/schedule/regenerate":
            start = body.get("start_date", "2026-01-01")
            count = generate_365_days_schedule(start_date=start)
            self._send_json({"success": True, "days_generated": count})

        # ซิงค์ราคากลางล่าสุดจากไฟล์ Excel
        elif path == "/api/benchmark/sync":
            count = sync_benchmark_from_excel()
            self._send_json({"success": True, "items_synced": count})

        # ล้างรายการจัดซื้อทั้งเดือน
        elif path == "/api/month/clear":
            month = body.get("month")
            clear_theme = body.get("clear_theme", False)
            if not month:
                self._send_json({"error": "กรุณาระบุเดือน (month) เช่น 2026-01"}, 400)
                return
            try:
                deleted = clear_month_purchases(month, clear_theme=clear_theme)
                self._send_json({"success": True, "deleted_items": deleted, "month": month})
            except Exception as e:
                self._send_json({"error": str(e)}, 400)

        # ล้างรายการจัดซื้อของวันนั้น
        elif path == "/api/day/clear":
            day_num = body.get("day")
            clear_theme = body.get("clear_theme", False)
            if not day_num:
                self._send_json({"error": "กรุณาระบุหมายเลขวัน (day)"}, 400)
                return
            try:
                deleted = clear_day_purchases(int(day_num), clear_theme=clear_theme)
                self._send_json({"success": True, "deleted_items": deleted, "day": int(day_num)})
            except Exception as e:
                self._send_json({"error": str(e)}, 400)

        # แก้ไขเมนูอาหาร / ธีมประจำวัน
        elif path == "/api/day/menu":
            day_num = body.get("day")
            theme = body.get("theme", "")
            notes = body.get("notes", "")
            template_key = body.get("template_key")
            apply_items = body.get("apply_items", False)

            if not day_num:
                self._send_json({"error": "กรุณาระบุหมายเลขวัน (day)"}, 400)
                return

            try:
                day_num = int(day_num)
                if apply_items and template_key and template_key in PROCUREMENT_TEMPLATES:
                    tpl = PROCUREMENT_TEMPLATES[template_key]
                    items_data = [
                        {"item_code": it["code"], "quantity": it["qty"], "notes": it.get("note", "")}
                        for it in tpl["items"]
                    ]
                    updated = update_daily_purchases(day_num, items_data, theme=theme or tpl["theme"], notes=notes)
                else:
                    updated = update_daily_metadata(day_num, theme=theme, notes=notes)

                self._send_json({"success": True, "day": updated})
            except Exception as e:
                self._send_json({"error": str(e)}, 400)

        else:
            self._send_json({"error": "ไม่พบคำสั่งนี้"}, 404)

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def _serve_bill_page(self, qs):
        day = int(qs.get("day", [1])[0])
        bill = get_bill_data_for_day(day)
        if bill:
            html = render_printable_bill_html(bill)
            self._send_html(html)
        else:
            self._send_html("<h1>ไม่พบข้อมูลวันที่ระบุ</h1>", 404)

    def _serve_dashboard(self):
        html = get_dashboard_html()
        self._send_html(html)


def get_dashboard_html():
    """ส่งคืนโค้ด HTML ของหน้า Dashboard ทั้งหมด โดยใช้ฟอนต์ Kanit และภาษาไทย 100%"""
    return r"""<!DOCTYPE html>
<html lang="th">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ระบบกำหนดรายการจัดซื้อและออกใบสั่งซื้อรายวัน (365 วัน)</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Kanit:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&family=Sarabun:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&display=swap" rel="stylesheet">
<style>
:root {
  --bg: #f3f4f6;
  --card: #ffffff;
  --primary: #1e40af;
  --primary-dark: #1e3a8a;
  --primary-light: #eff6ff;
  --accent: #0284c7;
  --success: #059669;
  --warning: #d97706;
  --danger: #dc2626;
  --text: #111827;
  --text-secondary: #4b5563;
  --border: #e5e7eb;
  --shadow: 0 1px 3px 0 rgba(0,0,0,0.08), 0 1px 2px -1px rgba(0,0,0,0.08);
  --shadow-lg: 0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -4px rgba(0,0,0,0.08);
  --radius: 8px;
  --mono: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body {
  font-family: 'Kanit', sans-serif;
  background: var(--bg);
  color: var(--text);
  font-size: 14.5px;
  line-height: 1.5;
  font-weight: 400;
}

/* Vector SVG Icons */
.ic {
  width: 16px;
  height: 16px;
  fill: none;
  stroke: currentColor;
  stroke-width: 2;
  stroke-linecap: round;
  stroke-linejoin: round;
  vertical-align: -3px;
  display: inline-block;
  flex-shrink: 0;
}
.ic-lg { width: 20px; height: 20px; vertical-align: -4px; }
.ic-xl { width: 24px; height: 24px; vertical-align: -5px; }

/* แถบด้านบน Topbar */
.topbar {
  background: linear-gradient(135deg, #1e3a8a 0%, #1e40af 50%, #2563eb 100%);
  color: #fff;
  padding: 0 28px;
  height: 64px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  box-shadow: 0 2px 10px rgba(30,58,138,0.25);
  position: sticky;
  top: 0;
  z-index: 100;
}
.topbar-title {
  font-size: 17px;
  font-weight: 600;
  letter-spacing: -0.2px;
  display: flex;
  align-items: center;
  gap: 10px;
}
.topbar-actions { display: flex; gap: 8px; align-items: center; }
.topbar-btn {
  background: rgba(255,255,255,0.14);
  color: #fff;
  border: 1px solid rgba(255,255,255,0.25);
  padding: 7px 14px;
  border-radius: 6px;
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
  transition: all 0.15s;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-family: 'Kanit', sans-serif;
}
.topbar-btn:hover { background: rgba(255,255,255,0.26); }
.topbar-btn-danger { background: rgba(220,50,50,0.55) !important; border-color: rgba(255,100,100,0.5) !important; }
.topbar-btn-danger:hover { background: rgba(220,50,50,0.80) !important; }

/* โครงสร้างหน้า */
.main-layout {
  display: grid;
  grid-template-columns: 330px 1fr;
  gap: 0;
  min-height: calc(100vh - 64px);
}
.sidebar {
  background: var(--card);
  border-right: 1px solid var(--border);
  overflow-y: auto;
  max-height: calc(100vh - 64px);
  position: sticky;
  top: 64px;
}
.content {
  padding: 24px 28px;
  overflow-y: auto;
}

/* แถบข้าง Sidebar */
.sidebar-section {
  padding: 18px;
  border-bottom: 1px solid var(--border);
}
.sidebar-section-title {
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  color: var(--text-secondary);
  font-weight: 600;
  margin-bottom: 12px;
}
.stat-row {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 6px 0;
  font-size: 13.5px;
}
.stat-label { color: var(--text-secondary); }
.stat-value { font-weight: 600; font-family: var(--mono); }
.stat-value.primary { color: var(--primary); font-size: 15px; }
.stat-value.success { color: var(--success); }

/* เมนูเดือน */
.month-nav {
  list-style: none;
  padding: 0;
}
.month-nav li {
  padding: 8px 12px;
  cursor: pointer;
  border-radius: 6px;
  margin-bottom: 2px;
  font-size: 13.5px;
  font-weight: 500;
  display: flex;
  justify-content: space-between;
  align-items: center;
  transition: all 0.12s;
}
.month-nav li:hover { background: var(--primary-light); color: var(--primary); }
.month-nav li.active { background: var(--primary); color: #fff; }
.month-nav .month-cost {
  font-family: var(--mono);
  font-size: 12.5px;
  font-weight: 600;
}

/* หมวดหมู่ */
.cat-pills {
  display: flex;
  flex-wrap: wrap;
  gap: 5px;
}
.cat-pill {
  font-size: 11.5px;
  padding: 4px 10px;
  border-radius: 14px;
  background: #f3f4f6;
  color: #374151;
  cursor: pointer;
  font-weight: 500;
  border: 1px solid transparent;
  transition: all 0.1s;
}
.cat-pill:hover, .cat-pill.active {
  background: var(--primary-light);
  color: var(--primary);
  border-color: var(--primary);
}

/* การ์ดเนื้อหา */
.card {
  background: var(--card);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  padding: 22px 26px;
  margin-bottom: 20px;
}
.card-title {
  font-size: 17px;
  font-weight: 600;
  margin-bottom: 18px;
  display: flex;
  align-items: center;
  gap: 8px;
}

/* แถวสรุปตัวเลข KPI */
.kpi-row {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 16px;
  margin-bottom: 20px;
}
.kpi-card {
  background: var(--card);
  border-radius: var(--radius);
  box-shadow: var(--shadow);
  padding: 16px 20px;
  display: flex;
  flex-direction: column;
}
.kpi-label {
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.4px;
  color: var(--text-secondary);
  font-weight: 600;
}
.kpi-value {
  font-size: 26px;
  font-weight: 700;
  font-family: var(--mono);
  margin-top: 4px;
  color: var(--text);
}
.kpi-sub {
  font-size: 12px;
  color: var(--text-secondary);
  margin-top: 2px;
}

/* ตารางปฏิทิน */
.day-grid {
  display: grid;
  grid-template-columns: repeat(7, 1fr);
  gap: 6px;
  margin-top: 12px;
}
.day-header {
  font-size: 12px;
  color: var(--text-secondary);
  font-weight: 600;
  text-align: center;
  padding: 6px 4px;
}
.day-cell {
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: 6px;
  padding: 6px 8px;
  font-size: 11.5px;
  cursor: pointer;
  transition: all 0.15s;
  min-height: 62px;
  position: relative;
}
.day-cell:hover {
  border-color: var(--primary);
  box-shadow: 0 0 0 2px rgba(30,64,175,0.15);
  transform: translateY(-1px);
}
.day-cell.selected {
  border-color: var(--primary);
  background: var(--primary-light);
  box-shadow: 0 0 0 2px rgba(30,64,175,0.3);
}
.day-cell .day-num {
  font-weight: 600;
  color: var(--text);
  font-size: 13px;
}
.day-cell .day-cost {
  font-family: var(--mono);
  font-size: 11px;
  font-weight: 600;
  color: var(--success);
  margin-top: 2px;
}
.day-cell .day-items-count {
  font-size: 10px;
  color: var(--text-secondary);
}
.day-cell .day-bar {
  position: absolute;
  bottom: 0;
  left: 0;
  right: 0;
  height: 3px;
  border-radius: 0 0 6px 6px;
  background: var(--primary);
  opacity: 0.55;
}

/* แผงรายละเอียดรายวัน */
.detail-panel { display: none; }
.detail-panel.visible { display: block; }
.detail-header {
  display: flex;
  justify-content: space-between;
  align-items: flex-start;
  margin-bottom: 18px;
}
.detail-day-title {
  font-size: 21px;
  font-weight: 700;
}
.detail-theme {
  font-size: 14px;
  color: var(--primary);
  font-weight: 500;
  margin-top: 3px;
}
.detail-meta {
  display: flex;
  gap: 16px;
  margin-top: 5px;
}
.detail-meta span {
  font-size: 13px;
  color: var(--text-secondary);
}

/* ตารางแสดงรายการ */
.items-table {
  width: 100%;
  border-collapse: collapse;
}
.items-table th {
  background: #f8fafc;
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.3px;
  color: var(--text-secondary);
  font-weight: 600;
  padding: 10px 8px;
  text-align: left;
  border-bottom: 2px solid var(--border);
}
.items-table td {
  padding: 9px 8px;
  border-bottom: 1px solid #f3f4f6;
  font-size: 13.5px;
  vertical-align: middle;
}
.items-table tr:hover td { background: #fafbfc; }
.items-table .text-right { text-align: right; }
.items-table .text-center { text-align: center; }
.items-table .mono { font-family: var(--mono); }
.items-table .item-code {
  font-family: var(--mono);
  font-size: 12px;
  color: var(--primary);
  font-weight: 600;
}

/* ป๊อปอัปเลือกวัตถุดิบ (Modal) */
.modal-overlay {
  display: none;
  position: fixed;
  top: 0; left: 0; right: 0; bottom: 0;
  background: rgba(0,0,0,0.5);
  z-index: 200;
  align-items: center;
  justify-content: center;
}
.modal-overlay.visible { display: flex; }
.modal {
  background: var(--card);
  border-radius: 12px;
  box-shadow: var(--shadow-lg);
  width: 860px;
  max-height: 85vh;
  display: flex;
  flex-direction: column;
}
.modal-header {
  padding: 18px 24px;
  border-bottom: 1px solid var(--border);
  display: flex;
  justify-content: space-between;
  align-items: center;
}
.modal-title { font-size: 17px; font-weight: 600; }
.modal-close {
  background: none;
  border: none;
  font-size: 18px;
  cursor: pointer;
  color: var(--text-secondary);
  padding: 4px 8px;
  border-radius: 4px;
}
.modal-close:hover { background: #f3f4f6; }
.modal-body {
  flex: 1;
  overflow-y: auto;
  padding: 16px 24px;
}
.modal-footer {
  padding: 14px 24px;
  border-top: 1px solid var(--border);
  display: flex;
  justify-content: space-between;
  align-items: center;
}

/* ช่องค้นหาและฟิลด์ฟอร์ม */
.search-input {
  width: 100%;
  padding: 10px 14px;
  border: 1px solid var(--border);
  border-radius: 6px;
  font-size: 14px;
  font-family: 'Kanit', sans-serif;
  margin-bottom: 12px;
  outline: none;
  transition: border-color 0.15s;
}
.search-input:focus {
  border-color: var(--primary);
  box-shadow: 0 0 0 3px rgba(30,64,175,0.1);
}

/* รายการเลือกวัตถุดิบ */
.ingredient-item {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 8px 10px;
  border-radius: 6px;
  cursor: pointer;
  transition: background 0.1s;
}
.ingredient-item:hover { background: #f9fafb; }
.ingredient-item input[type="checkbox"] {
  width: 17px;
  height: 17px;
  accent-color: var(--primary);
}
.ingredient-item .ing-name { font-weight: 500; flex: 1; }
.ingredient-item .ing-category {
  font-size: 11px;
  background: #f3f4f6;
  padding: 2px 7px;
  border-radius: 4px;
  color: var(--text-secondary);
  font-weight: 500;
}
.ingredient-item .ing-price {
  font-family: var(--mono);
  font-weight: 600;
  color: var(--success);
  font-size: 13.5px;
}
.ingredient-item .ing-unit {
  font-size: 12px;
  color: var(--text-secondary);
  width: 50px;
  text-align: right;
}
.qty-input {
  width: 65px;
  padding: 4px 6px;
  border: 1px solid var(--border);
  border-radius: 4px;
  font-size: 13.5px;
  text-align: center;
  font-family: var(--mono);
}

/* ปุ่มกด Buttons */
.btn {
  padding: 8px 16px;
  border-radius: 6px;
  font-size: 13.5px;
  font-weight: 500;
  cursor: pointer;
  border: none;
  transition: all 0.15s;
  display: inline-flex;
  align-items: center;
  gap: 6px;
  font-family: 'Kanit', sans-serif;
  text-decoration: none;
}
.btn-primary { background: var(--primary); color: #fff; }
.btn-primary:hover { background: var(--primary-dark); }
.btn-success { background: var(--success); color: #fff; }
.btn-success:hover { background: #047857; }
.btn-warning { background: var(--warning); color: #fff; }
.btn-warning:hover { background: #b45309; }
.btn-danger { background: var(--danger); color: #fff; }
.btn-danger:hover { background: #b91c1c; }
.btn-outline {
  background: transparent;
  border: 1px solid var(--border);
  color: var(--text);
}
.btn-outline:hover { background: #f9fafb; border-color: var(--primary); color: var(--primary); }
.btn-danger-outline {
  background: transparent;
  border: 1px solid #fecaca;
  color: var(--danger);
}
.btn-danger-outline:hover {
  background: #fef2f2;
  border-color: var(--danger);
}
.btn-warning-outline {
  background: transparent;
  border: 1px solid #fde68a;
  color: #b45309;
}
.btn-warning-outline:hover {
  background: #fffbeb;
  border-color: #f59e0b;
}
.btn-sm { padding: 5px 12px; font-size: 12.5px; }
.btn-xs { padding: 2px 8px; font-size: 11px; }

.day-theme-tag {
  font-size: 10px;
  color: var(--primary);
  font-weight: 500;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  margin-top: 2px;
  max-width: 100%;
}

/* แถบแท็บ (Tabs) */
.tabs {
  display: flex;
  border-bottom: 2px solid var(--border);
  margin-bottom: 20px;
}
.tab {
  padding: 10px 20px;
  font-size: 14.5px;
  font-weight: 500;
  color: var(--text-secondary);
  cursor: pointer;
  border-bottom: 2px solid transparent;
  margin-bottom: -2px;
  transition: all 0.15s;
  display: flex;
  align-items: center;
  gap: 6px;
}
.tab:hover { color: var(--primary); }
.tab.active {
  color: var(--primary);
  border-bottom-color: var(--primary);
  font-weight: 600;
}
.tab-content { display: none; }
.tab-content.active { display: block; }

/* การแจ้งเตือน Toast */
.toast-container {
  position: fixed;
  top: 76px;
  right: 24px;
  z-index: 300;
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.toast {
  background: #1f2937;
  color: #fff;
  padding: 12px 20px;
  border-radius: 8px;
  box-shadow: var(--shadow-lg);
  font-size: 13.5px;
  font-weight: 500;
  display: flex;
  align-items: center;
  gap: 8px;
  animation: slideIn 0.3s ease-out;
  max-width: 380px;
}
.toast.success { background: #065f46; }
.toast.error { background: #991b1b; }
@keyframes slideIn {
  from { transform: translateX(100%); opacity: 0; }
  to { transform: translateX(0); opacity: 1; }
}

@media (max-width: 1024px) {
  .main-layout { grid-template-columns: 1fr; }
  .sidebar { display: none; }
  .kpi-row { grid-template-columns: repeat(2, 1fr); }
  .day-grid { grid-template-columns: repeat(4, 1fr); }
}

/* Modern Beautiful Custom Popup Modal System */
.sw-popup-backdrop {
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
}
.sw-popup-backdrop.active {
  opacity: 1;
  visibility: visible;
}
.sw-popup-box {
  background: #ffffff;
  border-radius: 20px;
  box-shadow: 0 25px 50px -12px rgba(15, 23, 42, 0.25), 0 0 0 1px rgba(226, 232, 240, 0.9);
  width: 100%;
  max-width: 430px;
  padding: 28px 24px 22px 24px;
  transform: scale(0.9) translateY(12px);
  transition: transform 0.25s cubic-bezier(0.16, 1, 0.3, 1), opacity 0.2s ease;
  text-align: center;
  box-sizing: border-box;
  font-family: 'Kanit', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}
.sw-popup-backdrop.active .sw-popup-box {
  transform: scale(1) translateY(0);
}
.sw-popup-icon {
  width: 64px;
  height: 64px;
  border-radius: 50%;
  margin: 0 auto 16px auto;
  display: flex;
  align-items: center;
  justify-content: center;
  position: relative;
  animation: swPopBounce 0.4s cubic-bezier(0.175, 0.885, 0.32, 1.275);
}
@keyframes swPopBounce {
  0% { transform: scale(0.5); opacity: 0; }
  100% { transform: scale(1); opacity: 1; }
}
.sw-popup-icon.success { background: #dcfce7; color: #16a34a; }
.sw-popup-icon.error   { background: #fee2e2; color: #dc2626; }
.sw-popup-icon.warning { background: #fef3c7; color: #d97706; }
.sw-popup-icon.info    { background: #dbeafe; color: #2563eb; }
.sw-popup-icon svg {
  width: 34px;
  height: 34px;
  stroke: currentColor;
  stroke-width: 2.2;
  fill: none;
  stroke-linecap: round;
  stroke-linejoin: round;
}
.sw-popup-title {
  font-size: 19px;
  font-weight: 700;
  color: #0f172a;
  margin-bottom: 8px;
  line-height: 1.35;
}
.sw-popup-desc {
  font-size: 14.5px;
  color: #475569;
  line-height: 1.6;
  margin-bottom: 22px;
  white-space: pre-line;
  word-break: break-word;
}
.sw-popup-btn-row {
  display: flex;
  gap: 10px;
  justify-content: center;
}
.sw-popup-btn {
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
}
.sw-popup-btn:active {
  transform: scale(0.98);
}
.sw-popup-btn-primary {
  background: #2563eb;
  color: #ffffff;
  box-shadow: 0 4px 10px rgba(37, 99, 235, 0.25);
}
.sw-popup-btn-primary:hover {
  background: #1d4ed8;
  box-shadow: 0 6px 14px rgba(37, 99, 235, 0.35);
}
.sw-popup-btn-danger {
  background: #dc2626;
  color: #ffffff;
  box-shadow: 0 4px 10px rgba(220, 38, 38, 0.25);
}
.sw-popup-btn-danger:hover {
  background: #b91c1c;
  box-shadow: 0 6px 14px rgba(220, 38, 38, 0.35);
}
.sw-popup-btn-secondary {
  background: #f1f5f9;
  color: #475569;
  border: 1px solid #cbd5e1;
}
.sw-popup-btn-secondary:hover {
  background: #e2e8f0;
  color: #1e293b;
}
</style>
</head>
<body>

<!-- SVG Sprite สำหรับไอคอนเวกเตอร์ทั้งหมด -->
<svg xmlns="http://www.w3.org/2000/svg" style="display:none">
  <symbol id="ic-clipboard" viewBox="0 0 24 24"><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><rect x="8" y="2" width="8" height="4" rx="1"/><path d="M12 11h4M12 16h4M8 11h.01M8 16h.01"/></symbol>
  <symbol id="ic-chart" viewBox="0 0 24 24"><rect x="18" y="3" width="4" height="18" rx="1"/><rect x="10" y="8" width="4" height="13" rx="1"/><rect x="2" y="13" width="4" height="8" rx="1"/></symbol>
  <symbol id="ic-refresh" viewBox="0 0 24 24"><polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/></symbol>
  <symbol id="ic-zap" viewBox="0 0 24 24"><polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/></symbol>
  <symbol id="ic-calendar" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></symbol>
  <symbol id="ic-package" viewBox="0 0 24 24"><path d="m16.5 9.4-9-5.19"/><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/></symbol>
  <symbol id="ic-tag" viewBox="0 0 24 24"><path d="M12 2H2v10l9.29 9.29c.94.94 2.48.94 3.42 0l6.58-6.58c.94-.94.94-2.48 0-3.42L12 2Z"/><circle cx="7" cy="7" r="1"/></symbol>
  <symbol id="ic-receipt" viewBox="0 0 24 24"><path d="M4 2v20l2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1V2l-2 1-2-1-2 1-2-1-2 1-2-1-2 1Z"/><path d="M16 8h-6a2 2 0 1 0 0 4h4a2 2 0 1 1 0 4H8M12 6v2m0 8v2"/></symbol>
  <symbol id="ic-plus" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="M12 8v8M8 12h8"/></symbol>
  <symbol id="ic-printer" viewBox="0 0 24 24"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></symbol>
  <symbol id="ic-download" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></symbol>
  <symbol id="ic-cart" viewBox="0 0 24 24"><circle cx="8" cy="21" r="1"/><circle cx="19" cy="21" r="1"/><path d="M2.05 2.05h2l2.66 12.42a2 2 0 0 0 2 1.58h9.78a2 2 0 0 0 1.95-1.57l1.65-7.43H5.12"/></symbol>
  <symbol id="ic-save" viewBox="0 0 24 24"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><polyline points="17 21 17 13 7 13 7 21"/><polyline points="7 3 7 8 15 8"/></symbol>
  <symbol id="ic-file" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/></symbol>
  <symbol id="ic-check" viewBox="0 0 24 24"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></symbol>
  <symbol id="ic-x-circle" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="m15 9-6 6M9 9l6 6"/></symbol>
  <symbol id="ic-info" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10"/><path d="M12 16v-4M12 8h.01"/></symbol>
  <symbol id="ic-x" viewBox="0 0 24 24"><path d="M18 6 6 18M6 6l12 12"/></symbol>
  <symbol id="ic-chevron-left" viewBox="0 0 24 24"><polyline points="15 18 9 12 15 6"/></symbol>
  <symbol id="ic-chevron-right" viewBox="0 0 24 24"><polyline points="9 18 15 12 9 6"/></symbol>
  <symbol id="ic-arrow-left" viewBox="0 0 24 24"><path d="M19 12H5M12 19l-7-7 7-7"/></symbol>
  <symbol id="ic-edit" viewBox="0 0 24 24"><path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/><path d="m15 5 4 4"/></symbol>
  <symbol id="ic-trash" viewBox="0 0 24 24"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></symbol>
  <symbol id="ic-upload" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></symbol>
</svg>

<!-- แถบด้านบน TOPBAR -->
<div class="topbar">
  <div class="topbar-title">
    <svg class="ic ic-lg"><use href="#ic-clipboard"/></svg>
    ระบบกำหนดรายการจัดซื้อและออกใบสั่งซื้อรายวัน (365 วัน)
  </div>
  <div class="topbar-actions">
    <button class="topbar-btn" onclick="downloadBenchmark()">
      <svg class="ic"><use href="#ic-chart"/></svg> ดาวน์โหลด Excel ราคากลาง
    </button>
    <button class="topbar-btn topbar-btn-danger" onclick="resetAll365()">
      <svg class="ic"><use href="#ic-refresh"/></svg> รีเซ็ตข้อมูลทั้งหมด 365 วัน (เริ่มปีใหม่)
    </button>
    <button class="topbar-btn" onclick="syncBenchmark()">
      <svg class="ic"><use href="#ic-zap"/></svg> ซิงค์ราคาล่าสุด
    </button>
  </div>
</div>

<!-- ส่วนเนื้อหาหลัก -->
<div class="main-layout">

  <!-- แถบข้าง SIDEBAR -->
  <div class="sidebar" id="sidebar">
    <div class="sidebar-section">
      <div class="sidebar-section-title">สรุปยอดจัดซื้อทั้งปี (365 วัน)</div>
      <div class="stat-row">
        <span class="stat-label">ยอดรวมทั้งปี</span>
        <span class="stat-value primary" id="sidebar-total">—</span>
      </div>
      <div class="stat-row">
        <span class="stat-label">เฉลี่ยต่อวัน</span>
        <span class="stat-value" id="sidebar-avg">—</span>
      </div>
      <div class="stat-row">
        <span class="stat-label">ยอดสูงสุดต่อวัน</span>
        <span class="stat-value success" id="sidebar-peak">—</span>
      </div>
      <div class="stat-row">
        <span class="stat-label">ยอดต่ำสุดต่อวัน</span>
        <span class="stat-value" id="sidebar-min">—</span>
      </div>
    </div>

    <div class="sidebar-section">
      <div class="sidebar-section-title">เลือกดูตามเดือน</div>
      <ul class="month-nav" id="month-nav"></ul>
    </div>

    <div class="sidebar-section">
      <div class="sidebar-section-title">กรองตามหมวดหมู่วัตถุดิบ</div>
      <div class="cat-pills" id="cat-pills"></div>
    </div>
  </div>

  <!-- เนื้อหา CONTENT AREA -->
  <div class="content" id="content">

    <!-- แถวตัวเลขสรุป KPI -->
    <div class="kpi-row" id="kpi-row">
      <div class="kpi-card">
        <span class="kpi-label">ยอดจัดซื้อรวมทั้งปี</span>
        <span class="kpi-value" id="kpi-total">—</span>
        <span class="kpi-sub">คำนวณจากราคากลางคงที่ 365 วัน</span>
      </div>
      <div class="kpi-card">
        <span class="kpi-label">ยอดจัดซื้อเฉลี่ยต่อวัน</span>
        <span class="kpi-value" id="kpi-avg">—</span>
        <span class="kpi-sub">เฉลี่ยทุกวันในรอบปี</span>
      </div>
      <div class="kpi-card">
        <span class="kpi-label">จำนวนวัตถุดิบในแคตตาล็อก</span>
        <span class="kpi-value" id="kpi-items">—</span>
        <span class="kpi-sub">วัตถุดิบกำหนดราคากลางคงที่</span>
      </div>
      <div class="kpi-card">
        <span class="kpi-label">ระยะเวลาที่ครอบคลุม</span>
        <span class="kpi-value">365 วัน</span>
        <span class="kpi-sub">แผนการจัดซื้อครบทั้งปี</span>
      </div>
    </div>

    <!-- แท็บเมนู TABS -->
    <div class="tabs">
      <div class="tab active" data-tab="calendar" onclick="switchTab('calendar')">
        <svg class="ic"><use href="#ic-calendar"/></svg> ปฏิทินรายวัน 365 วัน
      </div>
      <div class="tab" data-tab="detail" onclick="switchTab('detail')">
        <svg class="ic"><use href="#ic-package"/></svg> รายละเอียดและแก้ไขรายการ
      </div>
      <div class="tab" data-tab="benchmark" onclick="switchTab('benchmark')">
        <svg class="ic"><use href="#ic-tag"/></svg> แคตตาล็อกราคากลาง (Excel)
      </div>
      <div class="tab" data-tab="bills" onclick="switchTab('bills')">
        <svg class="ic"><use href="#ic-receipt"/></svg> พิมพ์ใบสั่งซื้อราชการ
      </div>
    </div>

    <!-- แท็บ 1: ปฏิทิน 365 วัน -->
    <div class="tab-content active" id="tab-calendar">
      <div class="card">
        <div class="card-title" style="justify-content: space-between;">
          <span style="display: flex; align-items: center; gap: 8px;">
            <svg class="ic"><use href="#ic-calendar"/></svg> <span id="calendar-month-title">มกราคม 2569</span>
          </span>
          <div style="display: flex; gap: 6px;">
            <button class="btn btn-danger-outline btn-sm" onclick="clearCurrentMonth()" style="margin-right: 4px;">
              <svg class="ic"><use href="#ic-trash"/></svg> ล้างรายการทั้งเดือน
            </button>
            <button class="btn btn-outline btn-sm" onclick="prevMonth()">
              <svg class="ic"><use href="#ic-chevron-left"/></svg> เดือนก่อนหน้า
            </button>
            <button class="btn btn-outline btn-sm" onclick="nextMonth()">
              เดือนถัดไป <svg class="ic"><use href="#ic-chevron-right"/></svg>
            </button>
          </div>
        </div>
        <div class="day-grid" id="day-grid">
          <div class="day-header">จันทร์</div>
          <div class="day-header">อังคาร</div>
          <div class="day-header">พุธ</div>
          <div class="day-header">พฤหัสฯ</div>
          <div class="day-header">ศุกร์</div>
          <div class="day-header">เสาร์</div>
          <div class="day-header">อาทิตย์</div>
        </div>
      </div>
    </div>

    <!-- แท็บ 2: รายละเอียดและแก้ไขรายการจัดซื้อประจำวัน -->
    <div class="tab-content" id="tab-detail">
      <div class="card">
        <div id="detail-placeholder" style="text-align: center; padding: 48px 20px; color: var(--text-secondary);">
          <div style="margin-bottom: 14px;">
            <svg class="ic" style="width: 52px; height: 52px; color: var(--text-secondary);"><use href="#ic-package"/></svg>
          </div>
          <div style="font-size: 18px; font-weight: 600; color: var(--text);">กรุณาเลือกวันที่ต้องการดูรายละเอียด</div>
          <div style="font-size: 13.5px; margin-top: 4px;">คลิกเลือกวันใดก็ได้จากปฏิทิน หรือกรอกหมายเลขวันด้านล่าง</div>
          <div style="margin-top: 18px; display: flex; gap: 8px; justify-content: center; align-items: center;">
            <span style="font-size: 14px; font-weight: 500;">ไปยังวันที่:</span>
            <input type="number" id="goto-day-input" min="1" max="365" value="1"
              style="width: 80px; padding: 6px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 14px; text-align: center; font-family: var(--mono);">
            <button class="btn btn-primary btn-sm" onclick="goToDay()">ดูรายการ</button>
          </div>
        </div>
        <div class="detail-panel" id="detail-panel">
          <div class="detail-header">
            <div>
              <div class="detail-day-title" id="detail-title">วันที่ 1 — 1 มกราคม 2569</div>
              <div style="display: flex; align-items: center; gap: 8px; margin-top: 4px; flex-wrap: wrap;">
                <div class="detail-theme" id="detail-theme" style="margin-top: 0;">— ยังไม่ได้กำหนดเมนู —</div>
                <button class="btn btn-warning-outline btn-xs" onclick="openEditMenuModal()" title="แก้ไขชื่อเมนูอาหาร">
                  <svg class="ic"><use href="#ic-edit"/></svg> แก้ไขเมนู
                </button>
              </div>
              <div class="detail-meta">
                <span id="detail-items-count">0 รายการ</span>
                <span>•</span>
                <span id="detail-total-qty">0.00 หน่วยรวม</span>
                <span>•</span>
                <span id="detail-total-cost" style="font-weight: 700; color: var(--primary);">฿0.00</span>
              </div>
            </div>
            <div style="display: flex; gap: 8px; flex-wrap: wrap;">
              <button class="btn btn-primary btn-sm" onclick="openIngredientSelector()">
                <svg class="ic"><use href="#ic-plus"/></svg> เพิ่ม / ปรับรายการวัตถุดิบ
              </button>
              <button class="btn btn-success btn-sm" onclick="printDayBill()">
                <svg class="ic"><use href="#ic-printer"/></svg> พิมพ์ใบสั่งซื้อ
              </button>
              <button class="btn btn-outline btn-sm" onclick="exportDayXlsx()">
                <svg class="ic"><use href="#ic-chart"/></svg> ส่งออก Excel
              </button>
              <button class="btn btn-danger-outline btn-sm" onclick="clearCurrentDay()">
                <svg class="ic"><use href="#ic-trash"/></svg> เคลียร์รายการวันนี้
              </button>
            </div>
          </div>
          <!-- ปุ่มเลื่อนวัน -->
          <div style="display: flex; gap: 8px; margin-bottom: 14px;">
            <button class="btn btn-outline btn-sm" onclick="loadDay(currentDay - 1)">
              <svg class="ic"><use href="#ic-chevron-left"/></svg> วันก่อนหน้า
            </button>
            <button class="btn btn-outline btn-sm" onclick="loadDay(currentDay + 1)">
              วันถัดไป <svg class="ic"><use href="#ic-chevron-right"/></svg>
            </button>
          </div>
          <table class="items-table" id="detail-items-table">
            <thead>
              <tr>
                <th style="width: 45px;" class="text-center">ลำดับ</th>
                <th style="width: 85px;">รหัส</th>
                <th>รายการวัตถุดิบ</th>
                <th style="width: 140px;">หมวดหมู่</th>
                <th class="text-right" style="width: 70px;">จำนวน</th>
                <th class="text-center" style="width: 65px;">หน่วย</th>
                <th class="text-right" style="width: 110px;">ราคาต่อหน่วย (฿)</th>
                <th class="text-right" style="width: 110px;">รวมเงิน (฿)</th>
                <th class="text-center" style="width: 75px;">จัดการ</th>
              </tr>
            </thead>
            <tbody id="detail-items-body"></tbody>
            <tfoot>
              <tr style="font-weight: 600; background: #f8fafc;">
                <td colspan="4" class="text-right" style="padding: 12px 8px;">ยอดรวมประจำวัน:</td>
                <td class="text-right mono" id="detail-foot-qty">—</td>
                <td></td>
                <td></td>
                <td class="text-right mono" style="font-size: 16px; color: var(--primary);" id="detail-foot-cost">—</td>
                <td></td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>
    </div>

    <!-- แท็บ 3: แคตตาล็อกราคากลาง Benchmark (พร้อมปุ่มแก้ไขและเพิ่มวัตถุดิบ) -->
    <div class="tab-content" id="tab-benchmark">
      <div class="card">
        <div class="card-title" style="justify-content: space-between;">
          <span style="display: flex; align-items: center; gap: 8px;">
            <svg class="ic"><use href="#ic-tag"/></svg> แคตตาล็อกราคากลางคงที่ (อิงจากไฟล์ Excel)
          </span>
          <div style="display: flex; gap: 8px;">
            <button class="btn btn-primary btn-sm" onclick="openCreateBenchmarkModal()">
              <svg class="ic"><use href="#ic-plus"/></svg> เพิ่มวัตถุดิบใหม่
            </button>
            <button class="btn btn-warning-outline btn-sm" onclick="document.getElementById('excel-upload-input').click()">
              <svg class="ic"><use href="#ic-upload"/></svg> อัปโหลดไฟล์ Excel
            </button>
            <button class="btn btn-outline btn-sm" onclick="downloadBenchmark()">
              <svg class="ic"><use href="#ic-download"/></svg> ดาวน์โหลดไฟล์ Excel
            </button>
          </div>
          <input type="file" id="excel-upload-input" accept=".xlsx" style="display:none;" onchange="uploadBenchmarkExcel(this)">
        </div>
        <input type="text" id="benchmark-search" class="search-input" placeholder="ค้นหาวัตถุดิบตามชื่อ รหัส หรือหมวดหมู่..." oninput="filterBenchmark()">
        <table class="items-table">
          <thead>
            <tr>
              <th style="width: 80px;">รหัส</th>
              <th>ชื่อวัตถุดิบ</th>
              <th style="width: 130px;">หมวดหมู่</th>
              <th class="text-center" style="width: 65px;">หน่วย</th>
              <th class="text-right" style="width: 105px;">ราคาต่อหน่วย (฿)</th>
              <th style="width: 160px;">ผู้จำหน่าย / แหล่งที่มา</th>
              <th class="text-center" style="width: 75px;">สั่งขั้นต่ำ</th>
              <th class="text-center" style="width: 80px;">อายุการเก็บ</th>
              <th class="text-center" style="width: 85px;">จัดการ</th>
            </tr>
          </thead>
          <tbody id="benchmark-body"></tbody>
        </table>
      </div>
    </div>

    <!-- แท็บ 4: ออกและพิมพ์ใบสั่งซื้อ -->
    <div class="tab-content" id="tab-bills">
      <div class="card">
        <div class="card-title">
          <svg class="ic"><use href="#ic-receipt"/></svg> สร้างและพิมพ์ใบสั่งซื้อราชการ (แบบฟอร์มทางการ)
        </div>
        <p style="font-size: 13.5px; color: var(--text-secondary); margin-bottom: 18px;">
          ระบบสร้างใบสั่งซื้อราชการมาตรฐาน 2 หน้าครบถ้วน พร้อมเงื่อนไข 8 ข้อ และช่องลงลายมือชื่อตามระเบียบพัสดุ
        </p>

        <!-- ตั้งค่าข้อมูลหัวใบสั่งซื้อและผู้ลงนาม -->
        <div style="background: #eff6ff; padding: 16px 18px; border-radius: 8px; margin-bottom: 18px; border: 1px solid #bfdbfe; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;">
          <div>
            <div style="font-weight: 600; color: #1e40af; font-size: 15px; margin-bottom: 4px; display: flex; align-items: center; gap: 6px;">
              <svg class="ic" style="color: #1e40af;"><use href="#ic-edit"/></svg> ข้อมูลหัวใบสั่งซื้อและผู้ลงนาม (ผู้ขาย / ส่วนราชการ / ลายมือชื่อ)
            </div>
            <div style="font-size: 13px; color: var(--text-secondary);">
              แก้ไขชื่อและที่อยู่ผู้ขาย, ส่วนราชการ, ผู้สั่งซื้อ และผู้รับใบสั่งซื้อ (หน้าที่ 2) เพื่อใช้ในแบบฟอร์มพิมพ์ทั้งหมด
            </div>
          </div>
          <button class="btn btn-primary btn-sm" onclick="openBillSettingsModal()">
            <svg class="ic"><use href="#ic-edit"/></svg> แก้ไขข้อมูลใบสั่งซื้อ
          </button>
        </div>

        <!-- ใบสั่งซื้อรายวันเดี่ยว -->
        <div style="background: #f8fafc; padding: 18px; border-radius: 8px; margin-bottom: 16px; border: 1px solid var(--border);">
          <div style="font-weight: 600; margin-bottom: 8px; font-size: 15px;">
            <svg class="ic"><use href="#ic-file"/></svg> พิมพ์ใบสั่งซื้อรายวันเดี่ยว
          </div>
          <div style="display: flex; gap: 12px; align-items: center; flex-wrap: wrap;">
            <label style="font-size: 13.5px; font-weight: 500;">เลือกวันที่ (1–365):</label>
            <input type="number" id="bill-day" min="1" max="365" value="1"
              style="width: 85px; padding: 6px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 14px; text-align: center; font-family: var(--mono);">
            <button class="btn btn-primary btn-sm" onclick="openBillPage()">
              <svg class="ic"><use href="#ic-printer"/></svg> เปิดหน้าพิมพ์ใบสั่งซื้อ (Print / Save PDF)
            </button>
            <button class="btn btn-success btn-sm" onclick="exportBillXlsx()">
              <svg class="ic"><use href="#ic-chart"/></svg> ส่งออกไฟล์ Excel
            </button>
          </div>
        </div>

        <!-- ใบสั่งซื้อรายเดือน -->
        <div style="background: #f0fdf4; padding: 18px; border-radius: 8px; margin-bottom: 16px; border: 1px solid #bbf7d0;">
          <div style="font-weight: 600; margin-bottom: 10px; font-size: 15px; color: #166534;">
            <svg class="ic" style="color:#16a34a"><use href="#ic-calendar"/></svg> พิมพ์ใบสั่งซื้อรายเดือน (รวมทั้งเดือน)
          </div>
          <div style="display: flex; gap: 12px; align-items: center; flex-wrap: wrap; margin-bottom: 10px;">
            <label style="font-size: 13.5px; font-weight: 500;">เลือกเดือน:</label>
            <select id="bill-month-select" style="padding: 6px 12px; border: 1px solid var(--border); border-radius: 6px; font-size: 14px; font-family: inherit; min-width: 180px;">
              <option value="2026-01">มกราคม 2569 (วัน 1–31)</option>
              <option value="2026-02">กุมภาพันธ์ 2569 (วัน 32–59)</option>
              <option value="2026-03">มีนาคม 2569 (วัน 60–90)</option>
              <option value="2026-04">เมษายน 2569 (วัน 91–120)</option>
              <option value="2026-05">พฤษภาคม 2569 (วัน 121–151)</option>
              <option value="2026-06">มิถุนายน 2569 (วัน 152–181)</option>
              <option value="2026-07">กรกฎาคม 2569 (วัน 182–212)</option>
              <option value="2026-08">สิงหาคม 2569 (วัน 213–243)</option>
              <option value="2026-09">กันยายน 2569 (วัน 244–273)</option>
              <option value="2026-10">ตุลาคม 2569 (วัน 274–304)</option>
              <option value="2026-11">พฤศจิกายน 2569 (วัน 305–334)</option>
              <option value="2026-12">ธันวาคม 2569 (วัน 335–365)</option>
            </select>
            <button class="btn btn-success btn-sm" onclick="openMonthBillPage()">
              <svg class="ic"><use href="#ic-printer"/></svg> พิมพ์ใบสั่งซื้อเดือนนี้
            </button>
            <button class="btn btn-outline btn-sm" onclick="exportMonthBillXlsx()">
              <svg class="ic"><use href="#ic-chart"/></svg> ส่งออก Excel รายเดือน
            </button>
          </div>
          <div style="font-size: 12.5px; color: #6b7280;">
            <svg class="ic" style="width:13px;height:13px;color:#6b7280"><use href="#ic-info"/></svg>
            ดึงข้อมูลรายการจัดซื้อปัจจุบันจากฐานข้อมูลโดยตรง — สะท้อนยอดที่เพิ่ม/ลบ/แก้ไขล่าสุดเสมอ
          </div>
        </div>

        <!-- ใบสั่งซื้อรวมช่วงวัน -->
        <div style="background: #f8fafc; padding: 18px; border-radius: 8px; margin-bottom: 16px; border: 1px solid var(--border);">
          <div style="font-weight: 600; margin-bottom: 8px; font-size: 15px;">
            <svg class="ic"><use href="#ic-clipboard"/></svg> พิมพ์ใบสั่งซื้อรวมหลายวัน (กำหนดช่วงวันเอง)
          </div>
          <div style="display: flex; gap: 12px; align-items: center; flex-wrap: wrap;">
            <label style="font-size: 13.5px; font-weight: 500;">ตั้งแต่วันที่:</label>
            <input type="date" id="bill-start-date" min="2026-01-01" max="2026-12-31" value="2026-01-01"
              style="padding: 6px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 14px; font-family: inherit;">
            <label style="font-size: 13.5px; font-weight: 500;">ถึงวันที่:</label>
            <input type="date" id="bill-end-date" min="2026-01-01" max="2026-12-31" value="2026-01-07"
              style="padding: 6px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 14px; font-family: inherit;">
            <button class="btn btn-primary btn-sm" onclick="openRangeBillPage()">
              <svg class="ic"><use href="#ic-printer"/></svg> เปิดใบสั่งซื้อรวมช่วงวัน
            </button>
            <button class="btn btn-success btn-sm" onclick="exportRangeBillXlsx()">
              <svg class="ic"><use href="#ic-chart"/></svg> ส่งออก Excel รวม
            </button>
          </div>
          <div style="font-size: 12px; color: #6b7280; margin-top: 8px;" id="range-day-preview"></div>
        </div>

        <!-- ปุ่มลัดตามช่วงเวลา -->
        <div style="background: #f8fafc; padding: 18px; border-radius: 8px; border: 1px solid var(--border);">
          <div style="font-weight: 600; margin-bottom: 8px; font-size: 15px;">
            <svg class="ic"><use href="#ic-zap"/></svg> ทางลัด — พิมพ์ตามไตรมาส
          </div>
          <div style="display: flex; gap: 8px; flex-wrap: wrap;">
            <button class="btn btn-outline btn-sm" onclick="openRangeBill(1,90)">ไตรมาส 1 (วัน 1–90)</button>
            <button class="btn btn-outline btn-sm" onclick="openRangeBill(91,181)">ไตรมาส 2 (วัน 91–181)</button>
            <button class="btn btn-outline btn-sm" onclick="openRangeBill(182,273)">ไตรมาส 3 (วัน 182–273)</button>
            <button class="btn btn-outline btn-sm" onclick="openRangeBill(274,365)">ไตรมาส 4 (วัน 274–365)</button>
            <button class="btn btn-warning btn-sm" onclick="openRangeBill(1,365)">สรุปยอดทั้งปี (1–365)</button>
          </div>
        </div>
      </div>
    </div>

  </div>
</div>

<!-- ป๊อปอัปเลือกวัตถุดิบ (MODAL 1: SELECT INGREDIENTS FOR DAY) -->
<div class="modal-overlay" id="ingredient-modal">
  <div class="modal">
    <div class="modal-header">
      <div class="modal-title">
        <svg class="ic"><use href="#ic-cart"/></svg> เลือกวัตถุดิบสำหรับ <span id="modal-day-label">วันที่ 1</span>
      </div>
      <button class="modal-close" onclick="closeIngredientSelector()">
        <svg class="ic"><use href="#ic-x"/></svg>
      </button>
    </div>
    <div class="modal-body">
      <input type="text" id="modal-search" class="search-input" placeholder="ค้นหาวัตถุดิบตามชื่อ รหัส หรือหมวดหมู่..." oninput="filterModalIngredients()">
      <div id="modal-ingredients-list"></div>
    </div>
    <div class="modal-footer">
      <div id="modal-summary" style="font-size: 13.5px; color: var(--text-secondary);">
        เลือกแล้ว <strong id="modal-selected-count">0</strong> รายการ &bull;
        ประมาณการต้นทุน: <strong id="modal-est-cost" class="mono" style="color: var(--primary);">฿0.00</strong>
      </div>
      <div style="display: flex; gap: 8px;">
        <button class="btn btn-outline" onclick="closeIngredientSelector()">ยกเลิก</button>
        <button class="btn btn-primary" onclick="saveIngredientSelection()">
          <svg class="ic"><use href="#ic-save"/></svg> บันทึกรายการ
        </button>
      </div>
    </div>
  </div>
</div>

<!-- ป๊อปอัปแก้ไขและเพิ่มวัตถุดิบราคากลาง (MODAL 2: BENCHMARK EDIT & CREATE MODAL) -->
<div class="modal-overlay" id="benchmark-edit-modal">
  <div class="modal" style="width: 720px;">
    <div class="modal-header">
      <div class="modal-title" id="bm-modal-title">
        <svg class="ic"><use href="#ic-edit"/></svg> แก้ไขข้อมูลวัตถุดิบราคากลาง (Excel)
      </div>
      <button class="modal-close" onclick="closeBenchmarkEditModal()">
        <svg class="ic"><use href="#ic-x"/></svg>
      </button>
    </div>
    <div class="modal-body" style="padding: 20px 24px;">
      <input type="hidden" id="bm-is-create" value="0">
      
      <div style="display: grid; grid-template-columns: 1fr 2fr; gap: 14px; margin-bottom: 12px;">
        <div>
          <label style="font-size: 13px; font-weight: 500; display: block; margin-bottom: 4px;">รหัสวัตถุดิบ</label>
          <input type="text" id="bm-code" class="search-input" style="margin-bottom: 0;" placeholder="เช่น TH-041">
        </div>
        <div>
          <label style="font-size: 13px; font-weight: 500; display: block; margin-bottom: 4px;">ชื่อวัตถุดิบ *</label>
          <input type="text" id="bm-name" class="search-input" style="margin-bottom: 0;" placeholder="เช่น อกไก่ (หั่น), ซอสแดง...">
        </div>
      </div>

      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 12px;">
        <div>
          <label style="font-size: 13px; font-weight: 500; display: block; margin-bottom: 4px;">หมวดหมู่ *</label>
          <input type="text" id="bm-category" class="search-input" style="margin-bottom: 0;" list="bm-cat-suggestions" placeholder="เลือกหรือพิมพ์หมวดหมู่">
          <datalist id="bm-cat-suggestions">
            <option value="เนื้อสัตว์"></option>
            <option value="เนื้อสัตว์แปรรูป"></option>
            <option value="ผักสด"></option>
            <option value="เครื่องปรุงรส"></option>
            <option value="อาหารทะเล"></option>
            <option value="ไข่"></option>
            <option value="ผลิตภัณฑ์นม"></option>
            <option value="ข้าวและแป้ง"></option>
            <option value="ผลไม้"></option>
            <option value="เครื่องดื่ม"></option>
            <option value="ของหวาน"></option>
            <option value="ผลิตภัณฑ์ถั่วเหลือง"></option>
          </datalist>
        </div>
        <div>
          <label style="font-size: 13px; font-weight: 500; display: block; margin-bottom: 4px;">หน่วยนับ *</label>
          <input type="text" id="bm-unit" class="search-input" style="margin-bottom: 0;" list="bm-unit-suggestions" placeholder="เช่น กก., ขวด, ถุง, ฟอง...">
          <datalist id="bm-unit-suggestions">
            <option value="กก."></option>
            <option value="ขวด"></option>
            <option value="ขีด"></option>
            <option value="ถุง"></option>
            <option value="กระป๋อง"></option>
            <option value="มัด"></option>
            <option value="ฟอง"></option>
            <option value="ลูก"></option>
            <option value="หวี"></option>
            <option value="แถว"></option>
            <option value="หลอด"></option>
            <option value="กระปุก"></option>
          </datalist>
        </div>
      </div>

      <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 14px; margin-bottom: 12px;">
        <div>
          <label style="font-size: 13px; font-weight: 500; display: block; margin-bottom: 4px;">ราคาต่อหน่วย (฿) *</label>
          <input type="number" id="bm-price" step="0.25" min="0" class="search-input mono" style="margin-bottom: 0;" placeholder="0.00">
        </div>
        <div>
          <label style="font-size: 13px; font-weight: 500; display: block; margin-bottom: 4px;">ปริมาณสั่งขั้นต่ำ</label>
          <input type="number" id="bm-min-qty" step="0.5" min="0" class="search-input mono" style="margin-bottom: 0;" placeholder="1.0">
        </div>
        <div>
          <label style="font-size: 13px; font-weight: 500; display: block; margin-bottom: 4px;">อายุการเก็บ (วัน)</label>
          <input type="number" id="bm-shelf-life" step="1" min="1" class="search-input mono" style="margin-bottom: 0;" placeholder="7">
        </div>
      </div>

      <div style="margin-bottom: 12px;">
        <label style="font-size: 13px; font-weight: 500; display: block; margin-bottom: 4px;">ผู้จำหน่าย / แหล่งที่มา</label>
        <input type="text" id="bm-supplier" class="search-input" style="margin-bottom: 0;" placeholder="เช่น ร้านทัศนีย์ สุรเกียรติ">
      </div>

      <div>
        <label style="font-size: 13px; font-weight: 500; display: block; margin-bottom: 4px;">หมายเหตุ / รายละเอียดคุณลักษณะ</label>
        <input type="text" id="bm-notes" class="search-input" style="margin-bottom: 0;" placeholder="เช่น สด สะอาด ลอกหนังหั่นชิ้น, ไม่ช้ำ...">
      </div>
    </div>
    <div class="modal-footer">
      <div style="font-size: 12.5px; color: var(--text-secondary);">
        * ข้อมูลจะถูกอัปเดตลงไฟล์ <strong>benchmark_prices.xlsx</strong> และคำนวณราคาใหม่ทันที
      </div>
      <div style="display: flex; gap: 8px;">
        <button class="btn btn-danger-outline" id="bm-btn-delete" style="display:none;" onclick="deleteCurrentBenchmarkItem()">
          <svg class="ic"><use href="#ic-trash"/></svg> ลบรายการนี้
        </button>
        <button class="btn btn-outline" onclick="closeBenchmarkEditModal()">ยกเลิก</button>
        <button class="btn btn-primary" onclick="saveBenchmarkItem()">
          <svg class="ic"><use href="#ic-save"/></svg> บันทึกและอัปเดต Excel
        </button>
      </div>
    </div>
  </div>
</div>

<!-- ป๊อปอัปตั้งค่าข้อมูลใบสั่งซื้อ / ผู้ลงนาม (MODAL 3: BILL SETTINGS MODAL) -->
<div class="modal-overlay" id="bill-settings-modal">
  <div class="modal" style="width: 780px;">
    <div class="modal-header">
      <div class="modal-title">
        <svg class="ic"><use href="#ic-edit"/></svg> แก้ไขข้อมูลหัวใบสั่งซื้อและผู้ลงนาม
      </div>
      <button class="modal-close" onclick="closeBillSettingsModal()">
        <svg class="ic"><use href="#ic-x"/></svg>
      </button>
    </div>
    <div class="modal-body" style="padding: 20px 24px; max-height: 70vh; overflow-y: auto;">
      <!-- ข้อมูลผู้ขาย -->
      <div style="font-weight: 700; color: var(--primary); font-size: 14px; margin-bottom: 8px; border-bottom: 2px solid var(--border); padding-bottom: 4px;">
        🏢 ข้อมูลผู้ขาย (ด้านซ้ายบนของใบสั่งซื้อ)
      </div>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 10px;">
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">ชื่อผู้ขาย / ร้านค้า</label>
          <input type="text" id="dash-vendor_name" class="search-input" style="margin-bottom:0;">
        </div>
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">เบอร์โทรศัพท์ผู้ขาย</label>
          <input type="text" id="dash-vendor_tel" class="search-input" style="margin-bottom:0;">
        </div>
      </div>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 10px;">
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">ที่อยู่ผู้ขาย บรรทัด 1</label>
          <input type="text" id="dash-vendor_address_line1" class="search-input" style="margin-bottom:0;">
        </div>
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">ที่อยู่ผู้ขาย บรรทัด 2</label>
          <input type="text" id="dash-vendor_address_line2" class="search-input" style="margin-bottom:0;">
        </div>
      </div>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 10px;">
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">เลขประจำตัวผู้เสียภาษี</label>
          <input type="text" id="dash-vendor_tax_id" class="search-input" style="margin-bottom:0;">
        </div>
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">เลขที่บัญชีเงินฝากธนาคาร</label>
          <input type="text" id="dash-vendor_bank_acc" class="search-input" style="margin-bottom:0;">
        </div>
      </div>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 14px;">
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">ชื่อบัญชีเงินฝาก</label>
          <input type="text" id="dash-vendor_bank_name" class="search-input" style="margin-bottom:0;">
        </div>
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">ธนาคาร และ สาขา</label>
          <input type="text" id="dash-vendor_bank_branch" class="search-input" style="margin-bottom:0;">
        </div>
      </div>

      <!-- ข้อมูลส่วนราชการ -->
      <div style="font-weight: 700; color: var(--primary); font-size: 14px; margin-bottom: 8px; border-bottom: 2px solid var(--border); padding-bottom: 4px;">
        🏛️ ข้อมูลส่วนราชการและใบสั่งซื้อ (ด้านขวาบนของใบสั่งซื้อ)
      </div>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 10px;">
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">ชื่อส่วนราชการ</label>
          <input type="text" id="dash-dept_name" class="search-input" style="margin-bottom:0;">
        </div>
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">สังกัด / กรม</label>
          <input type="text" id="dash-dept_agency" class="search-input" style="margin-bottom:0;">
        </div>
      </div>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 10px;">
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">ที่อยู่ส่วนราชการ บรรทัด 1</label>
          <input type="text" id="dash-dept_address_line1" class="search-input" style="margin-bottom:0;">
        </div>
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">ที่อยู่ส่วนราชการ บรรทัด 2</label>
          <input type="text" id="dash-dept_address_line2" class="search-input" style="margin-bottom:0;">
        </div>
      </div>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 14px;">
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">เบอร์โทรศัพท์ส่วนราชการ</label>
          <input type="text" id="dash-dept_tel" class="search-input" style="margin-bottom:0;">
        </div>
        <div>
          <label style="font-size: 12.5px; font-weight: 500; display: block; margin-bottom: 3px;">ปีงบประมาณ</label>
          <input type="text" id="dash-fiscal_year" class="search-input" style="margin-bottom:0;">
        </div>
      </div>

      <!-- ผู้ลงนาม หน้า 2 -->
      <div style="font-weight: 700; color: var(--primary); font-size: 14px; margin-bottom: 8px; border-bottom: 2px solid var(--border); padding-bottom: 4px;">
        ✍️ ผู้ลงนามท้ายใบสั่งซื้อ (หน้าที่ 2)
      </div>
      <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 14px;">
        <div style="background: #eff6ff; padding: 12px; border-radius: 8px; border: 1px solid #bfdbfe;">
          <div style="color: #1e40af; font-weight: 600; font-size: 13.5px; margin-bottom: 8px;">ผู้สั่งซื้อ (Person placing the order)</div>
          <div style="margin-bottom: 8px;">
            <label style="font-size: 12px; font-weight: 500; display: block; margin-bottom: 3px;">ชื่อ-นามสกุล ผู้สั่งซื้อ</label>
            <input type="text" id="dash-buyer_name" class="search-input" style="margin-bottom:0; background:#fff;">
          </div>
          <div>
            <label style="font-size: 12px; font-weight: 500; display: block; margin-bottom: 3px;">ตำแหน่ง ผู้สั่งซื้อ</label>
            <input type="text" id="dash-buyer_position" class="search-input" style="margin-bottom:0; background:#fff;">
          </div>
        </div>
        <div style="background: #f0fdf4; padding: 12px; border-radius: 8px; border: 1px solid #bbf7d0;">
          <div style="color: #166534; font-weight: 600; font-size: 13.5px; margin-bottom: 8px;">ผู้รับใบสั่งซื้อ (Recipient of the order)</div>
          <div style="margin-bottom: 8px;">
            <label style="font-size: 12px; font-weight: 500; display: block; margin-bottom: 3px;">ชื่อ-นามสกุล ผู้รับใบสั่งซื้อ</label>
            <input type="text" id="dash-receiver_name" class="search-input" style="margin-bottom:0; background:#fff;">
          </div>
          <div>
            <label style="font-size: 12px; font-weight: 500; display: block; margin-bottom: 3px;">ตำแหน่ง ผู้รับใบสั่งซื้อ</label>
            <input type="text" id="dash-receiver_position" class="search-input" style="margin-bottom:0; background:#fff;">
          </div>
        </div>
      </div>
    </div>
    <div class="modal-footer">
      <div style="font-size: 12.5px; color: var(--text-secondary);">
        * ข้อมูลจะถูกบันทึกและใช้ในใบสั่งซื้อทุกวัน รวมถึงการพิมพ์และการส่งออก Excel
      </div>
      <div style="display: flex; gap: 8px;">
        <button class="btn btn-outline" onclick="closeBillSettingsModal()">ยกเลิก</button>
        <button class="btn btn-primary" onclick="saveBillSettings()">
          <svg class="ic"><use href="#ic-save"/></svg> บันทึกข้อมูลใบสั่งซื้อ
        </button>
      </div>
    </div>
  </div>
</div>

<!-- ป๊อปอัปแก้ไขเมนูอาหารประจำวัน (MODAL 4: EDIT DAILY MENU MODAL) -->
<div class="modal-overlay" id="menu-edit-modal">
  <div class="modal" style="width: 580px; max-width: 95vw;">
    <div class="modal-header">
      <div class="modal-title">
        <svg class="ic"><use href="#ic-edit"/></svg> แก้ไขเมนูอาหาร — <span id="menu-modal-day-label">วันที่ 1</span>
      </div>
      <button class="modal-close" onclick="closeEditMenuModal()">
        <svg class="ic"><use href="#ic-x"/></svg>
      </button>
    </div>
    <div class="modal-body" style="padding: 20px 24px;">
      <div style="margin-bottom: 16px;">
        <label style="display: block; font-weight: 600; margin-bottom: 6px; font-size: 13.5px; color: #1e293b;">
          ชื่อเมนู / รายการอาหารหลัก (ธีมประจำวัน):
        </label>
        <input type="text" id="menu-edit-theme-input" class="search-input" placeholder="เช่น แกงเขียวหวานไก่ &amp; ไข่พะโล้" style="width: 100%; font-size: 14px; margin-bottom: 0;">
      </div>

      <div style="margin-bottom: 18px;">
        <label style="display: block; font-weight: 600; margin-bottom: 6px; font-size: 13.5px; color: #1e293b;">
          บันทึกเพิ่มเติม / หมายเหตุ (Notes):
        </label>
        <input type="text" id="menu-edit-notes-input" class="search-input" placeholder="เช่น สำหรับเด็กโต 150 คน, งดเผ็ด" style="width: 100%; font-size: 14px; margin-bottom: 0;">
      </div>

      <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 14px; margin-bottom: 10px;">
        <div style="font-weight: 600; font-size: 13px; color: #334155; margin-bottom: 8px; display: flex; align-items: center; gap: 6px;">
          <svg class="ic" style="color: var(--primary);"><use href="#ic-tag"/></svg> หรือเลือกจากชุดเมนูแนะนำมาตรฐาน:
        </div>
        <select id="menu-template-select" class="search-input" style="width: 100%; margin-bottom: 8px; background: #fff;" onchange="onMenuTemplateSelected(this.value)">
          <option value="">-- คลิกเลือกชุดเมนูแนะนำ --</option>
        </select>

        <div id="menu-apply-items-container" style="display: none; margin-top: 10px;">
          <label style="display: flex; align-items: center; gap: 8px; font-size: 13px; color: #0f766e; cursor: pointer; background: #f0fdfa; border: 1px solid #99f6e4; padding: 10px 12px; border-radius: 8px;">
            <input type="checkbox" id="menu-apply-items-check" style="width: 16px; height: 16px; accent-color: #0d9488;">
            <span>นำรายการวัตถุดิบของชุดเมนูนี้มาปรับใช้แทนที่ของวันนี้อัตโนมัติ (<span id="menu-template-items-count">0</span> รายการ)</span>
          </label>
        </div>
      </div>
    </div>
    <div class="modal-footer">
      <div style="font-size: 12px; color: var(--text-secondary);">
        * ข้อมูลเมนูจะแสดงในตารางปฏิทินและรายละเอียดของวัน
      </div>
      <div style="display: flex; gap: 8px;">
        <button class="btn btn-outline" onclick="closeEditMenuModal()">ยกเลิก</button>
        <button class="btn btn-primary" onclick="saveDailyMenu()">
          <svg class="ic"><use href="#ic-save"/></svg> บันทึกเมนูอาหาร
        </button>
      </div>
    </div>
  </div>
</div>

<!-- ส่วนแสดงการแจ้งเตือน Toast -->
<div class="toast-container" id="toast-container"></div>

<script>
// จัดการสถานะในหน้าจอ
let summaryData = null;
let benchmarkItems = [];
let currentDay = null;
let currentMonth = 0;
let modalSelections = {};

const THAI_MONTH_NAMES = [
  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];

document.addEventListener('DOMContentLoaded', async () => {
  try {
    await Promise.all([loadSummary(), loadBenchmark()]);
    renderCalendar();
  } catch (err) {
    console.error('Initial load error:', err);
  }

  // ผูก event กับ date range picker เพื่อแสดง preview หมายเลขวัน
  const startDateEl = document.getElementById('bill-start-date');
  const endDateEl   = document.getElementById('bill-end-date');
  if (startDateEl) startDateEl.addEventListener('change', updateRangePreview);
  if (endDateEl)   endDateEl.addEventListener('change', updateRangePreview);
  updateRangePreview();
});

async function api(url, options = {}) {
  const res = await fetch(url, options);
  if (url.endsWith('.xlsx') || url.includes('export')) return res;
  const contentType = res.headers.get('content-type') || '';
  if (contentType.includes('json')) return res.json();
  return res.text();
}

// Modern Beautiful Popup Alert & Confirm System
function showPopupAlert({ title = 'แจ้งเตือน', message = '', type = 'info', btnText = 'ตกลง' }) {
  return new Promise((resolve) => {
    let container = document.getElementById('sw-popup-container');
    if (!container) {
      container = document.createElement('div');
      container.id = 'sw-popup-container';
      document.body.appendChild(container);
    }

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
          <div class="sw-popup-icon ${type}">
            <svg viewBox="0 0 24 24">${iconSvg}</svg>
          </div>
          <div class="sw-popup-title">${title}</div>
          <div class="sw-popup-desc">${message}</div>
          <div class="sw-popup-btn-row">
            <button class="sw-popup-btn sw-popup-btn-primary" id="sw-popup-ok-btn">${btnText}</button>
          </div>
        </div>
      </div>
    `;

    const overlay = document.getElementById('sw-popup-overlay');
    requestAnimationFrame(() => {
      if (overlay) overlay.classList.add('active');
    });

    let finished = false;
    const finish = () => {
      if (finished) return;
      finished = true;
      document.removeEventListener('keydown', keyHandler);
      if (overlay) overlay.classList.remove('active');
      setTimeout(() => { container.innerHTML = ''; resolve(true); }, 200);
    };

    const keyHandler = (e) => {
      if (e.key === 'Escape' || e.key === 'Enter') finish();
    };
    document.addEventListener('keydown', keyHandler);

    setTimeout(() => {
      const btn = document.getElementById('sw-popup-ok-btn');
      if (btn) {
        btn.focus();
        btn.onclick = finish;
      }
      if (overlay) {
        overlay.onclick = (e) => {
          if (e.target === overlay) finish();
        };
      }
    }, 20);
  });
}

function showPopupConfirm({ title = 'ยืนยันการทำรายการ', message = '', type = 'warning', confirmText = 'ยืนยัน', cancelText = 'ยกเลิก', isDanger = true }) {
  return new Promise((resolve) => {
    let container = document.getElementById('sw-popup-container');
    if (!container) {
      container = document.createElement('div');
      container.id = 'sw-popup-container';
      document.body.appendChild(container);
    }

    const iconSvg = isDanger || type === 'error'
      ? '<circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>'
      : '<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>';

    const btnClass = isDanger ? 'sw-popup-btn-danger' : 'sw-popup-btn-primary';

    container.innerHTML = `
      <div class="sw-popup-backdrop" id="sw-popup-overlay">
        <div class="sw-popup-box">
          <div class="sw-popup-icon ${isDanger ? 'error' : type}">
            <svg viewBox="0 0 24 24">${iconSvg}</svg>
          </div>
          <div class="sw-popup-title">${title}</div>
          <div class="sw-popup-desc">${message}</div>
          <div class="sw-popup-btn-row">
            <button class="sw-popup-btn sw-popup-btn-secondary" id="sw-popup-cancel-btn">${cancelText}</button>
            <button class="sw-popup-btn ${btnClass}" id="sw-popup-confirm-btn">${confirmText}</button>
          </div>
        </div>
      </div>
    `;

    const overlay = document.getElementById('sw-popup-overlay');
    requestAnimationFrame(() => {
      if (overlay) overlay.classList.add('active');
    });

    let finished = false;
    const finish = (result) => {
      if (finished) return;
      finished = true;
      document.removeEventListener('keydown', keyHandler);
      if (overlay) overlay.classList.remove('active');
      setTimeout(() => { container.innerHTML = ''; resolve(result); }, 200);
    };

    const keyHandler = (e) => {
      if (e.key === 'Escape') finish(false);
      else if (e.key === 'Enter') finish(true);
    };
    document.addEventListener('keydown', keyHandler);

    setTimeout(() => {
      const confBtn = document.getElementById('sw-popup-confirm-btn');
      const cancBtn = document.getElementById('sw-popup-cancel-btn');
      if (confBtn) {
        confBtn.focus();
        confBtn.onclick = () => finish(true);
      }
      if (cancBtn) cancBtn.onclick = () => finish(false);
      if (overlay) {
        overlay.onclick = (e) => {
          if (e.target === overlay) finish(false);
        };
      }
    }, 20);
  });
}

// Override window.alert & window.confirm for full system consistency
window.alert = function(msg) {
  const isErr = typeof msg === 'string' && (msg.includes('ผิดพลาด') || msg.includes('ไม่พบ') || msg.includes('กรุณา'));
  return showPopupAlert({
    title: isErr ? 'แจ้งเตือน' : 'สำเร็จ',
    message: String(msg),
    type: isErr ? 'warning' : 'success'
  });
};

window.confirm = function(msg) {
  return showPopupConfirm({
    title: 'ยืนยันการทำรายการ',
    message: String(msg),
    type: 'warning',
    confirmText: 'ยืนยัน',
    cancelText: 'ยกเลิก',
    isDanger: true
  });
};

function showToast(message, type = 'info') {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  const iconId = type === 'success' ? 'check' : type === 'error' ? 'x-circle' : 'info';
  toast.innerHTML = `<svg class="ic"><use href="#ic-${iconId}"/></svg> ${message}`;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 4000);
}

async function loadSummary() {
  summaryData = await api('/api/summary');
  updateKPIs();
  updateSidebar();
}

async function loadBenchmark() {
  const data = await api('/api/benchmark');
  benchmarkItems = data.items || [];
  document.getElementById('kpi-items').textContent = benchmarkItems.length + ' รายการ';
  renderBenchmarkTable();
}

function updateKPIs() {
  if (!summaryData || summaryData.error) return;
  const total = Number(summaryData.annual_total || 0);
  const avg = Number(summaryData.average_daily || 0);
  const elTotal = document.getElementById('kpi-total');
  const elAvg = document.getElementById('kpi-avg');
  if (elTotal) elTotal.textContent = '฿' + total.toLocaleString(undefined, {minimumFractionDigits: 2});
  if (elAvg) elAvg.textContent = '฿' + avg.toLocaleString(undefined, {minimumFractionDigits: 2});
}

function formatThaiDateShort(dateStr) {
  if (!dateStr) return '';
  const parts = dateStr.split('-');
  if (parts.length < 3) return dateStr;
  const day = parseInt(parts[2], 10);
  const mIdx = parseInt(parts[1], 10) - 1;
  const monthName = THAI_MONTH_NAMES[mIdx] || '';
  return `${day} ${monthName}`;
}

function updateSidebar() {
  if (!summaryData || summaryData.error) return;
  const total = Number(summaryData.annual_total || 0);
  const avg = Number(summaryData.average_daily || 0);
  const sbTotal = document.getElementById('sidebar-total');
  const sbAvg = document.getElementById('sidebar-avg');
  if (sbTotal) sbTotal.textContent = '฿' + total.toLocaleString(undefined, {minimumFractionDigits: 2});
  if (sbAvg) sbAvg.textContent = '฿' + avg.toFixed(2);

  const sbPeak = document.getElementById('sidebar-peak');
  const sbMin = document.getElementById('sidebar-min');

  if (summaryData.max_day && Number(summaryData.max_day.daily_total) > 0) {
    const dStr = formatThaiDateShort(summaryData.max_day.date);
    const amount = Number(summaryData.max_day.daily_total).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
    if (sbPeak) {
      sbPeak.textContent = `${dStr} (฿${amount})`;
      sbPeak.title = `วันที่ ${summaryData.max_day.day_number} (${summaryData.max_day.date})`;
    }
  } else {
    if (sbPeak) sbPeak.textContent = '—';
  }

  if (summaryData.min_day && Number(summaryData.min_day.daily_total) > 0) {
    const dStr = formatThaiDateShort(summaryData.min_day.date);
    const amount = Number(summaryData.min_day.daily_total).toLocaleString(undefined, {minimumFractionDigits: 2, maximumFractionDigits: 2});
    if (sbMin) {
      sbMin.textContent = `${dStr} (฿${amount})`;
      sbMin.title = `วันที่ ${summaryData.min_day.day_number} (${summaryData.min_day.date})`;
    }
  } else {
    if (sbMin) sbMin.textContent = '—';
  }

  const monthNav = document.getElementById('month-nav');
  monthNav.innerHTML = '';
  (summaryData.monthly_totals || []).forEach((m, idx) => {
    const li = document.createElement('li');
    li.className = idx === currentMonth ? 'active' : '';
    const parts = m.month.split('-');
    const mIdx = parseInt(parts[1]) - 1;
    const name = THAI_MONTH_NAMES[mIdx] || m.month;
    const bYear = parseInt(parts[0]) + 543;
    li.innerHTML = `<span>${name} ${bYear}</span><span class="month-cost">฿${m.month_total.toLocaleString()}</span>`;
    li.onclick = () => { currentMonth = idx; renderCalendar(); updateMonthNavActive(); };
    monthNav.appendChild(li);
  });

  const catPills = document.getElementById('cat-pills');
  catPills.innerHTML = '';
  (summaryData.category_totals || []).forEach(c => {
    const pill = document.createElement('span');
    pill.className = 'cat-pill';
    pill.textContent = c.category;
    pill.title = '฿' + c.total_spend.toLocaleString();
    catPills.appendChild(pill);
  });
}

function updateMonthNavActive() {
  document.querySelectorAll('.month-nav li').forEach((li, idx) => {
    li.className = idx === currentMonth ? 'active' : '';
  });
}

function renderCalendar() {
  if (!summaryData || summaryData.error || !Array.isArray(summaryData.daily_stats) || !summaryData.daily_stats.length) return;

  const monthly = summaryData.monthly_totals || [];
  if (currentMonth >= monthly.length) currentMonth = 0;
  const mInfo = monthly[currentMonth];
  if (!mInfo) return;

  const [year, mon] = mInfo.month.split('-').map(Number);
  const bYear = year + 543;
  document.getElementById('calendar-month-title').textContent = `${THAI_MONTH_NAMES[mon - 1]} ${bYear}`;

  const monthDays = summaryData.daily_stats.filter(d => d.date && d.date.startsWith(mInfo.month));
  const maxDailyCost = Math.max(...summaryData.daily_stats.map(d => d.daily_total), 1);

  const firstDay = new Date(year, mon - 1, 1);
  let dayOfWeek = firstDay.getDay();
  dayOfWeek = dayOfWeek === 0 ? 6 : dayOfWeek - 1;

  const grid = document.getElementById('day-grid');
  const headers = grid.querySelectorAll('.day-header');
  grid.innerHTML = '';
  headers.forEach(h => grid.appendChild(h));

  for (let i = 0; i < dayOfWeek; i++) {
    const empty = document.createElement('div');
    empty.style.minHeight = '62px';
    grid.appendChild(empty);
  }

  monthDays.forEach(d => {
    const dayDate = new Date(d.date + 'T00:00:00');
    const dayOfMonth = dayDate.getDate();
    const barWidth = (d.daily_total / maxDailyCost * 100).toFixed(0);

    const cell = document.createElement('div');
    cell.className = 'day-cell' + (d.day_number === currentDay ? ' selected' : '');
    cell.innerHTML = `
      <div class="day-num">${dayOfMonth}</div>
      <div class="day-cost">฿${d.daily_total.toFixed(0)}</div>
      <div class="day-items-count">${d.item_count} รายการ</div>
      ${d.theme ? `<div class="day-theme-tag" title="${d.theme}">${d.theme}</div>` : ''}
      <div class="day-bar" style="width: ${barWidth}%;"></div>
    `;
    cell.title = `วันที่ ${d.day_number}: ${d.date} (${d.day_of_week})\n฿${d.daily_total.toFixed(2)} — ${d.item_count} รายการ\n${d.theme || ''}`;
    cell.onclick = () => loadDay(d.day_number);
    grid.appendChild(cell);
  });

  updateMonthNavActive();
}

function prevMonth() {
  if (currentMonth > 0) { currentMonth--; renderCalendar(); }
}
function nextMonth() {
  const max = (summaryData && summaryData.monthly_totals) ? summaryData.monthly_totals.length - 1 : 11;
  if (currentMonth < max) { currentMonth++; renderCalendar(); }
}

async function clearCurrentMonth() {
  if (!summaryData || !summaryData.monthly_totals) return;
  const mInfo = summaryData.monthly_totals[currentMonth];
  if (!mInfo) return;
  const monthDisplay = THAI_MONTH_NAMES[parseInt(mInfo.month.split('-')[1]) - 1];
  const bYear = parseInt(mInfo.month.split('-')[0]) + 543;

  const ok = await showPopupConfirm({
    title: 'ล้างรายการจัดซื้อทั้งเดือน',
    message: `ต้องการล้างรายการจัดซื้อทั้งหมดของเดือน ${monthDisplay} ${bYear} ใช่หรือไม่?\n\nการกระทำนี้จะลบรายการวัตถุดิบทั้งหมดในเดือนนี้และไม่สามารถย้อนกลับได้`,
    confirmText: 'ล้างรายการทั้งเดือน',
    cancelText: 'ยกเลิก',
    isDanger: true
  });
  if (!ok) return;

  const res = await api('/api/month/clear', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ month: mInfo.month, clear_theme: false })
  });
  if (res && res.error) {
    showToast('เกิดข้อผิดพลาด: ' + res.error, 'error');
    return;
  }
  showToast(`ล้างรายการจัดซื้อเดือน ${monthDisplay} ${bYear} เรียบร้อยแล้ว (${res.deleted_items} รายการ)`, 'success');
  await loadSummary();
  renderCalendar();
  if (currentDay) await loadDay(currentDay);
}

async function clearCurrentDay() {
  if (!currentDay) {
    showToast('กรุณาเลือกวันที่ต้องการเคลียร์รายการ', 'warning');
    return;
  }

  const spec = await api(`/api/day/${currentDay}`);
  const itemCount = (spec && spec.items) ? spec.items.length : 0;
  const dayDate = (spec && spec.metadata) ? spec.metadata.date : '';
  const dayTheme = (spec && spec.metadata && spec.metadata.theme) ? ` (${spec.metadata.theme})` : '';

  let dateDesc = `วันที่ ${currentDay}`;
  if (dayDate) {
    const parts = dayDate.split('-');
    if (parts.length === 3) {
      const d = parseInt(parts[2]);
      const m = THAI_MONTH_NAMES[parseInt(parts[1]) - 1];
      const y = parseInt(parts[0]) + 543;
      dateDesc += ` [${d} ${m} ${y}]`;
    }
  }

  const ok = await showPopupConfirm({
    title: `ล้างรายการจัดซื้อ ${dateDesc}`,
    message: `ต้องการล้างรายการจัดซื้อทั้งหมดของ ${dateDesc}${dayTheme} ใช่หรือไม่?\n\nรายการวัตถุดิบทั้งหมด (${itemCount} รายการ) จะถูกลบออก และไม่สามารถย้อนกลับได้`,
    confirmText: 'ล้างรายการวันนี้',
    cancelText: 'ยกเลิก',
    isDanger: true
  });
  if (!ok) return;

  const res = await api('/api/day/clear', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ day: currentDay, clear_theme: false })
  });
  if (res && res.error) {
    showToast('เกิดข้อผิดพลาด: ' + res.error, 'error');
    return;
  }
  showToast(`ล้างรายการจัดซื้อวันที่ ${currentDay} เรียบร้อยแล้ว`, 'success');
  await loadDay(currentDay);
  await loadSummary();
  renderCalendar();
}

async function loadDay(dayNum) {
  if (dayNum < 1 || dayNum > 365) return;
  currentDay = dayNum;

  switchTab('detail');
  document.getElementById('detail-placeholder').style.display = 'none';
  document.getElementById('detail-panel').classList.add('visible');

  const spec = await api(`/api/day/${dayNum}`);
  if (!spec || spec.error) {
    showToast('ไม่พบข้อมูลวันที่เลือก', 'error');
    return;
  }

  const m = spec.metadata;
  const dParts = m.date.split('-');
  const bYear = parseInt(dParts[0]) + 543;
  const thaiDateFormatted = `${parseInt(dParts[2])} ${THAI_MONTH_NAMES[parseInt(dParts[1]) - 1]} ${bYear}`;

  document.getElementById('detail-title').textContent = `วันที่ ${dayNum} — ${thaiDateFormatted} (${m.day_of_week})`;
  document.getElementById('detail-theme').textContent = m.theme ? m.theme : '— ยังไม่ได้กำหนดเมนู —';
  document.getElementById('detail-items-count').textContent = spec.item_count + ' รายการ';
  document.getElementById('detail-total-qty').textContent = spec.total_quantity.toFixed(2) + ' หน่วยรวม';
  document.getElementById('detail-total-cost').textContent = '฿' + spec.total_cost.toFixed(2);

  const tbody = document.getElementById('detail-items-body');
  tbody.innerHTML = '';
  spec.items.forEach((it, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="text-center" style="color: var(--text-secondary); font-family: var(--mono);">${(idx + 1).toString().padStart(2, '0')}</td>
      <td class="item-code">${it.item_code}</td>
      <td><strong>${it.item_name}</strong>${it.notes ? '<br><span style="font-size:12px;color:var(--text-secondary);">' + it.notes + '</span>' : ''}</td>
      <td><span class="cat-pill">${it.category}</span></td>
      <td class="text-right mono" style="font-weight: 600;">${it.quantity.toFixed(2)}</td>
      <td class="text-center" style="color: var(--text-secondary);">${it.unit}</td>
      <td class="text-right mono">฿${it.benchmark_price.toFixed(2)}</td>
      <td class="text-right mono" style="font-weight: 600; color: var(--primary);">฿${it.subtotal_cost.toFixed(2)}</td>
      <td class="text-center">
        <button class="btn btn-danger-outline btn-sm" style="padding: 2px 7px; font-size: 11.5px;" onclick="deleteDayPurchaseItem(${it.id}, '${escapeJsStr(it.item_name)}')">
          <svg class="ic" style="width:12px;height:12px;"><use href="#ic-trash"/></svg> ลบ
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('detail-foot-qty').textContent = spec.total_quantity.toFixed(2);
  document.getElementById('detail-foot-cost').textContent = '฿' + spec.total_cost.toFixed(2);

  renderCalendar();
}

function goToDay() {
  const val = parseInt(document.getElementById('goto-day-input').value);
  if (val >= 1 && val <= 365) loadDay(val);
}

async function openIngredientSelector() {
  if (!currentDay) return;
  document.getElementById('modal-day-label').textContent = 'วันที่ ' + currentDay;

  const spec = await api(`/api/day/${currentDay}`);
  modalSelections = {};
  if (spec && spec.items) {
    spec.items.forEach(it => {
      modalSelections[it.item_code] = { selected: true, qty: it.quantity };
    });
  }

  renderModalIngredients();
  updateModalSummary();
  document.getElementById('ingredient-modal').classList.add('visible');
}

function closeIngredientSelector() {
  document.getElementById('ingredient-modal').classList.remove('visible');
}

function renderModalIngredients(filter = '') {
  const list = document.getElementById('modal-ingredients-list');
  list.innerHTML = '';

  const lowerFilter = filter.toLowerCase();
  const filtered = benchmarkItems.filter(b =>
    !filter ||
    b.name.toLowerCase().includes(lowerFilter) ||
    b.code.toLowerCase().includes(lowerFilter) ||
    b.category.toLowerCase().includes(lowerFilter)
  );

  filtered.forEach(b => {
    const sel = modalSelections[b.code] || { selected: false, qty: b.min_qty || 1 };

    const div = document.createElement('div');
    div.className = 'ingredient-item';
    div.innerHTML = `
      <input type="checkbox" ${sel.selected ? 'checked' : ''} data-code="${b.code}">
      <span class="ing-name">${b.name}</span>
      <span class="ing-category">${b.category}</span>
      <span class="ing-unit">${b.unit}</span>
      <span class="ing-price">฿${b.price.toFixed(2)}</span>
      <input type="number" class="qty-input" value="${sel.qty}" min="0.1" step="0.5" data-code-qty="${b.code}">
    `;

    const checkbox = div.querySelector('input[type="checkbox"]');
    const qtyInput = div.querySelector('.qty-input');

    checkbox.addEventListener('change', () => {
      modalSelections[b.code] = { selected: checkbox.checked, qty: parseFloat(qtyInput.value) || 1 };
      updateModalSummary();
    });
    qtyInput.addEventListener('input', () => {
      if (modalSelections[b.code]) modalSelections[b.code].qty = parseFloat(qtyInput.value) || 1;
      updateModalSummary();
    });

    list.appendChild(div);
  });
}

function filterModalIngredients() {
  const filter = document.getElementById('modal-search').value;
  renderModalIngredients(filter);
}

function updateModalSummary() {
  let count = 0, cost = 0;
  for (const [code, s] of Object.entries(modalSelections)) {
    if (s.selected) {
      count++;
      const bItem = benchmarkItems.find(b => b.code === code);
      if (bItem) cost += s.qty * bItem.price;
    }
  }
  document.getElementById('modal-selected-count').textContent = count;
  document.getElementById('modal-est-cost').textContent = '฿' + cost.toFixed(2);
}

async function saveIngredientSelection() {
  if (!currentDay) return;

  const items = [];
  for (const [code, s] of Object.entries(modalSelections)) {
    if (s.selected && s.qty > 0) {
      items.push({ item_code: code, quantity: s.qty });
    }
  }

  await api(`/api/day/${currentDay}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ items })
  });

  closeIngredientSelector();
  showToast(`บันทึกรายการวันที่ ${currentDay} เรียบร้อยแล้ว (${items.length} รายการ)`, 'success');
  await loadSummary();
  await loadDay(currentDay);
}

/* ==============================================================
   ฟังก์ชันจัดการแคตตาล็อกราคากลาง (Excel) & แก้ไขข้อมูล
============================================================== */

function renderBenchmarkTable(filter = '') {
  const tbody = document.getElementById('benchmark-body');
  tbody.innerHTML = '';

  const lowerFilter = filter.toLowerCase();
  const filtered = benchmarkItems.filter(b =>
    !filter ||
    b.name.toLowerCase().includes(lowerFilter) ||
    b.code.toLowerCase().includes(lowerFilter) ||
    b.category.toLowerCase().includes(lowerFilter) ||
    (b.supplier && b.supplier.toLowerCase().includes(lowerFilter))
  );

  filtered.forEach(b => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="item-code">${b.code}</td>
      <td><strong>${b.name}</strong>${b.notes ? '<br><span style="font-size:12px;color:var(--text-secondary);">' + b.notes + '</span>' : ''}</td>
      <td><span class="cat-pill">${b.category}</span></td>
      <td class="text-center">${b.unit}</td>
      <td class="text-right mono" style="font-weight: 600; color: var(--success);">฿${b.price.toFixed(2)}</td>
      <td style="font-size: 13px; color: var(--text-secondary);">${b.supplier || ''}</td>
      <td class="text-center mono">${b.min_qty}</td>
      <td class="text-center">${b.shelf_life} วัน</td>
      <td class="text-center" style="white-space: nowrap;">
        <button class="btn btn-outline btn-sm" style="padding: 3px 8px; font-size: 12px; margin-right: 4px;" onclick="openEditBenchmarkModal('${b.code}')">
          <svg class="ic" style="width:13px;height:13px;"><use href="#ic-edit"/></svg> แก้ไข
        </button>
        <button class="btn btn-danger-outline btn-sm" style="padding: 3px 8px; font-size: 12px;" onclick="deleteBenchmarkItem('${b.code}', '${escapeJsStr(b.name)}')">
          <svg class="ic" style="width:13px;height:13px;"><use href="#ic-trash"/></svg> ลบ
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function filterBenchmark() {
  const val = document.getElementById('benchmark-search').value;
  renderBenchmarkTable(val);
}

function openEditBenchmarkModal(code) {
  const item = benchmarkItems.find(b => b.code === code);
  if (!item) return;

  document.getElementById('bm-is-create').value = "0";
  document.getElementById('bm-modal-title').innerHTML = '<svg class="ic"><use href="#ic-edit"/></svg> แก้ไขข้อมูลวัตถุดิบราคากลาง: ' + item.code;
  const delBtn = document.getElementById('bm-btn-delete');
  if (delBtn) delBtn.style.display = 'inline-flex';
  
  const codeInput = document.getElementById('bm-code');
  codeInput.value = item.code;
  codeInput.readOnly = true;
  codeInput.style.background = '#f3f4f6';

  document.getElementById('bm-name').value = item.name || '';
  document.getElementById('bm-category').value = item.category || 'ทั่วไป';
  document.getElementById('bm-unit').value = item.unit || 'กก.';
  document.getElementById('bm-price').value = item.price;
  document.getElementById('bm-min-qty').value = item.min_qty || 1.0;
  document.getElementById('bm-shelf-life').value = item.shelf_life || 7;
  document.getElementById('bm-supplier').value = item.supplier || '';
  document.getElementById('bm-notes').value = item.notes || '';

  document.getElementById('benchmark-edit-modal').classList.add('visible');
}

function openCreateBenchmarkModal() {
  document.getElementById('bm-is-create').value = "1";
  document.getElementById('bm-modal-title').innerHTML = '<svg class="ic"><use href="#ic-plus"/></svg> เพิ่มวัตถุดิบใหม่เข้าแคตตาล็อกราคากลาง (Excel)';
  const delBtn = document.getElementById('bm-btn-delete');
  if (delBtn) delBtn.style.display = 'none';
  
  const nextNum = benchmarkItems.length + 1;
  const codeInput = document.getElementById('bm-code');
  codeInput.value = 'TH-' + String(nextNum).padStart(3, '0');
  codeInput.readOnly = false;
  codeInput.style.background = '#ffffff';

  document.getElementById('bm-name').value = '';
  document.getElementById('bm-category').value = 'ผักสด';
  document.getElementById('bm-unit').value = 'กก.';
  document.getElementById('bm-price').value = '0.00';
  document.getElementById('bm-min-qty').value = '1.0';
  document.getElementById('bm-shelf-life').value = '7';
  document.getElementById('bm-supplier').value = 'ร้านทัศนีย์ สุรเกียรติ';
  document.getElementById('bm-notes').value = '';

  document.getElementById('benchmark-edit-modal').classList.add('visible');
}

function closeBenchmarkEditModal() {
  document.getElementById('benchmark-edit-modal').classList.remove('visible');
}

async function saveBenchmarkItem() {
  const isCreate = document.getElementById('bm-is-create').value === "1";
  const code = document.getElementById('bm-code').value.trim();
  const name = document.getElementById('bm-name').value.trim();
  const category = document.getElementById('bm-category').value.trim();
  const unit = document.getElementById('bm-unit').value.trim();
  const price = parseFloat(document.getElementById('bm-price').value) || 0;
  const min_qty = parseFloat(document.getElementById('bm-min-qty').value) || 1.0;
  const shelf_life = parseInt(document.getElementById('bm-shelf-life').value) || 7;
  const supplier = document.getElementById('bm-supplier').value.trim();
  const notes = document.getElementById('bm-notes').value.trim();

  if (!name) {
    await showPopupAlert({ title: 'ข้อมูลไม่ครบถ้วน', message: 'กรุณากรอกชื่อวัตถุดิบ', type: 'warning' });
    return;
  }
  if (!category) {
    await showPopupAlert({ title: 'ข้อมูลไม่ครบถ้วน', message: 'กรุณากรอกหมวดหมู่วัตถุดิบ', type: 'warning' });
    return;
  }
  if (!unit) {
    await showPopupAlert({ title: 'ข้อมูลไม่ครบถ้วน', message: 'กรุณากรอกหน่วยนับ', type: 'warning' });
    return;
  }
  if (price <= 0) {
    const ok = await showPopupConfirm({
      title: 'ราคากลางเป็น 0.00 บาท',
      message: 'ราคาต่อหน่วยเป็น 0.00 บาท ต้องการดำเนินการต่อใช่หรือไม่?',
      confirmText: 'ดำเนินการต่อ',
      cancelText: 'กลับไปแก้ไข',
      isDanger: false
    });
    if (!ok) return;
  }

  const payload = {
    code, name, category, unit, price, min_qty, shelf_life, supplier, notes
  };

  const endpoint = isCreate ? '/api/benchmark/create' : '/api/benchmark/update';
  const result = await api(endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if (result.error) {
    showToast('เกิดข้อผิดพลาด: ' + result.error, 'error');
    return;
  }

  closeBenchmarkEditModal();
  showToast(`อัปเดต ${name} ลงไฟล์ Excel และฐานข้อมูลเรียบร้อยแล้ว`, 'success');
  
  await loadBenchmark();
  await loadSummary();
  if (currentDay) await loadDay(currentDay);
}

function escapeJsStr(str) {
  if (!str) return '';
  return String(str).replace(/'/g, "\\'").replace(/"/g, '&quot;');
}

async function deleteBenchmarkItem(code, name) {
  const ok = await showPopupConfirm({
    title: 'ยืนยันการลบวัตถุดิบราคากลาง',
    message: `คุณต้องการลบรายการวัตถุดิบ "${name}" (${code}) ออกจากแคตตาล็อกราคากลางและไฟล์ Excel ใช่หรือไม่?`,
    confirmText: 'ลบรายการนี้',
    cancelText: 'ยกเลิก',
    isDanger: true
  });
  if (!ok) return;
  const res = await api('/api/benchmark/delete', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ code })
  });
  if (res && res.error) {
    showToast('เกิดข้อผิดพลาดในการลบ: ' + res.error, 'error');
    return;
  }
  showToast(`ลบ "${name}" (${code}) ออกจากราคากลางและอัปเดต Excel สำเร็จ`, 'success');
  closeBenchmarkEditModal();
  await loadBenchmark();
  await loadSummary();
  if (currentDay) await loadDay(currentDay);
}

function deleteCurrentBenchmarkItem() {
  const code = document.getElementById('bm-code').value.trim();
  const name = document.getElementById('bm-name').value.trim();
  if (code) {
    deleteBenchmarkItem(code, name);
  }
}

async function deleteDayPurchaseItem(id, name) {
  const ok = await showPopupConfirm({
    title: 'ยืนยันการลบรายการจัดซื้อ',
    message: `คุณต้องการลบ "${name}" ออกจากรายการจัดซื้อของวันนี้ใช่หรือไม่?`,
    confirmText: 'ลบรายการ',
    cancelText: 'ยกเลิก',
    isDanger: true
  });
  if (!ok) return;
  const res = await api('/api/day/item/delete', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ id })
  });
  if (res && res.error) {
    showToast('เกิดข้อผิดพลาดในการลบ: ' + res.error, 'error');
    return;
  }
  showToast(`ลบ "${name}" เรียบร้อยแล้ว`, 'success');
  await loadSummary();
  if (currentDay) await loadDay(currentDay);
}

/* ==============================================================
   ฟังก์ชันออกใบสั่งซื้อ
============================================================== */

function printDayBill() {
  if (!currentDay) return;
  window.open(`/api/bill/html?day=${currentDay}`, '_blank');
}

function exportDayXlsx() {
  if (!currentDay) return;
  window.open(`/api/bill/export/xlsx?day=${currentDay}`, '_blank');
}

function openBillPage() {
  const day = parseInt(document.getElementById('bill-day').value);
  if (day >= 1 && day <= 365) window.open(`/api/bill/html?day=${day}`, '_blank');
}

function exportBillXlsx() {
  const day = parseInt(document.getElementById('bill-day').value);
  if (day >= 1 && day <= 365) window.open(`/api/bill/export/xlsx?day=${day}`, '_blank');
}

function openRangeBillPage() {
  const [start, end] = getRangeDayNumbers();
  if (!start || !end) return;
  window.open(`/api/bill/range/html?start=${start}&end=${end}`, '_blank');
}

function exportRangeBillXlsx() {
  const [start, end] = getRangeDayNumbers();
  if (!start || !end) return;
  window.open(`/api/bill/range/export/xlsx?start=${start}&end=${end}`, '_blank');
}

function openRangeBill(start, end) {
  window.open(`/api/bill/range/html?start=${start}&end=${end}`, '_blank');
}

// แปลงวันที่ที่เลือกเป็นหมายเลขวัน (1–365) โดยเทียบกับ 2026-01-01
function dateToDay(dateStr) {
  if (!dateStr) return null;
  const base = new Date('2026-01-01');
  const d    = new Date(dateStr);
  if (isNaN(d.getTime())) return null;
  const diff = Math.round((d - base) / 86400000) + 1;
  return (diff >= 1 && diff <= 365) ? diff : null;
}

function getRangeDayNumbers() {
  const startDate = document.getElementById('bill-start-date').value;
  const endDate   = document.getElementById('bill-end-date').value;
  const startDay  = dateToDay(startDate);
  const endDay    = dateToDay(endDate);
  if (!startDay || !endDay) {
    showPopupAlert({ title: 'กรุณาเลือกวันที่', message: 'วันที่ต้องอยู่ในช่วง 1 ม.ค. 2569 – 31 ธ.ค. 2569', type: 'warning', btnText: 'ตกลง' });
    return [null, null];
  }
  if (startDay > endDay) {
    showPopupAlert({ title: 'วันที่ไม่ถูกต้อง', message: 'วันเริ่มต้นต้องไม่เกินวันสิ้นสุด', type: 'warning', btnText: 'ตกลง' });
    return [null, null];
  }
  return [startDay, endDay];
}

function updateRangePreview() {
  const startDate = document.getElementById('bill-start-date').value;
  const endDate   = document.getElementById('bill-end-date').value;
  const preview   = document.getElementById('range-day-preview');
  if (!preview) return;
  const s = dateToDay(startDate);
  const e = dateToDay(endDate);
  if (s && e && s <= e) {
    preview.textContent = `= วันที่ ${s} ถึง ${e} (รวม ${e - s + 1} วัน)`;
    preview.style.color = '#16a34a';
  } else {
    preview.textContent = '';
  }
}

function openMonthBillPage() {
  const month = document.getElementById('bill-month-select').value;
  if (!month) return;
  window.open(`/api/bill/month/html?month=${encodeURIComponent(month)}`, '_blank');
}

function exportMonthBillXlsx() {
  const month = document.getElementById('bill-month-select').value;
  if (!month) return;
  window.open(`/api/bill/month/export/xlsx?month=${encodeURIComponent(month)}`, '_blank');
}

function downloadBenchmark() {
  window.open('/api/benchmark/download', '_blank');
}

async function uploadBenchmarkExcel(inputEl, directFile) {
  const file = directFile || (inputEl && inputEl.files ? inputEl.files[0] : null);
  if (!file) return;

  const fn = file.name.toLowerCase();
  if (!fn.endsWith('.xlsx')) {
    await showPopupAlert({
      title: 'รูปแบบไฟล์ไม่ถูกต้อง',
      message: 'กรุณาอัปโหลดไฟล์ Excel ในรูปแบบ .xlsx เท่านั้น (Excel Workbook)',
      type: 'warning'
    });
    if (inputEl) inputEl.value = '';
    return;
  }

  const ok = await showPopupConfirm({
    title: 'ยืนยันการอัปโหลดไฟล์ Excel ราคากลาง',
    message: `ต้องการอัปโหลดไฟล์ "${file.name}" ใช่หรือไม่?\n\nระบบจะนำเข้าข้อมูลวัตถุดิบและราคาต่อหน่วยจากไฟล์นี้ และอัปเดตแคตตาล็อกราคากลางทันที`,
    confirmText: 'อัปโหลดและอัปเดต',
    cancelText: 'ยกเลิก',
    isDanger: false
  });
  if (!ok) {
    if (inputEl) inputEl.value = '';
    return;
  }

  const progressEl = document.getElementById('excel-upload-progress');
  if (progressEl) progressEl.style.display = 'block';
  showToast('กำลังอัปโหลดและประมวลผลไฟล์ Excel...', 'info');

  try {
    const formData = new FormData();
    formData.append('file', file);

    const res = await fetch('/api/benchmark/upload', {
      method: 'POST',
      body: formData
    });

    const data = await res.json();
    if (!res.ok || data.error) {
      showToast('เกิดข้อผิดพลาด: ' + (data.error || 'ไม่สามารถอัปโหลดไฟล์ได้'), 'error');
      await showPopupAlert({
        title: 'อัปโหลดไม่สำเร็จ',
        message: data.error || 'ไม่สามารถประมวลผลไฟล์ Excel ได้ กรุณาตรวจสอบหัวตารางและข้อมูลในไฟล์',
        type: 'error'
      });
      return;
    }

    const count = data.items_loaded || data.items_synced || 0;
    showToast(`อัปโหลดไฟล์ Excel สำเร็จ! โหลดวัตถุดิบ ${count} รายการ`, 'success');
    await showPopupAlert({
      title: 'อัปโหลดไฟล์ Excel สำเร็จ',
      message: `อัปเดตราคากลางจากไฟล์ "${file.name}" เรียบร้อยแล้ว\n(นำเข้าข้อมูลวัตถุดิบทั้งหมด ${count} รายการ)`,
      type: 'success'
    });

    await loadBenchmark();
    await loadSummary();
    if (currentDay) await loadDay(currentDay);
    renderCalendar();
  } catch (err) {
    showToast('เกิดข้อผิดพลาดในการเชื่อมต่อ: ' + err.message, 'error');
  } finally {
    if (progressEl) progressEl.style.display = 'none';
    if (inputEl) inputEl.value = '';
  }
}

async function resetAll365() {
  const ok = await showPopupConfirm({
    title: '⚠️ รีเซ็ตข้อมูลทั้งหมด 365 วัน',
    message: 'การดำเนินการนี้จะ:\n\n• ลบรายการจัดซื้อทั้งหมดออก\n• ล้างชื่อเมนูของทุกวัน (365 วัน)\n\nใช้สำหรับเริ่มต้นปีใหม่เท่านั้น\nข้อมูลที่ลบจะไม่สามารถกู้คืนได้!\n\nต้องการดำเนินการต่อใช่หรือไม่?',
    confirmText: 'รีเซ็ตทั้งหมด',
    cancelText: 'ยกเลิก',
    isDanger: true
  });
  if (!ok) return;
  showToast('กำลังรีเซ็ตข้อมูลทั้งหมด 365 วัน...', 'info');
  try {
    const res = await api('/api/reset/all', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
    showToast('รีเซ็ตข้อมูลทั้งหมด 365 วันเรียบร้อยแล้ว!', 'success');
    await loadSummary();
    renderCalendar();
    if (currentDay) loadDay(currentDay);
  } catch(err) {
    showToast('เกิดข้อผิดพลาด: ' + err.message, 'error');
  }
}

async function regenerateSchedule() {
  const ok = await showPopupConfirm({
    title: 'สร้างตาราง 365 วันใหม่',
    message: 'ต้องการสร้างตารางการจัดซื้อ 365 วันใหม่ทั้งหมดใช่หรือไม่?\n\nการแก้ไขเดิมจะถูกรีเซ็ตกลับเป็นค่าเริ่มต้น',
    confirmText: 'สร้างตารางใหม่',
    cancelText: 'ยกเลิก',
    isDanger: true
  });
  if (!ok) return;
  showToast('กำลังสร้างตารางจัดซื้อ 365 วัน...', 'info');
  await api('/api/schedule/regenerate', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
  showToast('สร้างตารางจัดซื้อ 365 วันเรียบร้อยแล้ว!', 'success');
  await loadSummary();
  renderCalendar();
  if (currentDay) loadDay(currentDay);
}

async function syncBenchmark() {
  showToast('กำลังซิงค์ราคากลางจากไฟล์ Excel...', 'info');
  await api('/api/benchmark/sync', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' });
  showToast('ซิงค์ราคากลางจาก Excel เรียบร้อยแล้ว!', 'success');
  await loadBenchmark();
}

async function openBillSettingsModal() {
  const settings = await api('/api/bill/settings');
  if (settings) {
    document.getElementById('dash-vendor_name').value = settings.vendor_name || '';
    document.getElementById('dash-vendor_tel').value = settings.vendor_tel || '';
    document.getElementById('dash-vendor_address_line1').value = settings.vendor_address_line1 || '';
    document.getElementById('dash-vendor_address_line2').value = settings.vendor_address_line2 || '';
    document.getElementById('dash-vendor_tax_id').value = settings.vendor_tax_id || '';
    document.getElementById('dash-vendor_bank_acc').value = settings.vendor_bank_acc || '';
    document.getElementById('dash-vendor_bank_name').value = settings.vendor_bank_name || '';
    document.getElementById('dash-vendor_bank_branch').value = settings.vendor_bank_branch || '';
    document.getElementById('dash-dept_name').value = settings.dept_name || '';
    document.getElementById('dash-dept_agency').value = settings.dept_agency || '';
    document.getElementById('dash-dept_address_line1').value = settings.dept_address_line1 || '';
    document.getElementById('dash-dept_address_line2').value = settings.dept_address_line2 || '';
    document.getElementById('dash-dept_tel').value = settings.dept_tel || '';
    document.getElementById('dash-fiscal_year').value = settings.fiscal_year || '2569';
    document.getElementById('dash-buyer_name').value = settings.buyer_name || '';
    document.getElementById('dash-buyer_position').value = settings.buyer_position || '';
    document.getElementById('dash-receiver_name').value = settings.receiver_name || '';
    document.getElementById('dash-receiver_position').value = settings.receiver_position || '';
  }
  document.getElementById('bill-settings-modal').classList.add('visible');
}

function closeBillSettingsModal() {
  document.getElementById('bill-settings-modal').classList.remove('visible');
}

async function saveBillSettings() {
  const payload = {
    vendor_name: document.getElementById('dash-vendor_name').value.trim(),
    vendor_tel: document.getElementById('dash-vendor_tel').value.trim(),
    vendor_address_line1: document.getElementById('dash-vendor_address_line1').value.trim(),
    vendor_address_line2: document.getElementById('dash-vendor_address_line2').value.trim(),
    vendor_tax_id: document.getElementById('dash-vendor_tax_id').value.trim(),
    vendor_bank_acc: document.getElementById('dash-vendor_bank_acc').value.trim(),
    vendor_bank_name: document.getElementById('dash-vendor_bank_name').value.trim(),
    vendor_bank_branch: document.getElementById('dash-vendor_bank_branch').value.trim(),
    dept_name: document.getElementById('dash-dept_name').value.trim(),
    dept_agency: document.getElementById('dash-dept_agency').value.trim(),
    dept_address_line1: document.getElementById('dash-dept_address_line1').value.trim(),
    dept_address_line2: document.getElementById('dash-dept_address_line2').value.trim(),
    dept_tel: document.getElementById('dash-dept_tel').value.trim(),
    fiscal_year: document.getElementById('dash-fiscal_year').value.trim(),
    buyer_name: document.getElementById('dash-buyer_name').value.trim(),
    buyer_position: document.getElementById('dash-buyer_position').value.trim(),
    receiver_name: document.getElementById('dash-receiver_name').value.trim(),
    receiver_position: document.getElementById('dash-receiver_position').value.trim()
  };

  const res = await api('/api/bill/settings', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if (res && res.error) {
    showToast('เกิดข้อผิดพลาด: ' + res.error, 'error');
    return;
  }
  showToast('บันทึกข้อมูลหัวใบสั่งซื้อและผู้ลงนามเรียบร้อยแล้ว!', 'success');
  closeBillSettingsModal();
}

let cachedTemplates = null;

async function openEditMenuModal() {
  if (!currentDay) {
    showToast('กรุณาเลือกวันที่ต้องการแก้ไขเมนู', 'warning');
    return;
  }

  const spec = await api(`/api/day/${currentDay}`);
  if (!spec || spec.error) {
    showToast('ไม่พบข้อมูลวันที่เลือก', 'error');
    return;
  }

  const m = spec.metadata || {};
  const dParts = (m.date || '').split('-');
  let thaiFormatted = `วันที่ ${currentDay}`;
  if (dParts.length === 3) {
    const bYear = parseInt(dParts[0]) + 543;
    thaiFormatted = `วันที่ ${currentDay} (${parseInt(dParts[2])} ${THAI_MONTH_NAMES[parseInt(dParts[1]) - 1]} ${bYear})`;
  }

  document.getElementById('menu-modal-day-label').textContent = thaiFormatted;
  document.getElementById('menu-edit-theme-input').value = m.theme || '';
  document.getElementById('menu-edit-notes-input').value = m.notes || '';
  document.getElementById('menu-apply-items-check').checked = false;
  document.getElementById('menu-apply-items-container').style.display = 'none';

  // Load templates into dropdown
  if (!cachedTemplates) {
    const res = await api('/api/templates');
    if (res && res.templates) cachedTemplates = res.templates;
  }

  const select = document.getElementById('menu-template-select');
  select.innerHTML = '<option value="">-- คลิกเลือกชุดเมนูแนะนำมาตรฐาน --</option>';
  if (cachedTemplates) {
    Object.keys(cachedTemplates).forEach(k => {
      const t = cachedTemplates[k];
      const opt = document.createElement('option');
      opt.value = k;
      opt.textContent = `${t.name} (${t.item_count} รายการ)`;
      select.appendChild(opt);
    });
  }

  document.getElementById('menu-edit-modal').classList.add('visible');
  setTimeout(() => document.getElementById('menu-edit-theme-input').focus(), 100);
}

function closeEditMenuModal() {
  document.getElementById('menu-edit-modal').classList.remove('visible');
}

function onMenuTemplateSelected(key) {
  const container = document.getElementById('menu-apply-items-container');
  const countSpan = document.getElementById('menu-template-items-count');
  const themeInput = document.getElementById('menu-edit-theme-input');

  if (!key || !cachedTemplates || !cachedTemplates[key]) {
    container.style.display = 'none';
    return;
  }

  const tpl = cachedTemplates[key];
  themeInput.value = tpl.theme || tpl.name;
  countSpan.textContent = tpl.item_count;
  container.style.display = 'flex';
}

async function saveDailyMenu() {
  if (!currentDay) return;

  const theme = document.getElementById('menu-edit-theme-input').value.trim();
  const notes = document.getElementById('menu-edit-notes-input').value.trim();
  const templateKey = document.getElementById('menu-template-select').value;
  const applyItems = document.getElementById('menu-apply-items-check').checked;

  const payload = {
    day: currentDay,
    theme: theme,
    notes: notes,
    template_key: templateKey || null,
    apply_items: applyItems
  };

  const res = await api('/api/day/menu', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  });

  if (res && res.error) {
    showToast('เกิดข้อผิดพลาด: ' + res.error, 'error');
    return;
  }

  showToast(`บันทึกเมนูอาหารวันที่ ${currentDay} เรียบร้อยแล้ว!`, 'success');
  closeEditMenuModal();
  await loadDay(currentDay);
  await loadSummary();
  renderCalendar();
}

function switchTab(tabName) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelector(`.tab[data-tab="${tabName}"]`).classList.add('active');
  document.getElementById('tab-' + tabName).classList.add('active');
}
</script>

</body>
</html>"""


def run_server(port=PORT):
    """เริ่มต้นรัน HTTP Server สำหรับระบบจัดซื้อ"""
    print("=" * 64)
    print("  ระบบกำหนดรายการจัดซื้อและออกใบสั่งซื้อรายวัน (365 วัน)")
    print("  ฟอนต์ Kanit / ภาษาไทย 100% / ราคากลางคงที่จาก Excel")
    print("=" * 64)

    excel_path = init_benchmark_excel()
    print(f"[OK] ไฟล์ Excel ราคากลาง: {excel_path}")

    init_db()
    print(f"[OK] ฐานข้อมูลพร้อมใช้งาน: {DB_PATH}")

    init_empty_365_days()

    count = sync_benchmark_from_excel()
    print(f"[OK] ซิงค์รายการวัตถุดิบ {count} รายการจาก Excel")

    from database import get_connection
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) as cnt FROM daily_metadata")
    row_meta = cur.fetchone()
    cur.execute("SELECT COUNT(*) as cnt FROM daily_purchases")
    row_purchases = cur.fetchone()
    conn.close()

    meta_cnt = row_meta["cnt"] if row_meta else 0
    purchases_cnt = row_purchases["cnt"] if row_purchases else 0
    print(f"[CAL] ปฏิทิน {meta_cnt} วันพร้อมใช้งาน (มี {purchases_cnt} รายการจัดซื้อจริง — ปิดระบบสร้างข้อมูลจำลองแล้ว)")

    server = HTTPServer(("0.0.0.0", port), ProcurementHandler)
    print(f"\n[WEB] เซิร์ฟเวอร์ทำงานที่: http://localhost:{port}")
    print(f"   หน้าหลัก Dashboard: http://localhost:{port}/")
    print(f"   ใบสั่งซื้อวันที่ 1:  http://localhost:{port}/bill?day=1")
    print(f"   ใบสั่งซื้อวันที่ 297: http://localhost:{port}/bill?day=297")
    print(f"\n   กด Ctrl+C เพื่อหยุดการทำงาน\n")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[STOP] หยุดการทำงานเซิร์ฟเวอร์เรียบร้อย")
        server.server_close()


if __name__ == "__main__":
    run_server()

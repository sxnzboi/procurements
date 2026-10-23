"""
Supabase REST API Client for Procurement System.
Pure Python standard library (urllib.request, json, ssl).
Zero external dependencies. Works on Render, local, and containers.
"""

import os
import json
import ssl
import urllib.request
import urllib.parse
import urllib.error
from benchmark import DEFAULT_EXCEL_PATH, save_benchmark_prices_to_excel

SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://rqnhbzzenacplynzutkk.supabase.co")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY", "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJxbmhienplbmFjcGx5bnp1dGtrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTI3NTM0NjcsImV4cCI6MjEwODMyOTQ2N30.hdCWq35q8bB_gFybZceAPd8JYmaGTSoY4_i1nyIjjQY")

# SSL context that tolerates systems without local root certs
_SSL_CTX = ssl.create_default_context()
_SSL_CTX.check_hostname = False
_SSL_CTX.verify_mode = ssl.CERT_NONE

DEFAULT_BILL_SETTINGS = {
    "buyer_name": "นางสาวรานีซะ หะมิ",
    "buyer_position": "เจ้าหน้าที่พัสดุ / เจ้าหน้าที่จัดซื้อ",
    "receiver_name": "นางสาวทัศนีย์ สุรเกียรติ",
    "receiver_position": "ผู้ตรวจรับพัสดุ / หัวหน้างานโภชนาการ",
    "approver_name": "นายแพทย์วิชาญ อมรโรจน์",
    "approver_position": "ผู้อำนวยการโรงพยาบาล",
    "organization_name": "โรงพยาบาลสมเด็จพระยุพราชสายบุรี",
    "organization_address": "เลขที่ 11 หมู่ 1 ตำบลตะลุบัน อำเภอสายบุรี จังหวัดปัตตานี 94110",
    "organization_tax_id": "0994000164281",
    "department": "กลุ่มงานโภชนาการและโภชนบำบัด",
    "vendor_name": "นางสาวทัศนีย์ สุรเกียรติ",
    "vendor_shop": "ร้านทัศนีย์ สุรเกียรติ (จำหน่ายวัตถุดิบอาหารสดและแห้ง)",
    "vendor_tax_id": "3-9404-00123-45-1",
    "vendor_tel": "085-1523003",
    "vendor_address_line1": "หมู่บ้านปากตารอ 67/22 หมู่ที่ 11",
    "vendor_address_line2": "ตำบลบางเก่า อำเภอสายบุรี จังหวัดปัตตานี 94110",
    "tax_rate": "0.0",
    "payment_term_days": "30"
}

def is_supabase_configured():
    return bool(SUPABASE_URL and SUPABASE_KEY and SUPABASE_URL.startswith("http"))

def _api_request(endpoint, method="GET", data=None, params=None, extra_headers=None):
    url = f"{SUPABASE_URL.rstrip('/')}/rest/v1/{endpoint.lstrip('/')}"
    if params:
        query_str = urllib.parse.urlencode(params)
        url += f"?{query_str}"

    headers = {
        "apikey": SUPABASE_KEY,
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/json"
    }
    if extra_headers:
        headers.update(extra_headers)

    body = json.dumps(data).encode("utf-8") if data is not None else None
    req = urllib.request.Request(url, data=body, headers=headers, method=method)

    try:
        with urllib.request.urlopen(req, context=_SSL_CTX, timeout=10) as resp:
            content = resp.read().decode("utf-8")
            if content:
                try:
                    return json.loads(content)
                except json.JSONDecodeError:
                    return content
            return None
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8")
        raise RuntimeError(f"Supabase HTTP {e.code}: {err_msg}")
    except Exception as e:
        raise RuntimeError(f"Supabase connection error: {e}")

# ==================== BENCHMARK ITEMS ====================

def sb_get_all_benchmark_items():
    return _api_request("benchmark_items", params={"select": "*", "order": "category.asc,name.asc"}) or []

def sb_get_benchmark_item(code):
    res = _api_request("benchmark_items", params={"code": f"eq.{code}", "select": "*"})
    return res[0] if res else None

def sb_create_benchmark_item(data, excel_path=DEFAULT_EXCEL_PATH):
    code = data.get("code", "").strip()
    name = data.get("name", "").strip()
    category = data.get("category", "ทั่วไป").strip()
    unit = data.get("unit", "กก.").strip()
    price = float(data.get("price", 0.0))
    supplier = data.get("supplier", "ร้านทัศนีย์ สุรเกียรติ").strip()
    min_qty = float(data.get("min_qty", 1.0))
    shelf_life = int(data.get("shelf_life", 7))
    notes = data.get("notes", "").strip()

    record = {
        "code": code,
        "name": name,
        "category": category,
        "unit": unit,
        "price": price,
        "supplier": supplier,
        "min_qty": min_qty,
        "shelf_life": shelf_life,
        "notes": notes
    }
    _api_request("benchmark_items", method="POST", data=record, extra_headers={"Prefer": "return=representation"})
    
    # Sync Excel locally if file path exists
    try:
        all_items = sb_get_all_benchmark_items()
        save_benchmark_prices_to_excel(all_items, excel_path)
    except Exception:
        pass
    return sb_get_benchmark_item(code)

def sb_update_benchmark_item(code, data, excel_path=DEFAULT_EXCEL_PATH):
    patch_data = {}
    for k in ["name", "category", "unit", "supplier", "notes"]:
        if k in data and data[k] is not None:
            patch_data[k] = str(data[k]).strip()
    if "price" in data and data["price"] is not None:
        patch_data["price"] = float(data["price"])
    if "min_qty" in data and data["min_qty"] is not None:
        patch_data["min_qty"] = float(data["min_qty"])
    if "shelf_life" in data and data["shelf_life"] is not None:
        patch_data["shelf_life"] = int(data["shelf_life"])

    if patch_data:
        _api_request("benchmark_items", method="PATCH", data=patch_data, params={"code": f"eq.{code}"})

    try:
        all_items = sb_get_all_benchmark_items()
        save_benchmark_prices_to_excel(all_items, excel_path)
    except Exception:
        pass
    return sb_get_benchmark_item(code)

def sb_delete_benchmark_item(code, excel_path=DEFAULT_EXCEL_PATH):
    _api_request("benchmark_items", method="DELETE", params={"code": f"eq.{code}"})
    try:
        all_items = sb_get_all_benchmark_items()
        save_benchmark_prices_to_excel(all_items, excel_path)
    except Exception:
        pass
    return True

# ==================== DAILY PURCHASES & SPEC ====================

def sb_get_daily_purchase_spec(day_number):
    meta_res = _api_request("daily_metadata", params={"day_number": f"eq.{day_number}", "select": "*"})
    if not meta_res:
        return None
    meta = meta_res[0]

    items = _api_request("daily_purchases", params={
        "day_number": f"eq.{day_number}",
        "select": "*",
        "order": "category.asc,item_name.asc"
    }) or []

    total_cost = round(sum(float(it.get("subtotal_cost", 0.0)) for it in items), 2)
    total_qty = round(sum(float(it.get("quantity", 0.0)) for it in items), 2)
    item_count = len(items)

    cat_summary = {}
    for it in items:
        c = it.get("category", "ทั่วไป")
        cat_summary[c] = round(cat_summary.get(c, 0.0) + float(it.get("subtotal_cost", 0.0)), 2)

    return {
        "metadata": meta,
        "items": items,
        "total_cost": total_cost,
        "total_quantity": total_qty,
        "item_count": item_count,
        "category_summary": cat_summary
    }

def sb_update_daily_purchases(day_number, items_data, theme=None, notes=None):
    if theme is not None or notes is not None:
        patch = {}
        if theme is not None:
            patch["theme"] = theme
        if notes is not None:
            patch["notes"] = notes
        _api_request("daily_metadata", method="PATCH", data=patch, params={"day_number": f"eq.{day_number}"})

    # Delete existing items for day
    _api_request("daily_purchases", method="DELETE", params={"day_number": f"eq.{day_number}"})

    if items_data:
        benchmark_items = {b["code"]: b for b in sb_get_all_benchmark_items()}
        new_records = []
        for it in items_data:
            code = it["item_code"]
            qty = float(it.get("quantity", 1.0))
            item_note = it.get("notes", "")

            b_item = benchmark_items.get(code)
            if b_item:
                name = b_item["name"]
                cat = b_item["category"]
                unit = b_item["unit"]
                price = float(b_item["price"])
            else:
                name = it.get("item_name", "Custom Ingredient")
                cat = it.get("category", "General")
                unit = it.get("unit", "unit")
                price = float(it.get("benchmark_price", 0.0))

            subtotal = round(qty * price, 2)
            new_records.append({
                "day_number": day_number,
                "item_code": code,
                "item_name": name,
                "category": cat,
                "unit": unit,
                "quantity": qty,
                "benchmark_price": price,
                "subtotal_cost": subtotal,
                "notes": item_note
            })

        _api_request("daily_purchases", method="POST", data=new_records)

    return sb_get_daily_purchase_spec(day_number)

def sb_delete_daily_purchase_item(purchase_id):
    item_res = _api_request("daily_purchases", params={"id": f"eq.{purchase_id}", "select": "day_number"})
    if not item_res:
        return None
    day_number = item_res[0]["day_number"]
    _api_request("daily_purchases", method="DELETE", params={"id": f"eq.{purchase_id}"})
    return day_number

def sb_clear_month_purchases(month_str, clear_theme=True):
    # Find day_numbers in month
    days = _api_request("daily_metadata", params={"date": f"like.{month_str}%", "select": "day_number"}) or []
    day_numbers = [d["day_number"] for d in days]
    deleted_count = 0

    if day_numbers:
        in_str = f"in.({','.join(map(str, day_numbers))})"
        # Count existing items first
        existing = _api_request("daily_purchases", params={"day_number": in_str, "select": "id"}) or []
        deleted_count = len(existing)
        if deleted_count > 0:
            _api_request("daily_purchases", method="DELETE", params={"day_number": in_str})

    if clear_theme:
        _api_request("daily_metadata", method="PATCH", data={"theme": "", "notes": ""}, params={"date": f"like.{month_str}%"})

    return deleted_count

def sb_clear_day_purchases(day_number, clear_theme=False):
    existing = _api_request("daily_purchases", params={"day_number": f"eq.{day_number}", "select": "id"}) or []
    deleted_count = len(existing)
    if deleted_count > 0:
        _api_request("daily_purchases", method="DELETE", params={"day_number": f"eq.{day_number}"})
    if clear_theme:
        _api_request("daily_metadata", method="PATCH", data={"theme": "", "notes": ""}, params={"day_number": f"eq.{day_number}"})
    return deleted_count

def sb_reset_all_data_365():
    """Reset all 365 days data in Supabase: delete purchases and clear daily themes/notes."""
    try:
        _api_request("daily_purchases", method="DELETE", params={"id": "gt.0"})
    except Exception as e:
        print(f"[Supabase] delete all daily_purchases error: {e}")
    try:
        _api_request("daily_metadata", method="PATCH", data={"theme": "", "notes": ""}, params={"day_number": "gt.0"})
    except Exception as e:
        print(f"[Supabase] reset daily_metadata error: {e}")
    return True

def sb_update_daily_metadata(day_number, theme=None, notes=None):
    payload = {}
    if theme is not None:
        payload["theme"] = str(theme).strip()
    if notes is not None:
        payload["notes"] = str(notes).strip()
    if payload:
        _api_request("daily_metadata", method="PATCH", data=payload, params={"day_number": f"eq.{day_number}"})
    return sb_get_daily_purchase_spec(day_number)

# ==================== SUMMARY 365 DAYS ====================

def sb_get_365_days_summary():
    days = _api_request("daily_metadata", params={"select": "*", "order": "day_number.asc"}) or []
    purchases = _api_request("daily_purchases", params={"select": "day_number,item_code,item_name,category,unit,quantity,benchmark_price,subtotal_cost"}) or []

    # Aggregate by day
    purchases_by_day = {}
    for p in purchases:
        d = p["day_number"]
        if d not in purchases_by_day:
            purchases_by_day[d] = []
        purchases_by_day[d].append(p)

    daily_stats = []
    for m in days:
        d_num = m["day_number"]
        d_items = purchases_by_day.get(d_num, [])
        item_count = len(d_items)
        tot_qty = round(sum(float(x.get("quantity", 0)) for x in d_items), 2)
        tot_cost = round(sum(float(x.get("subtotal_cost", 0)) for x in d_items), 2)
        daily_stats.append({
            "day_number": d_num,
            "date": m["date"],
            "day_of_week": m["day_of_week"],
            "theme": m.get("theme", "") or "",
            "item_count": item_count,
            "total_quantity": tot_qty,
            "daily_total": tot_cost
        })

    annual_total = round(sum(d["daily_total"] for d in daily_stats), 2)
    avg_daily = round(annual_total / len(daily_stats), 2) if daily_stats else 0.0
    max_day = max(daily_stats, key=lambda d: d["daily_total"]) if daily_stats else None
    min_day = min(daily_stats, key=lambda d: d["daily_total"]) if daily_stats else None

    # Aggregate by month
    monthly_dict = {}
    for d in daily_stats:
        m_str = d["date"][:7]
        if m_str not in monthly_dict:
            monthly_dict[m_str] = {"month": m_str, "days_count": 0, "month_total": 0.0, "month_qty": 0.0}
        monthly_dict[m_str]["days_count"] += 1
        monthly_dict[m_str]["month_total"] = round(monthly_dict[m_str]["month_total"] + d["daily_total"], 2)
        monthly_dict[m_str]["month_qty"] = round(monthly_dict[m_str]["month_qty"] + d["total_quantity"], 2)

    monthly_totals = sorted(monthly_dict.values(), key=lambda x: x["month"])

    # Aggregate by ingredient
    ing_dict = {}
    for p in purchases:
        code = p["item_code"]
        if code not in ing_dict:
            ing_dict[code] = {
                "item_code": code,
                "item_name": p["item_name"],
                "category": p["category"],
                "unit": p["unit"],
                "benchmark_price": p["benchmark_price"],
                "annual_quantity": 0.0,
                "annual_spend": 0.0,
                "days_set": set()
            }
        ing_dict[code]["annual_quantity"] = round(ing_dict[code]["annual_quantity"] + float(p["quantity"]), 2)
        ing_dict[code]["annual_spend"] = round(ing_dict[code]["annual_spend"] + float(p["subtotal_cost"]), 2)
        ing_dict[code]["days_set"].add(p["day_number"])

    top_ingredients = []
    for ing in sorted(ing_dict.values(), key=lambda x: x["annual_spend"], reverse=True)[:10]:
        top_ingredients.append({
            "item_code": ing["item_code"],
            "item_name": ing["item_name"],
            "category": ing["category"],
            "unit": ing["unit"],
            "benchmark_price": ing["benchmark_price"],
            "annual_quantity": ing["annual_quantity"],
            "annual_spend": ing["annual_spend"],
            "days_purchased": len(ing["days_set"])
        })

    # Aggregate by category
    cat_dict = {}
    for p in purchases:
        c = p["category"]
        if c not in cat_dict:
            cat_dict[c] = {"category": c, "total_spend": 0.0, "total_qty": 0.0}
        cat_dict[c]["total_spend"] = round(cat_dict[c]["total_spend"] + float(p["subtotal_cost"]), 2)
        cat_dict[c]["total_qty"] = round(cat_dict[c]["total_qty"] + float(p["quantity"]), 2)

    category_totals = sorted(cat_dict.values(), key=lambda x: x["total_spend"], reverse=True)

    return {
        "daily_stats": daily_stats,
        "kpi": {
            "total_days": len(daily_stats),
            "annual_total_cost": annual_total,
            "average_daily_cost": avg_daily,
            "total_items": len(purchases),
            "max_spend_day": max_day,
            "min_spend_day": min_day
        },
        "top_ingredients": top_ingredients,
        "monthly_totals": monthly_totals,
        "category_totals": category_totals
    }

# ==================== BILL SETTINGS ====================

def sb_get_bill_settings():
    rows = _api_request("bill_settings", params={"select": "key,value"}) or []
    settings = dict(DEFAULT_BILL_SETTINGS)
    for r in rows:
        settings[r["key"]] = r["value"]
    return settings

def sb_update_bill_settings(settings_dict):
    records = []
    for k, v in settings_dict.items():
        if k in DEFAULT_BILL_SETTINGS or k.startswith("bill_") or k.startswith("order_") or k.startswith("deadline_") or k.startswith("cond_"):
            records.append({"key": k, "value": str(v).strip()})
    if records:
        _api_request("bill_settings", method="POST", data=records, extra_headers={"Prefer": "resolution=merge-duplicates"})
    return sb_get_bill_settings()

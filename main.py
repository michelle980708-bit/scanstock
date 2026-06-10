import io
import os
import sqlite3
import json
from datetime import datetime

import openpyxl
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

app = FastAPI()

# CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 静态文件
app.mount("/static", StaticFiles(directory="static"), name="static")

# ---------- 数据库 ----------
DB_NAME = "scanstock.db"

def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    
    # 创建商品表
    c.execute('''
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            barcode TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            category TEXT NOT NULL,
            quantity INTEGER DEFAULT 0,
            location TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # 创建库存变动表
    c.execute('''
        CREATE TABLE IF NOT EXISTS stock_movements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER NOT NULL,
            type TEXT NOT NULL,
            quantity INTEGER NOT NULL,
            reason TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (product_id) REFERENCES products(id)
        )
    ''')
    
    # 插入示例数据
    c.execute("SELECT COUNT(*) FROM products")
    if c.fetchone()[0] == 0:
        sample = [
            ("RE001", "盐酸", "试剂", 10, "试剂柜-01"),
            ("RE002", "酒精", "试剂", 20, "试剂柜-02"),
            ("CO001", "无粉手套", "耗材", 100, "耗材架-A"),
        ]
        c.executemany("INSERT INTO products (barcode, name, category, quantity, location) VALUES (?,?,?,?,?)", sample)
    
    # 创建字段配置表
    c.execute('''
        CREATE TABLE IF NOT EXISTS field_configs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            category TEXT NOT NULL,
            field_name TEXT NOT NULL,
            display_name TEXT NOT NULL,
            field_type TEXT DEFAULT 'text',
            is_required INTEGER DEFAULT 0,
            sort_order INTEGER DEFAULT 0,
            UNIQUE(category, field_name)
        )
    ''')
    
    # 添加 custom_fields 列
    try:
        c.execute("ALTER TABLE products ADD COLUMN custom_fields TEXT DEFAULT '{}'")
    except sqlite3.OperationalError:
        pass  # 列已存在
    
    conn.commit()
    conn.close()

init_db()

# ---------- 请求模型 ----------
class InboundRequest(BaseModel):
    barcode: str
    quantity: int
    reason: str = ""

class OutboundRequest(BaseModel):
    barcode: str
    quantity: int
    reason: str = ""

class CheckRequest(BaseModel):
    barcode: str
    actual_quantity: int
    reason: str = "盘点调整"

# ---------- 页面 ----------
@app.get("/", response_class=HTMLResponse)
def read_root():
    return HTMLResponse("<h1>ScanStock API is running</h1><p>访问 /static/simple_full.html 使用手机端</p>")

# ---------- API ----------
@app.get("/api/products")
def get_products():
    conn = get_db()
    rows = conn.execute("SELECT * FROM products ORDER BY id").fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.post("/api/inbound")
def inbound(req: InboundRequest):
    conn = get_db()
    prod = conn.execute("SELECT id, quantity FROM products WHERE barcode = ?", (req.barcode,)).fetchone()
    if not prod:
        raise HTTPException(404, "商品不存在")
    new_qty = prod["quantity"] + req.quantity
    conn.execute("UPDATE products SET quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_qty, prod["id"]))
    conn.execute("INSERT INTO stock_movements (product_id, type, quantity, reason) VALUES (?, 'inbound', ?, ?)",
                 (prod["id"], req.quantity, req.reason))
    conn.commit()
    conn.close()
    return {"message": "入库成功", "new_quantity": new_qty}

@app.post("/api/outbound")
def outbound(req: OutboundRequest):
    conn = get_db()
    prod = conn.execute("SELECT id, quantity FROM products WHERE barcode = ?", (req.barcode,)).fetchone()
    if not prod:
        raise HTTPException(404, "商品不存在")
    if prod["quantity"] < req.quantity:
        raise HTTPException(400, "库存不足")
    new_qty = prod["quantity"] - req.quantity
    conn.execute("UPDATE products SET quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_qty, prod["id"]))
    conn.execute("INSERT INTO stock_movements (product_id, type, quantity, reason) VALUES (?, 'outbound', ?, ?)",
                 (prod["id"], -req.quantity, req.reason))
    conn.commit()
    conn.close()
    return {"message": "出库成功", "new_quantity": new_qty}

@app.post("/api/check")
def check(req: CheckRequest):
    conn = get_db()
    prod = conn.execute("SELECT id, quantity FROM products WHERE barcode = ?", (req.barcode,)).fetchone()
    if not prod:
        raise HTTPException(404, "商品不存在")
    old = prod["quantity"]
    diff = req.actual_quantity - old
    if diff != 0:
        conn.execute("UPDATE products SET quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (req.actual_quantity, prod["id"]))
        conn.execute("INSERT INTO stock_movements (product_id, type, quantity, reason) VALUES (?, 'check', ?, ?)",
                     (prod["id"], diff, req.reason))
        conn.commit()
    conn.close()
    return {"message": "盘点记录已保存", "old": old, "new": req.actual_quantity, "diff": diff}

@app.get("/api/export/excel")
def export_products_to_excel(category: str = None):
    conn = get_db()
    if category:
        products = conn.execute("SELECT * FROM products WHERE category = ? ORDER BY id", (category,)).fetchall()
        fields = conn.execute("SELECT * FROM field_configs WHERE category = ? ORDER BY sort_order", (category,)).fetchall()
    else:
        products = conn.execute("SELECT * FROM products ORDER BY id").fetchall()
        fields = []
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{category or '所有商品'}"

    base_headers = ["条码", "商品名称", "分类", "库存数量", "存放位置"]
    custom_headers = [f["display_name"] for f in fields]
    headers = base_headers + custom_headers
    ws.append(headers)

    for p in products:
        custom = json.loads(p["custom_fields"] or "{}")
        row = [
            p["barcode"],
            p["name"],
            p["category"],
            p["quantity"],
            p["location"] or "",
        ]
        for f in fields:
            row.append(custom.get(f["field_name"], ""))
        ws.append(row)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={category or 'products'}.xlsx"}
    )

@app.post("/api/import/excel")
async def import_products_from_excel(file: UploadFile = File(...), category: str = None):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "只支持 .xlsx 或 .xls 文件")
    contents = await file.read()
    wb = openpyxl.load_workbook(io.BytesIO(contents))
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    base_mapping = {
        "条码": "barcode",
        "商品名称": "name",
        "分类": "category",
        "库存数量": "quantity",
        "存放位置": "location"
    }
    conn = get_db()
    fields = []
    if category:
        fields = conn.execute("SELECT * FROM field_configs WHERE category = ? ORDER BY sort_order", (category,)).fetchall()
    
    col_map = {}
    custom_col_map = {}
    for idx, h in enumerate(headers):
        if h in base_mapping:
            col_map[idx] = base_mapping[h]
        else:
            for f in fields:
                if f["display_name"] == h:
                    custom_col_map[idx] = f["field_name"]
                    break

    success_count = 0
    error_rows = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(row):
            continue
        product_data = {}
        custom_data = {}
        for col_idx, val in enumerate(row):
            if col_idx in col_map:
                product_data[col_map[col_idx]] = val
            elif col_idx in custom_col_map:
                custom_data[custom_col_map[col_idx]] = val
        
        barcode = product_data.get("barcode")
        name = product_data.get("name")
        if not barcode or not name:
            error_rows.append(f"第 {row_idx} 行缺少条码或商品名称")
            continue
        try:
            quantity = int(product_data.get("quantity", 0))
        except:
            quantity = 0
        category_val = product_data.get("category") or category
        if not category_val:
            error_rows.append(f"第 {row_idx} 行缺少分类")
            continue
        location = product_data.get("location", "")
        
        existing = conn.execute("SELECT id FROM products WHERE barcode = ?", (barcode,)).fetchone()
        if existing:
            conn.execute(
                "UPDATE products SET name=?, category=?, quantity=?, location=?, custom_fields=?, updated_at=CURRENT_TIMESTAMP WHERE barcode=?",
                (name, category_val, quantity, location, json.dumps(custom_data), barcode)
            )
        else:
            conn.execute(
                "INSERT INTO products (barcode, name, category, quantity, location, custom_fields) VALUES (?,?,?,?,?,?)",
                (barcode, name, category_val, quantity, location, json.dumps(custom_data))
            )
        success_count += 1
    conn.commit()
    conn.close()
    return {"message": f"导入完成，成功 {success_count} 条", "errors": error_rows or None}

@app.get("/api/field_configs/{category}")
def get_field_configs(category: str):
    conn = get_db()
    rows = conn.execute("SELECT * FROM field_configs WHERE category = ? ORDER BY sort_order", (category,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

class FieldConfigItem(BaseModel):
    field_name: str
    display_name: str
    field_type: str = "text"
    is_required: bool = False
    sort_order: int = 0

@app.post("/api/field_configs/{category}")
def save_field_configs(category: str, configs: list[FieldConfigItem]):
    conn = get_db()
    conn.execute("DELETE FROM field_configs WHERE category = ?", (category,))
    for idx, cfg in enumerate(configs):
        conn.execute(
            "INSERT INTO field_configs (category, field_name, display_name, field_type, is_required, sort_order) VALUES (?, ?, ?, ?, ?, ?)",
            (category, cfg.field_name, cfg.display_name, cfg.field_type, cfg.is_required, idx)
        )
    conn.commit()
    conn.close()
    return {"message": "保存成功"}
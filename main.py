import io
import os
import sqlite3
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
    c.execute("SELECT COUNT(*) FROM products")
    if c.fetchone()[0] == 0:
        sample = [
            ("RE001", "盐酸", "试剂", 10, "试剂柜-01"),
            ("RE002", "酒精", "试剂", 20, "试剂柜-02"),
            ("CO001", "无粉手套", "耗材", 100, "耗材架-A"),
        ]
        c.executemany("INSERT INTO products (barcode, name, category, quantity, location) VALUES (?,?,?,?,?)", sample)
    conn.commit()
    conn.close()

init_db()  # 启动时初始化

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
def export_products_to_excel():
    import io
    conn = get_db()
    rows = conn.execute("SELECT id, barcode, name, category, quantity, location FROM products ORDER BY id").fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Products"
    # 写表头
    ws.append(["ID", "条码", "名称", "分类", "库存数量", "存放位置"])
    for row in rows:
        ws.append([row["id"], row["barcode"], row["name"], row["category"], row["quantity"], row["location"]])

    # 保存到内存
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=products.xlsx"}
    )

@app.post("/api/import/excel")
async def import_products_from_excel(file: UploadFile = File(...)):
    # 检查文件后缀
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "只支持 .xlsx 或 .xls 文件")
    
    contents = await file.read()
    wb = openpyxl.load_workbook(io.BytesIO(contents))
    ws = wb.active
    
    conn = get_db()
    cursor = conn.cursor()
    success_count = 0
    error_rows = []
    
    # 假设第一行是表头，从第二行开始读取
    headers = []
    for col in ws.iter_cols(max_row=1, values_only=True):
        headers.append(col[0])
    
    # 映射列名到字段（支持中英文）
    mapping = {
        "条码": "barcode",
        "barcode": "barcode",
        "名称": "name",
        "name": "name",
        "分类": "category",
        "category": "category",
        "库存": "quantity",
        "库存数量": "quantity",
        "quantity": "quantity",
        "位置": "location",
        "location": "location",
    }
    
    for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if not any(row):  # 空行跳过
            continue
        row_data = dict(zip(headers, row))
        # 提取字段
        barcode = row_data.get("条码") or row_data.get("barcode")
        name = row_data.get("名称") or row_data.get("name")
        category = row_data.get("分类") or row_data.get("category")
        quantity = row_data.get("库存") or row_data.get("库存数量") or row_data.get("quantity")
        location = row_data.get("位置") or row_data.get("location")
        
        if not barcode or not name:
            error_rows.append(f"第 {row_idx} 行缺少条码或名称")
            continue
        
        # 数量处理
        try:
            quantity = int(quantity) if quantity is not None else 0
        except:
            error_rows.append(f"第 {row_idx} 行库存数量不是数字")
            continue
        
        # 检查商品是否存在（根据条码）
        existing = cursor.execute("SELECT id FROM products WHERE barcode = ?", (barcode,)).fetchone()
        if existing:
            # 更新
            cursor.execute(
                "UPDATE products SET name=?, category=?, quantity=?, location=?, updated_at=CURRENT_TIMESTAMP WHERE barcode=?",
                (name, category, quantity, location, barcode)
            )
        else:
            # 插入
            cursor.execute(
                "INSERT INTO products (barcode, name, category, quantity, location) VALUES (?,?,?,?,?)",
                (barcode, name, category, quantity, location)
            )
        success_count += 1
    
    conn.commit()
    conn.close()
    
    return {
        "message": f"导入完成，成功处理 {success_count} 条",
        "errors": error_rows if error_rows else None
    }


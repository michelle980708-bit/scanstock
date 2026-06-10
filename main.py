from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import os
import sqlite3
from datetime import datetime

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
    cursor = conn.cursor()
    cursor.execute('''
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
    cursor.execute('''
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
    cursor.execute("SELECT COUNT(*) FROM products")
    if cursor.fetchone()[0] == 0:
        sample = [
            ("RE001", "盐酸", "试剂", 10, "试剂柜-01"),
            ("RE002", "酒精", "试剂", 20, "试剂柜-02"),
            ("CO001", "无粉手套", "耗材", 100, "耗材架-A"),
        ]
        cursor.executemany(
            "INSERT INTO products (barcode, name, category, quantity, location) VALUES (?,?,?,?,?)",
            sample
        )
    conn.commit()
    conn.close()

# 启动时初始化
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

# ---------- 根路径 ----------
@app.get("/", response_class=HTMLResponse)
def read_root():
    return HTMLResponse("<h1>ScanStock API is running</h1><p>访问 /static/simple_full.html 使用手机端</p>")

@app.get("/api/products")
def get_products():
    conn = get_db()
    rows = conn.execute("SELECT * FROM products ORDER BY id").fetchall()
    conn.close()
    return [dict(r) for r in rows]

@app.post("/api/inbound")
def inbound(req: InboundRequest):
    conn = get_db()
    product = conn.execute("SELECT id, quantity FROM products WHERE barcode = ?", (req.barcode,)).fetchone()
    if not product:
        raise HTTPException(404, "商品不存在")
    new_qty = product["quantity"] + req.quantity
    conn.execute("UPDATE products SET quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_qty, product["id"]))
    conn.execute("INSERT INTO stock_movements (product_id, type, quantity, reason) VALUES (?, 'inbound', ?, ?)",
                 (product["id"], req.quantity, req.reason))
    conn.commit()
    conn.close()
    return {"message": "入库成功", "new_quantity": new_qty}

@app.post("/api/outbound")
def outbound(req: OutboundRequest):
    conn = get_db()
    product = conn.execute("SELECT id, quantity FROM products WHERE barcode = ?", (req.barcode,)).fetchone()
    if not product:
        raise HTTPException(404, "商品不存在")
    if product["quantity"] < req.quantity:
        raise HTTPException(400, "库存不足")
    new_qty = product["quantity"] - req.quantity
    conn.execute("UPDATE products SET quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (new_qty, product["id"]))
    conn.execute("INSERT INTO stock_movements (product_id, type, quantity, reason) VALUES (?, 'outbound', ?, ?)",
                 (product["id"], -req.quantity, req.reason))
    conn.commit()
    conn.close()
    return {"message": "出库成功", "new_quantity": new_qty}

@app.post("/api/check")
def check(req: CheckRequest):
    conn = get_db()
    product = conn.execute("SELECT id, quantity FROM products WHERE barcode = ?", (req.barcode,)).fetchone()
    if not product:
        raise HTTPException(404, "商品不存在")
    old = product["quantity"]
    diff = req.actual_quantity - old
    if diff != 0:
        conn.execute("UPDATE products SET quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", (req.actual_quantity, product["id"]))
        conn.execute("INSERT INTO stock_movements (product_id, type, quantity, reason) VALUES (?, 'check', ?, ?)",
                     (product["id"], diff, req.reason))
        conn.commit()
    conn.close()
    return {"message": "盘点记录已保存", "old": old, "new": req.actual_quantity, "diff": diff}
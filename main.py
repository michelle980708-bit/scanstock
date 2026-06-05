# main.py
from fastapi import FastAPI
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
import os
from database import get_db_connection  # 新增

app = FastAPI()

# 靜態檔案
app.mount("/static", StaticFiles(directory="static"), name="static")

# 根路徑：回傳 HTML
@app.get("/", response_class=HTMLResponse)
def read_root():
    index_path = os.path.join("static", "index.html")
    if not os.path.exists(index_path):
        return HTMLResponse(content="<h1>static/index.html 不存在</h1>", status_code=404)
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

# ---------- API 端點 ----------
@app.get("/api/products")
def get_products():
    """取得所有商品列表"""
    conn = get_db_connection()
    products = conn.execute("SELECT * FROM products ORDER BY id").fetchall()
    conn.close()
    # 將 sqlite3.Row 物件轉換為字典列表
    return [dict(product) for product in products]
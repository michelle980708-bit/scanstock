from fastapi import FastAPI, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import os
import json
from typing import Optional, List, Dict, Any
from database import get_db_connection
from fastapi import UploadFile, File
import openpyxl
from io import BytesIO

app = FastAPI()

# CORS 中间件（允许手机跨域访问）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 静态文件目录
app.mount("/static", StaticFiles(directory="static"), name="static")

# ========== 请求体模型定义 ==========
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

# ========== 网页根路径 ==========
@app.get("/", response_class=HTMLResponse)
def read_root():
    index_path = os.path.join("static", "index.html")
    if not os.path.exists(index_path):
        return HTMLResponse(content="<h1>static/index.html 不存在</h1>", status_code=404)
    with open(index_path, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

# ========== API 端点 ==========
@app.get("/api/products")
def get_products():
    conn = get_db_connection()
    products = conn.execute("SELECT * FROM products ORDER BY id").fetchall()
    conn.close()
    return [dict(p) for p in products]

@app.post("/api/inbound")
def inbound(request: InboundRequest):
    conn = get_db_connection()
    product = conn.execute(
        "SELECT id, quantity FROM products WHERE barcode = ?",
        (request.barcode,)
    ).fetchone()
    if not product:
        raise HTTPException(status_code=404, detail="商品不存在")
    new_qty = product["quantity"] + request.quantity
    conn.execute(
        "UPDATE products SET quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
        (new_qty, product["id"])
    )
    conn.execute(
        "INSERT INTO stock_movements (product_id, type, quantity, reason) VALUES (?, 'inbound', ?, ?)",
        (product["id"], request.quantity, request.reason)
    )
    conn.commit()
    conn.close()
    return {"message": "入库成功", "new_quantity": new_qty}

@app.post("/api/outbound")
def outbound(request: OutboundRequest):
    conn = get_db_connection()
    product = conn.execute(
        "SELECT id, quantity FROM products WHERE barcode = ?",
        (request.barcode,)
    ).fetchone()
    if not product:
        raise HTTPException(status_code=404, detail="商品不存在")
    if product["quantity"] < request.quantity:
        raise HTTPException(status_code=400, detail="库存不足")
    new_qty = product["quantity"] - request.quantity
    conn.execute(
        "UPDATE products SET quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
        (new_qty, product["id"])
    )
    conn.execute(
        "INSERT INTO stock_movements (product_id, type, quantity, reason) VALUES (?, 'outbound', ?, ?)",
        (product["id"], -request.quantity, request.reason)
    )
    conn.commit()
    conn.close()
    return {"message": "出库成功", "new_quantity": new_qty}

@app.post("/api/check")
def check_single(request: CheckRequest):
    conn = get_db_connection()
    product = conn.execute(
        "SELECT id, quantity FROM products WHERE barcode = ?",
        (request.barcode,)
    ).fetchone()
    if not product:
        raise HTTPException(status_code=404, detail="商品不存在")
    old_qty = product["quantity"]
    diff = request.actual_quantity - old_qty
    if diff != 0:
        conn.execute(
            "UPDATE products SET quantity = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
            (request.actual_quantity, product["id"])
        )
        conn.execute(
            "INSERT INTO stock_movements (product_id, type, quantity, reason) VALUES (?, 'check', ?, ?)",
            (product["id"], diff, request.reason)
        )
        conn.commit()
    conn.close()
    return {
        "message": "盘点记录已保存",
        "barcode": request.barcode,
        "old_quantity": old_qty,
        "actual_quantity": request.actual_quantity,
        "difference": diff
    }

# ========== 分类管理 API ==========
from pydantic import BaseModel
from typing import List, Optional, Dict, Any
import json

class CategoryCreate(BaseModel):
    name: str

class CategoryUpdate(BaseModel):
    name: str

class CustomFieldCreate(BaseModel):
    field_name: str
    field_type: str
    is_required: bool = False
    select_options: Optional[List[str]] = None

class CustomFieldUpdate(BaseModel):
    field_name: str
    field_type: str
    is_required: bool = False
    select_options: Optional[List[str]] = None
    display_order: int = 0

@app.get("/api/categories")
def get_categories():
    conn = get_db_connection()
    categories = conn.execute("SELECT * FROM categories ORDER BY id").fetchall()
    conn.close()
    return [dict(c) for c in categories]

@app.post("/api/categories")
def create_category(cat: CategoryCreate):
    conn = get_db_connection()
    try:
        conn.execute("INSERT INTO categories (name) VALUES (?)", (cat.name,))
        conn.commit()
        new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.close()
        return {"id": new_id, "name": cat.name}
    except sqlite3.IntegrityError:
        raise HTTPException(status_code=400, detail="分类名称已存在")

@app.put("/api/categories/{category_id}")
def update_category(category_id: int, cat: CategoryUpdate):
    conn = get_db_connection()
    result = conn.execute("UPDATE categories SET name = ? WHERE id = ?", (cat.name, category_id))
    conn.commit()
    conn.close()
    if result.rowcount == 0:
        raise HTTPException(status_code=404, detail="分类不存在")
    return {"message": "更新成功"}

@app.delete("/api/categories/{category_id}")
def delete_category(category_id: int):
    conn = get_db_connection()
    # 检查该分类下是否有商品
    product_count = conn.execute("SELECT COUNT(*) FROM products WHERE category_id = ?", (category_id,)).fetchone()[0]
    if product_count > 0:
        raise HTTPException(status_code=400, detail="该分类下还有商品，请先删除或转移商品")
    conn.execute("DELETE FROM categories WHERE id = ?", (category_id,))
    conn.commit()
    conn.close()
    return {"message": "删除成功"}

# ========== 自定义字段管理 API ==========
@app.get("/api/categories/{category_id}/fields")
def get_fields(category_id: int):
    conn = get_db_connection()
    fields = conn.execute(
        "SELECT * FROM custom_fields WHERE category_id = ? ORDER BY display_order",
        (category_id,)
    ).fetchall()
    conn.close()
    result = []
    for f in fields:
        d = dict(f)
        if d['select_options']:
            d['select_options'] = json.loads(d['select_options'])
        result.append(d)
    return result

@app.post("/api/categories/{category_id}/fields")
def create_field(category_id: int, field: CustomFieldCreate):
    conn = get_db_connection()
    opts = json.dumps(field.select_options) if field.select_options else None
    conn.execute(
        "INSERT INTO custom_fields (category_id, field_name, field_type, is_required, select_options) VALUES (?,?,?,?,?)",
        (category_id, field.field_name, field.field_type, field.is_required, opts)
    )
    conn.commit()
    new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.close()
    return {"id": new_id, **field.dict()}

@app.put("/api/fields/{field_id}")
def update_field(field_id: int, field: CustomFieldUpdate):
    conn = get_db_connection()
    opts = json.dumps(field.select_options) if field.select_options else None
    conn.execute(
        "UPDATE custom_fields SET field_name=?, field_type=?, is_required=?, select_options=?, display_order=? WHERE id=?",
        (field.field_name, field.field_type, field.is_required, opts, field.display_order, field_id)
    )
    conn.commit()
    conn.close()
    return {"message": "更新成功"}

@app.delete("/api/fields/{field_id}")
def delete_field(field_id: int):
    conn = get_db_connection()
    conn.execute("DELETE FROM custom_fields WHERE id = ?", (field_id,))
    conn.commit()
    conn.close()
    return {"message": "删除成功"}

# ========== 动态商品 API ==========
class ProductCreate(BaseModel):
    barcode: str
    name: str
    category_id: int
    quantity: int = 0
    location: Optional[str] = None
    custom_data: Optional[Dict[str, Any]] = None

class ProductUpdate(BaseModel):
    name: Optional[str] = None
    quantity: Optional[int] = None
    location: Optional[str] = None
    custom_data: Optional[Dict[str, Any]] = None

@app.get("/api/products/category/{category_id}")
def get_products_by_category(category_id: int):
    conn = get_db_connection()
    products = conn.execute(
        "SELECT * FROM products WHERE category_id = ? ORDER BY id",
        (category_id,)
    ).fetchall()
    conn.close()
    result = []
    for p in products:
        d = dict(p)
        if d['custom_data']:
            d['custom_data'] = json.loads(d['custom_data'])
        result.append(d)
    return result

@app.post("/api/products")
def create_product(product: ProductCreate):
    conn = get_db_connection()
    custom_json = json.dumps(product.custom_data) if product.custom_data else None
    try:
        conn.execute(
            "INSERT INTO products (barcode, name, category_id, quantity, location, custom_data) VALUES (?,?,?,?,?,?)",
            (product.barcode, product.name, product.category_id, product.quantity, product.location, custom_json)
        )
        conn.commit()
        new_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
        conn.close()
        return {"id": new_id, "message": "商品创建成功"}
    except sqlite3.IntegrityError:
        raise HTTPException(status_code=400, detail="条码已存在")

@app.put("/api/products/{product_id}")
def update_product(product_id: int, product: ProductUpdate):
    conn = get_db_connection()
    updates = []
    params = []
    if product.name is not None:
        updates.append("name = ?")
        params.append(product.name)
    if product.quantity is not None:
        updates.append("quantity = ?")
        params.append(product.quantity)
    if product.location is not None:
        updates.append("location = ?")
        params.append(product.location)
    if product.custom_data is not None:
        updates.append("custom_data = ?")
        params.append(json.dumps(product.custom_data))
    if not updates:
        return {"message": "无更新内容"}
    updates.append("updated_at = CURRENT_TIMESTAMP")
    query = f"UPDATE products SET {', '.join(updates)} WHERE id = ?"
    params.append(product_id)
    conn.execute(query, params)
    conn.commit()
    conn.close()
    return {"message": "更新成功"}

@app.delete("/api/products/{product_id}")
def delete_product(product_id: int):
    conn = get_db_connection()
    conn.execute("DELETE FROM products WHERE id = ?", (product_id,))
    conn.commit()
    conn.close()
    return {"message": "删除成功"}

@app.post("/api/import/excel/{category_id}")
async def import_excel(category_id: int, file: UploadFile = File(...)):
    # 读取Excel，根据category_id获取字段定义，解析并插入商品
    contents = await file.read()
    wb = openpyxl.load_workbook(BytesIO(contents))
    sheet = wb.active
    # 简化：假设第一行是列名，后面是数据
    # 需要实现字段映射逻辑
    return {"message": "导入功能示例"}

@app.get("/api/export/excel/{category_id}")
def export_excel(category_id: int):
    # 导出该分类下所有商品为Excel
    conn = get_db_connection()
    products = conn.execute("SELECT * FROM products WHERE category_id = ?", (category_id,)).fetchall()
    conn.close()
    wb = openpyxl.Workbook()
    ws = wb.active
    # 写入表头...
    # 返回 StreamingResponse
    from fastapi.responses import StreamingResponse
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=products.xlsx"})
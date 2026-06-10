import io
import os
import sqlite3
import json
import re
from datetime import datetime

import openpyxl
import pandas as pd
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.mount("/static", StaticFiles(directory="static"), name="static")

DB_NAME = "scanstock.db"

def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def safe_str(val):
    """将任何值安全转换为字符串，避免编码错误"""
    if val is None:
        return ""
    if isinstance(val, bytes):
        # 尝试常见编码
        for encoding in ['utf-8', 'gbk', 'big5', 'latin1']:
            try:
                return val.decode(encoding)
            except UnicodeDecodeError:
                continue
        # 最后丢弃无法解码的字符
        return val.decode('utf-8', errors='ignore')
    if isinstance(val, (int, float, datetime)):
        return str(val)
    if isinstance(val, str):
        return val
    return str(val)

def generate_field_name(display_name: str, existing_names: list) -> str:
    """根据显示名称生成唯一的英文字段名"""
    # 确保 display_name 是字符串
    display_name = safe_str(display_name)
    # 只保留字母数字和下划线，中文转为拼音？不，保留原样但转换特殊字符
    # 简单处理：将非字母数字替换为下划线
    base = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fff]', '_', display_name)
    base = re.sub(r'_+', '_', base).strip('_')
    if not base:
        base = "field"
    name = base
    counter = 1
    while name in existing_names:
        name = f"{base}_{counter}"
        counter += 1
    return name

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
    try:
        c.execute("ALTER TABLE products ADD COLUMN custom_fields TEXT DEFAULT '{}'")
    except:
        pass
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

class FieldConfigItem(BaseModel):
    display_name: str
    field_type: str = "text"
    is_required: bool = False

# ---------- 根页面 ----------
@app.get("/", response_class=HTMLResponse)
def read_root():
    return HTMLResponse("<h1>ScanStock API is running</h1><p>访问 /static/index.html 使用手机端</p>")

# ---------- 商品 API ----------
@app.get("/api/products")
def get_products(category: str = None):
    conn = get_db()
    if category:
        rows = conn.execute("SELECT * FROM products WHERE category = ? ORDER BY id", (category,)).fetchall()
    else:
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

# ---------- Excel 导入导出 ----------
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
        row = [p["barcode"], p["name"], p["category"], p["quantity"], p["location"] or ""]
        for f in fields:
            row.append(custom.get(f["field_name"], ""))
        ws.append(row)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": f"attachment; filename={category or 'products'}.xlsx"})

@app.post("/api/import/excel")
async def import_products_from_excel(file: UploadFile = File(...), category: str = None):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "只支持 .xlsx 或 .xls 文件")
    df = pd.read_excel(io.BytesIO(contents), engine='openpyxl')
    # 将 NaN 替换为空字符串
    df = df.fillna("")
    contents = await file.read()
    wb = openpyxl.load_workbook(io.BytesIO(contents))
    ws = wb.active
    # 安全获取表头
    headers = [safe_str(cell.value) for cell in ws[1]]
    base_mapping = {"条码": "barcode", "商品名称": "name", "分类": "category", "库存数量": "quantity", "存放位置": "location"}
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
        safe_row = [safe_str(cell) for cell in row]
        product_data = {}
        custom_data = {}
        for col_idx, val in enumerate(safe_row):
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
            conn.execute("UPDATE products SET name=?, category=?, quantity=?, location=?, custom_fields=?, updated_at=CURRENT_TIMESTAMP WHERE barcode=?",
                         (name, category_val, quantity, location, json.dumps(custom_data), barcode))
        else:
            conn.execute("INSERT INTO products (barcode, name, category, quantity, location, custom_fields) VALUES (?,?,?,?,?,?)",
                         (barcode, name, category_val, quantity, location, json.dumps(custom_data)))
        success_count += 1
    conn.commit()
    conn.close()
    return {"message": f"导入完成，成功 {success_count} 条", "errors": error_rows or None}

# ---------- 字段配置 API ----------
@app.post("/api/field_configs/{category}")
def save_field_configs(category: str, configs: list[FieldConfigItem]):
    try:
        conn = get_db()
        existing = conn.execute("SELECT field_name FROM field_configs WHERE category = ?", (category,)).fetchall()
        existing_names = [r["field_name"] for r in existing]
        conn.execute("DELETE FROM field_configs WHERE category = ?", (category,))
        for idx, cfg in enumerate(configs):
            field_name = generate_field_name(cfg.display_name, existing_names)
            conn.execute(
                "INSERT INTO field_configs (category, field_name, display_name, field_type, is_required, sort_order) VALUES (?,?,?,?,?,?)",
                (category, field_name, cfg.display_name, cfg.field_type, 1 if cfg.is_required else 0, idx)
            )
            existing_names.append(field_name)
        conn.commit()
        conn.close()
        return {"message": "保存成功"}
    except Exception as e:
        raise HTTPException(500, detail=f"保存失败: {str(e)}")

@app.post("/api/field_configs/import_excel")
async def import_field_configs_from_excel(file: UploadFile = File(...), category: str = None):
    if not category:
        raise HTTPException(400, "缺少分类参数")
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "只支持 .xlsx 或 .xls 文件")
    try:
        # 读取 Excel 文件到 DataFrame
        contents = await file.read()
        # 使用 pandas 读取，它会自动处理编码问题
        df = pd.read_excel(io.BytesIO(contents), engine='openpyxl')
        
        # 检查必要的列
        required_col = "显示名称"
        if required_col not in df.columns:
            raise HTTPException(400, f"Excel 缺少“{required_col}”列")
        
        configs = []
        for idx, row in df.iterrows():
            display_name = str(row.get("显示名称", "")).strip()
            if not display_name:
                continue
            field_type = "text"
            if "类型" in df.columns and pd.notna(row.get("类型")):
                ft = str(row["类型"]).lower()
                if ft in ["text", "number", "date"]:
                    field_type = ft
            is_required = False
            if "必填" in df.columns and pd.notna(row.get("必填")):
                val = str(row["必填"]).lower()
                is_required = val in ["是", "true", "1", "yes"]
            configs.append(FieldConfigItem(
                display_name=display_name,
                field_type=field_type,
                is_required=is_required
            ))
        
        # 保存配置到数据库（复用原有逻辑）
        conn = get_db()
        existing = conn.execute("SELECT field_name FROM field_configs WHERE category = ?", (category,)).fetchall()
        existing_names = [r["field_name"] for r in existing]
        conn.execute("DELETE FROM field_configs WHERE category = ?", (category,))
        for idx, cfg in enumerate(configs):
            field_name = generate_field_name(cfg.display_name, existing_names)
            conn.execute(
                "INSERT INTO field_configs (category, field_name, display_name, field_type, is_required, sort_order) VALUES (?,?,?,?,?,?)",
                (category, field_name, cfg.display_name, cfg.field_type, 1 if cfg.is_required else 0, idx)
            )
            existing_names.append(field_name)
        conn.commit()
        conn.close()
        return {"message": f"导入成功，共 {len(configs)} 个字段"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, detail=f"导入失败: {str(e)}")
    if not category:
        raise HTTPException(400, "缺少分类参数")
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "只支持 .xlsx 或 .xls 文件")
    try:
        contents = await file.read()
        wb = openpyxl.load_workbook(io.BytesIO(contents))
        ws = wb.active
        # 安全获取表头
        headers = [safe_str(cell.value) for cell in ws[1]]
        col_idx = {}
        for idx, h in enumerate(headers):
            if h in ["显示名称", "类型", "必填"]:
                col_idx[h] = idx
        if "显示名称" not in col_idx:
            raise HTTPException(400, "Excel 缺少“显示名称”列")
        configs = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row):
                continue
            # 安全处理每个单元格
            safe_row = [safe_str(cell) for cell in row]
            display_name = safe_row[col_idx["显示名称"]] if len(safe_row) > col_idx["显示名称"] else ""
            if not display_name:
                continue
            display_name = display_name.strip()
            field_type = "text"
            if "类型" in col_idx and len(safe_row) > col_idx["类型"] and safe_row[col_idx["类型"]]:
                ft = safe_row[col_idx["类型"]].lower()
                if ft in ["text", "number", "date"]:
                    field_type = ft
            is_required = False
            if "必填" in col_idx and len(safe_row) > col_idx["必填"] and safe_row[col_idx["必填"]]:
                val = safe_row[col_idx["必填"]].lower()
                is_required = val in ["是", "true", "1", "yes"]
            configs.append(FieldConfigItem(
                display_name=display_name,
                field_type=field_type,
                is_required=is_required
            ))
        # 保存配置
        conn = get_db()
        existing = conn.execute("SELECT field_name FROM field_configs WHERE category = ?", (category,)).fetchall()
        existing_names = [r["field_name"] for r in existing]
        conn.execute("DELETE FROM field_configs WHERE category = ?", (category,))
        for idx, cfg in enumerate(configs):
            field_name = generate_field_name(cfg.display_name, existing_names)
            conn.execute(
                "INSERT INTO field_configs (category, field_name, display_name, field_type, is_required, sort_order) VALUES (?,?,?,?,?,?)",
                (category, field_name, cfg.display_name, cfg.field_type, 1 if cfg.is_required else 0, idx)
            )
            existing_names.append(field_name)
        conn.commit()
        conn.close()
        return {"message": f"导入成功，共 {len(configs)} 个字段"}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, detail=f"导入失败: {str(e)}")
    
@app.get("/api/field_configs/template")
def download_field_config_template():
    # 创建一个新的 Excel 工作簿
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "字段配置模板"
    
    # 写入表头
    ws.append(["显示名称", "类型", "必填"])
    
    # 写入示例数据（可选，方便用户理解）
    ws.append(["有效期", "date", "是"])
    ws.append(["保存温度", "number", "否"])
    ws.append(["备注", "text", "否"])
    
    # 将工作簿保存到内存中的字节流
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    # 返回 StreamingResponse，告诉浏览器这是一个附件下载
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=field_config_template.xlsx"}
    )

@app.get("/api/field_configs/{category}")
def get_field_configs(category: str):
    conn = get_db()
    rows = conn.execute("SELECT * FROM field_configs WHERE category = ? ORDER BY sort_order", (category,)).fetchall()
    conn.close()
    return [{"display_name": r["display_name"], "field_type": r["field_type"], "is_required": bool(r["is_required"])} for r in rows]

@app.get("/api/export/product_template")
def download_product_template():
    """下载商品导入模板（基础字段）"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "商品导入模板"
    headers = ["条码", "商品名称", "分类", "库存数量", "存放位置"]
    ws.append(headers)
    # 示例行
    ws.append(["RE001", "示例商品", "试剂", "10", "位置A"])
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=product_import_template.xlsx"}
    )
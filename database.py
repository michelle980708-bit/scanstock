# database.py
import sqlite3
import json
from datetime import datetime

DB_NAME = "scanstock.db"

def get_db_connection():
    """取得資料庫連線"""
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row  # 讓查詢結果可以用欄位名稱存取
    return conn

def init_db():
    """初始化資料庫：建立需要的表格"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 商品表格（暫時只有基本欄位，後續會依分類擴充）
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            barcode TEXT UNIQUE NOT NULL,      -- 條碼（唯一）
            name TEXT NOT NULL,                 -- 商品名稱
            category TEXT NOT NULL,             -- 分類（試劑/耗材/物品）
            quantity INTEGER DEFAULT 0,         -- 庫存數量
            location TEXT,                      -- 存放位置（中文）
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # 插入一些範例資料（讓資料庫不空）
    cursor.execute("SELECT COUNT(*) FROM products")
    count = cursor.fetchone()[0]
    if count == 0:
        sample_products = [
            ("RE001", "鹽酸", "試劑", 10, "試劑櫃-01"),
            ("RE002", "酒精", "試劑", 20, "試劑櫃-02"),
            ("CO001", "無粉手套", "耗材", 100, "耗材架-A"),
        ]
        cursor.executemany(
            "INSERT INTO products (barcode, name, category, quantity, location) VALUES (?, ?, ?, ?, ?)",
            sample_products
        )
        print("已插入範例商品資料")
    
    conn.commit()
    conn.close()
    print("資料庫初始化完成")

# 如果直接執行這個檔案，就執行初始化
if __name__ == "__main__":
    init_db()
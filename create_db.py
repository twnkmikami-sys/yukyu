import sqlite3

# =====================================
# DB接続
# =====================================

conn = sqlite3.connect("leave.db")

cursor = conn.cursor()

# =====================================
# 社員マスタ
# =====================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS employees (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    employee_id TEXT,
    name TEXT,
    kana TEXT,
    hire_date TEXT,
    branch TEXT
)
""")

# =====================================
# 有給データ
# =====================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS leave_data (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    employee_id TEXT,
    name TEXT,

    granted_days REAL,
    used_days REAL,
    remain_days REAL
)
""")

# =====================================
# 有給取得履歴
# =====================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS leave_history (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    employee_id TEXT,
    name TEXT,

    leave_date TEXT,
    leave_days REAL
)
""")

# =====================================
# 有給付与履歴
# =====================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS grant_history (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    employee_id TEXT,
    name TEXT,

    grant_date TEXT,
    grant_days REAL,

    expire_date TEXT,

    used_days REAL,

    expired_flag TEXT
)
""")

# =====================================
# 出勤データ
# =====================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS attendance (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    employee_id TEXT,
    name TEXT,

    work_date TEXT,

    attendance_type TEXT,

    attendance_days REAL
)
""")

# =====================================
# 自動付与管理
# =====================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS auto_grants (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    employee_id TEXT,

    grant_year INTEGER,

    grant_date TEXT
)
""")

# =====================================
# ユーザー管理
# =====================================

cursor.execute("""
CREATE TABLE IF NOT EXISTS users (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    username TEXT,

    password TEXT,

    role TEXT,

    employee_id TEXT
)
""")

# =====================================
# 保存
# =====================================

conn.commit()

conn.close()

print("DB作成完了")
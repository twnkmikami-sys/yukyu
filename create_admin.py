
import sqlite3
import hashlib

conn = sqlite3.connect("leave.db")

cursor = conn.cursor()

# パスワード
password = "admin123"

hash_pass = hashlib.sha256(
    password.encode()
).hexdigest()

cursor.execute("""
INSERT INTO users (

    username,
    password,
    role,
    employee_id

)
VALUES (?, ?, ?, ?)
""", (

    "admin",
    hash_pass,
    "admin",
    "ADMIN"

))

conn.commit()

print("管理者作成完了")


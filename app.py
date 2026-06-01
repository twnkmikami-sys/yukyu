import streamlit as st
import hashlib
import pandas as pd
import sqlite3
import zipfile
import io
import os
import requests
from datetime import datetime, timedelta
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.worksheet.page import PageMargins
from openpyxl import Workbook
from openpyxl.utils import get_column_letter

# =====================================
# ⚠️ ここに自分のBOXアップロードURLを入力！
# =====================================
BOX_UPLOAD_URL = "https://town-security.box.com/s/wn5iypq63p183l3wwl03due4m0lh9fo7"  # ← ここを自分のURLに置き換えてください

# =====================================
# DB接続
# =====================================
conn = sqlite3.connect(
    "leave.db",
    check_same_thread=False
)
cursor = conn.cursor()

# =====================================
# テーブル作成
# =====================================
cursor.execute("""
CREATE TABLE IF NOT EXISTS employees (
    employee_id TEXT,
    name TEXT,
    kana TEXT,
    hire_date TEXT,
    branch TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS leave_data (
    employee_id TEXT,
    name TEXT,
    granted_days REAL,
    used_days REAL,
    remain_days REAL
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS leave_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id TEXT,
    name TEXT,
    leave_date TEXT,
    leave_days REAL
)
""")

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

cursor.execute("""
CREATE TABLE IF NOT EXISTS grant_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id TEXT,
    name TEXT,
    grant_date TEXT,
    expire_date TEXT,
    used_days REAL,
    expired_flag TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS auto_grants (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    employee_id TEXT,
    grant_year REAL,
    grant_date TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT,
    password TEXT,
    role TEXT,
    employee_id TEXT
)
""")
conn.commit()

# =====================================
# ログイン処理
# =====================================
st.title("有給休暇管理システム")
login_user = st.text_input(
    "ユーザー名",
    key="login_user"
)
login_pass = st.text_input(
    "パスワード",
    type="password",
    key="login_pass"
)

if st.button("ログイン"):
    hash_pass = hashlib.sha256(
        login_pass.encode()
    ).hexdigest()
    user_df = pd.read_sql(
        "SELECT * FROM users WHERE username=? AND password=?",
        conn,
        params=(login_user, hash_pass)
    )
    if len(user_df) > 0:
        st.session_state["login"] = True
        st.session_state["role"] = user_df.iloc[0]["role"]
        st.session_state["login_emp_id"] = user_df.iloc[0]["employee_id"]
        st.success("ログイン成功")
        st.rerun()
    else:
        st.error("ユーザー名またはパスワードが違います")

# ログイン状態チェック
if "login" not in st.session_state or not st.session_state["login"]:
    st.stop()

# =====================================
# メインメニュー
# =====================================
st.sidebar.title("メニュー")
menu = st.sidebar.selectbox(
    "機能を選択",
    ["Excel個人台帳出力", "全社員Excel一括出力"]
)

# =====================================
# 📤 BOXへファイルアップロード共通関数
# =====================================
def upload_to_box(buffer, filename):
    """メモリ上のファイルデータをBOXのアップロードURLへ送信"""
    try:
        buffer.seek(0)
        files = {
            'file': (
                filename,
                buffer,
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
                if filename.endswith('.xlsx') else 'application/zip'
            )
        }
        response = requests.post(BOX_UPLOAD_URL, files=files, timeout=30)
        if response.status_code == 200:
            return True, f"✅ BOXへ自動保存完了: {filename}"
        else:
            return False, f"⚠️ BOX保存失敗(コード:{response.status_code})"
    except Exception as e:
        return False, f"⚠️ BOX保存エラー: {str(e)}"

# =====================================
# Excel個人台帳出力（プレビュー+BOX自動保存）
# =====================================
if menu == "Excel個人台帳出力":
    st.header("Excel個人台帳出力")

    # 状態を記憶：データ保持用
    if "excel_data" not in st.session_state:
        st.session_state["excel_data"] = None
    if "excel_filename" not in st.session_state:
        st.session_state["excel_filename"] = None
    if "preview_personal" not in st.session_state:
        st.session_state["preview_personal"] = None

    # 社員選択
    emp_list_df = pd.read_sql("SELECT employee_id, name FROM employees", conn)
    emp_list = emp_list_df["name"].tolist()
    selected_employee = st.selectbox("社員を選択", emp_list)
    employee_id = emp_list_df[emp_list_df["name"] == selected_employee]["employee_id"].iloc[0]

    if st.button("Excel作成", key="excel_output"):
        # 各種データ取得
        employee_row = pd.read_sql("SELECT * FROM employees WHERE employee_id=?", conn, params=(employee_id,))
        employee_leave = pd.read_sql("SELECT * FROM leave_data WHERE employee_id=?", conn, params=(employee_id,))
        employee_history = pd.read_sql("SELECT * FROM leave_history WHERE employee_id=?", conn, params=(employee_id,))
        employee_attendance = pd.read_sql("SELECT * FROM attendance WHERE employee_id=?", conn, params=(employee_id,))
        grant_export = pd.read_sql("SELECT * FROM grant_history WHERE employee_id=?", conn, params=(employee_id,))

        # Workbook 作成
        wb = Workbook()
        ws = wb.active
        ws.title = "有給台帳"

        # 色設定
        yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
        red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
        green_fill = PatternFill(start_color="99FF99", end_color="99FF99", fill_type="solid")

        # 印刷設定
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)

        # タイトル
        ws.merge_cells("A1:E1")
        ws["A1"] = "有給休暇管理台帳"
        ws["A1"].font = Font(bold=True, size=18)
        ws["A1"].alignment = Alignment(horizontal="center")

        # 社員情報
        ws["A3"] = "社員ID"
        ws["B3"] = employee_id
        ws["A4"] = "氏名"
        ws["B4"] = selected_employee
        ws["A5"] = "所属"
        ws["B5"] = employee_row.iloc[0]["branch"]
        ws["A6"] = "入社日"
        ws["B6"] = employee_row.iloc[0]["hire_date"]

        # 有給情報
        leave_row = employee_leave.iloc[0]
        ws["D3"] = "付与日数"
        ws["E3"] = float(leave_row["granted_days"])
        ws["D4"] = "使用日数"
        ws["E4"] = float(leave_row["used_days"])
        ws["D5"] = "残日数"
        ws["E5"] = float(leave_row["remain_days"])

        # 出勤率計算
        attendance_rate = 0
        if len(employee_attendance) > 0:
            work_days = employee_attendance[
                employee_attendance["attendance_type"].isin(["出勤", "有給"])
            ]["attendance_days"].sum()
            total_days = employee_attendance[
                employee_attendance["attendance_type"] != "休日"
            ]["attendance_days"].sum()
            if total_days > 0:
                attendance_rate = round(work_days / total_days * 100, 1)

        ws["D9"] = "出勤率"
        ws["E9"] = f"{attendance_rate}%"
        if attendance_rate >= 80:
            ws["E9"].fill = green_fill
        else:
            ws["E9"].fill = red_fill

        # 有給履歴
        ws["A9"] = "取得日"
        ws["B9"] = "取得日数"
        row_no = 10
        history_data = []
        for _, row in employee_history.iterrows():
            ws.cell(row=row_no, column=1).value = row["leave_date"]
            ws.cell(row=row_no, column=2).value = row["leave_days"]
            history_data.append({"取得日": row["leave_date"], "取得日数": row["leave_days"]})
            row_no += 1

        # 消滅警告
        warning_list = []
        warning_row = max(row_no + 2, 15)
        ws[f"D{warning_row}"] = "消滅予定"
        today_date = datetime.today().date()
        warning_row += 1
        for _, row in grant_export.iterrows():
            if row["expired_flag"] == "消滅済":
                continue
            expire_date = datetime.strptime(row["expire_date"], "%Y-%m-%d").date()
            remain_days = float(row["grant_days"]) - float(row["used_days"])
            if remain_days <= 0:
                continue
            days_left = (expire_date - today_date).days
            if days_left <= 30:
                ws[f"D{warning_row}"] = "期限接近"
                ws[f"E{warning_row}"] = str(expire_date)
                ws[f"F{warning_row}"] = remain_days
                ws[f"D{warning_row}"].fill = yellow_fill
                ws[f"E{warning_row}"].fill = yellow_fill
                ws[f"F{warning_row}"].fill = yellow_fill
                warning_list.append({"状態": "期限接近", "期限日": expire_date, "残日数": remain_days})
                warning_row += 1

        # 列幅
        ws.column_dimensions["A"].width = 20
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 5
        ws.column_dimensions["D"].width = 20
        ws.column_dimensions["E"].width = 20
        ws.column_dimensions["F"].width = 15

        # 枠線
        thin = Side(border_style="thin", color="000000")
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=6):
            for cell in row:
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # --- プレビュー用データを保存 ---
        st.session_state["preview_personal"] = {
            "社員情報": [
                {"項目": "社員ID", "内容": employee_id},
                {"項目": "氏名", "内容": selected_employee},
                {"項目": "所属", "内容": employee_row.iloc[0]["branch"]},
                {"項目": "入社日", "内容": employee_row.iloc[0]["hire_date"]}
            ],
            "有給情報": [
                {"項目": "付与日数", "内容": float(leave_row["granted_days"])},
                {"項目": "使用日数", "内容": float(leave_row["used_days"])},
                {"項目": "残日数", "内容": float(leave_row["remain_days"])},
                {"項目": "出勤率", "内容": f"{attendance_rate}%"}
            ],
            "有給取得履歴": pd.DataFrame(history_data),
            "消滅警告リスト": pd.DataFrame(warning_list)
        }

        # Excelデータをメモリ保存
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        filename = f"{selected_employee}_有給台帳_{datetime.today().strftime('%Y%m%d')}.xlsx"
        st.session_state["excel_data"] = excel_buffer
        st.session_state["excel_filename"] = filename

        # --- 📤 BOXへ自動保存実行 ---
        success, msg = upload_to_box(excel_buffer, filename)
        if success:
            st.success("✅ Excel作成完了")
            st.info(msg)
        else:
            st.warning(msg)

    # --- プレビュー表示エリア ---
    if st.session_state["preview_personal"] is not None:
        st
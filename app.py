"""
แอปฟอร์มลงทะเบียนข้อมูลพนักงาน
ให้พนักงานกรอกข้อมูลเอง แล้วบันทึกลงไฟล์ Excel (CSR_NAME.xlsx) บน GitHub โดยตรง
ผ่าน GitHub API เพื่อให้ข้อมูลอยู่ถาวร ไม่หายเมื่อแอป restart
"""

import base64
import io
import re
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

# ---------- ตั้งค่า GitHub repo ที่จะใช้เก็บไฟล์ ----------
GITHUB_OWNER = "ITBASz"
GITHUB_REPO = "CSR"
GITHUB_BRANCH = "main"
GITHUB_FILE_PATH = "CSR_NAME.xlsx"
GITHUB_API_URL = (
    f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/contents/{GITHUB_FILE_PATH}"
)

SHEET_NAME = "พนักงาน"

COLUMNS = [
    "วันที่-เวลาลงทะเบียน",
    "ชื่อ-นามสกุล",
    "ชื่อแผนก",
    "Size เสื้อ",
    "เบอร์โทร",
    "โรคประจำตัว",
]

SHIRT_SIZES = ["S", "M", "L", "XL", "2XL", "3XL"]

DEPARTMENTS = [
    "ทรัพยากรมนุษย์",
    "คลังสินค้า",
    "บัญชีและการเงิน",
    "ขายและการตลาด",
    "วิจัยและพัฒนา",
    "ประกันคุณภาพ",
    "จัดซื้อ",
    "ความปลอดภัย ฯ",
    "วิศวกรรม",
    "ผลิต",
    "ออฟฟิศผลิต",
    "คอปเปอร์ซัลเฟต",
    "คอปเปอร์ออกไซด์",
    "ผลิตภัณฑ์พิเศษ",
    "แอลเอฟซี LFC",
    "คาบอเนต",
]

st.set_page_config(page_title="ลงทะเบียนข้อมูลพนักงาน", page_icon="📝", layout="centered")


def github_headers() -> dict:
    """ส่วนหัว request สำหรับเรียก GitHub API พร้อม token จาก Secrets"""
    token = st.secrets["github_token"]
    return {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github+json",
    }


def load_data():
    """ดึงไฟล์ Excel ล่าสุดจาก GitHub มาอ่าน คืนค่า (DataFrame, sha ของไฟล์)
    ถ้ายังไม่มีไฟล์บน GitHub ให้คืนตารางเปล่าและ sha เป็น None"""
    resp = requests.get(
        GITHUB_API_URL, headers=github_headers(), params={"ref": GITHUB_BRANCH}
    )
    if resp.status_code == 404:
        return pd.DataFrame(columns=COLUMNS), None
    resp.raise_for_status()
    data = resp.json()
    content = base64.b64decode(data["content"])
    df = pd.read_excel(io.BytesIO(content), sheet_name=SHEET_NAME, dtype=str).fillna("")
    return df, data["sha"]


def save_data(df: pd.DataFrame, sha):
    """เขียนตารางทั้งหมดกลับเป็นไฟล์ Excel แล้ว commit ขึ้น GitHub ทันที"""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=SHEET_NAME)

    payload = {
        "message": f"อัปเดตข้อมูลพนักงาน {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "content": base64.b64encode(buffer.getvalue()).decode("utf-8"),
        "branch": GITHUB_BRANCH,
    }
    if sha:
        payload["sha"] = sha

    resp = requests.put(GITHUB_API_URL, headers=github_headers(), json=payload)
    resp.raise_for_status()


def is_valid_phone(phone: str) -> bool:
    """ตรวจสอบรูปแบบเบอร์โทรแบบหลวมๆ: ตัวเลข 9-10 หลัก อาจมีขีดคั่นได้"""
    digits_only = re.sub(r"[-\s]", "", phone)
    return digits_only.isdigit() and 9 <= len(digits_only) <= 10


st.title("📝 ลงทะเบียนข้อมูลพนักงาน")
st.write("กรุณากรอกข้อมูลของท่านให้ครบถ้วน")

with st.form("registration_form", clear_on_submit=True):
    full_name = st.text_input("ชื่อ-นามสกุล *", placeholder="เช่น สมชาย ใจดี")
    department = st.selectbox("ชื่อแผนก *", options=DEPARTMENTS)
    shirt_size = st.selectbox("Size เสื้อ *", options=SHIRT_SIZES, index=2)
    phone = st.text_input("เบอร์โทร *", placeholder="เช่น 0812345678")
    medical_condition = st.text_area(
        "โรคประจำตัว (ถ้ามี)",
        placeholder="กรอกเฉพาะถ้ามี หากไม่มีเว้นว่างไว้ได้",
        help="ข้อมูลนี้จะถูกเก็บเป็นความลับ ใช้เพื่อการดูแลด้านความปลอดภัยเท่านั้น",
    )

    submitted = st.form_submit_button("✅ บันทึกข้อมูล", use_container_width=True)

    if submitted:
        errors = []
        if not full_name.strip():
            errors.append("กรุณากรอกชื่อ-นามสกุล")
        if not phone.strip():
            errors.append("กรุณากรอกเบอร์โทร")
        elif not is_valid_phone(phone):
            errors.append("รูปแบบเบอร์โทรไม่ถูกต้อง กรุณากรอกตัวเลข 9-10 หลัก")

        if errors:
            for e in errors:
                st.error(e)
        else:
            try:
                df, sha = load_data()
                new_row = {
                    "วันที่-เวลาลงทะเบียน": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "ชื่อ-นามสกุล": full_name.strip(),
                    "ชื่อแผนก": department.strip(),
                    "Size เสื้อ": shirt_size,
                    "เบอร์โทร": phone.strip(),
                    "โรคประจำตัว": medical_condition.strip(),
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                save_data(df, sha)
                st.success(f"บันทึกข้อมูลของคุณ {full_name.strip()} เรียบร้อยแล้ว 🎉")
            except Exception as e:
                st.error(f"บันทึกข้อมูลไม่สำเร็จ กรุณาลองใหม่อีกครั้ง ({e})")

st.divider()
try:
    st.caption(f"จำนวนผู้ลงทะเบียนทั้งหมด: {len(load_data()[0])} คน")
except Exception:
    st.caption("ไม่สามารถโหลดจำนวนผู้ลงทะเบียนได้ในขณะนี้")

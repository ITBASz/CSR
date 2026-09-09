"""
แอปฟอร์มลงทะเบียนข้อมูลพนักงาน
ให้พนักงานกรอกข้อมูลเอง แล้วบันทึกลงไฟล์ Excel (employees.xlsx)
ถ้ายังไม่มีไฟล์ แอปจะสร้างไฟล์ใหม่ให้อัตโนมัติตอนมีคนกรอกครั้งแรก
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import os
import re

EXCEL_FILE = "CSR_NAME.xlsx"
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


def load_data() -> pd.DataFrame:
    """โหลดข้อมูลเดิมจากไฟล์ ถ้ายังไม่มีไฟล์ให้คืนตารางเปล่า"""
    if not os.path.exists(EXCEL_FILE):
        return pd.DataFrame(columns=COLUMNS)
    return pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, dtype=str).fillna("")


def save_data(df: pd.DataFrame):
    """บันทึกตารางทั้งหมดกลับลงไฟล์ Excel (สร้างไฟล์ใหม่ถ้ายังไม่มี)"""
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False, sheet_name=SHEET_NAME)


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
        if not department.strip():
            errors.append("กรุณากรอกชื่อแผนก")
        if not phone.strip():
            errors.append("กรุณากรอกเบอร์โทร")
        elif not is_valid_phone(phone):
            errors.append("รูปแบบเบอร์โทรไม่ถูกต้อง กรุณากรอกตัวเลข 9-10 หลัก")

        if errors:
            for e in errors:
                st.error(e)
        else:
            df = load_data()
            new_row = {
                "วันที่-เวลาลงทะเบียน": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "ชื่อ-นามสกุล": full_name.strip(),
                "ชื่อแผนก": department.strip(),
                "Size เสื้อ": shirt_size,
                "เบอร์โทร": phone.strip(),
                "โรคประจำตัว": medical_condition.strip(),
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            save_data(df)
            st.success(f"บันทึกข้อมูลของคุณ {full_name.strip()} เรียบร้อยแล้ว 🎉")

st.divider()
st.caption(f"จำนวนผู้ลงทะเบียนทั้งหมด: {len(load_data())} คน")

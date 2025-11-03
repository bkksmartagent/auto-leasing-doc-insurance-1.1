import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx2pdf import convert
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import tempfile
import os
import re
import io
from datetime import datetime, timedelta
from babel.dates import format_date
from num2words import num2words
from PIL import Image
import uuid
from modules.helpers import (
    replace_text_in_paragraphs, replace_text_in_tables, replace_image_placeholder_in_paragraphs,
    parse_thai_date_str, convert_en_date_to_thai, ordinal, date_data, bahttext,
    get_image_size, safe_filename
)

def contract_form():
    data = {}
    st.header("Leasing Contract and Continue Leasing")
    st.markdown("**General Details**")
    data["start_date"] = st.text_input("วันเริ่มสัญญา", placeholder="เช่น 1 มกราคม 2568", key="start_date")
    data["end_date"] = st.text_input("วันสิ้นสุดสัญญา", placeholder="เช่น 31 ธันวาคม 2569", key="end_date")
    data["en_contract_period"] = st.text_input("Contract period", placeholder="e.g. 1 Year and 6 Months", key="en_contract_period")
    data["th_contract_period"] = st.text_input("ระยะเวลาสัญญา", placeholder="เช่น 1 ปี 6 เดือน", key="th_contract_period")

    st.markdown("**Landlord Personal Details**")
    data["landlord_en_name"] = st.text_input("Name", placeholder="e.g. Mr.Name Surname", key="landlord_en_name")
    data["landlord_th_name"] = st.text_input("ชื่อผู้ให้เช่า", placeholder="e.g. นายชื่อ นามสกุล", key="landlord_th_name")
    data["landlord_idcard"] = st.text_input("ID card/passport no.", placeholder="e.g. 1234435567832", key="landlord_idcard")
    data["landlord_en_address"] = st.text_input("Address", placeholder="e.g. 111 This Road, That Sub-District and District, Bangkok 10110", key="landlord_en_address")
    data["landlord_th_address"] = st.text_input("ที่อยู่ตามบัตรประชาชน", placeholder="e.g. 111 ถนนที่นี้ แขวงที่นั้น เขตที่นู้น กรุงเทพฯ 10110", key="landlord_th_address")
    data["photo1"] = st.file_uploader("แนบรูปบัตรประชาชน (ผู้ให้เช่า)", type=["jpg", "jpeg", "png"])

    st.markdown("**Tenant Personal Details**")
    data["tenant_en_name"] = st.text_input("Name", placeholder="e.g. Ms.Name Surname", key="tenant_en_name")
    data["tenant_th_name"] = st.text_input("ชื่อผู้เช่า", placeholder="e.g. นางสาวชื่อ นามสกุล", key="tenant_th_name")
    data["tenant_idcard"] = st.text_input("ID card/passport no.", placeholder="e.g. 1234435567832", key="tenant_idcard")
    data["tenant_en_nationality"] = st.text_input("Nationality", placeholder="e.g. Thai", key="tenant_en_nationality")
    data["tenant_th_nationality"] = st.text_input("สัญชาติ", placeholder="e.g. ไทย", key="tenant_th_nationality")
    data["tenant_en_address"] = st.text_input("Address", placeholder="e.g. 111 This Road, That Sub-District and District, Bangkok 10110", key="tenant_en_address")
    data["tenant_th_address"] = st.text_input("ที่อยู่ติดต่อได้", placeholder="e.g. 111 ถนนที่นี้ แขวงที่นั้น เขตที่นู้น กรุงเทพฯ 10110", key="tenant_th_address")
    data["photo2"] = st.file_uploader("แนบรูปบัตรประชาชน (ผู้เช่า)", type=["jpg", "jpeg", "png"])

    st.markdown("**The Premises Details**")
    data["unit_number"] = st.text_input("Unit/room no.", placeholder="e.g. 654/1", key="unit_number")
    data["en_building_name"] = st.text_input("Project name", placeholder="e.g. The Line Sathorn", key="en_building_name")
    data["th_building_name"] = st.text_input("ชื่อโครงการ", placeholder="e.g. เดอะ ไลน์ สาทร", key="th_building_name")
    data["en_building_address"] = st.text_input("Address", placeholder="e.g. 111 This Road, That Sub-District and District, Bangkok 10110", key="en_building_address")
    data["th_building_address"] = st.text_input("ที่อยู่โครงการ", placeholder="e.g. 111 ถนนที่นี้ แขวงที่นั้น เขตที่นู้น กรุงเทพฯ 10110", key="th_building_address")

    data["floor_number"] = st.text_input("Floor no.", placeholder="e.g. 3", key="floor_number")
    data["area_size"] = st.text_input("Area size (sqm)", placeholder="in square meters (sqm)", key="area_size")
    data["rent_price"] = st.text_input("Monthly rental rate", placeholder="e.g. 15000", key="rent_price")
    data["deposit_price"] = st.text_input("Deposit rate", placeholder="e.g. 30000", key="deposit_price")

    st.markdown("**Landlord's Bank Account Details**")
    data["bank_name"] = st.text_input("Bank", placeholder="e.g. Kasikorn Bank", key="bank_name")
    data["account_number"] = st.text_input("Account no.", placeholder="e.g. 01444702454", key="account_number")
    data["account_name"] = st.text_input("Account name", placeholder="e.g. Mrs.Name Surname", key="account_name")

    st.markdown("**Continue Leasing Details**")
    data["contract_year"] = st.text_input("This agreement is for the [ordinal] year", placeholder="e.g. 3", key="contract_year")
    contract_year = data.get("contract_year", "")

    if not data.get("tenant_th_name") or not data.get("tenant_idcard") or not data.get("tenant_en_name"):
        st.error("Please complete all required fields.")
        return {}

    if not data.get("en_building_name") or not data.get("unit_number"):
        st.error("Project name and room number required.")
        return {}

    start_date = parse_thai_date_str(data.get("start_date"))
    end_date = parse_thai_date_str(data.get("end_date"))
    if not start_date or not end_date:
        st.error("รูปแบบวันที่ไม่ถูกต้อง กรุณากรอกใหม่ เช่น 1 มกราคม 2568")
        return {}

    try:
        rent = float(data.get("rent_price", "0"))
        deposit = float(data.get("deposit_price", "0"))
    except:
        st.error("Rental and deposit must be numbers.")
        return {}

    start_data = date_data(start_date)
    end_data = date_data(end_date)
    data.update({f"start_{k}": v for k, v in start_data.items()})
    data.update({f"end_{k}": v for k, v in end_data.items()})

    data["start_day_en"] = format_date(start_date, format="d MMMM y", locale="en")
    data["end_day_en"] = format_date(end_date, format="d MMMM y", locale="en")
    data["start_day_th"] = format_date(start_date, format="d MMMM y", locale="th")
    data["end_day_th"] = format_date(end_date, format="d MMMM y", locale="th")
    data["start_day_month_year_th"] = format_date(start_date, format="d MMMM y", locale="th")
    data["rent_price_en"] = num2words(rent, lang="en").title() + " Baht Only"
    data["rent_price_th"] = bahttext(rent)

    plus_5 = start_date + timedelta(days=5)
    data["day_plus_5"] = format_date(plus_5, format="d MMMM y", locale="th")
    data["day_plus_5_ordinal"] = ordinal(plus_5.day)

    data["deposit_price_en"] = num2words(deposit, lang="en").title() + " Baht Only"
    data["deposit_price_th"] = bahttext(deposit)

    try:
        contract_year_num = int(data["contract_year"])
        data["contract_year_ordinal"] = ordinal(contract_year_num)
    except Exception:
        data["contract_year_ordinal"] = data["contract_year"]

    from io import BytesIO

    if data["photo1"] is not None:
        data["photo1"].seek(0)
        data["photo1"] = BytesIO(data["photo1"].read())

    if data["photo2"] is not None:
        data["photo2"].seek(0)
        data["photo2"] = BytesIO(data["photo2"].read())

    return data

def booking_form():
    data = {}
    st.header("🗓️ Booking Leasing")
    st.markdown("Please provide more information to proceed with the lease booking.")
    data["en_building_name"] = st.text_input("Project Name", placeholder="e.g. Ideo Q", key="booking_en_building_name")
    data["booking_date"] = st.text_input("Booking Date", placeholder="e.g. 1 January 2025", key="booking_date")

    st.markdown("**Property Details**")
    data["landlord_en_name"] = st.text_input("Landlord Name", placeholder="e.g. Mrs.Name Surname", key="booking_landlord_en_name")
    data["floor_number"] = st.text_input("Block/Floor", placeholder="e.g. 9", key="booking_floor_number")
    data["unit_number"] = st.text_input("Unit Number", placeholder="e.g. 371/5", key="booking_unit_number")
    data["building_number"] = st.text_input("Building", placeholder="e.g. B", key="booking_building_number")
    data["area_size"] = st.text_input("Build-Up Area", placeholder="e.g. 31.5", key="booking_area_size")

    st.markdown("**Leasing & Purchase Price Details**")
    data["rent_price_full"] = st.text_input("Price Before Discount", placeholder="e.g. 17000", key="rent_price_full")
    data["rent_discount"] = st.text_input("Discount", placeholder="e.g. 2000", key="rent_discount")
    data["rent_price"] = st.text_input("Net Price", placeholder="e.g. 10000", key="booking_rent_price")
    data["deposit_price"] = st.text_input("Deposit rate", placeholder="e.g. 30000", key="booking_deposit_price")

    st.markdown("**Booking Details**")
    data["tenant_en_name"] = st.text_input("Tenant Name", placeholder="e.g. Mr.Name Surname", key="booking_tenant_en_name")
    data["tenant_idcard"] = st.text_input("ID Number", placeholder="e.g. 1-2344-56678-90-1", key="booking_tenant_idcard")
    data["tenant_birth"] = st.text_input("Date of Birth", placeholder="e.g. 2 February 1999", key="tenant_birth")
    data["tenant_en_nationality"] = st.text_input("Nationality", placeholder="e.g. Chinese", key="booking_tenant_en_nationality")
    data["tenant_en_address"] = st.text_input("Address", placeholder="e.g. 111 This Road, That Sub-District and District, Bangkok 10110", key="booking_tenant_en_address")

    st.markdown("**Move In Details**")
    data["en_move_in_date"] = st.text_input("Move In Date", placeholder="e.g. 3 March 2025", key="en_move_in_date")

    st.markdown("**Bank Transfer Details**")
    data["account_name_booked"] = st.text_input("Account name", placeholder="e.g. Mrs.Name Surname", key="account_name_booked")
    data["en_bank_name"] = st.text_input("Bank Name", placeholder="e.g. Kasikorn Bank", key="en_bank_name")
    data["th_bank_name"] = st.text_input("ชื่อธนาคาร", placeholder="e.g. ธนาคารกสิกรไทย จำกัด (มหาชน)", key="th_bank_name")
    data["account_no"] = st.text_input("Account no.", placeholder="e.g. 01444702454", key="account_no")

    st.markdown("**Remark(s) Details**")
    data["remarks"] = st.text_input("Remark(s)", placeholder="e.g. The landlord agrees to repaint the bedroom wall in white before handover.", key="remarks")

    if not data.get("en_bank_name") or not data.get("account_no"):
        st.error("Please complete all required fields.")
        return {}

    if not data.get("booking_date") or not data.get("en_move_in_date"):
        st.error("Invalid date format. Please use the format: e.g. 1 January 2568.")
        return {}

    th_move_in_date = ""
    if data.get("en_move_in_date"):
        try:
            th_move_in_date = convert_en_date_to_thai(data.get("en_move_in_date"))
            if "Invalid date format." in th_move_in_date:
                st.error("Invalid date format. Please use the format: e.g. 1 January 2568.")
                return {}
        except Exception:
            st.error("Error occured while converting date to thai")
            return {}

    data["th_move_in_date"] = th_move_in_date

    return data

def furniture_form():
    data = {}
    st.header("🪑 Furniture Lists")
    data["en_building_name"] = st.text_input("Project Name", placeholder="e.g. Happy Condo", key="furniture_en_building_name")
    data["floor_number"] = st.text_input("Block/Floor", placeholder="e.g. 3", key="furniture_floor_number")
    data["unit_number"] = st.text_input("Unit Number", placeholder="e.g. 20", key="furniture_unit_number")
    data["landlord_en_name"] = st.text_input("Landlord Name", placeholder="e.g. Mr.Name Surname", key="furniture_landlord_en_name")
    data["landlord_th_name"] = st.text_input("ชื่อผู้ให้เช่า", placeholder="e.g. นายชื่อ นามสกุล", key="furniture_landlord_th_name")
    data["tenant_en_name"] = st.text_input("Tenant Name", placeholder="e.g. Ms.Name Surname", key="furniture_tenant_en_name")
    data["tenant_th_name"] = st.text_input("ชื่อผู้เช่า", placeholder="e.g. นางสาวชื่อ นามสกุล", key="furniture_tenant_th_name")

    return data

def run_full_form():
    selected = st.session_state.get("selected_docs", {})

    if not selected:
        st.warning("⚠️ Please select at least one document type from the homepage.")
        return {}

    form_data = {}
    if selected.get("contract"):
        contract_data = contract_form()
        if contract_data:
            st.session_state["contract_data"] = contract_data
            form_data["contract"] = contract_data

    if selected.get("booking"):
        booking_data = booking_form()
        if booking_data:
            st.session_state["booking_data"] = booking_data
            form_data["booking"] = booking_data

    if selected.get("furniture"):
        furniture_data = furniture_form()
        if furniture_data:
            st.session_state["furniture_data"] = furniture_data
            st.session_state["furniture_list"] = []
            form_data["furniture"] = furniture_data

    return form_data

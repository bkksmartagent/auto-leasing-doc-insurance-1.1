import sys
import os
import time
import streamlit as st
#from docx2pdf import convert
from modules.inputs import contract_form, booking_form, furniture_form
from modules.helpers import generate_contract, generate_booking, generate_furniture, safe_filename

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

st.title("📝 Fill in Your Document(s) Details")
selected = st.session_state.get("selected_docs", {})

if not selected:
    st.warning("⚠️ Please go back and select document types.")
    st.stop()

output_dir = "output"
os.makedirs(output_dir, exist_ok=True)

# ---------- Fill Form Phase ----------
if selected.get("contract"):
    with st.form("contract_form"):
        contract_data = contract_form()
        if st.form_submit_button("Submit Contract"):
            st.session_state["contract_data"] = contract_data
            st.success("Contract data saved.")

if selected.get("booking"):
    with st.form("booking_form"):
        booking_data = booking_form()
        if st.form_submit_button("Submit Booking"):
            st.session_state["booking_data"] = booking_data
            st.success("Booking data saved.")

if selected.get("furniture"):
    with st.form("furniture_form"):
        furniture_data = furniture_form()
        submitted = st.form_submit_button("Submit Furniture List")
        if submitted:
            st.session_state["furniture_data"] = furniture_data
            st.success("Furniture data saved.")

    st.markdown("---")
    st.subheader("📋 Add Furniture Items")

    if "furniture_list" not in st.session_state:
        st.session_state.furniture_list = []

    if "upload_counter" not in st.session_state:
        st.session_state.upload_counter = 0

    def add_item():
        if st.session_state.get("new_image") is not None:
            st.session_state.furniture_list.append({
                "remark": st.session_state.new_remark,
                "image": st.session_state.new_image,
            })
            st.session_state.new_remark = ""
            st.session_state.upload_counter += 1
        else:
            st.warning("Please upload an image to proceed.")

    with st.form("furniture_add_item"):
        remark = st.text_input("Remark", key="new_remark")
        image = st.file_uploader("Upload Picture", type=["png", "jpg", "jpeg"], key="new_image")
        submitted_add = st.form_submit_button("Add Item", on_click=add_item)

    if "furniture_list" in st.session_state:
        for i, item in enumerate(st.session_state.furniture_list):
            cols = st.columns([1, 2, 4, 1])
            cols[0].markdown(f"<div style='text-align: center; font-weight: bold;'>{i + 1}</div>", unsafe_allow_html=True)
            if item["image"] is not None:
                cols[1].image(item["image"], width=80)
            cols[2].write(item["remark"])
            if cols[3].button("✕", key=f"del_{i}"):
                st.session_state.furniture_list.pop(i)
                st.rerun()


# ---------- Generate Document Phase -----------
if st.button("📄 Issue the Document(s)"):
    missing = False
    generated_files = []

    contract_data = st.session_state.get("contract_data", {})
    booking_data = st.session_state.get("booking_data", {})
    furniture_data = st.session_state.get("furniture_data", {})
    furniture_list = st.session_state.get("furniture_list", [])

    if selected.get("contract") and not contract_data:
        missing = True
    if selected.get("booking") and not booking_data:
        missing = True
    if selected.get("furniture") and (not furniture_data or not furniture_list):
        missing = True

    if missing:
        st.warning("⚠️ Please fill and submit all selected forms first.")
        st.stop()

    st.session_state["generated_files"] = []

    if selected.get("contract"):
        filename_base = safe_filename(f"Leasing_Contract_and_Continue_Leasing_{contract_data.get('en_building_name', '')}_{contract_data.get('unit_number', '')}")
        docx_path = os.path.join(output_dir, f"{filename_base}.docx")
        #pdf_path = os.path.join(output_dir, f"{filename_base}.pdf")
        generate_contract(contract_data, "templates/contract_template.docx", docx_path)
        #try:
            #convert(docx_path, pdf_path)
        #except Exception:
            #st.warning("⚠️ Contract PDF conversion failed.")
        generated_files.append((filename_base, docx_path, None))

    if selected.get("booking"):
        filename_base = safe_filename(f"Booking_Leasing_{booking_data.get('en_building_name', '')}_{booking_data.get('unit_number', '')}")
        docx_path = os.path.join(output_dir, f"{filename_base}.docx")
        #pdf_path = os.path.join(output_dir, f"{filename_base}.pdf")
        generate_booking(booking_data, "templates/booking_template.docx", docx_path)
        #try:
            #convert(docx_path, pdf_path)
        #except Exception:
            #st.warning("⚠️ Booking PDF conversion failed.")
        generated_files.append((filename_base, docx_path, None))

    if selected.get("furniture"):
        filename_base = safe_filename(f"Furniture_Lists_{furniture_data.get('en_building_name', '')}_{furniture_data.get('unit_number', '')}")
        docx_path = os.path.join(output_dir, f"{filename_base}.docx")
        #pdf_path = os.path.join(output_dir, f"{filename_base}.pdf")
        generate_furniture("templates/furniture_template.docx", docx_path, furniture_data, furniture_list)
        #try:
            #convert(docx_path, pdf_path)
        #except Exception:
            #st.warning("⚠️ Furniture PDF conversion failed.")
        generated_files.append((filename_base, docx_path, None))

    st.session_state["generated_files"] = generated_files
    st.success("🎉 All selected documents have been issued.")

# ---------- Show Download Buttons -----------
if "generated_files" in st.session_state and st.session_state["generated_files"]:
    st.markdown("### 📥 Download Your Files")
    for filename_base, docx_path, pdf_path in st.session_state["generated_files"]:
        if os.path.exists(docx_path):
            with open(docx_path, "rb") as f:
                st.download_button(f"📄 Download {filename_base}.docx", f, file_name=f"{filename_base}.docx")
        #if os.path.exists(pdf_path):
            #with open(pdf_path, "rb") as f:
                #st.download_button(f"📕 Download {filename_base}.pdf", f, file_name=f"{filename_base}.pdf")

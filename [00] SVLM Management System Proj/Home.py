import streamlit as st
import os
from datetime import datetime, timedelta

st.title("📑 SVLM Auto-Issued System")
st.markdown("**Powered by Siamvilai Development Co., Ltd.**", help="Internal document generation system.")
st.subheader("Select the document you would like to issue.")

contract = st.checkbox("Leasing Contract and Continue Contract")
booking = st.checkbox("Booking Leasing")
furniture = st.checkbox("Furniture Lists")

if st.button("Get Started"):
    if not any([contract, booking, furniture]):
        st.warning("⚠️ Select at least one document to get started.")
    else:
        st.session_state.selected_docs = {
            "contract": contract,
            "booking": booking,
            "furniture": furniture
        }
        st.switch_page("pages/Issuance.py")

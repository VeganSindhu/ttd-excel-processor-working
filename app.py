import streamlit as st
import tempfile
import os
from ttd_filler_logic import generate_output

st.set_page_config(page_title="TTD Excel Processor", layout="centered")

st.title("📦 TTD Excel Processor")
st.write("Upload **TTD Orders** and **TTD Postal** files to generate output.")

orders_file = st.file_uploader("Upload TTD Orders Excel", type=["xlsx"])
postal_file = st.file_uploader("Upload TTD Postal Excel", type=["xlsx"])

if orders_file and postal_file:
    if st.button("Generate Output"):
        with tempfile.TemporaryDirectory() as tmp:
            orders_path = os.path.join(tmp, "orders.xlsx")
            postal_path = os.path.join(tmp, "postal.xlsx")

            with open(orders_path, "wb") as f:
                f.write(orders_file.read())

            with open(postal_path, "wb") as f:
                f.write(postal_file.read())

            output_path, count = generate_output(orders_path, postal_path)

            st.success(f"✅ Processed {count} articles")

            with open(output_path, "rb") as f:
                st.download_button(
                    "⬇️ Download Output Excel",
                    f,
                    file_name="Matching_Output.xlsx"
                )

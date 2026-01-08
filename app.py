import streamlit as st
import tempfile
import os

from ttd_filler_logic import generate_output

st.set_page_config(page_title="TTD Postal Processor", layout="centered")

st.title("📦 TTD Postal Excel Processor")

st.write("Upload the required files and download the processed output.")

# ---------------- FILE UPLOADS ----------------

orders_file = st.file_uploader(
    "Upload TTD Orders Excel",
    type=["xlsx"]
)

postal_file = st.file_uploader(
    "Upload TTD Postal Excel",
    type=["xlsx"]
)

# ---------------- PROCESS ----------------

if st.button("Process Files"):
    if not orders_file or not postal_file:
        st.error("Please upload both Orders and Postal files.")
    else:
        with st.spinner("Processing... Please wait"):
            with tempfile.TemporaryDirectory() as tmpdir:
                orders_path = os.path.join(tmpdir, "orders.xlsx")
                postal_path = os.path.join(tmpdir, "postal.xlsx")
                output_path = os.path.join(tmpdir, "Matching_Output.xlsx")

                # Save uploaded files
                with open(orders_path, "wb") as f:
                    f.write(orders_file.read())
                with open(postal_path, "wb") as f:
                    f.write(postal_file.read())

                # FIXED INTERNAL FILES
                template_path = "TTD Template.xlsx"
                volumetric_path = "Volumetric Measurement.xlsx"

                # Run your logic
                generate_output(
                    orders_path,
                    postal_path,
                    template_path,
                    volumetric_path,
                    output_path
                )

                # Download button
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="⬇ Download Processed Excel",
                        data=f,
                        file_name="TTD_Postal_Output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                st.success("Processing completed successfully!")

import streamlit as st
from track_3_stream import process_excel  # Import your function
import os

GITHUB_TOKEN = os.getenv("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")
repo_url =  f"https://{GITHUB_TOKEN}@github.com/lynnmumbi/track3_report_analyzer.git"


# Streamlit UI
st.title("Track 3 Report Analyzer")
st.write("Hey you😃, please upload your Excel file and download the processed version.")

# Upload file button
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])

if uploaded_file:
    if st.button("click me to Analyze 😊🚀"):
        with st.spinner("Processing..."):
            try:
                processed_file = process_excel(uploaded_file)  # Calls your function
                st.session_state.processed_file = processed_file
                st.success("Analysis complete! Download your file below.")
            except Exception as e:
                st.error(f"Error: {e}")


# Download button (Only appears if analysis is done)
if "processed_file" in st.session_state and st.session_state.processed_file:
    st.download_button(
        label="Download Processed File",
        data=st.session_state.processed_file,
        file_name="Analyzed_Workbook.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

import streamlit as st
from portfolio_engine import process_portfolio

st.set_page_config(
    page_title="Yahoo Portfolio Analyzer",
    page_icon="📈",
    layout="centered"
)

st.title("📈 Yahoo Portfolio Analyzer")
st.write("Upload your Excel portfolio file, run analysis, and download the updated file.")

uploaded_file = st.file_uploader(
    "Choose your Excel file",
    type=["xlsx"]
)

if uploaded_file is not None:

    st.success("Excel file uploaded successfully.")

    if st.button("Run analysis"):

        with st.spinner("Processing Yahoo Finance data... please wait..."):

            try:
                output_file = process_portfolio(uploaded_file)

                st.success("Analysis completed successfully.")

                st.download_button(
                    label="Download updated Excel file",
                    data=output_file,
                    file_name="Portfolio_Updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Error: {e}")

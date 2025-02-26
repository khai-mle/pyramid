import os
import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from docx import Document

st.set_page_config(layout="wide")

# Custom styling
st.markdown("""
    <style>
        body {
            font-family: Arial, sans-serif;
            color: black;
            background-color: white;
        }
        .stApp {
            background-color: white;
        }
        .stTitle, .stHeader, .stSubheader {
            color: red;
        }
        .stSelectbox div[data-baseweb="select"], .stTextInput, .stTextArea, .stButton {
            border: none !important;
            box-shadow: none !important;
        }
    </style>
""", unsafe_allow_html=True)

# Load company data
def load_company_data():
    try:
        df = pd.read_csv("companies.csv")
        required_columns = ["ticker", "company name", "description", "sector", "industry", "sub-industry", "tag"]
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"Missing required columns in companies.csv: {', '.join(missing_columns)}")
            return None
        return df
    except Exception as e:
        st.error(f"Error loading companies.csv: {str(e)}")
        return None

company_df = load_company_data()

# Search functionality
st.title("Prospect 10-K Analysis")
st.subheader("Tick the company you want to analyze")
search_query = st.text_input("Enter company name or ticker:")

if company_df is not None:
    if search_query:
        filtered_df = company_df[
            company_df["company name"].str.contains(search_query, case=False, na=False) |
            company_df["ticker"].str.contains(search_query, case=False, na=False)
        ]
    else:
        filtered_df = company_df.copy()
    
    # Apply filters for sector, industry, and sub-industry
    col1, col2, col3 = st.columns(3)
    
    with col1:
        sector = st.selectbox("Select Sector", ["All"] + sorted(company_df["sector"].dropna().unique().tolist()))
    with col2:
        industry_options = company_df[company_df["sector"] == sector]["industry"].dropna().unique().tolist() if sector != "All" else company_df["industry"].dropna().unique().tolist()
        industry = st.selectbox("Select Industry", ["All"] + sorted(industry_options))
    with col3:
        sub_industry_options = company_df[(company_df["industry"] == industry) & (company_df["sector"] == sector)]["sub-industry"].dropna().unique().tolist() if industry != "All" else company_df["sub-industry"].dropna().unique().tolist()
        sub_industry = st.selectbox("Select Sub-Industry", ["All"] + sorted(sub_industry_options))
    
    # Apply filters to the DataFrame
    if sector != "All":
        filtered_df = filtered_df[filtered_df["sector"] == sector]
    if industry != "All":
        filtered_df = filtered_df[filtered_df["industry"] == industry]
    if sub_industry != "All":
        filtered_df = filtered_df[filtered_df["sub-industry"] == sub_industry]
    
    # Display search results in an interactive table
    st.subheader("Company Database")
    gb = GridOptionsBuilder.from_dataframe(filtered_df[["ticker", "company name", "description"]])
    gb.configure_selection('multiple', use_checkbox=True, header_checkbox=True)
    gb.configure_default_column(wrapText=True, sortable=True)
    grid_options = gb.build()
    grid_response = AgGrid(
        filtered_df, gridOptions=grid_options, update_mode=GridUpdateMode.SELECTION_CHANGED,
        height=400, width='100%', fit_columns_on_grid_load=True
    )
    
    selected_rows = grid_response.get("selected_rows", [])
    if isinstance(selected_rows, list) and all(isinstance(row, dict) for row in selected_rows):
        selected_tickers = [row.get("ticker", "") for row in selected_rows]
    else:
        selected_tickers = []
    
    # Additional question box
    st.subheader("Additional Inquiry")
    user_question = st.text_area("What else do you want to ask about the selected companies?")
    
    # Submit button
    if st.button("Submit Inquiry"):
        st.success("Your inquiry has been submitted successfully!")
    
    # Download report button
    if selected_tickers and os.path.exists("report.docx"):
        with open("report.docx", "rb") as file:
            st.download_button(
                label="Download Report",
                data=file,
                file_name="company_analysis_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

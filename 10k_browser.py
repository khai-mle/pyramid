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
        /* Style for submit button */
        div[data-testid="stButton"] > button[kind="primary"] {
            background-color: red;
            color: white;
            font-weight: bold;
        }
        /* Remove extra spacing */
        div[data-testid="column"] {
            padding: 0 !important;
            margin: 0 !important;
        }
        /* Make service selection buttons fill their container */
        #service-selection div[data-testid="stButton"] > button {
            display: block;
            width: 100%;
            height: 100%;
            padding: 20px; /* adjust padding as needed */
        }
    </style>
""", unsafe_allow_html=True)

# Load company data
def load_company_data():
    try:
        df = pd.read_csv("companies.csv")
        return df
    except Exception as e:
        st.error(f"Error loading companies.csv: {str(e)}")
        return None

company_df = load_company_data()

# Load BearingPoint services and sub-categories
def load_service_data():
    try:
        df = pd.read_csv("Multi-Level_Table_of_BearingPoint_Service_Lines.csv")
        return df
    except Exception as e:
        st.error(f"Error loading service data: {str(e)}")
        return None

service_df = load_service_data()

# Service Selection
st.title("Prospect 10-K Analysis")
st.subheader("BearingPoint Service Line Selector")
if service_df is not None:
    service_options = service_df["Service Line"].unique().tolist()
    
    # Initialize selected_service in session state if not already present
    if "selected_service" not in st.session_state:
        st.session_state.selected_service = None
    
    # Wrap service selection buttons in a dedicated container for custom styling
    st.markdown("<div id='service-selection'>", unsafe_allow_html=True)
    
    # Create a single row of columns for all service options
    cols = st.columns(len(service_options))
    
    # Use Streamlit's native buttons
    for i, service in enumerate(service_options):
        with cols[i]:
            if st.button(service, key=f"service_{i}", use_container_width=True):
                st.session_state.selected_service = service

    st.markdown("</div>", unsafe_allow_html=True)

    # Display sub-categories if a service is selected
    if st.session_state.selected_service:
        sub_category_options = service_df[service_df["Service Line"] == st.session_state.selected_service]["Sub-Category"].unique().tolist()
        selected_sub_categories = st.multiselect("Select Specializations:", sub_category_options)

# Company Selection
st.subheader("Corporate Domain Selector")
if company_df is not None:
    col1, col2, col3 = st.columns(3)
    with col1:
        sector = st.selectbox("Sector Class", ["All"] + sorted(company_df["sector"].dropna().unique().tolist()))
    with col2:
        if sector != "All":
            industry_options = sorted(company_df[company_df["sector"] == sector]["industry"].dropna().unique().tolist())
        else:
            industry_options = sorted(company_df["industry"].dropna().unique().tolist())
        industry = st.selectbox("Industry", ["All"] + industry_options)
    with col3:
        if sector != "All" and industry != "All":
            sub_industry_options = sorted(company_df[(company_df["sector"] == sector) & (company_df["industry"] == industry)]["sub-industry"].dropna().unique().tolist())
        else:
            sub_industry_options = sorted(company_df["sub-industry"].dropna().unique().tolist())
        sub_industry = st.selectbox("Business Vertical", ["All"] + sub_industry_options)
    
    search_query = st.text_input("Enter Company Name or Ticker:")
    
    # Apply filters dynamically
    filtered_df = company_df.copy()
    if sector != "All":
        filtered_df = filtered_df[filtered_df["sector"] == sector]
    if industry != "All":
        filtered_df = filtered_df[filtered_df["industry"] == industry]
    if sub_industry != "All":
        filtered_df = filtered_df[filtered_df["sub-industry"] == sub_industry]
    if search_query:
        filtered_df = filtered_df[
            filtered_df["company name"].str.contains(search_query, case=False, na=False) |
            filtered_df["ticker"].str.contains(search_query, case=False, na=False)
        ]
    
    st.write("Tick the companies you want to analyze")
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
    
    # Additional Inquiry
    st.subheader("What else would you like to know?")
    user_question = st.text_area("Additional considerations")
    
    # Submit button - make it red
    if st.button("Submit", key="submit_button", 
                 type="primary",  # This will make it use the primary button style
                 help="Submit your inquiry", 
                 use_container_width=True):
        st.success("Your inquiry has been submitted successfully!")
    
    # Download report button
    if selected_tickers and os.path.exists("report.docx"):
        with open("report.docx", "rb") as file:
            st.download_button(
                label="Secure Dossier",
                data=file,
                file_name="corporate_analysis_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

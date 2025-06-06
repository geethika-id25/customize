# app.py
import streamlit as st
import pandas as pd
import plotly.express as px
import re
from io import BytesIO

# Function to clean and normalize column names
def normalize_columns(df):
    """Normalize column names by converting to lowercase and replacing special characters."""
    df.columns = [re.sub(r'[^a-zA-Z0-9]', '_', col.lower().strip()) for col in df.columns]
    return df

# Function to infer column types
def infer_column_types(df):
    """Infer column types (numeric, categorical, binary) dynamically."""
    types = {}
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            types[col] = 'numeric'
        elif df[col].nunique() <= 10 or df[col].dtype == 'bool' or df[col].isin([0, 1, 'yes', 'no']).all():
            types[col] = 'categorical'
        else:
            types[col] = 'text'
    return types

# Function to process natural language queries
def process_query(df, query, col_types):
    """Process user query and return text response or chart."""
    query = query.lower().strip()
    
    # Handle missing values
    df = df.fillna({'numeric': df.select_dtypes(include='number').mean(),
                    'categorical': 'missing',
                    'text': 'missing'})
    
    # Statistical summaries
    if 'average' in query or 'mean' in query:
        for col in df.columns:
            if col_types[col] == 'numeric' and col in query:
                result = df[col].mean()
                return f"Average {col.replace('_', ' ')}: {result:.2f}", None
    
    # Filtered queries (e.g., "how many customers are under 30")
    elif 'how many' in query or 'count' in query:
        for col in df.columns:
            if col in query:
                match = re.search(r'(\w+)\s*(<|>|<=|>=|=)\s*(\d+\.?\d*)', query)
                if match:
                    col, op, value = match.groups()
                    value = float(value)
                    if op == '<':
                        result = df[df[col] < value].shape[0]
                    elif op == '>':
                        result = df[df[col] > value].shape[0]
                    elif op == '=':
                        result = df[df[col] == value].shape[0]
                    return f"Count of {col.replace('_', ' ')} {op} {value}: {result}", None
    
    # Grouped queries (e.g., "compare sales by region")
    elif 'compare' in query or 'by' in query:
        for col in df.columns:
            if col_types[col] == 'categorical' and col in query:
                group_col = col
                agg_col = next((c for c in df.columns if col_types[c] == 'numeric' and c in query), None)
                if agg_col:
                    result = df.groupby(group_col)[agg_col].sum().reset_index()
                    fig = px.bar(result, x=group_col, y=agg_col, 
                                title=f"{agg_col.replace('_', ' ')} by {group_col.replace('_', ' ')}",
                                color_discrete_sequence=['#36A2EB'])
                    return result.to_markdown(), fig
    
    # Visualization queries (e.g., "show a bar chart of job")
    elif 'bar chart' in query or 'distribution' in query:
        for col in df.columns:
            if col_types[col] == 'categorical' and col in query:
                counts = df[col].value_counts().reset_index()
                counts.columns = [col, 'count']
                fig = px.bar(counts, x=col, y='count', 
                            title=f"Distribution of {col.replace('_', ' ')}",
                            color_discrete_sequence=['#FF6384'])
                return counts.to_markdown(), fig
    
    return "Sorry, I couldn't understand the query. Please try rephrasing!", None

# Streamlit app
st.set_page_config(page_title="NeoStats Assistant", layout="wide")
st.title("NeoStats Conversational Assistant")
st.markdown("Upload an Excel file (.xlsx) and ask questions about your data in plain English.")

# File upload
uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"], key="file_uploader")
if uploaded_file:
    try:
        # Read and process Excel file
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        df = normalize_columns(df)
        col_types = infer_column_types(df)
        
        # Display data preview
        st.subheader("Data Preview")
        st.dataframe(df.head(5), use_container_width=True)
        
        # Query input
        st.subheader("Ask a Question")
        query = st.text_input("Enter your question (e.g., 'What is the average income?' or 'Show a bar chart of job')")
        if query:
            with st.spinner("Processing your query..."):
                text_response, chart = process_query(df, query, col_types)
                
                # Display results
                if text_response:
                    st.markdown("**Result:**")
                    st.markdown(text_response)
                if chart:
                    st.markdown("**Visualization:**")
                    st.plotly_chart(chart, use_container_width=True)
                
    except Exception as e:
        st.error(f"Error processing file or query: {str(e)}")
else:
    st.info("Please upload an Excel file to begin.")

# Sidebar with instructions
with st.sidebar:
    st.header("How to Use")
    st.markdown("""
    1. Upload a valid .xlsx file (up to 500 rows, 10-20 columns).
    2. Ask questions in plain English, such as:
       - 'What is the average income?'
       - 'How many customers are under 30?'
       - 'Compare sales by region'
       - 'Show a bar chart of job'
    3. View results as text, tables, or charts.
    """)

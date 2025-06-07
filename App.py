import streamlit as st
import pandas as pd
import openai
import io
import plotly.express as px

# Set your OpenAI key here (or use environment variable)
openai.api_key = st.secrets["OPENAI_API_KEY"]

# Function to clean column names
def clean_column_names(df):
    df.columns = [col.strip().replace(' ', '_').lower() for col in df.columns]
    return df

# Load and clean Excel file
def load_excel(file):
    df = pd.read_excel(file)
    df = clean_column_names(df)
    return df

# Prompt engineering for LLM
def generate_prompt(df, user_question):
    sample = df.head(3).to_csv(index=False)
    prompt = f"""
You are a data analyst. A user has uploaded a dataset with the following preview:

import streamlit as st
import pandas as pd
import openai  # Required for ChatGPT integration
import matplotlib.pyplot as plt
import io
import os
import base64
from dotenv import load_dotenv

# Load environment variables
load_dotenv()
openai.api_key = os.getenv("OPENAI_API_KEY")

st.set_page_config(page_title="Excel Chatbot", layout="wide")
st.title("📊 Excel-Based Chat Assistant")

# Helper functions
def read_excel(file):
    df = pd.read_excel(file)
    df.columns = [col.strip().lower().replace(" ", "_") for col in df.columns]
    return df

def plot_to_base64(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png")
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")

# Upload and read Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file:
    df = read_excel(uploaded_file)
    st.write("Data Preview:", df.head())

    # User input
    user_query = st.text_input("Ask a question about the data")

    if user_query:
        # Construct prompt
        prompt = f"""
You are a data analyst. Here is a preview of the dataset:

{df.head(20).to_string(index=False)}

Please answer the following question based on the data above:
"""
        prompt += f"\nQuestion: {user_query}\n"

        # Query OpenAI
        try:
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3
            )
            answer = response['choices'][0]['message']['content']
            st.markdown("### Answer:")
            st.write(answer)
        except Exception as e:
            st.error(f"Error: {e}")
else:
    st.info("Please upload an Excel (.xlsx) file to begin.")

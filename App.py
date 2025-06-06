import streamlit as st
import pandas as pd
import openai
import matplotlib.pyplot as plt
import seaborn as sns
import io

# Set your OpenAI API Key securely
openai.api_key = st.secrets["OPENAI_API_KEY"]

# Normalize column names
def normalize_columns(df):
    df.columns = [col.strip().lower().replace(" ", "_").replace("-", "_") for col in df.columns]
    return df

# Convert user question + df to LLM prompt
def create_prompt(question, df):
    preview = df.head(3).to_csv(index=False)
    return f"""
You are a smart data analyst assistant. A user uploaded an Excel sheet. Here's a preview:
{preview}

Now, answer this question based on the uploaded data:
{question}

If a visual (bar chart, histogram, line chart) is suitable, describe it and provide a Python code using matplotlib/seaborn.
"""

# Streamlit UI
st.title("📊 Excel Insight Chatbot")
st.write("Upload your Excel file and ask questions in plain English.")

uploaded_file = st.file_uploader("Upload Excel (.xlsx) file", type=["xlsx"])
question = st.text_input("Ask your question")

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df = normalize_columns(df)
    st.write("Data Preview", df.head())

    if question:
        with st.spinner("Thinking..."):
            prompt = create_prompt(question, df)

            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",  # or "gpt-4" if available
                messages=[{"role": "system", "content": "You are a helpful data assistant."},
                          {"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=800
            )
            answer = response["choices"][0]["message"]["content"]
            st.markdown("**Answer:**")
            st.markdown(answer)

            # Optional: Try to extract and run matplotlib code if included
            if "```python" in answer:
                try:
                    code = answer.split("```python")[1].split("```")[0]
                    exec_globals = {"df": df, "plt": plt, "sns": sns}
                    exec(code, exec_globals)
                    st.pyplot(plt.gcf())
                    plt.clf()
                except Exception as e:
                    st.error(f"Error generating chart: {e}")


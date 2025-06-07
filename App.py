%pip install -q unstructured langchain
%pip install -q "unstructured[all-docs]"
import pandas as pd
excel_path = "Book2.xlsx"
if excel_path:
    df = pd.read_excel(excel_path)
    data = df.to_string(index=False)
else:
    print("Upload an Excel file")
from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.embeddings import OllamaEmbeddings
from langchain_community.vectorstores import Chroma

text_splitter = RecursiveCharacterTextSplitter(chunk_size=7500, chunk_overlap=100)
chunks = text_splitter.split_text(data)

embedding_model = OllamaEmbeddings(model="nomic-embed-text", show_progress=False)
vector_db = Chroma.from_texts(
    texts=chunks, 
    embedding=embedding_model,
    collection_name="local-rag"
    from langchain_community.chat_models import ChatOllama

local_model = "llama2"
llm = ChatOllama(model=local_model)
from langchain.prompts import PromptTemplate

QUERY_PROMPT = PromptTemplate(
    input_variables=["question"],
    template="""You are an AI assistant. Answer the user's questions based on the column names: 
    Id, order_id, name, sales, refund, and status. Original question: {question}"""
)
from langchain.retrievers.multi_query import MultiQueryRetriever

retriever = MultiQueryRetriever.from_llm(
    vector_db.as_retriever(), 
    llm,
    prompt=QUERY_PROMPT
)
from langchain.prompts import ChatPromptTemplate
from langchain_core.runnables import RunnablePassthrough
from langchain_core.output_parsers import StrOutputParser

template = """Answer the question based ONLY on the following context:
{context}
Question: {question}
"""

prompt = ChatPromptTemplate.from_template(template)

chain = (
    {"context": retriever, "question": RunnablePassthrough()}
    | prompt
    | llm
    | StrOutputParser()
)
raw_result = chain.invoke("How many rows are there?")
final_result = f"{raw_result}\n\nIf you have more questions, feel free to ask!"
print(final_result)

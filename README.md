**Excel Query Chatbot with AI**
This project is about  users to upload Excel files, ask natural language queries about the data, and receive actionable insights or visualizations. The chatbot uses Google’s Generative AI API (Gemini) to interpret queries and generate Python code for data analysis and visualization.

In today's data-driven world, working with Excel files is an everyday task for many professionals. But what if you could interact with your data using natural language instead of complex formulas or endless scrolling? That's where BChat Excel comes in!

What is BChat Excel?

BChat Excel is a web-based application built with Streamlit that allows you to upload an Excel file and ask questions about the data in natural language. Powered by cutting-edge language models and embeddings from LangChain, BChat Excel processes your Excel data, enabling you to have a seamless conversational experience with your datasets.

🛠️ Key Features:
Upload & Preview Data:

Upload any Excel file, and BChat Excel will instantly provide you with a preview of the data.
Intelligent Text Processing:

Utilizes the Recursive Character Text Splitter to handle and process large amounts of data by splitting it into manageable chunks.
Embeddings for Efficient Search:

Leverages the power of Ollama Embeddings for creating vector representations of your data, making it possible to perform rapid and accurate text-based searches.
Advanced Language Model Integration:

Uses the Llama2 model from Ollama, fine-tuned to understand and respond to queries related to your data.
Multi-Query Retriever:

Implements a MultiQueryRetriever to refine search results and provide accurate answers based on the context of your data.
Conversational Interface:

The chat-like interface mimics a conversation with an AI assistant, making it easy to ask follow-up questions or dive deeper into your dataset.


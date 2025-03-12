import streamlit as st
import openai
import docx
import pandas as pd
import PyPDF2
import pptx
from io import BytesIO
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Sidebar with logo and link
with st.sidebar:
    st.image("image.png", width=150)
    st.markdown(
        f'<a href="https://fanalysis-uhevotzeiesw3nczetzsmy.streamlit.app/" target="_blank">'
        f'<img src="image.png" width="150"></a>',
        unsafe_allow_html=True
    )

st.image("https://www.innominds.com/hubfs/Innominds-201612/img/nav/Innominds-Logo.png", width=200)
st.title("How Can I Assist You?🤖")

# OpenAI API Configuration
openai.api_key = "14560021aaf84772835d76246b53397a"
openai.api_base = "https://amrxgenai.openai.azure.com/"
openai.api_type = 'azure'
openai.api_version = '2024-02-15-preview'
deployment_name = 'gpt'

# Initialize session state
if "messages" not in st.session_state:
    st.session_state.messages = []
if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None
if "extracted_text" not in st.session_state:
    st.session_state.extracted_text = ""
if "new_upload" not in st.session_state:
    st.session_state.new_upload = False

# Function to extract text from uploaded files
def extract_text(file):
    text = ""
    if file.type == "application/pdf":
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    elif file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        doc = Document(file)
        for para in doc.paragraphs:
            text += para.text + "\n"
    elif file.type == "application/vnd.openxmlformats-officedocument.presentationml.presentation":
        ppt = pptx.Presentation(file)
        for slide in ppt.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text += shape.text + "\n"
    elif file.type == "text/csv":
        df = pd.read_csv(file)
        text += df.to_string()
    elif file.type == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        df = pd.read_excel(file)
        text += df.to_string()
    return text.strip()

# Display chat messages in a conversational format
for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        if "file" in message and message["file"]:
            with st.expander(f"📄 Uploaded: {message['file']}"):
                st.markdown("_File content hidden from chat but used in response._")
        st.markdown(message["content"])

# Chat input section with "+" button for file upload
col1, col2 = st.columns([5, 1])
with col1:
    user_input = st.chat_input("Enter your query:")
with col2:
    if st.button("➕"):  # Clicking this opens the file uploader
        st.session_state.new_upload = True

# Persistent file uploader
if st.session_state.new_upload:
    uploaded_file = st.file_uploader("Upload File", type=["pdf", "docx", "pptx", "csv", "xlsx"], key="file_uploader")
    
    if uploaded_file:
        st.session_state.uploaded_file = uploaded_file.name
        st.session_state.extracted_text = extract_text(uploaded_file)
        st.session_state.new_upload = False  # Reset after successful upload

# Display uploaded file info with expander
if st.session_state.uploaded_file:
    with st.expander("📂 Uploaded File"):
        st.markdown(f"**{st.session_state.uploaded_file}**")
        st.text_area("Extracted Text (Hidden in Chat)", st.session_state.extracted_text, height=150)

if user_input and user_input.strip():
    # Prepare full conversation history for context
    conversation_history = "\n".join(
        [f"{msg['role']}: {msg['content']}" for msg in st.session_state.messages]
    )

    # Combine extracted text and history in the prompt
    combined_prompt = conversation_history + f"\nUser: {user_input}"
    if st.session_state.uploaded_file:  # Ensure document context is used
        combined_prompt = f"Document Context:\n{st.session_state.extracted_text}\n\nChat History:\n{conversation_history}\n\nUser: {user_input}"
    
    # Append user message to history
    st.session_state.messages.append({
        "role": "user", "content": user_input, 
        "file": st.session_state.uploaded_file if st.session_state.uploaded_file else None
    })
    
    # Generate response using OpenAI API
    response = openai.ChatCompletion.create(
        engine=deployment_name,
        messages=[{"role": "system", "content": combined_prompt}],
        temperature=0.7,
        max_tokens=2000
    )
    
    ai_response = response["choices"][0]["message"]["content"]
    
    # Append AI response to history
    st.session_state.messages.append({"role": "assistant", "content": ai_response})
    
    st.rerun()

# Function to create a DOCX file from chat history
def create_docx():
    doc = Document()
    doc.add_heading("Chatbot Conversation History", level=1).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    for msg in st.session_state.messages:
        p = doc.add_paragraph()
        run = p.add_run(msg["role"].capitalize() + "\n")
        run.bold = True
        run.font.size = Pt(14)
        if "file" in msg and msg["file"]:
            doc.add_paragraph(f"📄 {msg['file']} (Content Hidden)")
        p.add_run(msg["content"]).font.size = Pt(12)
        p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        doc.add_paragraph("\n")
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Download chat history as DOCX
if st.button("Download Chat History"):
    docx_file = create_docx()
    st.download_button(label="Download Chat History", data=docx_file, file_name="chat_history.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

import json
from transformers import pipeline
from docx import Document

def extract_text_from_docx(docx_path, text_file_path):
    """Extracts text from a DOCX file and saves it to a text file."""
    doc = Document(docx_path)
    text = "\n".join(para.text for para in doc.paragraphs if para.text.strip())  

    with open(text_file_path, "w", encoding="utf-8") as text_file:
        text_file.write(text)
    print(f"Extracted text saved to {text_file_path}")

    return text

summarizer = pipeline("summarization", model="facebook/bart-large-cnn")

def summarize_batch(text_chunks):
    """Summarizes multiple chunks at once."""
    summaries = []
    for chunk in text_chunks:
        summary = summarize_text(chunk)
        summaries.append(summary)
    return summaries

def summarize_text(text):
    """Summarizes a single text chunk."""
    summaries = summarizer(text_chunks, max_length=130, min_length=30, do_sample=False)
    return [summary['summary_text'] for summary in summaries]

def chunk_text(text, chunk_size=1000):
    """Chunks the text into smaller pieces of specified size."""
    chunks = [text[i:i + chunk_size] for i in range(0, len(text), chunk_size)]
    return chunks

def save_summary_to_json(summaries, output_json_path):
    """Saves the summaries to a JSON file."""
    data = {"summaries": summaries}
    with open(output_json_path, "w", encoding="utf-8") as json_file:
        json.dump(data, json_file, indent=4, ensure_ascii=False)
    print(f"Summary saved to {output_json_path}")

docx_path = "Business Requirements Document - IP Automation.docx"
text_file_path = "extracted_text.txt"
output_json_path = "summarized_text.json"


extracted_text = extract_text_from_docx(docx_path, text_file_path)


text_chunks = chunk_text(extracted_text)


summaries = summarize_batch(text_chunks)


save_summary_to_json(summaries, output_json_path)

import os
import pandas as pd
import fitz  # PyMuPDF
from docx import Document

class ContentExtractor:

    def __init__(self, directory='Data/'):
        self.directory = directory

    def extract_text_from_txt(self, file_path):
        with open(file_path, "r", encoding="utf-8") as file:
            return file.read()

    def extract_text_from_pdf(self, file_path):
        text = ""
        doc = fitz.open(file_path)
        for page in doc:
            text += page.get_text("text") + "\n"
        return text

    def extract_text_from_docx(self, file_path):
        doc = Document(file_path)
        return "\n".join([para.text for para in doc.paragraphs])

    def extract_text_from_csv(self, file_path):
        df = pd.read_csv(file_path)
        return df.to_string(index=False)

    def extract_file(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".txt":
            return self.extract_text_from_txt(file_path)
        elif ext == ".pdf":
            return self.extract_text_from_pdf(file_path)
        elif ext == ".docx":
            return self.extract_text_from_docx(file_path)
        elif ext == ".csv":
            return self.extract_text_from_csv(file_path)
        else:
            return f"⚠ Unsupported file format: {ext}"

    def extract_from_directory(self):
        contents = {}
        for file_name in os.listdir(self.directory):
            file_path = os.path.join(self.directory, file_name)
            if os.path.isfile(file_path):
                contents[file_name] = self.extract_file(file_path)
        return contents


if __name__ == "__main__":
    extractor = ContentExtractor(directory="Data/")
    extracted_content = extractor.extract_from_directory()

    for file, content in extracted_content.items():
        print(f"\n📄 Extracted content from {file}:\n{content[:500]}...\n")

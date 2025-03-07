# ConvertEase 🚀

**Seamless conversions, anytime.**  

ConvertEase is an all-in-one file conversion web application that simplifies converting your documents, images, and presentations. With a sleek, modern UI and powerful backend functionality, ConvertEase lets you transform your files with just a few clicks.

---

## Features ✨

- **📷 JPG to PDF:** Convert image files (JPG, JPEG, PNG) into high-quality PDFs.  
- **📄 PDF to JPG:** Extract pages from PDFs and save them as JPG images.  
- **📑 PPT to PDF:** Convert PowerPoint presentations to PDF using LibreOffice.  
- **📝 Word to PDF:** Transform Word documents (DOC/DOCX) into PDFs.  
- **📊 Excel to PDF:** Convert Excel spreadsheets (XLS/XLSX) into PDFs.  
- **🔎 PDF OCR:** Extract text from scanned PDFs using Tesseract OCR.  
- **✍️ PDF to Word:** Convert PDFs into editable Word documents.  
- **🖼️ PDF to PPT:** Convert PDFs into editable PowerPoint presentations.

---

## Technologies Used 🛠️

- **Python & Flask**  
- **HTML, CSS & JavaScript**  
- **Pillow** – Image processing  
- **pdf2image** – Converting PDF pages to images  
- **Tesseract OCR** – Optical Character Recognition  
- **Poppler** – PDF rendering  
- **LibreOffice** – Document conversion  
- **Ghostscript** – PDF compression (if used)

---

## Installation & Setup 📦

1. **Clone the Repository**  
   Get the project files from the GitHub repository and navigate into the project folder.

2. **Create & Activate a Virtual Environment**  
   - **Windows:** Set up a virtual environment and activate it.  
   - **macOS/Linux:** Set up a virtual environment and activate it.

3. **Install Python Dependencies**  
   Install the required Python packages listed in the project.

4. **Install External Tools & Update Paths**  
   - **Poppler:** Download Poppler for Windows, extract it, and update the path in `app.py`.  
   - **Tesseract OCR:** Download Tesseract OCR and ensure it’s installed at `C:\Program Files\Tesseract-OCR\tesseract.exe` (path set in `app.py`).  
   - **LibreOffice:** Download LibreOffice and update the path in `app.py` if needed.  
   - **Ghostscript:** Download Ghostscript and ensure it’s accessible in your system PATH for PDF compression.

---

## Usage 🚀

1. **Run the Application**  
   Start the ConvertEase web application on your local machine.

2. **Open Your Browser**  
   Go to `http://127.0.0.1:5000` in your web browser.

3. **Convert Your Files**  
   Choose a conversion type, upload your file, and download the converted result.

---

Enjoy seamless file conversions with ConvertEase!

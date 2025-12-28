# PDF to Word Replica – Mediation Application Form

This project programmatically recreates a structured legal PDF document as a Microsoft Word (.docx) file using Python.  
The goal is to accurately replicate the layout, spacing, alignment, and structure of the original PDF.

## Tech Stack
- Python 3.x
- python-docx
- Flask

## Approach
- Studied the PDF layout and identified table-based structure.
- Recreated the document manually using tables and merged cells for layout accuracy.
- Tuned margins, column widths, spacing, and font size to closely match the original PDF.
- Wrapped the generator in a lightweight Flask app for easy execution and deployment.

## How to Run Locally
```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py

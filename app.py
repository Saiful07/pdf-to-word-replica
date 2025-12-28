from flask import Flask, send_file
from doc_generator import create_doc
import os

app = Flask(__name__)

@app.route("/")
def generate_doc():
    output_file = "mediation_application_form.docx"
    create_doc(output_file)
    return send_file(output_file, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)

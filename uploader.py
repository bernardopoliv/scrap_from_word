import os
import io

import docx
from flask import Flask, flash, request, redirect, url_for, send_file

from scrap_from_word import (
    get_values,
    export
)
# http://flask.pocoo.org/docs/1.0/patterns/fileuploads/

# UPLOAD_FOLDER = '/path/to/the/uploads'
ALLOWED_EXTENSIONS = set(['docx', 'doc'])

app = Flask(__name__)
app.secret_key = b'_5#andwqehas,mdnewr2738rhksdjffdy2L"F4Q8z\n\xec]/'

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if not request.form.get('doc_type', '') in ('AG', 'RE', 'CO'):
            return 'You need to specify a document type.'
        doc_type = request.form['doc_type']

        # check if the post request has the file key
        if 'file' not in request.files:
            return 'File not selected.'

        file = request.files['file']

        if file.filename == '':
            return 'File not selected.'

        if file and allowed_file(file.filename):
            doc = docx.Document(file)

            data = get_values(doc_type, doc)
            xlsx = io.BytesIO()  # in-memory file for saving
            export(data, xlsx)  # writes on xlsx

            xlsx.seek(0)  # returns buffer to 0

            # Sends the processed file as response
            return send_file(
                xlsx, 
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                as_attachment=True,
                attachment_filename='workbook.xlsx'
            )

    return '''
    <!doctype html>
    <title>Upload new docx</title>
    <h1>Upload new docx</h1>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file>
      <fieldset>
        <legend>Select document type</legend>

        <div>
            <input type="radio" id="agr"
                   name="doc_type" value="AG" checked />
            <label for="agr">Agricultural</label>
        </div>

        <div>
            <input type="radio" id="com"
                   name="doc_type" value="CO" />
            <label for="com">Commercial</label>
        </div>

        <div>
            <input type="radio" id="res"
                   name="doc_type" value="RE" />
            <label for="res">Residential</label>
        </div>

    </fieldset>
      <input type=submit value=Upload>
    </form>
    '''

# We only need this for local development.
if __name__ == '__main__':
    app.run()
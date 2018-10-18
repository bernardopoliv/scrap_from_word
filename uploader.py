import os
from flask import Flask, flash, request, redirect, url_for

# http://flask.pocoo.org/docs/1.0/patterns/fileuploads/

# UPLOAD_FOLDER = '/path/to/the/uploads'
ALLOWED_EXTENSIONS = set(['docx', 'doc'])

app = Flask(__name__)
# app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':

        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)

        file = request.files['file']

        # if user does not select file, browser also
        # submit an empty part without filename
        if file.filename == '':
            flash('No selected file.')
            return redirect(request.url)

        if file and allowed_file(file.filename):
            return file.read()

            # filename = secure_filename(file.filename)
            # file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            # return redirect(url_for('uploaded_file',
                                    # filename=filename))

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
                   name="doc_type" value="agr" checked />
            <label for="agr">Agricultural</label>
        </div>

        <div>
            <input type="radio" id="com"
                   name="doc_type" value="com" />
            <label for="com">Commercial</label>
        </div>

        <div>
            <input type="radio" id="res"
                   name="doc_type" value="res" />
            <label for="res">Residential</label>
        </div>

    </fieldset>
      <input type=submit value=Upload>
    </form>
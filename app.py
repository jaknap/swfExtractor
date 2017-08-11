import os, glob
from flask import Flask, render_template, request, redirect, url_for, send_from_directory
from werkzeug import secure_filename
import zipfile
import shutil
import os
from os import listdir
from os.path import isfile, join
import time
from path import path
import ffmpy

# Initialize the Flask application
app = Flask(__name__)

# This is the path to the upload directory
app.config['UPLOAD_FOLDER'] = 'uploads/'
mypath = 'uploads/'
# These are the extension that we are accepting to be uploaded
app.config['ALLOWED_EXTENSIONS'] = set(['pptx'])

# For a given file, return whether it's an allowed type or not
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in app.config['ALLOWED_EXTENSIONS']

# This route will show a form to perform an AJAX request
# jQuery is loaded to execute the request and update the
# value of the operation
@app.route('/')
def index():
    return render_template('index.html')


# Route that will process the file upload
@app.route('/upload', methods=['POST'])
def upload():
    # Get the name of the uploaded file
    file = request.files['file']
    # Check if the file is one of the allowed types/extensions
    if file and allowed_file(file.filename):
        # Make the filename safe, remove unsupported chars
        filename = secure_filename(file.filename)
        print(filename)
        dest_dir = 'uploads/'
        # the upload folder we setup
        
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        for fn in glob.iglob(os.path.join('uploads/', '*.pptx')):
            os.rename(fn, fn[:-4] + 'zip')

        extractZip()
        copytree1(mypath+'/ppt/media',mypath,symlinks=False, ignore=None)
        removePics()
        copytree2('C:/Users/puchil/Documents/upload/uploads/','C:/Users/puchil/Documents/upload/',symlinks=False, ignore=None)
        returnVal = swfconv()
        removeGifs()
        removeFolders()
        # Redirect the user to the uploaded_file route, which
        # will basicaly show on the browser the uploaded file
        
        #return "ok",200
        return redirect(url_for('uploaded_file',
                                filename=returnVal))


def extractZip():
    onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]
    for f in onlyfiles:
      if f.endswith(".zip"):
        file_name = "C:/Users/puchil/Documents/upload/uploads/"+f
        zip_ref = zipfile.ZipFile(file_name,'r')
        zip_ref.extractall(mypath)
        zip_ref.close()



def copytree1(src, dst, symlinks=False, ignore=None):
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks, ignore)
        else:
            shutil.copy2(s, d)

def removePics():
    d = path(mypath)

    #JPG removal    
    filesjpg = d.walkfiles("*.jpg")
    for file in filesjpg:
        file.remove()

    #PNG removal
    filespng = d.walkfiles("*.png")
    for file in filespng:
        file.remove()

def copytree2(src, dst, symlinks=False, ignore=None):
    onlyfiles = [f for f in listdir('C:/Users/puchil/Documents/upload/uploads') if isfile(join('C:/Users/puchil/Documents/upload/uploads', f))]
    gifList=[]
    for f in onlyfiles:
      if f.endswith(".gif"):
        s = os.path.join(src, f)
        d = os.path.join(dst, f)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks, ignore)
        else:
            shutil.copy2(s, d)


   
def swfconv():
    onlyfiles = [f for f in listdir('C:/Users/puchil/Documents/upload/') if isfile(join('C:/Users/puchil/Documents/upload/', f))]
    gifList=[]
    swfList=[]
    for f in onlyfiles:
      if f.endswith(".gif"):
        file_name = f
        gifList.append(file_name)

    for file in gifList:
        ff = ffmpy.FFmpeg(
        inputs={file: None},
        outputs={'output '+file[:-4]+'.swf': '-qscale:v 3'})
        ff.run()


    onlyfiles_new = [f for f in listdir('C:/Users/puchil/Documents/upload/') if isfile(join('C:/Users/puchil/Documents/upload/', f))]
    for f1 in onlyfiles_new:
      if f1.endswith(".swf"):
        file_name1 = f1
        swfList.append(file_name1)

    with zipfile.ZipFile('swfCollection.zip', 'w') as myzip:
        for f in swfList:   
            myzip.write(f)


    onlyfiles_zip = [f for f in listdir('C:/Users/puchil/Documents/upload/') if isfile(join('C:/Users/puchil/Documents/upload/', f))]
    zpList=[]
    for f in onlyfiles_zip:
      if f.endswith(".zip"):
        s = os.path.join('C:/Users/puchil/Documents/upload/', f)
        d = os.path.join('C:/Users/puchil/Documents/upload/uploads', f)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks, ignore)
        else:
            shutil.copy2(s, d)

    zipname="swfCollection.zip"
    return zipname


def removeGifs():
    d = path(mypath)

    #Gif removal
    filesjpg = d.walkfiles("*.gif")
    for file in filesjpg:
        file.remove()

    r = path('C:/Users/puchil/Documents/upload/')
    fileswf = r.walkfiles("*.swf")
    for file in fileswf:
        file.remove()
    filesgif = r.walkfiles("*.gif")
    for file in filesgif:
        file.remove()
    fileszip = r.walkfiles("*.zip")
    for file in fileszip:
        if file[-17:] != "swfCollection.zip":
            file.remove()

def removeFolders():
    shutil.rmtree(mypath+'/_rels')
    shutil.rmtree(mypath+'/docProps')
    shutil.rmtree(mypath+'/ppt')




@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'],
                               filename)

if __name__ == '__main__':
    app.run(
        host="0.0.0.0",
        port=int("80"),
        debug=True
    )

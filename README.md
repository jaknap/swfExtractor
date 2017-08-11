# swfExtractor

A Flask based web utility to get swf files from your pptx files.

Download the ffpmeg command-line utility for your OS:<br>
https://www.ffmpeg.org/download.html


Py Packages needed:<br>
ffmpy: pip install ffmpy<br>
flask: pip install flask

Windows running instructions:<br>
set FLASK_APP=app.py<br>
flask run 

Your app should be running on port 80

-Upload the pptx file on the webpage by clicking the upload button.<br>
-Once uploaded completely on the server, a zip file will be downloaded on the client-side 
 with the swf files having animation from the presentation file. 

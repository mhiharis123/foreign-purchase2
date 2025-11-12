from flask import Flask, send_from_directory
import webbrowser
import threading
import time
import os
import sys

app = Flask(__name__)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Serve static files
@app.route('/')
def index():
    return send_from_directory(resource_path('.'), 'index.html')

@app.route('/<path:path>')
def serve_file(path):
    return send_from_directory(resource_path('.'), path)

def open_browser():
    # Give the server a second to start
    time.sleep(1)
    webbrowser.open('http://127.0.0.1:5000/')

if __name__ == '__main__':
    # Start the browser-opening thread
    threading.Thread(target=open_browser).start()
    
    # Start the Flask server without debug mode for production
    app.run(debug=False)
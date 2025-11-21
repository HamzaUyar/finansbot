import os
import sys
import time
import threading
import webview
import socket
from streamlit.web import cli as stcli

def get_free_port():
    """Find a free port to run Streamlit on."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        s.listen(1)
        port = s.getsockname()[1]
    return port

import subprocess

def run_streamlit(port):
    """Run Streamlit in a subprocess."""
    # Path to the app
    app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app", "streamlit_app.py")
    
    # Run streamlit as a subprocess
    cmd = [
        sys.executable, "-m", "streamlit", "run", app_path,
        "--server.port", str(port),
        "--server.headless", "true",
        "--global.developmentMode", "false",
        "--server.enableXsrfProtection", "false",
        "--server.enableCORS", "false",
        "--browser.gatherUsageStats", "false"
    ]
    process = subprocess.Popen(cmd)
    return process

def start_webview(port, streamlit_process):
    """Start the webview window."""
    # Wait a bit for Streamlit to start
    time.sleep(2)
    
    try:
        webview.create_window(
            "Konsolidasyon Raporu Aracı",
            f"http://localhost:{port}",
            width=1200,
            height=800,
            resizable=True,
            text_select=True,
            zoomable=True
        )
        webview.start()
    finally:
        # Kill streamlit when window closes
        streamlit_process.terminate()

if __name__ == "__main__":
    port = get_free_port()
    
    # Enable downloads
    webview.settings['ALLOW_DOWNLOADS'] = True
    
    # Start Streamlit in a subprocess
    process = run_streamlit(port)
    
    # Start the desktop window
    start_webview(port, process)

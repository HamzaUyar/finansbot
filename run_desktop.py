import multiprocessing
import os
import sys
import time
import socket
import urllib.request
import urllib.error

# Windows frozen executable için kritik - sonsuz process döngüsünü önler
if __name__ == "__main__":
    multiprocessing.freeze_support()


def get_free_port():
    """Find a free port to run Streamlit on."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('', 0))
        s.listen(1)
        port = s.getsockname()[1]
    return port


def is_frozen():
    """Check if we're running as a frozen executable (PyInstaller)."""
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')


def get_base_path():
    """Get the base path for resources."""
    if is_frozen():
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def wait_for_server(port, timeout=30):
    """Wait for the Streamlit server to be ready."""
    start_time = time.time()
    url = f"http://localhost:{port}/_stcore/health"
    
    while time.time() - start_time < timeout:
        try:
            response = urllib.request.urlopen(url, timeout=1)
            if response.status == 200:
                return True
        except (urllib.error.URLError, urllib.error.HTTPError, Exception):
            pass
        time.sleep(0.5)
    
    return False


def run_streamlit_server(port):
    """Run Streamlit server directly (for frozen executable mode)."""
    try:
        # Set environment variables for frozen mode
        base_path = get_base_path()
        
        # Add base path to Python path so imports work
        if base_path not in sys.path:
            sys.path.insert(0, base_path)
        
        app_path = os.path.join(base_path, "app", "streamlit_app.py")
        
        # Verify the app file exists
        if not os.path.exists(app_path):
            print(f"ERROR: App file not found at {app_path}")
            print(f"Base path: {base_path}")
            print(f"Contents: {os.listdir(base_path) if os.path.exists(base_path) else 'N/A'}")
            return
        
        # Import and run streamlit
        from streamlit.web import cli as stcli
        
        sys.argv = [
            "streamlit", "run", app_path,
            "--server.port", str(port),
            "--server.headless", "true",
            "--global.developmentMode", "false",
            "--server.enableXsrfProtection", "false",
            "--server.enableCORS", "false",
            "--browser.gatherUsageStats", "false",
            "--server.fileWatcherType", "none"
        ]
        stcli.main()
    except Exception as e:
        print(f"Streamlit error: {e}")
        import traceback
        traceback.print_exc()


def run_streamlit_subprocess(port):
    """Run Streamlit in a subprocess (for development mode)."""
    import subprocess
    
    base_path = get_base_path()
    app_path = os.path.join(base_path, "app", "streamlit_app.py")
    
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


def show_error_window(message):
    """Show an error message in a webview window."""
    import webview
    
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
                margin: 0;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            }}
            .error-box {{
                background: white;
                padding: 40px;
                border-radius: 12px;
                box-shadow: 0 10px 40px rgba(0,0,0,0.3);
                text-align: center;
                max-width: 500px;
            }}
            h1 {{ color: #e74c3c; margin-bottom: 20px; }}
            p {{ color: #555; line-height: 1.6; }}
        </style>
    </head>
    <body>
        <div class="error-box">
            <h1>⚠️ Hata</h1>
            <p>{message}</p>
        </div>
    </body>
    </html>
    """
    
    webview.create_window("Hata", html=html, width=600, height=400)
    webview.start()


def start_webview(port, streamlit_process=None):
    """Start the webview window."""
    import webview
    
    # Wait for Streamlit to actually start
    print(f"Waiting for Streamlit server on port {port}...")
    
    if not wait_for_server(port, timeout=30):
        show_error_window(
            "Streamlit sunucusu başlatılamadı. Lütfen uygulamayı tekrar çalıştırın. "
            "Sorun devam ederse, antivirüs yazılımınızı kontrol edin."
        )
        if streamlit_process:
            streamlit_process.terminate()
        return
    
    print("Streamlit server is ready!")
    
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
        if streamlit_process:
            streamlit_process.terminate()


def main():
    import webview
    import threading
    
    port = get_free_port()
    print(f"Using port: {port}")
    print(f"Frozen: {is_frozen()}")
    print(f"Base path: {get_base_path()}")
    
    # Enable downloads
    webview.settings['ALLOW_DOWNLOADS'] = True
    
    if is_frozen():
        # Frozen mode: Streamlit'i ayrı thread'de çalıştır
        streamlit_thread = threading.Thread(
            target=run_streamlit_server,
            args=(port,),
            daemon=True
        )
        streamlit_thread.start()
        
        # Give thread a moment to start
        time.sleep(1)
        
        # Desktop window'u başlat
        start_webview(port)
    else:
        # Development mode: subprocess kullanabiliriz
        process = run_streamlit_subprocess(port)
        start_webview(port, streamlit_process=process)


if __name__ == "__main__":
    main()

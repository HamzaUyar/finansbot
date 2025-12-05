import multiprocessing
import os
import sys
import time
import socket

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


def run_streamlit_server(port):
    """Run Streamlit server directly (for frozen executable mode)."""
    # Streamlit'i doğrudan import edip çalıştırıyoruz
    # Bu subprocess döngüsünü önler
    from streamlit.web import cli as stcli
    
    app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app", "streamlit_app.py")
    
    # Frozen mode'da __file__ farklı çalışabilir
    if is_frozen():
        # PyInstaller'da _MEIPASS geçici dizini kullanılır
        app_path = os.path.join(sys._MEIPASS, "app", "streamlit_app.py")
    
    sys.argv = [
        "streamlit", "run", app_path,
        "--server.port", str(port),
        "--server.headless", "true",
        "--global.developmentMode", "false",
        "--server.enableXsrfProtection", "false",
        "--server.enableCORS", "false",
        "--browser.gatherUsageStats", "false"
    ]
    stcli.main()


def run_streamlit_subprocess(port):
    """Run Streamlit in a subprocess (for development mode)."""
    import subprocess
    
    app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "app", "streamlit_app.py")
    
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


def start_webview(port, streamlit_process=None, streamlit_thread=None):
    """Start the webview window."""
    import webview
    
    # Wait a bit for Streamlit to start
    time.sleep(3)
    
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
        if streamlit_process:
            streamlit_process.terminate()


def main():
    import webview
    import threading
    
    port = get_free_port()
    
    # Enable downloads
    webview.settings['ALLOW_DOWNLOADS'] = True
    
    if is_frozen():
        # Frozen mode: Streamlit'i ayrı thread'de çalıştır
        # subprocess kullanamayız çünkü exe kendini tekrar açar
        streamlit_thread = threading.Thread(
            target=run_streamlit_server,
            args=(port,),
            daemon=True
        )
        streamlit_thread.start()
        
        # Desktop window'u başlat
        start_webview(port, streamlit_thread=streamlit_thread)
    else:
        # Development mode: subprocess kullanabiliriz
        process = run_streamlit_subprocess(port)
        start_webview(port, streamlit_process=process)


if __name__ == "__main__":
    main()

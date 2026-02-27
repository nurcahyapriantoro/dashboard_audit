import sys
import os
import streamlit.web.cli as stcli

def resolve_path(path):
    if getattr(sys, 'frozen', False):
        basedir = sys._MEIPASS
    else:
        basedir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(basedir, path)

if __name__ == "__main__":
    # Point to the internal app.py
    app_path = resolve_path("app.py")
    
    # Mock the command line arguments for streamlit
    sys.argv = [
        "streamlit",
        "run",
        app_path,
        "--global.developmentMode=false",
        "--server.headless=false",  # Set to false to auto-open browser
        "--browser.gatherUsageStats=false", # Disable usage stats
        "--server.port=8501",
    ]
    
    sys.exit(stcli.main())

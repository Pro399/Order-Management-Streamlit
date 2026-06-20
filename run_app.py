import os
import sys
import threading
import webbrowser
import time
import traceback
import multiprocessing

def open_browser():
    """Waits for the server to boot, then opens the browser."""
    time.sleep(3)
    webbrowser.open("http://localhost:8501")

def main():
    try:
        from streamlit.web import cli as stcli
        
        # Detect the path of the current script or executable
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))

        # Construct the path to app.py
        app_path = os.path.join(application_path, 'app.py')

        # Force environment variables to strictly disable Dev Mode
        os.environ["STREAMLIT_GLOBAL_DEVELOPMENT_MODE"] = "false"

        print("Starting Item Invoicing System...")
        print("Please wait, opening in your web browser...")
        print("--------------------------------------------------")
        print("⚠️ DO NOT CLOSE THIS WINDOW while using the app.")
        print("Closing this window will shut down the local server.")
        print("--------------------------------------------------")

        # Start a background thread to open the browser
        threading.Thread(target=open_browser, daemon=True).start()

        # Run the Streamlit app internally with strict port and mode overrides
        sys.argv = [
            "streamlit", 
            "run", 
            app_path, 
            "--server.port=8501", 
            "--server.headless=true", 
            "--global.developmentMode=false"
        ]
        sys.exit(stcli.main())
        
    except Exception as e:
        print("\n" + "="*50)
        print("❌ CRITICAL ERROR PREVENTED APP FROM STARTING:")
        print("="*50)
        traceback.print_exc()
        print("="*50)
        input("\nPress ENTER to close this window...")

if __name__ == "__main__":
    # CRITICAL: Prevents infinite loop / fork-bombing in PyInstaller on Windows
    multiprocessing.freeze_support()
    main()
import sys
import os
from streamlit.web import cli as stcli

def run_streamlit():
    # 定位打包后的main.py
    script_path = os.path.join(os.path.dirname(sys.executable), "main.py")
    sys.argv = ["streamlit", "run", script_path, "--server.headless=true", "--server.port=8501", "--browser.gatherUsageStats=false"]
    stcli.main()

if __name__ == "__main__":
    run_streamlit()
import os
import subprocess
import sys

# 1. Path to actual logic file
script_path = "proc_workbench.py"

# 2. Tell Streamlit to run on port 8080 (HICP default)
print("Starting Procurement Workbench via Streamlit...")
subprocess.run([
    "streamlit", 
    "run", 
    script_path, 
    "--server.port", "8080", 
    "--server.address", "0.0.0.0",
    "--browser.gatherUsageStats", "false"
])
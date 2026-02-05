import subprocess
import os

# This tells the cloud: "Don't run me, run proc_workbench.py using Streamlit instead!"
if __name__ == "__main__":
    subprocess.run([
        "streamlit", 
        "run", 
        "proc_workbench.py", 
        "--server.port", "8080", 
        "--server.address", "0.0.0.0"
    ])
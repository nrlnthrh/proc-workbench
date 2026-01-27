import subprocess
import os

# The cloud start streamlit using the correct port
subprocess.run([
    "streamlit", "run", "app.py", 
    "--server.port", "8080", 
    "--server.address", "0.0.0.0"
])

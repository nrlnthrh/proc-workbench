#!/bin/bash 
# we use port 8080 because that is what OpenShift Services using 
streamlit run proc_workbench.py --server.port 8080 --server.address 0.0.0.0

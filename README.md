# Procurement Workbench v9.0

## Overview
Procurement Workbench is a Streamlit-based web application designed to automate the validation and analysis of procurement data. It uses a **Hybrid Rule Engine** that combines hardcoded business logic with dynamic rules loaded from external configuration files.

## Features
- **SMD Analysis:** Regional validation (APAC/EU), duplicate detection, and dynamic field-rule enforcement via Excel configuration.
- **PO Analysis:** Intelligent PO status classification (Service vs. Material vs. Direct) and audit of 16+ categories including PCN, UOM, and Split Accounting.
- **Email Validation:** Data integrity checks for vendor contact lists.
- **Automated Reporting:** Generates professionally formatted Excel workbooks with color-coded errors and dashboard summaries.

## Installation
1. Ensure Python 3.9+ is installed.
2. Install dependencies:
   <> bash 
   pip install streamlit pandas numpy xlsxwriter

# HOW TO RUN
<> bash
streamlit run proc_workbench.py

# CONFIGURATION
The tool relies on two primary configuration files: 
   **1. SMD_Rules_Config.xlsx:** Contains rules for mandatory fields, postal codes, and reference lists. 
   **2. PO_Rules_Config.xlsx:** Contains the logic matrix for PO status and standard rule parameters.

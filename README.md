PROCleans

📄 Overview
**PROCleans** is a high-performance data auditing suite built with Streamlit. It automates the validation of procurement datasets by utilizing a **Hybrid Rule Engine**, combining hardcoded financial logic with dynamic, user-definable rules managed via Excel configuration files.

✨ Key Features
- **SMD Analysis:** Regional validation (APAC), postal code integrity, and Synertrade-to-Vendor duplicate detection.
- **PO Analysis:** Automated classification (Indirect Service/Indirect Material/Direct) and audit of 16+ categories including PCN, UOM, and Text/Special Character checks.
- **Email Validation:** Chronological ID logic to ensure primary contact default flags are correctly assigned.
- **Automated Reporting:** Generates professional formatted excel file (.xlsx) reports with dashboard summaries and color-coded error highlighting. 

⚙️ Installation & Setup
1. Prerequisites
   - Python 3.9 or higher
   - Microsoft Excel (to manage the configuration file)

2. Install dependencies:
   Open your terminal/command prompt and run:
     ---------------------------------------------------------------
     | <> bash                                                    |
     | ---------------------------------------------------------- |
     | pip install streamlit pandas numpy xlsxwriter openpyxl     |
     ---------------------------------------------------------------  

3. Project Structure
   To ensure the app runs correctly, organize your folder like this: 
   __________________________________________________________________
   | <> Text                                                        |
   |------------------------------------------------------------    |
   | PROCleans/                                                     |
   | |- proc_workbench.py           # The main application code     |
   | |- SMD_Rules_Config.xlsx       # Rules for SMD Analysis        |
   | |- PO_Rules_Config.xlsx        # Logic matrix for PO Analysis  |
   | |- requirements.txt            # List of dependencies          |
   |________________________________________________________________|

   # HOW TO RUN
   In your terminal, navigate to the project folder and execute:
     ---------------------------------------------------------------
     | <> bash                                                    |
     | ---------------------------------------------------------- |
     | streamlit run proc_workbench.py                            |
     ---------------------------------------------------------------  

   # CONFIGURATION
   The workbench is designed to be "No-Code" for daily updates. You can chance logic without touching the Python script by editing the following files: 
      **1. SMD_Rules_Config.xlsx:** Define mandatory fields, allowed values in dropdowns (Reference Lists), and postal code lengths by country. 
      **2. PO_Rules_Config.xlsx:** Manage the Logic Matrix (setting status to Open/Review/Close based on the PO requirements) and update banned requester lists or PCN codes.

📊 Usage Workflow
   1. Launch the app vie the command line. 
   2. Select your desired module (SMD, PO, or Email) from the Sidebar.
   3. Upload the corresponding Rules Config file first.
   4. Upload your Raw Data file.
   5. Click *"Run Analysis"* and download the generated Excel report.

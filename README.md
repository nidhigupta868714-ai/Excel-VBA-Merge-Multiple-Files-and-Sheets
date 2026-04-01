## Excel VBA – Merge Multiple Files and Sheets into One Master Sheet

Excel VBA automation to merge data from multiple Excel files and all worksheets into a single master sheet while keeping the header only once.

## 📊 Project Overview

This project demonstrates how Excel VBA automation can be used to consolidate data from multiple Excel files and all their worksheets into a single master sheet.

The macro prompts the user to select a folder that contains Excel files. It automatically opens each file, loops through every worksheet, and merges the data into a newly created sheet in the current workbook. Since all source files share the same structure, the header is copied only once to maintain a clean and organized dataset.

This solution eliminates manual copy-paste work and is highly useful for MIS reporting, data management, and repetitive data consolidation tasks.

## 🚀 Key Features
- Folder picker to select source files
- Reads data from all worksheets of all Excel files
- Merges data into a new sheet automatically
- Ensures the header appears only once
- Handles any number of files and rows dynamically
- Saves significant manual effort and time

## 🛠️ Skills Demonstrated
- VBA Automation
- File Handling using VBA
- Looping through Workbooks and Worksheets
- Dynamic Range Detection
- Data Consolidation Techniques

## ▶️ How to Use
- Open Excel and press Alt + F11
- Insert a new Module
- Paste the VBA code
- Run the macro
- Select the folder containing the Excel files
- All data will be combined into a newly created master sheet

## ✅ Use Cases
- MIS Reporting
- Data consolidation from multiple files
- Monthly/Weekly report merging
- Reducing repetitive manual Excel work

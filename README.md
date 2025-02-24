# **Automated Data Discrepancy Check Tool**

## **Project Overview**

This Python-based automation tool streamlines the detection and reporting of data discrepancies in parts material information data. It ensures accuracy in these datasets by automating data ***cleaning, comparison, and reporting processes***. The tool ***generates structured reports*** on inconsistencies, improving efficiency and reducing manual errors.

## **Key Features**

1. **Automated Data Extraction & Cleaning**
    - Reads and processes Excel files (`.xlsx`, `.xlsm`) and CSV files from multiple sources.
    - Applies data cleaning techniques such as removing duplicates, formatting date fields, standardizing part numbers, and handling missing values.
    - Ensures consistency in column names, order, and data types across datasets.
2. **Data Comparison & Discrepancy Detection**
    - Compares multiple datasets, including MF Final Order vs. OSP Final Order, MF Forecast vs. OSP Forecast, and Order Plan files.
    - Identifies row mismatches, incorrect values, missing data, and discrepancies in key fields such as `PARTNO`, `DEST_CODE`, and `CONT_CODE`.
    - Uses various comparison algorithms to detect order quantity mismatches, cumulative sum inconsistencies, and incorrect final order values.
3. **Structured Discrepancy Reporting**
    - Generates discrepancy reports in CSV format, categorized into unmatched entries and incorrect values.
    - Saves reports in dynamically created folders, ensuring proper file organization.
    - Provides observations and notes for further review.
4. **User-Friendly Interaction & Execution**
    - Offers a menu-based selection system, allowing users to choose specific file comparisons.
    - Includes prompts for additional user input, such as exporting reports or handling cumulative sum calculations.
    - Organizes files efficiently by creating daily timestamped folders for better tracking.

## **Impact & Benefits**

- **Increased Accuracy:** Eliminates manual errors in packaging data reconciliation.
- **Efficiency Gains:** Saves significant time by automating daily file comparisons and reporting.
- **Improved Process Transparency:** Provides clear, structured reports for better decision-making.
- **Proactive Issue Resolution:** Helps identify and correct data discrepancies quickly.

This project significantly enhances the data validation process, ensuring more reliable data handling and operational efficiency.

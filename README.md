This project automates the end-to-end processing of NYC traffic accident data, integrating data preprocessing, database management, and analytical reporting. It requires Microsoft Access, Excel, and Python (including the pandas library).

The workflow consists of three primary components: data cleaning, database ingestion, and report generation. The script clean_data.py preprocesses raw accident data, standardizing formats and handling missing values. The VBA macro import_to_access.vba facilitates structured data import into MS Access, ensuring optimized storage. The generate_report.vba macro extracts insights via SQL queries and compiles them into structured reports in Excel.

Users must verify file paths, ensure database accessibility, and adjust macro security settings to prevent execution errors. This framework establishes a robust, automated pipeline for systematic traffic accident analysis.

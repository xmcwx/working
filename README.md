# Master Excel Tool

## Overview

The Master Excel Tool is a powerful Go application designed to process multiple Excel spreadsheets and consolidate specific data into a master output file. It's particularly useful for identifying and reporting on systems with vulnerable or end-of-life operating systems across multiple clients.

## Key Features

- **Batch Processing**: Processes multiple Excel files from a specified input directory.
- **Selective Data Extraction**: Allows users to choose which columns to extract from the source files.
- **Client Identification**: Extracts client names from a designated "Account" sheet or derives them from filenames.
- **Vulnerability Detection**: Identifies end-of-life Windows versions and Linux systems.
- **Consolidated Reporting**: Creates a master Excel file with separate sheets for each client, focusing on vulnerable systems.
- **Safe Client Logging**: Generates a separate log of clients without detected vulnerabilities.
- **Detailed Logging**: Provides comprehensive logging of the entire process for auditing and troubleshooting.

## How It Works

1. The tool scans an input directory for Excel files.
2. For each file, it extracts the client name and relevant data from specified columns.
3. It identifies systems with end-of-life Windows versions or Linux kernels.
4. The tool creates a new sheet in the master Excel file for each client, populating it with data on vulnerable systems.
5. If no vulnerabilities are found for a client, it logs this information separately.
6. The process includes error handling, timeout management, and detailed logging throughout.

## Use Cases

- IT Security Audits: Quickly identify vulnerable systems across multiple client environments.
- Compliance Reporting: Generate reports on operating system versions for regulatory compliance.
- Asset Management: Track and report on the operating system landscape across various clients or departments.

## Output

- A master Excel file containing consolidated data on vulnerable systems.
- A text file listing clients with no detected vulnerabilities.
- A detailed log file for process auditing and troubleshooting.

This tool streamlines the process of identifying and reporting on vulnerable systems across multiple Excel spreadsheets, making it an invaluable asset for IT security professionals and system administrators.



# Sheet2Report

## Overview

Sheet2Report is a Go application designed to process Excel spreadsheets and generate comprehensive reports based on specific criteria. It's particularly useful for analyzing and summarizing data across multiple sheets, focusing on key metrics.

## Key Features

- **Multi-Sheet Processing**: Analyzes data from multiple sheets within a single Excel file.
- **Flexible Data Extraction**: Configurable to extract and process specific columns from each sheet.
- **Vulnerability Detection**: Identifies systems with end-of-life operating systems or other specified vulnerabilities.
- **Summary Generation**: Creates a summary sheet with key metrics and findings.
- **Detailed Reporting**: Generates individual report sheets for each analyzed sheet, highlighting important data points.
- **Data Aggregation**: Consolidates information across sheets for comprehensive analysis.
- **Error Handling**: Robust error checking and logging for reliable operation.

## How It Works

1. The tool opens a specified Excel file and iterates through its sheets.
2. For each sheet, it extracts relevant data based on predefined column mappings.
3. It processes the data, identifying vulnerabilities and calculating key metrics.
4. A summary sheet is created, providing an overview of findings across all sheets.
5. Individual report sheets are generated for each analyzed sheet, containing detailed information.
6. The tool handles various data formats and potential errors, ensuring reliable output.

## Use Cases

- **IT Asset Management**: Track and report on system vulnerabilities across different departments or locations.
- **Compliance Reporting**: Generate reports on system statuses for regulatory compliance.
- **Security Audits**: Quickly identify and summarize vulnerable systems within an organization.
- **Resource Planning**: Analyze system distribution and identify areas needing upgrades or replacements.

## Output

- An updated Excel file containing:
  - A summary sheet with aggregated data and key findings.
  - Individual report sheets for each analyzed sheet, highlighting important information.
  - The original data sheets, preserved for reference.

## Benefits

- **Time-Saving**: Automates the process of analyzing and reporting on large datasets.
- **Consistency**: Ensures uniform analysis and reporting across all data sheets.
- **Customizable**: Can be adapted to focus on different data points or criteria as needed.
- **Scalable**: Capable of processing files with multiple sheets and large amounts of data.

Sheet2Report streamlines the process of extracting, analyzing, and reporting on complex Excel data, making it an invaluable tool for IT professionals, auditors, and managers dealing with large-scale system information.

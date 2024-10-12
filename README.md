# working

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

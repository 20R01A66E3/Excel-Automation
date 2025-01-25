# Excel Automation

## Description

The **Excel Automation Utility** is a Python-based tool for processing and enhancing Excel spreadsheets. This project simplifies repetitive tasks such as modifying cell values based on mathematical operations and generating bar charts to visualize data. With robust error handling, it ensures smooth operation even when encountering invalid inputs or missing files.

## Features
- **Process and Correct Data**: Apply mathematical operations (addition, subtraction, multiplication, division) to a specified column and save results in a new column.
- **Generate Bar Charts**: Create bar charts to visualize data from the processed spreadsheet.
- **Error Handling**: Handles missing files, invalid sheet names, empty cells, non-numeric values, and division by zero gracefully.

## Installation

1. **Prerequisites**:
   - Python 3.6 or later
   - `openpyxl` library

2. **Install Dependencies**:
   Run the following command to install the required library:
   ```bash
   pip install openpyxl
Download the Script: Download the Automation.py file to your project directory.
## Usage
### Processing and Correcting Data
The process_sheets function processes a specific column in an Excel file and applies a mathematical operation to modify its values.

#### Function Parameters:
filename: The name of the Excel file to process.
sheet_name: The sheet in the workbook to process.
column_to_correct: The column number (starting from 1) to apply the operation.
correction_value: The value to use in the operation.
operation_symbol: The mathematical operation (+, -, *, /).

### Creating a Bar Chart
The create_chart function generates a bar chart based on the processed data.

#### Function Parameters:
filename: The name of the Excel file to process.
sheet_name: The sheet in the workbook to create the chart.

## Contributing
Contributions are welcome! If you have suggestions or find bugs, please open an issue or submit a pull request.

#### How to Contribute:
Fork the repository.
Create a new branch for your feature (git checkout -b feature-name).
Commit your changes (git commit -m "Add feature name").
Push to the branch (git push origin feature-name).
Open a pull request.

## Author Information
Author: Sairam Konakanchi

Contact: https://github.com/20R01A66E3


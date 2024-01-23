# trademe_sold_forecast_per_sku
Sales Data Processing Tool
Description

This Python script processes sales data from a specified input Excel file and generates a summary in a new Excel file. The tool calculates various statistics for each SKU, including the first and last sale dates, total sales quantity, daily average sales, and monthly average sales. The resulting data can be useful for inventory management and sales forecasting.
Prerequisites

    Python 3.x
    Required Python packages: openpyxl, datetime

Usage

    Ensure you have Python 3.x installed on your system.
    Install required packages by running:

    bash

pip install openpyxl

Run the script by executing the following command in your terminal:

bash

    python sales_data_processing.py

    The script will prompt you to enter the input file path (e.g., C:\\Users\\ThinkPad\\SynologyDrive\\Trademe\\sold.xlsx) and the desired output file path (e.g., C:\\Users\\ThinkPad\\SynologyDrive\\Trademe\\result.xlsx).

Input File Format

The script expects an Excel file with specific columns. Make sure your input file adheres to the required format.
Output File

The processed data will be saved in a new Excel file specified by the user.
Note

    If you encounter a FileNotFoundError, double-check the file path and ensure the file exists.
    The script uses date and time information from the input file to calculate statistics.
    Customize the input and output file paths as needed for your specific use case.

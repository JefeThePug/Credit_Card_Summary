# Credit Card Summary Web Scraping Tool

## Overview

This tool is designed to scrape credit card purchase data from a bank's website, categorize the purchases, and generate a financial summary in an Excel file. The summary can then be exported as a PDF and emailed to a specified recipient.

## Features

- **Web Scraping**: Collects credit card purchases for a given month from the bank's website using HTML parsing with Beautiful Soup.
- **Categorization**: Organizes purchases into predefined categories based on a customizable dictionary. Unrecognized purchases can be updated later.
- **Excel Integration**: Generates an Excel file that allows for manual adjustments and summary viewing of categorized purchases.
- **VBA Automation**: Includes VBA scripts to facilitate data management, PDF generation, and emailing functionality.

## Components

### 1. `Get My Data.app`

This app was created in Apple's `Automator` 
- Captures the HTML content of the bank's page.
- Stores it as a text file in the project's directory.
- Launches the `html_script.py` python script to begin parsing the HTML.
- Launches the excel file (`ExcelPasterApp.xlsm`) to continue the process.

### 2. `html_script.py`

This Python script performs the following tasks:

- Reads HTML content from the file created by `Get My App.app`.
- Parses the content to extract purchase data.
- Categorizes each purchase based on predefined categories.
- Writes the organized data into an Excel file with specified formatting. 

### 3. `ExcelPasterApp.xlsm`

The Excel file contains several VBA scripts that automate tasks such as:

- Clearing previous data upon opening or closing the workbook.
- Copying data from one worksheet to another based on user input.
- Saving the summary as a PDF and emailing it using the `emailpdf.py` script by way of `PythonCommand.scpt`.

### 4. `PythonCommand.scpt`

This AppleScript is used to execute the Python email script from within Excel:

> **Note**: For this to work, you must copy the script file `PythonCommand.scpt` and paste it in the directory: `/Users/<username>/Library/Application Scripts/com.microsoft.Excel/` for the computer you wish to run this on.

### 5. `emailpdf.py`

This Python script handles the emailing of the generated PDF. It utilizes the SMTP protocol to send the PDF as an attachment to a specified email address. The script requires the email address and password to authenticate with the email server.

## Usage Instructions

### Setup

1. Ensure you have Python installed with the following packages:
   - `pandas`
   - `beautifulsoup4`
   - `openpyxl`

2. Copy the `PythonCommand.scpt` file into the appropriate directory.

3. Customize the `charges` dictionary in `html_script.py` to reflect your specific categorization needs.
> **Note**: This tool is intended for use with one specific credit card company and its charges reflect the frequent charge of the user it was made for. A significant amount of customization to the python code and excel file will be required if you intend to implement this elsewhere.

### Running the Tool

1. Run `Get My Data.app`
2. When the `ExcelPasterApp.xlsm` opens, navigate to the appropriate month.
3. Make small manual adjustments to the data as required.
> For the most part, the data should already be organized but there may be a few unique purchases that are not recognized.
4. Send this PDF to your email for your records by clicking the "Save Summary" button on the Excel sheet.

### Email Configuration

- Update the `sender` variable in `emailpdf.py` with your email address.
- In the Excel's VBA, replace the `;password` in the `AppleScriptTask` with a semicolon and then your authenticated password. Eg. `;Abc123`
- Also in the same VBA, Update the `email` variable to be the recipient's email.

## Notes

- Ensure that your email provider allows SMTP access and that you have enabled any necessary settings (e.g., less secure app access).
- The script assumes that the structure of the HTML from the bank's website does not change. Any modifications to the HTML structure may require updates to the parsing logic in `html_script.py`.

## License

This project is licensed under the MIT License. See the LICENSE file for more details.

## Acknowledgments

- This tool is intended for personal use and should not be used for commercial purposes without proper modifications and testing.

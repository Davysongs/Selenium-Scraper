# Selenium Web Scraping Project

This project provides a Python script to scrape data from a list of URLs specified in a DOCX file using Selenium. The extracted data is then stored in an Excel file in a structured format. The script extracts basic information such as titles, text, and hyperlinks from each web page.

## Table of Contents

- [Project Overview](#project-overview)
- [Requirements](#requirements)
- [Installation](#installation)
- [Setup and Configuration](#setup-and-configuration)
- [Script Workflow](#script-workflow)
- [Exporting Data to Excel](#exporting-data-to-excel)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## Project Overview

This project aims to automate web scraping tasks using Selenium. It reads URLs from a DOCX file, visits each URL, and extracts relevant information. The data is then exported to an Excel file for further analysis or sharing.

## Requirements

To run this project, ensure you have the following installed:

- **Python** (Version 3.6 or later): The scripting language used for this project.
- **Selenium**: A popular framework for automating web browsers.
- **ChromeDriver**: Selenium's driver for automating Google Chrome.
- **Pandas**: A data manipulation library used to create and export data to Excel.
- **python-docx**: A library to read DOCX files.
- **openpyxl**: A library to handle Excel files.

## Installation

To install the necessary packages, use the following command:

```bash
pip install selenium pandas python-docx openpyxl
```

Additionally, ensure you have downloaded ChromeDriver and set it up with the correct path. You can download ChromeDriver from the [official website](https://sites.google.com/a/chromium.org/chromedriver/).

## Setup and Configuration

1. **Setting Up ChromeDriver**:
   - Download the appropriate version of ChromeDriver based on your Chrome browser version.
   - Extract the downloaded file and place it in a directory accessible to your Python script.
   - Update the `chromedriver_path` variable in the script with the correct path to ChromeDriver.

2. **Reading the DOCX File**:
   - Ensure you have a DOCX file containing a list of URLs. The script expects URLs to be in separate lines or paragraphs.
   - Update the `doc_path` variable in the script with the correct path to the DOCX file.

3. **Excel Output File**:
   - Specify the path where you want to save the Excel file. Update the `output_path` variable accordingly.

## Script Workflow

The script follows these steps:

1. **Initialize Selenium WebDriver**:
   - Set up Selenium with ChromeDriver and headless mode (optional).
   - Configure WebDriver with the required options.

2. **Read URLs from the DOCX File**:
   - Use `python-docx` to read the file and extract URLs into a list.

3. **Scrape Data from Each URL**:
   - Open each URL using Selenium.
   - Extract relevant information such as page title, body text, and hyperlinks.
   - Store the extracted data in a list of dictionaries.

4. **Export Data to Excel**:
   - Convert the list of dictionaries into a Pandas DataFrame.
   - Export the DataFrame to an Excel file using `openpyxl`.

## Exporting Data to Excel

The data extracted from the URLs is stored in an Excel file with structured rows and columns. This allows for easy manipulation and analysis using Excel or other data analysis tools.

The following information is stored in the Excel file:
- **URL**: The original URL from which the data was scraped.
- **Title**: The title of the web page.
- **Body Text**: A summary or excerpt from the body of the page.
- **Links**: A list of hyperlinks found on the page.

## Troubleshooting

If you encounter issues, consider the following:

- **ChromeDriver Version Mismatch**:
  Ensure that the ChromeDriver version matches your Chrome browser version.
- **Script Errors**:
  Check the error messages in the console for specific issues.
- **Timeouts**:
  If the script takes too long, try adjusting the `time.sleep()` durations to ensure pages fully load.

## Contributing

Contributions to this project are welcome. If you'd like to suggest improvements or report issues, please open a GitHub issue or submit a pull request with your changes.

## License

This project is open source and licensed under the [MIT License](https://opensource.org/licenses/MIT). You are free to use, modify, and distribute the code, but remember to attribute the original authors.

```

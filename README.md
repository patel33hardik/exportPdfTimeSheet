# exportPdfTimeSheet

## Overview

### Case
As part of my weekly routine at work, every Thursday I am responsible for opening an Excel sheet. Once opened, I perform a filter to extract the relevant data. I then export the filtered data as a PDF file and apply a digital signature to it. Lastly, I send the PDF document via email to the Account department for their records. This task is crucial for ensuring accurate and up-to-date financial records are maintained.

### Solution
The "Export pdf timesheet" script is used for copying an Excel file from either a remote or local drive. It uses the `win32com.client` Python library to open the file and add macros. These macros automatically change specific cell values and perform filters, and at the end, the filtered results are exported to a PDF file.

## How to Use

1. **Prerequisites**
   - Ensure Python is installed on your system.
   - Install the required libraries using the following command:
     ```bash
     pip install win32com
     ```

2. **Clone the Repository**
   ```bash
   git clone https://github.com/patel33hardik/exportPdfTimeSheet.git
   cd exportPdfTimeSheet

## License
This project is licensed under the MIT License.
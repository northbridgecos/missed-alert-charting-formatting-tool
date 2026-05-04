# Missed Alert Charting Formatting Tool
Formats and aggregates missed alert charting days report into a visually intuitive excel file.  
Processes locally on your machine. No data is shared outside of your local environment.

## Download
[**Click here to download the latest version**](https://github.com/northbridgecos/missed-alert-charting-formatting-tool/archive/refs/tags/v1.0.0.zip)

### Preview
![App](/assets/app.png)

### Original
![Original](/assets/original.png)

### Formatted
![Formatted](/assets/formatted.png)

## To Use
1. Download the Yardi report
- Yardi CRM & Carestream
- Yardi EHR
- Reports
- Scheduling
- Missed Monthly Alert Charting Report
- Apply filters, Specify community and month
- Excel
2. Launch the Alert Charting Tool app
3. Press 'Select File & Run'
4. Choose input file (downloaded from Yardi in step 1)
5. Choose output file name and location

## Built with 
* **Python 3.x** - Core programming language.
* **Tkinter** - For the graphical user interface.
* **Pandas** - Used for data manipulation and Excel aggregation.
* **Openpyxl** - Engine for Excel formatting.
* **PyInstaller** - Used to compile the source into a standalone executable.

## Developers
1. Make sure Python is installed
2. Clone the repository  
`git clone https://github.com/northbridgecos/missed-alert-charting-formatting-tool.git`
3. Navigate to the folder  
`cd missed-alert-charting-formatting-tool`
4. Install necessary libraries  
`pip install pandas openpyxl`
or
`pip install -r requirements.txt`
5. Run the application  
`python main.py`

## Architecture
* `main.py`: Handles the Tkinter GUI logic, file dialogs, and user interactions.
* `process.py`: Contains the processing logic. Reads the Yardi report, cleans the data, and formats the Excel output.

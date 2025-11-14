# Cash Sheet Auto-Filler

This is a Python script designed to automate the daily process of filling out cash sheet workbooks. It reads daily sales reports from two different POS systems (Infor and Tavlo), parses the data, and automatically fills in the corresponding weekly cash sheet Excel file.

## Features

- Parses **Infor (.csv)** reports and **Tavlo (.xls/XML)** reports.
- Automatically locates the correct cash sheet workbook and worksheet based on the date (e.g., "Wednesday").
- Maps report data to the corresponding row in the cash sheet workbook (e.g., "Crimson Corner", "U.S.S.").
- Translates tender types from both POS systems into a standard format for the cash sheet.
- Validates the filled data by checking the "Over/Short" column in the cash sheet.
- Provides a simple `run_autofiller.bat` file for one-click execution on Windows.

## How It Works

The program works in a three-step process:

1. **Read**: The `main.py` script asks for a date, scans the `reports` folder for all `.csv` and `.xls` files, and determines the correct weekday.
2. **Parse**: It uses the appropriate parser (`infor_parser.py` or `tavlo_parser.py`) to read the report file, find the location name, and extract all financial data (sales, tax, tenders, count).
3. **Write**: It uses `excel_autofiller.py` to:
   - Locate the correct cash sheet workbook (e.g., `Crimson Corner.xlsx`) using the map in `config.py`.
   - Open the correct worksheet (e.g., "Wednesday").
   - Find the correct row for that location (e.g., "Thirst").
   - Fill in all the parsed data into the correct columns.
   - Save the file and re-check its calculations for errors.

## Requirements

- Python 3.x
- All libraries listed in `requirements.txt`

## ⚙️ Installation (For a New Setup)

1. Clone this repository:
   ```bash
   git clone https://github.com/Seaw24/Cash-sheet-autofill.git
   ```
2. Navigate into the project folder:
   ```bash
   cd Cash-sheet-autofill
   ```
3. (Recommended) Create and activate a virtual environment (this helps isolate dependencies):
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```
4. Install the required libraries listed in `requirements.txt`:
   ```bash
   pip install -r requirements.txt
   ```
5. Proceed to the **Configuration** section.

## 🔧 Configuration (One-Time Setup)

## 🔧 Configuration (One-Time Setup)

Before you can run the program, you **must** edit the `src/config.py` file to match your computer's paths and casheet setup.

1. **`REPORTS_FOLDER` & `CASHEET_FOLDER`**:

   - `REPORTS_FOLDER` is now set automatically; it finds the reports folder in your main project directory (one level above `src`). You should not need to change this.
   - `CASHEET_FOLDER` must be updated manually (e.g., every week). Paste the **full absolute path** to your current week's casheet folder.
   - **Important:** The path must start with an `r` to be a "raw string."
   - Example:
     ```
     CASHEET_FOLDER = r"C:\Users\MyName\My Documents\casheets_11-17-25"
     ```

2. **`REPORTS_CASHSHEET_MAP`**:

   - This is the "brain" of the program. It links the location name from the report to the correct cash sheet file and row.
   - **Format**: `"Location Name From Report": ("Cash Sheet File Base Name", "Location Name in Sheet")`
   - Example:
     ```python
     "Union Shake Smart 1-1": ('UnionShakeSmart', "U.S.S.")
     ```
   - This tells the program: "When you see a report from `Union Shake Smart 1-1`, find a file in the `casheet` folder named `UnionShakeSmart`, and then find the row inside that file named `U.S.S.`."

3. **Tender Maps (`INFOR_TENDERS_MAP`, `TAVLO_TENDERS_MAP`)**:

   - These maps translate the tender names from the report (e.g., 'AT Flex Dollars') to the cash sheet column names (e.g., 'flex').
   - Edit these only if your reports or cash sheets change.

4. **`FILL_COL_MAP`**:
   - This map tells the autofiller which column number to use for each data point.
   - Example:
     ```python
     "total_sales": 3
     ```
   - This means "put total sales in column C."

## 🚀 How to Use (Daily Workflow)

Follow these steps every day to run the autofiller:

1. **Clear Old Reports**:

   - Navigate to your `reports` folder and **delete all files** from the previous day.

2. **Update Cash Sheet Folder (First Day of Week Only)**:

   - If this is the start of a new week, create your new weekly cash sheet folder (e.g., `casheet 11-17-25`).
   - Open `config.py` and paste the new folder's path into the `CASHEET_FOLDER` variable.

3. **Download New Reports**:

   - Download all new reports into your empty `reports` folder. **Do not rename the files**.
   - **Infor (.csv)**: Use "Direct Export" and select these four sections:
     - `Summary`
     - `Tax Collected`
     - `Tenders`
     - `Service Performance`
   - **Tavlo (.xls)**: Download the file normally.

4. **Run the Program**:

   - Double-click the `run_autofiller.bat` file. A terminal window will open.

5. **Enter the Date**:

   - The program will ask: `Which cash sheet day are you filling in?`
   - Type the date in **MM/DD/YYYY** format (e.g., `10/22/2025`) and press Enter.
   - _(Tip: If you just press Enter, it will automatically run for yesterday.)_

6. **Verify**:
   - The script will print its progress in the terminal.
   - When it's finished, the `casheet` folder will **open automatically**.
   - Open the Excel files to check that the data was filled in correctly.

## 📁 File Structure

```
/Cash-sheet-autofill/
├── run_autofiller.bat    # <== Click this to run
├── requirements.txt      # Python libraries
├── .gitignore
├── LICENSE
├── README.md             # This file
│
├── reports/              # (Empty) Download daily reports here
└── casheet/              # (Empty) Put your weekly casheet files here
|
└── src/                  # <== All the Python code is in here
    ├── main.py           # Main script, contains all logic
    ├── config.py         # **IMPORTANT: All paths and maps are here**
    ├── infor_parser.py   # Logic for Infor (.csv) files
    ├── tavlo_parser.py   # Logic for Tavlo (.xls/xml) files
    └── excel_autofiller.py # Logic for writing to Excel
```

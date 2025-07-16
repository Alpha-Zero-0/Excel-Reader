# Excel Reader Tool

A simple web-based Excel reader tool that allows users to review and edit Excel files row by row, specifically focusing on an "approved" column.

## Features

- 📊 Upload Excel files (.xlsx, .xls)
- 📋 Display data row by row with column headers
- ✅ Edit "approved" column (yes/no) for each row
- 🎯 Navigate between rows (first, previous, next, last, jump to specific row)
- 📈 Real-time progress tracking and statistics
- 💾 Download updated Excel file
- 🔄 Support for multiple sheets

## Installation

1. Make sure you have Python installed on your system
2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Method 1: Using the startup script (Recommended)
1. **On macOS/Linux:** Double-click `start_app.sh` or run it in terminal
2. **On Windows:** Double-click `start_app.bat`
3. The application will start and show you the URL to open in your browser

### Method 2: Manual startup
1. Open terminal/command prompt
2. Navigate to the Excel Reader folder
3. Run: `streamlit run excel_reader.py`
4. Open your web browser and go to the URL shown (usually `http://localhost:8501`)

### Using the application:

### Using the application:

1. Upload your Excel file using the file uploader

2. If your file has multiple sheets, select the one you want to work with

3. Navigate through rows using the navigation buttons

4. Review each row's data and update the "approved" status as needed

5. Download the updated Excel file when you're done

## Notes

- If your Excel file doesn't have an "approved" column, one will be automatically added with default "no" values
- The tool tracks your changes and shows summary statistics
- You can preview all data at once or work row by row
- The original file is not modified - you download a new updated version

## Requirements

- Python 3.7+
- streamlit
- pandas
- openpyxl
- xlsxwriter
- xlrd

## Sharing with Others

To share this tool with people who don't have VS Code:

1. **Option 1 - Local Installation:**
   - Share the folder containing these files
   - Provide installation instructions above

2. **Option 2 - Streamlit Cloud (Free):**
   - Upload this code to GitHub
   - Deploy for free on [Streamlit Cloud](https://streamlit.io/cloud)
   - Share the public URL

3. **Option 3 - Other Cloud Services:**
   - Deploy to Heroku, Railway, or similar platforms
   - Share the public URL

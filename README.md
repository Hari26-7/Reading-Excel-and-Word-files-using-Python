# Excel, Word and Python File Processing

## Project overview
A beginner-friendly Python project that reads an Excel registration workbook, cleans and organizes the data, saves a cleaned Excel file, and creates a Word report automatically.

## Files
- `registration_data.xlsx` — fictional sample registration data with spacing, capitalization, a duplicate, and a missing email.
- `process_files.py` — automation script.
- `cleaned_registration_data.xlsx` — generated cleaned workbook.
- `registration_summary.docx` — generated Word report.
- `output_screenshot.png` — screenshot-style image showing the successful terminal output.
- `requirements.txt` — required Python packages.

## Run the project
1. Install Python 3.10 or later.
2. Open a terminal in this folder.
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Run:
   ```bash
   python process_files.py
   ```

The script creates or updates `cleaned_registration_data.xlsx` and `registration_summary.docx` in the same folder.

## Cleaning steps
- Trims extra spaces.
- Converts email addresses to lowercase.
- Standardizes department names.
- Removes duplicate registrations based on email.
- Sorts by department and student name.
- Flags missing email values for review.

**Note:** The registration entries are fictional sample data.

# Portfolio Updater from Excel

Automate your portfolio website generation by maintaining your data in a structured Excel file.

## Project Structure

- `update_resume_website.py`: Main script to synchronize the portfolio website with Excel data.
- `extract_excel_data.py`: Core module for extracting and parsing data from the Excel file.
- `inspect_excel.py`: Utility to list Excel sheets and inspect raw row data.
- `verify_extraction.py`: Utility to verify the data extraction logic and see parsed output.
- `portfolio-sakib/`: Directory containing the portfolio website HTML and CSS files.

## How to use

1. **Install Requirements**

    ```bash
    pip install openpyxl beautifulsoup4
    ```

2. **Prepare your Excel**
    Ensure your data is in `Résumé__Abu.Musa.Sakib.xlsx`.

3. **Update Website**

    ```bash
    python update_resume_website.py
    ```

4. **Inspect & Verify (Optional)**

    ```bash
    # Generate sheets.txt and excel_inspect.txt
    python inspect_excel.py

    # Generate verify_output.txt
    python verify_extraction.py
    ```

## Key Features

- **Full Site Sync**: Automatically updates all 5 pages of the portfolio website.
- **Dynamic Projects**: Populates the "Projects" page from the `07-project-roles` sheet.
- **Thematic Skill Cards**: Maps skill categories to relevant images automatically.
- **Common Elements**: Synchronizes navigation and footer (including social links and copyright) across all pages.
- **Smart Formatting**: Handles complex Excel structures and formats dates correctly.

To keep your portfolio website in sync with your Excel data:

```bash
python update_resume_website.py
```

This script:

- **Inspects**: Performs a row-level check of your Excel structure.
- **Verifies**: Validates the extracted data in the console.
- **Syncs**: Automatically updates the portfolio website with your latest info, experience, and skills.

## Deployment

Simply run the synchronization script whenever you update your Excel file, and your website will be ready to go!

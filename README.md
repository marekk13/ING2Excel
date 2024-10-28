# ING2Excel
Automated processing and insertion of transactions from ING Bank Śląski bank statements listed in CSV format to an Excel file.

## Table of Contents
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Description](#description)

## Prerequisites
1. **Python Libraries**:
   - `pandas`: For processing CSV data.
   - `openpyxl`: For working with Excel files.
  
2. **CSV Statement from ING**:
   - Download the bank statement directly from the ING bank website, keeping the original filename, `Lista_transakcji_nrXXXX.csv` (where `XXXX` is the statement number). Select 'Pobierz historię do pliku' in the 'Historia' section and then specify CSV file format.



    <img width="350" alt="image" src="https://github.com/user-attachments/assets/b0abea1f-278c-4144-98fc-e81b6695e9da">

    <img width="250" alt="image" src="https://github.com/user-attachments/assets/d9e5d033-56bd-4a2c-9140-0f5bf36e58f1">
   
    - Place the file in the directory specified by `folder_path` in the script.

2. **Biedronka project**
   - Some methods from my previous (project)[https://github.com/marekk13/biedronka] were used. The needed file is included in the repository.

## Installation
1. Install the required libraries:
   ```bash
   pip install -r requirements.txt
   ```

2. Clone this repository:
   ```bash
   git clone https://github.com/marekk13/ING2Excel.git
   cd ING2Excel 
   ```

3. Place the CSV transaction statement file in the specified directory (`folder_path`).

## Description
This project automates the process of importing bank transaction statements downloaded from ING and exporting cleaned, categorized data to an organized Excel sheet. The script selects the latest statement, processes it, updates transaction titles for clarity, and categorizes expenses, providing insights into monthly expenditures.

### Key Features
1. **Automatic Latest File Selection**:
   - Automatically selects the latest CSV file that begins with `Lista_transakcji_nr` and ends with `.csv` for each run, making it easy to add new data without manual file selection.

2. **Data Cleaning and Processing**:
   - Filters out transactions in non-PLN currencies.
   - Fills in missing values and distinguishes between expenses and income based on the `Kwota transakcji (waluta rachunku)` column.
   - Standardizes transaction dates to `YYYY-MM-DD`.

3. **Customizable Transaction Titles for BLIK Transfers**:
   - The `sub_blik_payment_titles` function customizes default BLIK transaction titles by using the user-defined transfer title, providing more detailed information in Excel.
   
4. **Title Mapping for Card Payments**:
   - The `sub_card_payment_titles` function modifies titles for card transactions based on specific keywords, such as `WWW.BILET.INTERCITY.PL` or `OLX_`, allowing custom names to appear in the Excel sheet.
   - **Note**: You can expand the `mapping` variable to include more patterns. Adding custom entries to `mapping` enhances record readability in Excel, allowing more transactions to be automatically categorized by specific names.
5. **Expense Categorization**:
   - The script automatically assigns categories like `transport`, `spożywcze` (groceries), `media`, and more based on transaction titles, making it easier to analyze spending trends in Excel.
   - **Note**: You can add entries to the `category_mapping` variable in `payment_category` to ensure that additional transaction titles are categorized automatically, reducing the need for manual classification in Excel.
6. **Define Custom Categories** 
   - The `ExcelDataInserter` class allows you to define custom categories directly in its constructor, making it flexible to tailor the output for specific types of expenses.

7. **Monthly Layout Export to Excel**:
   - Automatically generates monthly sheets for transactions based on their dates.
   - Organizes income and expenses separately, with sum calculations for each month.
   - Adds categorized expenses, monthly totals, and balance calculations.

### Data Ingestion and Processing Workflow
1. **CSV Data Loading**: 
   - The script automatically selects the most recent CSV file, imports, and processes it, removing any unnecessary columns.

2. **Data Cleaning**:
   - Non-PLN transactions are excluded, data is standardized, and any missing values are handled appropriately.

3. **Transaction Title Customization**:
   - `sub_blik_payment_titles`: Replaces generic BLIK transaction titles with user-defined titles, adding more descriptive information in the final export.
   - `sub_card_payment_titles`: Adjusts card transaction titles based on preset mappings, enhancing readability.

4. **Excel Export with Monthly Structure**:
   - Processed data is grouped into monthly sheets, with calculations for income and expenses, including summaries and category-based totals for expenses.
  
### Excel Sheet Columns
#### Expenses detailed
  - `Date`: Transaction date.
  - `Description`: Custom transaction titles, including user-defined BLIK titles and adjusted card payments.
  - `Amount`: Transaction amount, with currency formatting.
  - `Category`: Expense category.

#### Incomes (Wpływy)
 - Specific income amount
 - Specific income name

#### Expenses grouped by category (Wydatki wg kategorii)
  - Amount spent by category
  - Category name

#### Monthly balance (bilans)
  - Spendings subtracted from incomes in financial format (PLN).
![Screenshot 2024-10-27 171141](https://github.com/user-attachments/assets/5d512b0d-fa5e-4393-a124-f0e9762d44f7)


This project is designed specifically for ING account users and requires the input data to be in the specified CSV format. Before running, ensure that the input file is correctly placed and named according to the requirements. 

# Expenses_Scripts

A simple Python expense tracking application that stores expenses in JSON format.

## Features

- ✅ **Smart AI categorization** - Automatically categorizes 100+ merchants
- ✅ **Custom rules** - QT/QuikTrip/Caseys: <$50 = misc, ≥$50 = transportation  
- ✅ **Auto CSV export** - Creates clean spreadsheet-ready files
- ✅ **Auto download** - Files saved directly to your computer
- ✅ **Bank format support** - Handles complex bank transaction descriptions
- ✅ **BOM handling** - Automatically processes files with UTF-8 BOM characters
- ✅ **Enhanced error handling** - Robust processing of malformed CSV data
- ✅ **9 categories** - Food, Transportation, Shopping, Entertainment, Healthcare, Financial, Transfer, Misc, Other

## Prerequisites

- Python 3.7+ (uses standard library only)
- Bank CSV export file with transaction data

## Quick Setup & Run

1. **Upload your bank CSV file to VS Code**
2. **Run the command:**

   ```bash
   python expenses.py your_bank_file.csv
   ```

3. **Check your Downloads folder** for the categorized CSV file

**That's it!** Your expenses are now categorized and ready for analysis.

## Setup

1. **Clone or navigate to the repository:**

   ```bash
   cd /workspaces/Expenses_Scripts
   ```

2. **Install dependencies (optional - no external deps required):**

   ```bash
   pip install -r requirements.txt
   ```

## Quick Start

**Run with your bank CSV file:**

```bash
python expenses.py your_bank_export.csv
```

**Example:**

```bash
python expenses.py October_Spending.csv
```

The script will automatically:

- ✅ Categorize all transactions using AI-powered rules
- ✅ Export a clean CSV file ready for Excel/Google Sheets  
- ✅ Download the file directly to your computer's Downloads folder

## Usage

### Option 1: Load expenses from your CSV file

```bash
python expenses.py your_expenses.csv
```

### Option 2: Run with test data (no CSV file)

```bash
python expenses.py
```

### CSV File Format

The script automatically detects and handles various CSV formats, including complex bank statements:

**Supported column names:**

- **Amount**: `amount`, `cost`, `price`, `total`, `value`, `amount ($)`
- **Category**: `category`, `class`, `group`
- **Description**: `action`, `description`, `desc`, `note`, `memo`, `details`
- **Date**: `date`, `timestamp`, `created`, `when`, `run date`

**Bank Statement Support:**

- Automatically handles UTF-8 BOM characters
- Skips empty rows and malformed headers  
- Prioritizes transaction details from "Action" fields
- Filters out deposits and income transactions
- Converts negative amounts to positive expenses

**Example simple CSV format:**

```csv
date,amount,category,description
2025-10-20,15.75,Food,Lunch at restaurant
2025-10-21,42.00,Transportation,Uber ride
2025-10-22,125.00,Shopping,Groceries
```

**Example bank statement format:**

```csv
Run Date,Action,Symbol,Description,Type,Amount ($),Cash Balance ($)
10/28/2025,"DEBIT CARD PURCHASE WALMART.COM",,"No Description",Cash,-109.56,Processing
10/28/2025,"DIRECT DEBIT QUIKTRIP CORP",,"No Description",Cash,-69.12,Processing
```

### What it does

- Loads expenses from your CSV file (if provided)
- Automatically cleans and processes bank statement formats
- Converts and stores them in JSON format
- Shows summary with totals by category
- Handles different CSV formats automatically
- Exports categorized results to Excel with charts
- Falls back to test data if no CSV provided

### Sample output with CSV

```text
Loading expenses from CSV file: bank_statement.csv
Detected CSV columns: ['Run Date', 'Action', 'Symbol', 'Description', 'Type', 'Amount ($)', 'Cash Balance ($)']
Successfully loaded 220 expenses from bank_statement.csv

=== EXPENSE SUMMARY ===
Total expenses: 220
Total amount: $10,808.21

By category:
  Other: $5,387.12 (49.8%)
  Groceries: $1,770.59 (16.4%)
  Shopping: $756.34 (7.0%)
  Food: $738.27 (6.8%)
  Transportation: $722.53 (6.7%)
  Entertainment: $454.14 (4.2%)
  Misc: $396.80 (3.7%)
  Healthcare: $340.00 (3.1%)
  Financial: $242.42 (2.2%)

✅ Exported 220 categorized expenses to: bank_statement_categorized.xlsx
📊 Created 11 category tabs plus summary and chart tabs
📥 Downloading to your computer's Downloads folder...
```

## Files

- `expenses.py` - Main expense tracking script with CSV support
- `expenses.json` - Generated data file (created after first run)
- `requirements.txt` - Python dependencies (empty - uses standard library only)

## CSV File Transfer

To use your CSV file from your computer:

1. **Upload to workspace**: Drag and drop your CSV file into the VS Code file explorer
2. **Or copy via terminal**:

   ```bash
   # If you have the file path on your system
   cp /path/to/your/expenses.csv /workspaces/Expenses_Scripts/
   ```

3. **Then run the script**:

   ```bash
   python expenses.py your_expenses.csv
   ```

## Development

To extend the script, you can:

- Add methods for filtering/reporting expenses
- Implement expense editing/deletion  
- Add data validation
- Create a CLI interface with argparse
- Modify categorization rules in `_categorize_merchant()` method
- Add new vendor patterns to the vendor database

## Common Commands

**Process bank CSV and download results:**

```bash
python expenses.py October_Spending.csv
```

**Run with test data (no CSV file):**

```bash
python expenses.py
```

## GitHub Codespace Management

**Stop the current codespace:**

```bash
gh codespace stop
```

**List all your codespaces:**

```bash
gh codespace list
```

**Stop a specific codespace:**

```bash
gh codespace stop -c bug-free-winner-rvjjjppjrx4356r4
```

**Alternative ways to stop:**

- From VS Code: `Ctrl+Shift+P` → "Codespaces: Stop Current Codespace"
- From GitHub web: Visit [github.com/codespaces](https://github.com/codespaces) → Find your codespace → "..." menu → "Stop codespace"

# Credit Card Categorizer

A Python tool for categorizing and analyzing credit card transactions with Excel export and preservation of manual edits.

## Features

- ✅ **Multi-card support** – Handles multiple credit card CSVs at once
- ✅ **Smart vendor categorization** – 5 categories: Work, Personal, Budgeted, Birthday/Christmas, Horse
- ✅ **Purchaser detection** – Identifies who made each purchase by card/account
- ✅ **Excel export** – Creates a workbook with All Transactions, per-category tabs, and summary
- ✅ **Preservation mode** – Re-running on a categorized Excel file preserves all manual changes, deletions, and notes
- ✅ **Payment filtering** – Excludes payment/transfer transactions automatically
- ✅ **Notes column** – Add manual notes to any transaction and keep them across runs
- ✅ **Auto download** – Files saved directly to your computer

## Prerequisites

- Python 3.7+
- openpyxl (install with `pip install openpyxl`)
- Credit card CSV export files (Chase, Amex, Discover, Citi, Capital One, etc.)

## Quick Setup & Run

1. **Upload your credit card CSV files to VS Code**
2. **Run the command:**

   ```bash
   python credit_card_categorizer.py
   ```

   Or to re-process an existing categorized Excel file (preserves manual changes):

   ```bash
   python credit_card_categorizer.py credit_card_categorized_YYYY-MM-DD.xlsx
   ```

3. **Check your Downloads folder** for the categorized Excel file

**That's it!** Your credit card transactions are now categorized and ready for review.

## Setup

1. **Clone or navigate to the repository:**

   ```bash
   cd /workspaces/Expenses_Scripts
   ```

2. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Option 1: Process all CSV files in the directory

```bash
python credit_card_categorizer.py
```

### Option 2: Re-process an existing categorized Excel file

```bash
python credit_card_categorizer.py credit_card_categorized_YYYY-MM-DD.xlsx
```

- This preserves all manual changes, deletions, and notes in the Excel file.

### Output

- **All Transactions** tab: All transactions, editable, with category and notes columns
- **Category tabs**: One tab for each category (Work, Personal, Budgeted, Birthday/Christmas, Horse)
- **Category Summary**: Purchaser breakdowns and totals

### Excel File Naming

- First run: `credit_card_categorized_YYYY-MM-DD.xlsx`
- Subsequent runs: `credit_card_final_YYYY-MM-DD.xlsx`

## File Structure

- `credit_card_categorizer.py` – Main script
- `credit_card_categorized_YYYY-MM-DD.xlsx` – Output Excel file
- `requirements.txt` – Python dependencies

## Development

To extend the script, you can:

- Add new vendor patterns or categories
- Adjust purchaser detection logic
- Improve Excel formatting or add charts
- Add new summary or reporting features

## Common Commands

**Process all credit card CSVs and download results:**

```bash
python credit_card_categorizer.py
```

**Re-process a categorized Excel file (preserve manual changes):**

```bash
python credit_card_categorizer.py credit_card_categorized_YYYY-MM-DD.xlsx
```

**Run with test data (no CSV file):**

```bash
python credit_card_categorizer.py
```

## Tips

- Edit the All Transactions tab in Excel to change categories or add notes, then re-run the script to preserve your changes.
- Delete rows in the All Transactions tab to remove transactions from future runs.
- The script will always preserve your manual edits when re-processing a categorized Excel file.

---

**Questions?** Open an issue or check the script for more details.

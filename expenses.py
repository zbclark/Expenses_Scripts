import json  
import datetime  
import csv
import sys
import re
import os
import subprocess
from collections import defaultdict

# Try to import openpyxl for Excel functionality
try:
    from openpyxl import Workbook
    from openpyxl.chart import PieChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_AVAILABLE = True
except ImportError:
    EXCEL_AVAILABLE = False
    print("⚠️  Note: openpyxl not available. Install with 'pip install openpyxl' for Excel features.")  
  
class ExpenseTracker:  
    def __init__(self, filename="expenses.json"):  
        self.filename = filename  
        self.expenses = self.load_expenses()  
      
    def load_expenses(self):  
        """Load expenses from file, create empty list if file doesn't exist"""  
        try:  
            with open(self.filename, 'r') as file:  
                return json.load(file)  
        except FileNotFoundError:  
            return []  
      
    def save_expenses(self):  
        """Save expenses to file"""  
        with open(self.filename, 'w') as file:  
            json.dump(self.expenses, file, indent=2)  
      
    def add_expense(self, amount, category, description=""):  
        """Add a new expense"""  
        expense = {  
            "date": datetime.date.today().isoformat(),  
            "amount": float(amount),  
            "category": category.lower(),  
            "description": description  
        }  
        self.expenses.append(expense)  
        self.save_expenses()  
        print(f"Added expense: ${amount} for {category}")  

    def load_from_csv(self, csv_filename):
        """Load expenses from a CSV file with enhanced bank statement support"""
        # Clear existing expenses to avoid duplicates
        self.expenses = []
        loaded_count = 0
        try:
            # First, clean the CSV file to handle BOM and empty lines
            cleaned_filename = self._clean_csv_file(csv_filename)
            
            with open(cleaned_filename, 'r', newline='', encoding='utf-8') as csvfile:
                # Try to detect the delimiter
                sample = csvfile.read(1024)
                csvfile.seek(0)
                sniffer = csv.Sniffer()
                delimiter = sniffer.sniff(sample).delimiter
                
                reader = csv.DictReader(csvfile, delimiter=delimiter)
                
                # Print detected columns for debugging
                print(f"Detected CSV columns: {reader.fieldnames}")
                
                for row_num, row in enumerate(reader, 1):
                    try:
                        # Try different common column name variations
                        amount = self._get_csv_value(row, ['amount', 'cost', 'price', 'total', 'value', 'amount ($)'])
                        category = self._get_csv_value(row, ['category', 'class', 'group'])  # Remove 'type' for bank statements
                        description = self._get_csv_value(row, ['action', 'description', 'desc', 'note', 'memo', 'details'])  # Prioritize 'action' for bank statements
                        date = self._get_csv_value(row, ['date', 'timestamp', 'created', 'when', 'run date'])
                        
                        if amount is not None:
                            # Clean up amount (remove $ and commas, handle negatives)
                            amount_str = str(amount).replace('$', '').replace(',', '')
                            try:
                                amount_float = float(amount_str)
                                # Skip deposits (positive amounts in bank statements)
                                if amount_float > 0:
                                    continue
                                amount_float = abs(amount_float)  # Make expenses positive
                            except ValueError:
                                print(f"Skipping row {row_num} with invalid amount: {amount}")
                                continue
                            
                            # Always parse bank transaction for bank statements
                            if description:
                                parsed_category, merchant = self._parse_bank_transaction(description, amount_float)
                                category = parsed_category
                                # Use merchant name as description if it's cleaner
                                if len(merchant) < len(description):
                                    description = merchant
                            
                            # Skip income/transfer transactions
                            if category and category.lower() == 'income':
                                continue
                            
                            # Format date properly
                            formatted_date = self._format_date(date) if date else datetime.date.today().isoformat()
                            
                            expense = {
                                "date": formatted_date,
                                "amount": amount_float,
                                "category": (category or 'other').lower(),
                                "description": description or ''
                            }
                            self.expenses.append(expense)
                            loaded_count += 1
                            
                    except Exception as e:
                        print(f"Error processing row {row_num}: {e}")
                        continue
                        
            self.save_expenses()
            print(f"Successfully loaded {loaded_count} expenses from {csv_filename}")
            return loaded_count
            
        except FileNotFoundError:
            print(f"Error: CSV file '{csv_filename}' not found")
            return 0
        except Exception as e:
            print(f"Error reading CSV file: {e}")
            return 0

    def load_from_excel(self, excel_filename):
        """Load expenses from an Excel file"""
        if not EXCEL_AVAILABLE:
            print("❌ Excel support requires openpyxl. Installing now...")
            try:
                subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl"], check=True)
                print("✅ openpyxl installed successfully! Please run the script again.")
                return 0
            except subprocess.CalledProcessError:
                print("❌ Failed to install openpyxl. Please install manually: pip install openpyxl")
                return 0
        
        # Clear existing expenses to avoid duplicates
        self.expenses = []
        loaded_count = 0
        
        try:
            from openpyxl import load_workbook
            
            # Load the workbook
            wb = load_workbook(excel_filename, read_only=True)
            
            # Look for the "All Transactions" sheet first, or use the first available sheet
            sheet = None
            if "All Transactions" in wb.sheetnames:
                sheet = wb["All Transactions"]
                print(f"Using 'All Transactions' sheet")
            else:
                # Find a sheet that looks like transaction data (has Date, Amount, Description columns)
                for sheet_name in wb.sheetnames:
                    test_sheet = wb[sheet_name]
                    headers = []
                    for cell in test_sheet[1]:
                        if cell.value:
                            headers.append(str(cell.value).lower().strip())
                        else:
                            headers.append('')
                    
                    # Check if this looks like transaction data
                    has_date = any('date' in h for h in headers)
                    has_amount = any(word in ' '.join(headers) for word in ['amount', 'transaction', 'debit', 'credit'])
                    
                    if has_date and has_amount:
                        sheet = test_sheet
                        print(f"Using sheet: '{sheet_name}' (detected transaction data)")
                        break
                
                # If no suitable sheet found, use the first one
                if sheet is None:
                    sheet = wb.worksheets[0]
                    print(f"Using first available sheet: '{sheet.title}'")
            
            # Get headers from first row
            headers = []
            for cell in sheet[1]:
                if cell.value:
                    headers.append(str(cell.value).lower().strip())
                else:
                    headers.append('')
            
            print(f"Detected Excel columns: {headers}")
            
            # Process each row (skip header)
            for row_num, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
                if not any(row):  # Skip empty rows
                    continue
                    
                # Create a row dictionary
                row_dict = {}
                for i, value in enumerate(row):
                    if i < len(headers) and value is not None:
                        row_dict[headers[i]] = str(value) if value else ''
                
                # Skip rows that don't look like transactions
                if not row_dict:
                    continue
                
                # Parse the transaction - handle both raw bank data and already-categorized data
                try:
                    # Check if this is already categorized data (has category column)
                    if 'category' in row_dict and row_dict['category']:
                        # This is already processed data from our own export
                        amount_str = str(row_dict.get('amount', '0')).replace('$', '').replace(',', '')
                        try:
                            amount_float = float(amount_str)
                        except ValueError:
                            continue
                            
                        expense = {
                            "date": row_dict.get('date', ''),
                            "amount": amount_float,
                            "category": row_dict.get('category', 'other').lower(),
                            "description": row_dict.get('description', '')
                        }
                    else:
                        # This is raw bank data, needs parsing
                        amount = row_dict.get('amount ($)', '') or row_dict.get('amount', '')
                        description = row_dict.get('action', '') or row_dict.get('description', '')
                        date = row_dict.get('run date', '') or row_dict.get('date', '')
                        
                        if not amount or not description:
                            continue
                            
                        # Clean up amount
                        amount_str = str(amount).replace('$', '').replace(',', '')
                        try:
                            amount_float = float(amount_str)
                            if amount_float > 0:  # Skip deposits
                                continue
                            amount_float = abs(amount_float)
                        except ValueError:
                            continue
                        
                        # Parse bank transaction
                        parsed_category, merchant = self._parse_bank_transaction(description, amount_float)
                        if parsed_category.lower() == 'income':  # Skip transfers
                            continue
                            
                        expense = {
                            "date": date if date else datetime.date.today().isoformat(),
                            "amount": amount_float,
                            "category": parsed_category.lower(),
                            "description": merchant or description
                        }
                    
                    if expense:
                        self.expenses.append(expense)
                        loaded_count += 1
                except Exception as e:
                    print(f"Warning: Could not parse row {row_num}: {e}")
                    continue
            
            wb.close()
            self.save_expenses()
            print(f"Successfully loaded {loaded_count} expenses from {excel_filename}")
            return loaded_count
            
        except FileNotFoundError:
            print(f"Error: Excel file '{excel_filename}' not found")
            return 0
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return 0
    
    def _get_csv_value(self, row, possible_keys):
        """Helper to find value from CSV row using different possible column names"""
        for key in possible_keys:
            # Try exact match (case-insensitive)
            for row_key in row.keys():
                if row_key and row_key.lower().strip() == key.lower():
                    value = row[row_key]
                    if value is not None:
                        value = str(value).strip()
                        return value if value else None
                # Also try partial matches for compound column names
                if row_key and key.lower() in row_key.lower():
                    value = row[row_key]
                    if value is not None:
                        value = str(value).strip()
                        return value if value else None
        return None

    def _clean_csv_file(self, csv_filename):
        """Clean CSV file to handle BOM characters and empty lines"""
        import tempfile
        import os
        
        # Create a temporary cleaned file
        temp_fd, temp_filename = tempfile.mkstemp(suffix='.csv', text=True)
        
        try:
            with open(csv_filename, 'r', encoding='utf-8-sig') as input_file:
                with os.fdopen(temp_fd, 'w', encoding='utf-8') as output_file:
                    lines = input_file.readlines()
                    
                    # Find and write the header line
                    header_written = False
                    for line in lines:
                        stripped_line = line.strip()
                        if stripped_line and ('Run Date' in stripped_line or 'date' in stripped_line.lower() or 'amount' in stripped_line.lower()):
                            output_file.write(stripped_line + '\n')
                            header_written = True
                            break
                    
                    # Write all data lines after the header
                    if header_written:
                        reading_data = False
                        for line in lines:
                            stripped_line = line.strip()
                            # Skip empty lines
                            if not stripped_line:
                                continue
                            # Start reading data after header
                            if 'Run Date' in stripped_line or 'date' in stripped_line.lower():
                                reading_data = True
                                continue
                            # Write data lines
                            if reading_data:
                                output_file.write(stripped_line + '\n')
                    
            return temp_filename
            
        except Exception as e:
            # If cleaning fails, return original filename
            os.close(temp_fd)
            os.unlink(temp_filename)
            return csv_filename

    def _format_date(self, date_str):
        """Format date string to ISO format (YYYY-MM-DD)"""
        if not date_str:
            return datetime.date.today().isoformat()
        
        try:
            # Handle MM/DD/YYYY format
            if '/' in date_str:
                parts = date_str.split('/')
                if len(parts) == 3:
                    month, day, year = parts
                    return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
            
            # Handle other formats - add more as needed
            return date_str
            
        except Exception:
            return datetime.date.today().isoformat()

    def show_summary(self):
        """Display a summary of all expenses"""
        if not self.expenses:
            print("No expenses to summarize.")
            return
        
        from collections import defaultdict
        
        totals = defaultdict(float)
        for expense in self.expenses:
            totals[expense['category']] += expense['amount']
        
        # Sort by amount
        sorted_totals = sorted(totals.items(), key=lambda x: x[1], reverse=True)
        total_amount = sum(totals.values())
        
        print(f"\n=== EXPENSE SUMMARY ===")
        print(f"Total expenses: {len(self.expenses)}")
        print(f"Total amount: ${total_amount:,.2f}")
        print(f"\nBy category:")
        
        for category, amount in sorted_totals:
            percentage = (amount / total_amount) * 100
            print(f"  {category.title()}: ${amount:,.2f} ({percentage:.1f}%)")

    def export_to_csv(self, output_filename=None):
        """Export expenses to a clean CSV file"""
        if not self.expenses:
            print("No expenses to export.")
            return None
        
        if not output_filename:
            # Generate filename based on current JSON filename
            base_name = self.filename.replace('.json', '')
            output_filename = f"{base_name}_categorized.csv"
        
        try:
            with open(output_filename, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['date', 'amount', 'category', 'description']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                
                writer.writeheader()
                for expense in sorted(self.expenses, key=lambda x: x['date']):
                    writer.writerow({
                        'date': expense['date'],
                        'amount': f"{expense['amount']:.2f}",
                        'category': expense['category'].title(),
                        'description': expense['description']
                    })
            
            print(f"📁 Exported {len(self.expenses)} expenses to: {output_filename}")
            return output_filename
            
        except Exception as e:
            print(f"Error exporting to CSV: {e}")
            return None

    def _parse_bank_transaction(self, action_text, amount=0):
        """Parse bank transaction to extract merchant and category with smart rules"""
        action = action_text.upper()
        
        # Skip deposits/credits and ALL transfers
        if any(word in action for word in ['CHECK RECEIVED', 'DEPOSIT', 'CREDIT', 'TRANSFER FROM', 'TRANSFER TO', 'TRANSFERRED FROM', 'TRANSFERRED TO']):
            return 'income', action_text
        
        # Skip direct debits that are transfers/payments between accounts
        if any(word in action for word in ['DIRECT DEBIT VENMO', 'DIRECT DEBIT PAYPAL', 'DIRECT DEBIT ZELLE']):
            return 'income', action_text
        
        # Extract merchant from common patterns
        if 'DEBIT CARD PURCHASE' in action:
            # Look for merchant name after date pattern
            # Find merchant name between date and location/reference numbers
            match = re.search(r'\d{2}/\d{2} (.+?) \d{12}', action)
            if match:
                merchant = match.group(1).strip()
            else:
                # Fallback: extract text after "OUTSTAND AUTH MM/DD "
                match = re.search(r'OUTSTAND AUTH \d{2}/\d{2} (.+?) \d{12}', action)
                if match:
                    merchant = match.group(1).strip()
                else:
                    merchant = action_text
        elif 'DIRECT DEBIT' in action:
            # Extract company name from direct debit
            match = re.search(r'DIRECT DEBIT (.+?) PREAUTHPMT', action)
            if match:
                merchant = match.group(1).strip()
            else:
                merchant = action_text
        else:
            merchant = action_text
        
        # Get category using enhanced categorization
        category = self._categorize_merchant(merchant, amount)
        
        return category, merchant

    def _categorize_merchant(self, merchant, amount):
        """Enhanced merchant categorization with smart rules and vendor database"""
        merchant_lower = merchant.lower()
        
        # Define gas stations for amount-based categorization
        gas_stations = [
            'quiktrip', 'qt ', 'caseys', 'casey\'s', 'shell', 'exxon', 'mobil', 'chevron', 
            'bp', 'conoco', 'phillips 66', 'texaco', 'valero', 'citgo', '7-eleven', 
            'circle k', 'wawa', 'sheetz', 'loves', 'pilot', 'flying j', 'ta travel', 
            'speedway', 'marathon', 'sunoco', 'gulf', 'cenex'
        ]
        
        # Gas station logic: <$50 = misc, ≥$50 = transportation, BUT "OUTSIDE" = always transportation
        if any(station in merchant_lower for station in gas_stations):
            # Special case: QT/Quiktrip with "OUTSIDE" is always transportation regardless of amount
            if ('qt' in merchant_lower or 'quiktrip' in merchant_lower) and 'outside' in merchant_lower:
                return 'transportation'
            elif amount < 50:
                return 'misc'
            else:
                return 'transportation'
        
        # Sonic is always food
        if 'sonic' in merchant_lower:
            return 'food'
        
        # Walmart is always groceries
        if any(walmart in merchant_lower for walmart in ['walmart', 'wal-mart', 'walmart.com']):
            return 'groceries'
        
        # Enhanced vendor database with common merchants
        vendor_categories = {
            # Fast Food & Restaurants
            'food': [
                'taco bell', 'chick-fil-a', 'chick fil a', 'mcdonalds', 'mcdonald\'s', 'burger king',
                'subway', 'kfc', 'pizza hut', 'dominos', 'papa johns', 'wendy\'s', 'wendys',
                'arbys', 'arby\'s', 'popeyes', 'chipotle', 'panera', 'starbucks', 'dunkin',
                'ihop', 'dennys', 'denny\'s', 'waffle house', 'cracker barrel', 'olive garden',
                'applebees', 'chilis', 'tgi fridays', 'outback', 'red lobster', 'panda express',
                'qdoba', 'five guys', 'in-n-out', 'whataburger', 'culvers', 'culver\'s',
                'jimmy johns', 'jimmy john\'s', 'jersey mikes', 'firehouse subs', 'blaze pizza',
                'cafe', 'restaurant', 'diner', 'grill', 'bistro', 'eatery', 'coffee', 'bakery'
            ],
            # Transportation (non-gas station)
            'transportation': [
                'uber', 'lyft', 'taxi', 'parking', 'toll', 'metro', 'bus', 'train', 'airline',
                'airport', 'rental car', 'hertz', 'enterprise', 'budget', 'avis', 'national'
            ],
            # Groceries & Food Shopping
            'groceries': [
                'kroger', 'safeway', 'publix', 'whole foods', 'trader joes', 'trader joe\'s',
                'aldi', 'food lion', 'giant', 'stop shop', 'harris teeter', 'wegmans',
                'grocery', 'supermarket', 'fresh market', 'food store', 'butcher', 'deli'
            ],
            # Retail & Shopping (excluding groceries/walmart)
            'shopping': [
                'target', 'costco', 'sams club', 'sam\'s club', 'bjs', 'amazon',
                'best buy', 'home depot', 'lowes', 'lowe\'s', 'menards', 'ace hardware',
                'cvs', 'walgreens', 'rite aid', 'dollar general', 'dollar tree', 'family dollar',
                'macys', 'macy\'s', 'nordstrom', 'jcpenney', 'kohl\'s', 'kohls', 'sears',
                'old navy', 'gap', 'banana republic', 'tj maxx', 'marshalls', 'ross',
                'bed bath beyond', 'bath body works', 'victoria secret', 'victoria\'s secret',
                'store', 'shop', 'market', 'general', 'supply'
            ],
            # Healthcare & Medical
            'healthcare': [
                'cvs pharmacy', 'walgreens pharmacy', 'pharmacy', 'hospital', 'clinic',
                'medical', 'doctor', 'dentist', 'dental', 'optometry', 'vision', 'urgent care',
                'pediatric', 'therapy', 'physical therapy', 'chiropractic', 'veterinary', 'vet'
            ],
            # Entertainment & Recreation
            'entertainment': [
                'movie', 'cinema', 'theater', 'theatre', 'netflix', 'spotify', 'apple music',
                'disney', 'hulu', 'amazon prime', 'xbox', 'playstation', 'nintendo', 'steam',
                'game', 'arcade', 'bowling', 'golf', 'gym', 'fitness', 'spa', 'salon',
                'amusement', 'zoo', 'aquarium', 'museum', 'concert', 'ticket', 'event'
            ],
            # Utilities & Services
            'utilities': [
                'electric', 'electricity', 'gas company', 'water', 'sewer', 'trash', 'waste',
                'internet', 'cable', 'phone', 'cellular', 'verizon', 'att', 'at&t', 't-mobile',
                'sprint', 'comcast', 'xfinity', 'spectrum', 'cox', 'dish', 'directv'
            ],
            # Financial & Professional Services
            'financial': [
                'bank', 'credit union', 'atm', 'paypal', 'venmo', 'zelle', 'cashapp', 'cash app',
                'square', 'stripe', 'insurance', 'tax', 'accounting', 'legal', 'lawyer',
                'attorney', 'notary', 'real estate', 'mortgage', 'loan'
            ]
        }
        
        # Check vendor database
        for category, vendors in vendor_categories.items():
            if any(vendor in merchant_lower for vendor in vendors):
                return category
        
        # Pattern-based categorization for specific transaction types
        if any(word in merchant_lower for word in ['atm', 'cash advance', 'withdrawal']):
            return 'cash'
        elif any(word in merchant_lower for word in ['subscription', 'monthly', 'annual']):
            return 'subscription'
        elif 'sq *' in merchant_lower or 'square *' in merchant_lower:
            # Square payments - try to categorize by business name after SQ *
            sq_business = merchant_lower.replace('sq *', '').strip()
            # Recursively categorize the business name
            return self._categorize_merchant(sq_business, amount)
        
        # Default fallback
        return 'other'

    def display_summary(self):
        """Display a summary of all expenses"""
        if not self.expenses:
            print("No expenses recorded yet.")
            return
            
        total = sum(expense['amount'] for expense in self.expenses)
        categories = defaultdict(float)
        
        for expense in self.expenses:
            categories[expense['category']] += expense['amount']
        
        print(f"\n=== EXPENSE SUMMARY ===")
        print(f"Total expenses: {len(self.expenses)}")
        print(f"Total amount: ${total:.2f}")
        print(f"\nBy category:")
        for category, amount in sorted(categories.items()):
            print(f"  {category.title()}: ${amount:.2f}")
        
        print(f"\nAll expenses:")
        # Sort expenses by date ascending for display
        sorted_expenses = self._sort_expenses_by_date_ascending(self.expenses)
        for expense in sorted_expenses:
            print(f"  {expense['date']} | ${expense['amount']:.2f} | {expense['category']} | {expense['description']}")

    def _sort_expenses_by_date_ascending(self, expenses):
        """Sort expenses by date in ascending order (earliest first), then by amount descending"""
        from datetime import datetime
        
        def sort_key(expense):
            try:
                # Parse MM/DD/YYYY format
                parsed_date = datetime.strptime(expense['date'], '%m/%d/%Y')
            except ValueError:
                try:
                    # Try YYYY-MM-DD format as fallback
                    parsed_date = datetime.strptime(expense['date'], '%Y-%m-%d')
                except ValueError:
                    # If parsing fails, use the string date (fallback)
                    parsed_date = expense['date']
            return (parsed_date, -expense['amount'])  # Negative amount for descending order within same date
        
        return sorted(expenses, key=sort_key)

    def export_to_excel(self, output_filename=None, input_filename=None):
        """Export categorized expenses to an Excel file with multiple tabs and charts"""
        if not self.expenses:
            print("No expenses to export.")
            return None
        
        if not EXCEL_AVAILABLE:
            print("❌ Excel export requires openpyxl. Installing now...")
            try:
                subprocess.run([sys.executable, "-m", "pip", "install", "openpyxl"], check=True)
                print("✅ openpyxl installed successfully! Please run the script again.")
                return None
            except subprocess.CalledProcessError:
                print("❌ Failed to install openpyxl. Please install manually: pip install openpyxl")
                return None
        
        if output_filename is None:
            if input_filename:
                # Generate filename based on input filename
                import os
                base_name = os.path.splitext(os.path.basename(input_filename))[0]
                output_filename = f"{base_name}_categorized.xlsx"
            else:
                # Fallback to date range if no input filename provided
                dates = [expense['date'] for expense in self.expenses]
                
                # Parse dates properly for correct min/max calculation
                from datetime import datetime
                parsed_dates = []
                for date_str in dates:
                    try:
                        parsed_date = datetime.strptime(date_str, '%m/%d/%Y')
                        parsed_dates.append(parsed_date)
                    except ValueError:
                        try:
                            parsed_date = datetime.strptime(date_str, '%Y-%m-%d')
                            parsed_dates.append(parsed_date)
                        except ValueError:
                            continue
                
                if parsed_dates:
                    min_date = min(parsed_dates).strftime('%m-%d-%Y')
                    max_date = max(parsed_dates).strftime('%m-%d-%Y')
                else:
                    min_date = min(dates).replace('/', '-')
                    max_date = max(dates).replace('/', '-')
                    
                output_filename = f"categorized_expenses_{min_date}_to_{max_date}.xlsx"
        
        try:
            # Create workbook
            wb = Workbook()
            
            # Remove default sheet
            wb.remove(wb.active)
            
            # Create summary sheet with all transactions
            summary_sheet = wb.create_sheet("All Transactions")
            
            # Add headers with styling
            headers = ['Date', 'Amount', 'Category', 'Description', 'Notes']
            for col, header in enumerate(headers, 1):
                cell = summary_sheet.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True, size=12)
                cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                cell.font = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center")
            
            # Sort and add all transactions by date ascending
            sorted_expenses = self._sort_expenses_by_date_ascending(self.expenses)
            for row, expense in enumerate(sorted_expenses, 2):
                summary_sheet.cell(row=row, column=1, value=expense['date'])
                summary_sheet.cell(row=row, column=2, value=expense['amount'])
                summary_sheet.cell(row=row, column=3, value=expense['category'].title())
                summary_sheet.cell(row=row, column=4, value=expense['description'])
                summary_sheet.cell(row=row, column=5, value='')  # Empty notes column
            
            # Auto-adjust column widths
            for column in summary_sheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                summary_sheet.column_dimensions[column_letter].width = adjusted_width
            
            # Group expenses by category to get existing categories
            categories = defaultdict(list)
            for expense in self.expenses:
                categories[expense['category']].append(expense)
            
            # Add additional empty categories that we want tabs for
            additional_categories = ['horse', 'birthday/christmas']
            for add_cat in additional_categories:
                if add_cat not in categories:
                    categories[add_cat] = []
            
            # Create a sheet for each category with dynamic formulas (including empty ones)
            for category, expenses in categories.items():
                # Sanitize sheet name for Excel (remove invalid characters)
                safe_sheet_name = category.title().replace('/', '-').replace('\\', '-').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
                sheet = wb.create_sheet(safe_sheet_name)
                
                # Add headers with styling for category sheets
                category_headers = ['Date', 'Amount', 'Category', 'Description', 'Notes']
                for col, header in enumerate(category_headers, 1):
                    cell = sheet.cell(row=1, column=col, value=header)
                    cell.font = Font(bold=True, size=12)
                    cell.fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.alignment = Alignment(horizontal="center")
                
                # Add simple formula-based approach for dynamic updates
                # Use basic Excel functions that work across versions
                max_row = len(self.expenses) + 1  # +1 for header row
                
                # Add a simple instruction row for this category
                sheet.cell(row=2, column=1, value=f"=IF(COUNTIF('All Transactions'!$C$2:$C${max_row},\"{category.title()}\")>0,\"Category data below\",\"No {category.title()} transactions\")")
                
                # Add the current matching transactions as static data
                matching_expenses = [expense for expense in self.expenses if expense['category'] == category]
                
                # Add the current matching transactions starting from row 3
                for row_num, expense in enumerate(matching_expenses, 3):
                    sheet.cell(row=row_num, column=1, value=expense['date'])
                    sheet.cell(row=row_num, column=2, value=expense['amount'])
                    sheet.cell(row=row_num, column=3, value=expense['category'].title())
                    sheet.cell(row=row_num, column=4, value=expense['description'])
                    sheet.cell(row=row_num, column=5, value='')  # Empty notes column
                
                # Add instructions for empty categories
                if not matching_expenses:
                    sheet.cell(row=3, column=1, value="To add transactions to this category:")
                    sheet.cell(row=4, column=1, value="1. Go to 'All Transactions' tab")
                    sheet.cell(row=5, column=1, value="2. Change any transaction's Category column")
                    sheet.cell(row=6, column=1, value="3. Re-run the script to update")
                
                # Auto-adjust column widths
                for column in sheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    sheet.column_dimensions[column_letter].width = adjusted_width
            
            # Create chart sheet with dynamic formulas
            chart_sheet = wb.create_sheet("Category Chart")
            
            # Get unique categories for dynamic chart
            unique_categories = list(categories.keys())
            max_row = len(self.expenses) + 1  # +1 for header row
            
            # Add chart data headers
            chart_sheet['A1'] = 'Category'
            chart_sheet['B1'] = 'Total Amount'
            chart_sheet['A1'].font = Font(bold=True)
            chart_sheet['B1'].font = Font(bold=True)
            
            # Add dynamic SUMIF formulas for chart totals (widely compatible)
            chart_row = 2
            for category in sorted(unique_categories):
                # Category name (use sanitized name for display)
                display_name = category.title().replace('/', '-').replace('\\', '-').replace('?', '').replace('*', '').replace('[', '').replace(']', '')
                chart_sheet.cell(row=chart_row, column=1, value=display_name)
                
                # Use SUMIF formula which is widely supported across Excel versions
                sum_formula = f'=SUMIF(\'All Transactions\'!$C:$C,"{category.title()}",\'All Transactions\'!$B:$B)'
                chart_sheet.cell(row=chart_row, column=2, value=sum_formula)
                
                chart_row += 1
            
            # Create pie chart
            pie = PieChart()
            labels = Reference(chart_sheet, min_col=1, min_row=2, max_row=chart_row-1)
            data = Reference(chart_sheet, min_col=2, min_row=1, max_row=chart_row-1)
            pie.add_data(data, titles_from_data=True)
            pie.set_categories(labels)
            pie.title = "Expenses by Category"
            pie.height = 15
            pie.width = 20
            
            # Add data labels
            pie.dataLabels = DataLabelList()
            pie.dataLabels.showPercent = True
            pie.dataLabels.showVal = True
            
            # Add chart to sheet
            chart_sheet.add_chart(pie, "D3")
            
            # Auto-adjust column widths for chart sheet
            chart_sheet.column_dimensions['A'].width = 20
            chart_sheet.column_dimensions['B'].width = 15
            
            # Save the workbook
            wb.save(output_filename)
            
            print(f"\n✅ Exported {len(self.expenses)} categorized expenses to: {output_filename}")
            print(f"📊 Created {len(categories)} category tabs plus summary and chart tabs")
            
            # Show summary
            print(f"\nExported categories:")
            for category in sorted(categories.keys()):
                count = len(categories[category])
                total = sum(expense['amount'] for expense in categories[category])
                print(f"  {category.title()}: {count} transactions, ${total:.2f}")
            
            # Auto-download the file
            self._download_to_computer(output_filename)
            
            return output_filename
            
        except Exception as e:
            print(f"Error exporting to Excel: {e}")
            return None

    def _download_to_computer(self, filename):
        """Download the CSV file to the user's local computer"""
        try:
            # Get the absolute path to the file
            file_path = os.path.abspath(filename)
            
            # Check if file exists
            if not os.path.exists(file_path):
                print(f"❌ File not found: {file_path}")
                return
            
            print(f"\n📥 Downloading {filename} to your computer...")
            
            # Use VS Code's built-in download functionality
            # This works in development containers and codespaces
            download_command = f'code --download "{file_path}"'
            
            try:
                # Try VS Code download first
                result = subprocess.run(download_command, shell=True, capture_output=True, text=True, timeout=10)
                if result.returncode == 0:
                    print(f"✅ Successfully downloaded {filename} to your Downloads folder!")
                    print(f"💡 Check your browser's Downloads folder for the file.")
                    return
            except (subprocess.TimeoutExpired, subprocess.CalledProcessError, FileNotFoundError):
                pass
            
            # Fallback: Provide instructions for manual download
            print(f"📋 Manual download instructions:")
            print(f"   1. In VS Code, look at the left sidebar (File Explorer)")
            print(f"   2. Right-click on '{filename}'")
            print(f"   3. Select 'Download' from the context menu")
            print(f"   4. The file will be saved to your Downloads folder")
            print(f"   5. Alternative: Click the file to open it, then Ctrl+S to save")
            
            # Also try using the browser download method if in a web environment
            if os.environ.get('CODESPACE_NAME') or os.environ.get('GITPOD_WORKSPACE_ID'):
                print(f"\n🌐 Alternative: Open the file in browser:")
                print(f"   File location: /workspaces/Expenses_Scripts/{filename}")
                print(f"   Or use: File menu → Download → navigate to {filename}")
                
        except Exception as e:
            print(f"❌ Error during download: {e}")
            print(f"📁 File saved in workspace at: /workspaces/Expenses_Scripts/{filename}")
            print(f"💡 You can manually download it using VS Code's file explorer.")
  
# Main execution
if __name__ == "__main__":  
    tracker = ExpenseTracker()
    
    # Check if file provided as command line argument
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
        
        # Detect file type based on extension
        file_extension = input_file.lower().split('.')[-1]
        
        if file_extension in ['xlsx', 'xls']:
            print(f"Loading expenses from Excel file: {input_file}")
            loaded = tracker.load_from_excel(input_file)
        elif file_extension == 'csv':
            print(f"Loading expenses from CSV file: {input_file}")
            loaded = tracker.load_from_csv(input_file)
        else:
            print(f"❌ Unsupported file type: .{file_extension}")
            print("📝 Supported file types: .csv, .xlsx, .xls")
            loaded = 0
        
        if loaded > 0:
            tracker.display_summary()
            
            # Automatically export to Excel with multiple tabs and charts
            exported_file = tracker.export_to_excel(input_filename=input_file)
            if exported_file:
                print(f"\n🎉 Ready to use! Your categorized expenses are now in: {exported_file}")
                print(f"📊 Multiple tabs available: All Transactions, individual category tabs, and Category Chart")
                print(f"💡 You can open this file in Excel, Google Sheets, or any spreadsheet program.")
        else:
            print("No expenses were loaded from the input file.")
    else:
        # Original test behavior - add sample expenses
        print("No input file provided. Running with test data...")
        print("Usage: python expenses.py <filename>")
        print("Example: python expenses.py my_expenses.csv")
        print("Example: python expenses.py my_expenses.xlsx")
        print("📝 Supported file types: .csv, .xlsx, .xls\n")
        
        # Add a few test expenses  
        tracker.add_expense(12.50, "Food", "Lunch at cafe")  
        tracker.add_expense(45.00, "Transportation", "Gas")  
        tracker.add_expense(8.99, "Food", "Coffee")  
          
        # Show what we have so far  
        tracker.display_summary()
        
        # Export test data 
        tracker.export_to_excel()  
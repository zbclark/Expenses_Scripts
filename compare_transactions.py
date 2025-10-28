#!/usr/bin/env python3
"""
Compare October.csv with the categorized output to find missing transactions
"""

import csv
import os
from datetime import datetime

def clean_amount(amount_str):
    """Clean and convert amount string to float"""
    if not amount_str:
        return 0.0
    # Remove any commas and convert to float
    return abs(float(str(amount_str).replace(',', '')))

def parse_csv_transactions():
    """Parse transactions from October.csv"""
    transactions = []
    
    with open('October.csv', 'r', encoding='utf-8-sig') as file:
        # Skip empty lines at the beginning
        content = file.read().strip()
        lines = content.split('\n')
        
        # Find the header line
        header_line = None
        for i, line in enumerate(lines):
            if line.strip() and 'Run Date' in line:
                header_line = i
                break
        
        if header_line is None:
            print("Could not find header line in CSV")
            return []
        
        # Parse from header line onwards
        csv_content = '\n'.join(lines[header_line:])
        reader = csv.DictReader(csv_content.splitlines())
        
        for row in reader:
            if not row.get('Run Date') or not row.get('Description'):
                continue
                
            # Extract date
            date_str = row['Run Date'].strip()
            if not date_str:
                continue
                
            # Extract amount
            amount_str = row.get('Amount ($)', '').strip()
            if not amount_str:
                continue
                
            amount = clean_amount(amount_str)
            if amount == 0:
                continue
            
            # Extract description
            description = row['Description'].strip()
            if not description or description == 'No Description':
                # Use the Action field instead
                description = row.get('Action', '').strip()
            
            transaction = {
                'date': date_str,
                'description': description,
                'amount': amount,
                'raw_row': dict(row)
            }
            transactions.append(transaction)
    
    return transactions

def get_processed_transactions_from_script():
    """Get the transactions that were processed by running the script again in dry-run mode"""
    # We'll capture the output from the script to see what it processed
    processed_transactions = []
    
    # Read the expenses.json file that was created
    try:
        import json
        with open('expenses.json', 'r') as f:
            data = json.load(f)
            
        # data is a list of transactions
        for expense in data:
            processed_transactions.append({
                'date': expense.get('date', ''),
                'description': expense.get('description', ''),
                'amount': float(expense.get('amount', 0)),
                'category': expense.get('category', '')
            })
    except FileNotFoundError:
        print("expenses.json not found - need to run the script first")
        return []
    
    return processed_transactions

def compare_transactions():
    """Compare CSV transactions with processed transactions"""
    print("🔍 Comparing October.csv with processed transactions...")
    print("=" * 60)
    
    # Get transactions from both sources
    csv_transactions = parse_csv_transactions()
    processed_transactions = get_processed_transactions_from_script()
    
    print(f"📊 CSV file contains: {len(csv_transactions)} transactions")
    print(f"📊 Processed file contains: {len(processed_transactions)} transactions")
    print()
    
    # Create lookup sets for comparison
    processed_lookup = set()
    for trans in processed_transactions:
        # Create a key based on amount and description
        key = (trans['amount'], trans['description'].lower().strip())
        processed_lookup.add(key)
    
    # Find missing transactions
    missing_transactions = []
    
    for csv_trans in csv_transactions:
        key = (csv_trans['amount'], csv_trans['description'].lower().strip())
        if key not in processed_lookup:
            missing_transactions.append(csv_trans)
    
    # Report results
    if missing_transactions:
        print(f"❌ Found {len(missing_transactions)} transactions in CSV that are NOT in processed file:")
        print("=" * 60)
        
        for i, trans in enumerate(missing_transactions, 1):
            print(f"{i}. Date: {trans['date']}")
            print(f"   Amount: ${trans['amount']:.2f}")
            print(f"   Description: {trans['description']}")
            print(f"   Raw row: {trans['raw_row']}")
            print()
    else:
        print("✅ All CSV transactions are included in the processed file!")
    
    # Also check for duplicates in processed
    processed_amounts = [t['amount'] for t in processed_transactions]
    csv_amounts = [t['amount'] for t in csv_transactions]
    
    print(f"💰 Total amount in CSV: ${sum(csv_amounts):.2f}")
    print(f"💰 Total amount processed: ${sum(processed_amounts):.2f}")
    print(f"💰 Difference: ${abs(sum(csv_amounts) - sum(processed_amounts)):.2f}")

if __name__ == "__main__":
    compare_transactions()
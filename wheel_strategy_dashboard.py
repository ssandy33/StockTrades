#!/usr/bin/env python3
"""
Wheel Strategy Transactions Dashboard Script
Generates comprehensive Excel dashboard with YTD summaries and charts

Requirements (install with: pip install pandas matplotlib xlsxwriter):
- pandas
- matplotlib
- xlsxwriter

Usage: python wheel_strategy_dashboard.py

User Configuration:
- OS: macOS Sequoia 15.6
- Python Version: 3.13.2
- Input CSV: /Users/shawnsandy/Code Repos/StockTrades/Designated_Bene_Individual_XXX885_Transactions_20250811-030603.csv
- Output Folder: /Users/shawnsandy/Code Repos/StockTrades/
- Analysis Year: 2025
- Month-Year Labels: Y
"""

import os
import sys
import re
from datetime import datetime
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
import xlsxwriter
from io import BytesIO
import warnings
warnings.filterwarnings('ignore')

# Configuration from user input
INPUT_CSV = "/Users/shawnsandy/Code Repos/StockTrades/Designated_Bene_Individual_XXX885_Transactions_20250811-030603.csv"
OUTPUT_FOLDER = "/Users/shawnsandy/Code Repos/StockTrades/"
ANALYSIS_YEAR = 2025
USE_MONTH_LABELS = True
OUTPUT_FILE = os.path.join(OUTPUT_FOLDER, "wheel_all_transactions_with_dashboard_and_ytd_summary_reverted.xlsx")

def load_and_clean(csv_path):
    """Load CSV and clean the data with robust error handling"""
    try:
        # Read CSV
        df = pd.read_csv(csv_path)
        print(f"Loaded {len(df)} rows from CSV")
        
        # Normalize headers: strip spaces and replace spaces with underscores
        df.columns = [col.strip().replace(' ', '_') for col in df.columns]
        
        # Parse TradeDate from Date column (extract first MM/DD/YYYY)
        def extract_trade_date(date_str):
            if pd.isna(date_str):
                return None
            date_str = str(date_str)
            # Look for MM/DD/YYYY pattern
            match = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', date_str)
            if match:
                month, day, year = match.groups()
                try:
                    return pd.to_datetime(f"{year}-{month}-{day}")
                except:
                    return None
            return None
        
        df['TradeDate'] = df['Date'].apply(extract_trade_date)
        
        # Convert numeric columns (remove $ and commas)
        numeric_columns = ['Price', 'Fees_&_Comm', 'Amount']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = df[col].astype(str).str.replace('$', '').str.replace(',', '')
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Handle missing columns gracefully
        for col in numeric_columns:
            if col not in df.columns:
                print(f"Warning: Column '{col}' not found, creating with zeros")
                df[col] = 0
                
        return df
        
    except Exception as e:
        print(f"Error loading/cleaning data: {e}")
        sys.exit(1)

def derive_fields(df):
    """Derive helper fields for analysis"""
    
    # IsOption: Check if Symbol or Description indicates option
    def is_option(row):
        text = str(row.get('Symbol', '')) + ' ' + str(row.get('Description', ''))
        text = text.upper()
        return any(pattern in text for pattern in [' CALL', ' PUT', ' C ', ' P ', 'OPTION'])
    
    df['IsOption'] = df.apply(is_option, axis=1)
    
    # Ticker: Extract from Symbol or Description
    def extract_ticker(row):
        symbol = str(row.get('Symbol', ''))
        if symbol and symbol != 'nan':
            # Get text before first space
            ticker = symbol.split()[0] if ' ' in symbol else symbol
            return ticker.upper()
        
        # Fallback to Description
        desc = str(row.get('Description', ''))
        if desc and desc != 'nan':
            # Look for all-caps token
            words = desc.split()
            for word in words:
                if word.isupper() and len(word) <= 5 and word.isalpha():
                    return word
        return 'UNKNOWN'
    
    df['Ticker'] = df.apply(extract_ticker, axis=1)
    
    # Premium_Net: Amount if IsOption, else 0
    df['Premium_Net'] = df.apply(lambda row: row['Amount'] if row['IsOption'] else 0, axis=1)
    
    # CashFlow: Just the Amount (signed)
    df['CashFlow'] = df['Amount']
    
    # Category classification
    def classify_category(row):
        action = str(row.get('Action', '')).upper()
        desc = str(row.get('Description', '')).upper()
        combined = action + ' ' + desc
        
        if row['IsOption']:
            if any(kw in combined for kw in ['SELL TO OPEN', 'STO']):
                return 'Options - STO'
            elif any(kw in combined for kw in ['BUY TO CLOSE', 'BTC']):
                return 'Options - BTC'
            elif 'ASSIGNED' in combined:
                return 'Options - Assigned'
            elif 'EXPIRED' in combined:
                return 'Options - Expired'
            else:
                return 'Options - Other'
        else:
            if any(kw in combined for kw in ['BUY', 'BOUGHT', 'PURCHASE']):
                return 'Stock - Buy'
            elif any(kw in combined for kw in ['SELL', 'SOLD']):
                return 'Stock - Sell'
            elif 'INTEREST' in combined:
                return 'Interest'
            else:
                return 'Other'
    
    df['Category'] = df.apply(classify_category, axis=1)
    
    # Sort by TradeDate
    df = df.sort_values('TradeDate', na_position='last')
    
    return df

def build_summary_tables(df, ytd_year):
    """Build summary and YTD summary tables"""
    
    # All-time Summary
    summary_data = {
        'Metric': [],
        'Value': []
    }
    
    # Total Premium
    total_premium = df['Premium_Net'].sum()
    summary_data['Metric'].append('Total Premium (options)')
    summary_data['Value'].append(total_premium)
    
    # Category breakdown
    category_cashflow = df.groupby('Category')['CashFlow'].sum().sort_values(ascending=False)
    for cat, val in category_cashflow.items():
        summary_data['Metric'].append(f'CashFlow - {cat}')
        summary_data['Value'].append(val)
    
    # Premium by ticker (top 10)
    ticker_premium = df[df['Premium_Net'] != 0].groupby('Ticker')['Premium_Net'].sum().sort_values(ascending=False).head(10)
    for ticker, val in ticker_premium.items():
        summary_data['Metric'].append(f'Premium - {ticker}')
        summary_data['Value'].append(val)
    
    summary_df = pd.DataFrame(summary_data)
    
    # YTD Summary
    ytd_df = df[df['TradeDate'].dt.year == ytd_year].copy()
    
    ytd_data = {
        'Metric': [],
        'Value': []
    }
    
    # YTD Options Premium
    ytd_options_premium = ytd_df['Premium_Net'].sum()
    ytd_data['Metric'].append('YTD Options Premium')
    ytd_data['Value'].append(ytd_options_premium)
    
    # YTD Net Stock Trade P/L
    stock_sell = ytd_df[ytd_df['Category'] == 'Stock - Sell']['CashFlow'].sum()
    stock_buy = ytd_df[ytd_df['Category'] == 'Stock - Buy']['CashFlow'].sum()
    ytd_stock_pl = stock_sell + stock_buy
    ytd_data['Metric'].append('YTD Net Stock Trade P/L')
    ytd_data['Value'].append(ytd_stock_pl)
    
    # YTD Net P/L (Premiums + Stock)
    ytd_net_pl = ytd_options_premium + ytd_stock_pl
    ytd_data['Metric'].append('YTD Net P/L (Premiums + Stock)')
    ytd_data['Value'].append(ytd_net_pl)
    
    # YTD Interest
    ytd_interest = ytd_df[ytd_df['Category'] == 'Interest']['CashFlow'].sum()
    ytd_data['Metric'].append('YTD Interest')
    ytd_data['Value'].append(ytd_interest)
    
    # YTD Fees
    ytd_fees = ytd_df['Fees_&_Comm'].sum()
    ytd_data['Metric'].append('YTD Fees')
    ytd_data['Value'].append(ytd_fees)
    
    # YTD Net P/L (Incl Interest & Fees)
    ytd_total_pl = ytd_net_pl + ytd_interest - ytd_fees
    ytd_data['Metric'].append('YTD Net P/L (Incl Interest & Fees)')
    ytd_data['Value'].append(ytd_total_pl)
    
    ytd_summary_df = pd.DataFrame(ytd_data)
    
    return summary_df, ytd_summary_df, ytd_df

def build_dashboard_data(ytd_df, use_month_labels):
    """Build dashboard tables and charts"""
    
    if len(ytd_df) == 0:
        print("Warning: No YTD data available for dashboard")
        return None, None, None
    
    # Monthly Premium table
    ytd_df['Month'] = ytd_df['TradeDate'].dt.to_period('M')
    monthly_premium = ytd_df.groupby('Month')['Premium_Net'].sum().reset_index()
    monthly_premium['Month'] = monthly_premium['Month'].dt.to_timestamp()
    monthly_premium = monthly_premium.sort_values('Month')
    monthly_premium['Cumulative_Premium'] = monthly_premium['Premium_Net'].cumsum()
    
    # Create charts
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8))
    
    # Chart A: Monthly Premium Income (bar chart)
    if use_month_labels:
        x_labels = [d.strftime('%b %Y') for d in monthly_premium['Month']]
    else:
        x_labels = [d.strftime('%Y-%m-%d') for d in monthly_premium['Month']]
    
    ax1.bar(range(len(monthly_premium)), monthly_premium['Premium_Net'], color='steelblue')
    ax1.set_xticks(range(len(monthly_premium)))
    ax1.set_xticklabels(x_labels, rotation=45, ha='right')
    ax1.set_title('Monthly Premium Income (YTD)', fontsize=14, fontweight='bold')
    ax1.set_ylabel('Premium ($)', fontsize=12)
    ax1.grid(True, alpha=0.3)
    ax1.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
    
    # Add value labels on bars
    for i, v in enumerate(monthly_premium['Premium_Net']):
        ax1.text(i, v, f'${v:,.0f}', ha='center', va='bottom' if v >= 0 else 'top')
    
    # Chart B: Cumulative Premium Income (line chart)
    ax2.plot(range(len(monthly_premium)), monthly_premium['Cumulative_Premium'], 
             marker='o', linewidth=2, markersize=8, color='darkgreen')
    ax2.set_xticks(range(len(monthly_premium)))
    ax2.set_xticklabels(x_labels, rotation=45, ha='right')
    ax2.set_title('Cumulative Premium Income (YTD)', fontsize=14, fontweight='bold')
    ax2.set_ylabel('Cumulative Premium ($)', fontsize=12)
    ax2.grid(True, alpha=0.3)
    
    # Add value labels on line
    for i, v in enumerate(monthly_premium['Cumulative_Premium']):
        ax2.text(i, v, f'${v:,.0f}', ha='center', va='bottom')
    
    plt.tight_layout()
    
    # Save figure to bytes buffer
    buf = BytesIO()
    plt.savefig(buf, format='png', dpi=100)
    buf.seek(0)
    chart_image = buf.getvalue()
    plt.close()
    
    # Format monthly table for display
    monthly_display = monthly_premium.copy()
    if use_month_labels:
        monthly_display['Month'] = monthly_display['Month'].dt.strftime('%b %Y')
    else:
        monthly_display['Month'] = monthly_display['Month'].dt.strftime('%Y-%m-%d')
    
    return monthly_display, chart_image, fig

def write_excel(df, summary_df, ytd_summary_df, monthly_df, chart_image, output_path):
    """Write all data to Excel with proper formatting"""
    
    # Create workbook
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    workbook = writer.book
    
    # Define formats
    currency_format = workbook.add_format({'num_format': '$#,##0.00'})
    number_format = workbook.add_format({'num_format': '#,##0.00'})
    date_format = workbook.add_format({'num_format': 'mm/dd/yyyy'})
    header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
    
    # 1. Transactions sheet
    df.to_excel(writer, sheet_name='Transactions', index=False)
    transactions_sheet = writer.sheets['Transactions']
    
    # Format columns
    for i, col in enumerate(df.columns):
        col_letter = chr(65 + i) if i < 26 else 'A' + chr(65 + i - 26)
        
        if col in ['Amount', 'Premium_Net', 'CashFlow', 'Fees_&_Comm']:
            transactions_sheet.set_column(f'{col_letter}:{col_letter}', 15, currency_format)
        elif col == 'TradeDate':
            transactions_sheet.set_column(f'{col_letter}:{col_letter}', 12, date_format)
        elif col == 'Price':
            transactions_sheet.set_column(f'{col_letter}:{col_letter}', 10, number_format)
        else:
            transactions_sheet.set_column(f'{col_letter}:{col_letter}', 15)
    
    # 2. Summary sheet
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    summary_sheet = writer.sheets['Summary']
    summary_sheet.set_column('A:A', 30)
    summary_sheet.set_column('B:B', 15, currency_format)
    
    # 3. YTD Summary sheet
    ytd_summary_df.to_excel(writer, sheet_name='YTD Summary', index=False)
    ytd_sheet = writer.sheets['YTD Summary']
    ytd_sheet.set_column('A:A', 35)
    ytd_sheet.set_column('B:B', 15, currency_format)
    
    # 4. Dashboard sheet
    if monthly_df is not None and chart_image is not None:
        dashboard_sheet = workbook.add_worksheet('Dashboard')
        
        # Write monthly table
        dashboard_sheet.write('A1', 'Monthly Premium Analysis (YTD)', header_format)
        dashboard_sheet.write('A3', 'Month', header_format)
        dashboard_sheet.write('B3', 'Premium', header_format)
        dashboard_sheet.write('C3', 'Cumulative', header_format)
        
        for i, row in monthly_df.iterrows():
            dashboard_sheet.write(i + 4, 0, row['Month'])
            dashboard_sheet.write(i + 4, 1, row['Premium_Net'], currency_format)
            dashboard_sheet.write(i + 4, 2, row['Cumulative_Premium'], currency_format)
        
        dashboard_sheet.set_column('A:A', 15)
        dashboard_sheet.set_column('B:C', 15)
        
        # Insert chart image
        dashboard_sheet.insert_image('E3', 'chart.png', {'image_data': BytesIO(chart_image)})
    
    # Save workbook
    writer.close()
    print(f"\nExcel file saved: {output_path}")

def main():
    """Main execution function"""
    print("=" * 60)
    print("Wheel Strategy Transactions Dashboard Generator")
    print("=" * 60)
    
    # Check if input file exists
    if not os.path.exists(INPUT_CSV):
        print(f"Error: Input file not found: {INPUT_CSV}")
        sys.exit(1)
    
    # Step 1: Load and clean data
    print("\n1. Loading and cleaning data...")
    df = load_and_clean(INPUT_CSV)
    
    # Step 2: Derive fields
    print("2. Deriving analysis fields...")
    df = derive_fields(df)
    
    # Step 3: Build summary tables
    print(f"3. Building summary tables (YTD year: {ANALYSIS_YEAR})...")
    summary_df, ytd_summary_df, ytd_df = build_summary_tables(df, ANALYSIS_YEAR)
    
    # Step 4: Build dashboard
    print("4. Creating dashboard charts...")
    monthly_df, chart_image, fig = build_dashboard_data(ytd_df, USE_MONTH_LABELS)
    
    # Step 5: Write Excel file
    print("5. Writing Excel file...")
    write_excel(df, summary_df, ytd_summary_df, monthly_df, chart_image, OUTPUT_FILE)
    
    # Print summary
    print("\n" + "=" * 60)
    print("PROCESSING COMPLETE")
    print("=" * 60)
    print(f"Rows processed: {len(df)}")
    print(f"YTD rows ({ANALYSIS_YEAR}): {len(ytd_df)}")
    
    if len(ytd_df) > 0:
        ytd_premium_total = ytd_df['Premium_Net'].sum()
        ytd_net_pl = ytd_summary_df[ytd_summary_df['Metric'] == 'YTD Net P/L (Incl Interest & Fees)']['Value'].values[0]
        print(f"YTD Premium Total: ${ytd_premium_total:,.2f}")
        print(f"YTD Net P/L (All): ${ytd_net_pl:,.2f}")
    
    print(f"\nOutput saved to: {OUTPUT_FILE}")
    print("\nDone!")

if __name__ == "__main__":
    main()
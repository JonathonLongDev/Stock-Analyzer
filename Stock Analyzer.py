import yfinance as yf
import pandas as pd
import xlwings as xw
import time
from yfinance.exceptions import YFRateLimitError

# Ask for stock symbol
symbol = input("Enter stock symbol (e.g., AAPL): ").upper() #.upper() # Converts all lowercase letters to uppercase
tickers = [symbol]
sheet_name = symbol  # Sheet name matches the input symbol
template_sheet_name = "Business Template Sheet"  # Name of the template sheet. Change this if your template sheet has a different name.

# Insert the File path that leads to your excel file!
file_path = r'Insert your direct file path to excel sheet'

summary_data = []

# Helper functions
def safe_get(info, key):
    return info.get(key, 'Data not available') if info.get(key) is not None else 'Data not available'

def format_percent(value):
    return f"{value * 100:.2f}%" if isinstance(value, (float, int)) else value

def get_net_income(symbol):
    try:
        stock = yf.Ticker(symbol)
        income_stmt = stock.financials
        if income_stmt.empty:
            return 'Data not available'
        latest_column = income_stmt.columns[0]
        net_income = income_stmt.loc['Net Income', latest_column]
        return f"${net_income:,.0f}"
    except Exception as e:
        print(f"Error fetching net income for {symbol}: {e}")
        return 'Data not available'

# Function to fetch stock info with retry mechanism
def fetch_stock_info(symbol, max_retries=3):
    stock = yf.Ticker(symbol)
    for attempt in range(max_retries):
        try:
            info = stock.info
            time.sleep(1)  # Add a 1-second delay to avoid rate limiting
            return info
        except YFRateLimitError:
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt * 10  # Exponential backoff: 10s, 20s, 40s
                print(f"Rate limit hit. Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                print(f"Failed to fetch data for {symbol} after {max_retries} attempts.")
                return {}
        except Exception as e:
            print(f"Error fetching data for {symbol}: {e}")
            return {}
    return {}

# Initialize variable for inside ownership
inside_ownership = 'Data not available'

# Get data
for symbol in tickers:
    info = fetch_stock_info(symbol)
    if not info:
        print(f"Skipping {symbol} due to fetch failure.")
        continue

    inside_ownership = format_percent(safe_get(info, 'heldPercentInsiders'))

    summary_data.append({
        'Company Name': safe_get(info, 'shortName'),
        'Symbol': symbol,
        'Sector / Industry': f"{safe_get(info, 'sector')} / {safe_get(info, 'industry')}",
        'Business Model Summary': safe_get(info, 'longBusinessSummary'),
        'Revenue (Most Recent FY)': safe_get(info, 'totalRevenue'),
        'Net Income (Most Recent FY)': get_net_income(symbol),
        'Revenue Growth Rate (YoY)': format_percent(safe_get(info, 'revenueGrowth')),
        'Profit Margin': format_percent(safe_get(info, 'profitMargins')),
        'Cash on Balance Sheet': safe_get(info, 'totalCash'),
        'Debt Level': safe_get(info, 'totalDebt'),
        'Dividend Status': format_percent(safe_get(info, 'dividendYield')) if isinstance(safe_get(info, 'dividendYield'), (float, int)) else 'Data not available',
        'Current Stock Price': safe_get(info, 'currentPrice'),
        'Market Cap': safe_get(info, 'marketCap'),
        'EPS': safe_get(info, 'trailingEps'),
        'EBITDA': safe_get(info, 'ebitda'),
        'Price Target Range': f"{safe_get(info, 'targetLowPrice')} - {safe_get(info, 'targetHighPrice')}",
        'Source Link': f'=HYPERLINK("https://finance.yahoo.com/quote/{symbol}", "Yahoo Finance")'
    })

# Create transposed DataFrame
df = pd.DataFrame(summary_data)
if df.empty:
    print("No data to write to Excel.")
    exit()

df = df[['Company Name', 'Symbol', 'Sector / Industry', 'Business Model Summary', 'Revenue (Most Recent FY)',
            'Net Income (Most Recent FY)', 'Revenue Growth Rate (YoY)', 'Profit Margin', 'Cash on Balance Sheet',
            'Debt Level', 'Dividend Status', 'Current Stock Price', 'Market Cap', 'EPS', 'EBITDA',
            'Price Target Range', 'Source Link']]
df = df.T

# Use xlwings to write to Excel
app = xw.App(visible=False)
wb = app.books.open(file_path)

# --- New Sheet Creation ---
# Check if the template sheet exists
try:
    template_sheet = wb.sheets[template_sheet_name]
except KeyError:
    print(f"Error: Template sheet '{template_sheet_name}' not found.  Please make sure it exists in the Excel file.")
    wb.close()
    app.quit()
    exit()

# Create a new sheet by copying the template sheet
new_sheet = template_sheet.copy(name=sheet_name, before=template_sheet)  # creates a copy before the template

# Select the new sheet
ws = wb.sheets[sheet_name]
# --- End New Sheet Creation ---

# Clear the relevant data range (starting at A1)
ws.range("A1").resize(df.shape[0], df.shape[1]).clear_contents()

# Write new data to the sheet
ws.range("A1").value = df

# Write inside ownership to B25 with label in A25
ws.range("A25").value = "Inside Ownership"
ws.range("B25").value = inside_ownership

# Write financial ratios to B29–B32 with labels
financial_labels = [
    "Return on Equity (ROE)",
    "Return on Assets (ROA)",
    "Operating Profit Margin",
    "Gross Profit Margin"
]
financial_values = [
    format_percent(safe_get(info, 'returnOnEquity')),
    format_percent(safe_get(info, 'returnOnAssets')),
    format_percent(safe_get(info, 'operatingMargins')),
    format_percent(safe_get(info, 'grossMargins'))
]

for i, (label, value) in enumerate(zip(financial_labels, financial_values), start=29):
    ws.range(f"A{i}").value = label
    ws.range(f"B{i}").value = value

# Write ratios to B36–B38 with labels
more_labels = [
    "Debt to Equity Ratio",
    "Current Ratio",
    "Quick Ratio"
]
more_values = [
    safe_get(info, 'debtToEquity'),
    safe_get(info, 'currentRatio'),
    safe_get(info, 'quickRatio')
]

for i, (label, value) in enumerate(zip(more_labels, more_values), start=36):
    ws.range(f"A{i}").value = label
    ws.range(f"B{i}").value = value

# Write valuation ratios to B42–B44 with labels in A42–A44
valuation_labels = [
    "P/E Ratio",
    "P/B Ratio",
    "P/S Ratio"
]
valuation_values = [
    safe_get(info, 'trailingPE'),
    safe_get(info, 'priceToBook'),
    safe_get(info, 'priceToSalesTrailing12Months')
]

for i, (label, value) in enumerate(zip(valuation_labels, valuation_values), start=42):
    ws.range(f"A{i}").value = label
    ws.range(f"B{i}").value = value

# Update B12 value by dividing it by 100
b12_value = ws.range("B12").value
if isinstance(b12_value, (float, int)):  # Ensure the value is numeric
    ws.range("B12").value = b12_value / 100

# Auto-fit for readability
ws.autofit()

# --- Update "Stock Mastersheet" ---
mastersheet = wb.sheets["Stock Mastersheet"]

# Find the next available row in the mastersheet
last_row = mastersheet.range("A" + str(mastersheet.api.Rows.Count)).end('up').row + 1

# Insert data into the "Stock Mastersheet"
mastersheet.range(f"A{last_row}").value = safe_get(info, 'shortName')  # Company Name
mastersheet.range(f"B{last_row}").value = symbol  # Stock Symbol
mastersheet.range(f"C{last_row}").value = f'=HYPERLINK("#{sheet_name}!A1", "{sheet_name}")'  # Link to the newly created sheet

# Save the workbook and close
wb.save()
wb.close()
app.quit()

print(f"✅ Data successfully written to sheet '{sheet_name}' and added to 'Stock Mastersheet'.\n{file_path}")

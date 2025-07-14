# Stock-Analyzer
This script uses yFinance to fetch financial information from yahoo finance for publicly traded companies and inputs them into an excel sheet.

What the Script Does

Gathers Financial Data: It uses the yfinance library to fetch a wide range of financial information about a publicly traded company from Yahoo Finance. This includes everything from the company's name and business summary to detailed financial metrics like revenue, net income, profit margins, and various valuation ratios (P/E, P/B, etc.).

Creates a New Excel Sheet: The script works with a pre-existing Excel file that you specify. It's designed to find a "template" sheet within your Excel workbook, create a copy of it, and rename the new sheet to match the stock symbol you entered (e.g., "AAPL").

Populates the Sheet: It then populates this new sheet with all the gathered financial data, neatly organizing it with labels for clarity.

Updates a Master Sheet: In addition to creating a detailed sheet for the individual stock, the script also updates a "Stock Mastersheet." It adds a new row with the company's name, its stock symbol, and a convenient hyperlink that takes you directly to the newly created detailed sheet for that stock.

Error Handling: The script includes error handling to manage situations where data might not be available or if it encounters issues with rate limiting from Yahoo Finance, making it more robust.

Comments- Please make sure to download the Business Template Sheet Excel file to use

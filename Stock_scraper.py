import yfinance as yf
from datetime import datetime, timedelta
import pandas as pd
import time
import sys
import os
import csv

def get_closing_prices_yfinance(stock_symbols):
    """
    Fetches real-time closing prices and other stock data for a list of companies
    using the yfinance library.
    """
    closing_prices = {}
    for company, symbol in stock_symbols.items():
        print(f"Fetching {company} ({symbol}) using yfinance...")
        try:
            ticker = yf.Ticker(symbol)
            data = ticker.info

            last_done = data.get('regularMarketPrice')
            prev_close = data.get('previousClose')
            day_high = data.get('dayHigh')
            day_low = data.get('dayLow')
            
            # Use 'regularMarketPrice' as the primary price, fallback to 'currentPrice'
            # and 'previousClose' to ensure data availability.
            if last_done is None:
                last_done = data.get('currentPrice')
            if prev_close is None:
                prev_close = data.get('regularMarketPreviousClose')

            if last_done is not None and prev_close is not None:
                change = last_done - prev_close
                change_pct = (change / prev_close) * 100 if prev_close != 0 else 0

                closing_prices[company] = {
                    'company_name': company,
                    'symbol': symbol,
                    'prev_close': round(prev_close, 3),
                    'last_done': round(last_done, 3),
                    'high': round(day_high, 3) if day_high is not None else 0.0,
                    'low': round(day_low, 3) if day_low is not None else 0.0,
                    'change': round(change, 3),
                    'change_pct': round(change_pct, 2),
                    'current_date': datetime.now().strftime('%Y-%m-%d'),
                    'prev_date': (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
                }
                print(f"✓ {company}: Prev: RM{prev_close:.3f} → Last Done: RM{last_done:.3f}")
            else:
                print(f"✗ No price data for {company}. Might be delisted or inactive.")
                
        except Exception as e:
            print(f"✗ Error fetching {company} with yfinance: {e}")
            sys.exc_info()
            continue
        
        # Add a short delay to avoid overwhelming the server
        time.sleep(1)

    return closing_prices

def save_to_csv(data_to_save):
    """Saves stock data to a CSV file in a specified folder."""
    if not data_to_save:
        print("No data to save.")
        return False

    # Check if running locally (Windows path exists) or in GitHub Actions
    local_folder = r'E:\Investment list for Malaysia stocks\Closing price'
    
    if os.path.exists(os.path.dirname(local_folder)) or os.name == 'nt':
        # Running locally on Windows
        folder_path = local_folder
        print(f"💻 Running locally - saving to: {folder_path}")
    else:
        # Running in GitHub Actions or other cloud environment
        folder_path = os.getcwd()
        print(f"☁️ Running in cloud - saving to current directory: {folder_path}")
    
    os.makedirs(folder_path, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f'closing_prices_{timestamp}.csv'
    file_path = os.path.join(folder_path, filename)
    
    try:
        # Use a DataFrame to handle the saving
        df = pd.DataFrame.from_dict(data_to_save, orient='index')
        df.to_csv(file_path, index=False)
        print(f"\n✅ Data saved successfully to: {file_path}")
        return True
    except Exception as e:
        print(f"❌ Error saving to CSV: {e}")
        return False

def is_market_day():
    """Check if today is a weekday (market day)"""
    today = datetime.now()
    if today.weekday() >= 5:  # Saturday = 5, Sunday = 6
        print(f"Today is {today.strftime('%A')} - Market is closed")
        return False
    return True

def run_stock_scraper():
    """
    Main function to run the stock scraping and reporting.
    """
    print(f"🚀 Malaysian Stock Tracker - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Check if it's a market day
    if not is_market_day():
        print("📅 Skipping - Market is closed on weekends")
        return

    stock_symbols = {
        "BAuto": "5248.KL", "Bursa": "1818.KL", "CIMB": "1023.KL", "Dayang": "5141.KL",
        "GenM": "4715.KL", "HapSeng": "3034.KL", "HLC": "5274.KL", "HSPlant": "5138.KL",
        "IGBCR": "5299.KL", "IGBREIT": "5227.KL", "IJM": "3336.KL", "Inari": "0166.KL",
        "KIPREIT": "5280.KL", "Kim Loong Resources Berhad": "5027.KL", "Magnum Berhad": "3859.KL", "Maxis": "6012.KL",
        "MayBank": "1155.KL", "MBSB": "1171.KL", "PAVREIT": "5212.KL", "Public Bank Berhad": "1295.KL",
        "Petronas Chemicals Group Berhad": "5183.KL", "PetGas": "6033.KL", "PPB": "4065.KL", "RHBBank": "1066.KL", "Sentral REIT": "5123.KL",
        "Sime": "4197.KL", "SToto": "1562.KL", "Tenaga": "5347.KL", "UOADev": "5200.KL",
        "UOAREIT": "5110.KL", "YTLPowr": "6742.KL"
    }

    stock_data = get_closing_prices_yfinance(stock_symbols)

    if not stock_data:
        print("\nFailed to retrieve any stock data. Exiting.")
        return

    # Create a DataFrame from the dictionary
    df = pd.DataFrame.from_dict(stock_data, orient='index')

    # Reorder columns to place 'last_done' after 'prev_close' and remove 'volume'
    df = df[['company_name', 'symbol', 'prev_close', 'last_done', 'high', 'low', 'change', 'change_pct']]

    # Format the output with a clean table
    print("\n" + "=" * 95)
    print(f"Market Report for {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 95)
    
    # Print the table header
    print(f"{'Company':<30} | {'Symbol':<8} | {'Prev. Close':>11} | {'Last Done':>11} | {'High':>11} | {'Low':>11} | {'Change':>10} | {'% Change':>10}")
    print("-" * 95)
    
    # Print data for each stock
    for index, row in df.iterrows():
        change_text = f"{row['change']:+.3f}"
        change_pct_text = f"{row['change_pct']:+.2f}%"

        print(f"{row['company_name']:<30} | "
              f"{row['symbol']:<8} | "
              f"RM{row['prev_close']:<9.3f} | "
              f"RM{row['last_done']:<9.3f} | "
              f"RM{row['high']:<9.3f} | "
              f"RM{row['low']:<9.3f} | "
              f"{change_text:>10} | "
              f"{change_pct_text:>10}")
        
    print("-" * 95)
    print("Disclaimer: Data is provided for informational purposes only and is not intended for trading purposes.")
    print("=" * 95)
    
    # Call the function to save the data to a CSV file
    save_to_csv(stock_data)
    print("🎉 Stock tracking completed!")

if __name__ == "__main__":
    run_stock_scraper()
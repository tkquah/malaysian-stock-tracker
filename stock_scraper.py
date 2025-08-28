import yfinance as yf
from datetime import datetime, timedelta
import pandas as pd
import time
import sys
import os
import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

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
    """Saves stock data to a CSV file."""
    if not data_to_save:
        print("No data to save.")
        return None

    # Save to current directory (works for both local and cloud)
    folder_path = os.getcwd()
    os.makedirs(folder_path, exist_ok=True)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f'malaysian_stocks_{timestamp}.csv'
    file_path = os.path.join(folder_path, filename)
    
    try:
        # Use a DataFrame to handle the saving
        df = pd.DataFrame.from_dict(data_to_save, orient='index')
        df.to_csv(file_path, index=False)
        print(f"\n✅ Data saved successfully to: {file_path}")
        return file_path
    except Exception as e:
        print(f"⌐ Error saving to CSV: {e}")
        return None

def send_email_with_attachment(csv_file_path, stock_summary):
    """Send email with CSV attachment using Gmail SMTP"""
    
    # Email configuration - UPDATED TO USE YOUR GMAIL
    sender_email = "tkquahinv@gmail.com"  # Changed to your Gmail
    sender_password = "your_app_password_here"  # Will be set via GitHub secrets
    recipient_email = "tkquahinv@gmail.com"  # Sending to yourself
    
    # Create message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = f"Malaysian Stock Report - {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    
    # Create email body with stock summary
    body = f"""
Hello,

Your Malaysian stock market report is ready!

📊 STOCK SUMMARY ({datetime.now().strftime('%Y-%m-%d %H:%M:%S')}):

{stock_summary}

🔍 Detailed data is attached as a CSV file.

📈 Top Gainers and Losers will be highlighted in the attachment.

Best regards,
Malaysian Stock Tracker Bot

---
This is an automated message. Data is for informational purposes only.
    """
    
    msg.attach(MIMEText(body, 'plain'))
    
    # Attach CSV file
    if csv_file_path and os.path.exists(csv_file_path):
        try:
            with open(csv_file_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {os.path.basename(csv_file_path)}'
            )
            msg.attach(part)
            print("✅ CSV file attached to email")
            
        except Exception as e:
            print(f"⌐ Error attaching file: {e}")
    
    # Send email
    try:
        # Use environment variable for password (more secure)
        password = os.environ.get('GMAIL_APP_PASSWORD', sender_password)
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)
        
        text = msg.as_string()
        server.sendmail(sender_email, recipient_email, text)
        server.quit()
        
        print(f"✅ Email sent successfully to {recipient_email}")
        return True
        
    except Exception as e:
        print(f"⌐ Error sending email: {e}")
        return False

def generate_stock_summary(stock_data):
    """Generate a text summary of stock performance"""
    if not stock_data:
        return "No stock data available."
    
    df = pd.DataFrame.from_dict(stock_data, orient='index')
    
    # Sort by percentage change
    df_sorted = df.sort_values('change_pct', ascending=False)
    
    summary = []
    summary.append("TOP 5 GAINERS:")
    summary.append("-" * 40)
    for i, (_, row) in enumerate(df_sorted.head(5).iterrows()):
        summary.append(f"{i+1}. {row['company_name'][:20]:<20} {row['change_pct']:+6.2f}%")
    
    summary.append("\nTOP 5 LOSERS:")
    summary.append("-" * 40)
    for i, (_, row) in enumerate(df_sorted.tail(5).iterrows()):
        summary.append(f"{i+1}. {row['company_name'][:20]:<20} {row['change_pct']:+6.2f}%")
    
    summary.append(f"\nTotal stocks tracked: {len(df)}")
    summary.append(f"Gainers: {len(df[df['change_pct'] > 0])}")
    summary.append(f"Losers: {len(df[df['change_pct'] < 0])}")
    summary.append(f"Unchanged: {len(df[df['change_pct'] == 0])}")
    
    return "\n".join(summary)

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
    print(f"🚀 Malaysian Stock Tracker with Email - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
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

    print(f"📊 Fetching data for {len(stock_symbols)} Malaysian stocks...")
    stock_data = get_closing_prices_yfinance(stock_symbols)

    if not stock_data:
        print("\nFailed to retrieve any stock data. Exiting.")
        return

    # Create a DataFrame from the dictionary
    df = pd.DataFrame.from_dict(stock_data, orient='index')

    # Reorder columns to place 'last_done' after 'prev_close'
    df = df[['company_name', 'symbol', 'prev_close', 'last_done', 'high', 'low', 'change', 'change_pct']]

    # Display summary in console
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
    
    # Save to CSV file
    print("\n💾 Saving data to CSV...")
    csv_file_path = save_to_csv(stock_data)
    
    # Generate summary for email
    stock_summary = generate_stock_summary(stock_data)
    
    # Send email with attachment
    if csv_file_path:
        print("\n📧 Sending email with CSV attachment...")
        email_sent = send_email_with_attachment(csv_file_path, stock_summary)
        
        if email_sent:
            print("✅ Report successfully sent to tkquahinv@gmail.com")
        else:
            print("⌐ Failed to send email, but CSV file is saved locally")
    
    print("🎉 Malaysian Stock Tracker completed!")

if __name__ == "__main__":
    run_stock_scraper()

import os
import pandas as pd
from datetime import datetime, time, timedelta
import yfinance as yf
import tkinter as tk
from tkinter import filedialog
import rumps
from AppKit import NSApplication, NSStatusBar, NSTextField, NSColor, NSView, NSMakeRect, NSVariableStatusItemLength
import pandas_market_calendars as mcal

def get_file_path():
    home_dir = os.path.expanduser('~')
    file_path = os.path.join(home_dir, 'stocks.xls')  # Assuming the file is named 'stocks.xls'

    if os.path.exists(file_path):
        return file_path
    else:
        print(f"File not found: {file_path}")
        exit(1)

class StockApp(rumps.App):
    def __init__(self):
        super(StockApp, self).__init__("")
        self.file_path = get_file_path()
        print(f"Using file: {self.file_path}")
        self.df = pd.read_excel(self.file_path)
        self.ticker_symbols = self.df['Ticker'].tolist()
        self.shares = dict(zip(self.df['Ticker'], self.df['Shares']))
        self.cached_closing_prices = {}  # Dictionary to cache closing prices

        self.menu = ['Update']
        self.status_bar = NSStatusBar.systemStatusBar()
        self.status_item = self.status_bar.statusItemWithLength_(NSVariableStatusItemLength)

        self.timer = rumps.Timer(self.on_tick, 5)  # Update every 5 seconds
        self.timer.start()

    def get_nyse_schedule(self, date):
        nyse = mcal.get_calendar('NYSE')
        schedule = nyse.schedule(start_date=date, end_date=date)
        if not schedule.empty:
            return schedule.iloc[0]  # Return the schedule for the given date
        return None

    def get_previous_business_day(self, date):
        previous_day = date - timedelta(days=1)
        while self.get_nyse_schedule(previous_day) is None:
            previous_day -= timedelta(days=1)
        return previous_day

    def on_tick(self, sender):
        print("on_tick called")
        total_portfolio_value_today = 0
        total_portfolio_value_yesterday = 0
        today = pd.Timestamp.today().normalize()
        print(f"Today's Date: {today}")

        # Get NYSE schedule for today
        nyse_schedule_today = self.get_nyse_schedule(today)
        print(f"NYSE Schedule Today: {nyse_schedule_today}")
        if nyse_schedule_today is None:
            # Handle non-trading days (e.g., weekends or holidays)
            print("NYSE is closed today.")
            previous_business_day = self.get_previous_business_day(today)
            print(f"Previous Business Day: {previous_business_day}")
            nyse_schedule_previous = self.get_nyse_schedule(previous_business_day)
            print(f"NYSE Schedule Previous: {nyse_schedule_previous}")
            if nyse_schedule_previous is None:
                print("No previous business day found.")
                return

            # Use the previous business day's regular session closing price and extended session price
            for ticker_symbol in self.ticker_symbols:
                try:
                    ticker = yf.Ticker(ticker_symbol)
                    
                    # Get the previous business day's regular session closing price
                    prev_day_data = ticker.history(start=previous_business_day, end=previous_business_day + timedelta(days=1), interval="1d")
                    print(f"Previous Day Data for {ticker_symbol}: {prev_day_data}")
                    if not prev_day_data.empty:
                        closing_price = prev_day_data['Close'].iloc[-1]
                        self.cached_closing_prices[ticker_symbol] = closing_price  # Update cache
                    else:
                        closing_price = None

                    # Get the previous business day's extended session price as the current price
                    extended_session_data = ticker.history(start=previous_business_day, end=previous_business_day + timedelta(days=1), interval="1m", prepost=True).tail(1)
                    print(f"Extended Session Data for {ticker_symbol}: {extended_session_data}")
                    if not extended_session_data.empty:
                        current_price = extended_session_data['Close'].iloc[-1]
                    else:
                        current_price = None

                    # Calculate portfolio values
                    if closing_price is not None:
                        total_value_yesterday = closing_price * self.shares[ticker_symbol]
                        total_portfolio_value_yesterday += total_value_yesterday

                    if current_price is not None:
                        total_value_today = current_price * self.shares[ticker_symbol]
                        total_portfolio_value_today += total_value_today

                except Exception as e:
                    print(f"Error occurred for {ticker_symbol}: {e}")

            # Calculate and display the change in portfolio value
            change_in_value = total_portfolio_value_today - total_portfolio_value_yesterday
            formatted_change_in_value = "${:,.2f}".format(change_in_value)
            print(formatted_change_in_value)
            
            self.update_status_bar(formatted_change_in_value, change_in_value)
            return

        # If NYSE is open, proceed with the original logic
        nyse_closing_time = nyse_schedule_today['market_close'].to_pydatetime().time()
        current_time = datetime.now().time()
        print(f"NYSE Closing Time: {nyse_closing_time}, Current Time: {current_time}")

        for ticker_symbol in self.ticker_symbols:
            try:
                ticker = yf.Ticker(ticker_symbol)
                if current_time >= nyse_closing_time:  # If it's after NYSE closing time
                    # Fetch the current day's close price
                    data = ticker.history(start=today, end=today + timedelta(days=1), interval="1d")
                    print(f"Data for {ticker_symbol}: {data}")
                    if not data.empty:
                        closing_price = data['Close'].iloc[-1]
                        self.cached_closing_prices[ticker_symbol] = closing_price  # Update cache
                    else:
                        closing_price = None
                else:
                    # Use the cached closing price if available
                    if ticker_symbol in self.cached_closing_prices:
                        closing_price = self.cached_closing_prices[ticker_symbol]
                    else:
                        # Fetch the previous business day's closing price and cache it
                        previous_business_day = today - timedelta(days=1)
                        while self.get_nyse_schedule(previous_business_day) is None:
                            previous_business_day -= timedelta(days=1)
                        prev_day_data = ticker.history(start=previous_business_day, end=previous_business_day + timedelta(days=1)).tail(1)
                        print(f"Previous Day Data for {ticker_symbol}: {prev_day_data}")
                        if not prev_day_data.empty:
                            closing_price = prev_day_data['Close'].iloc[-1]
                            self.cached_closing_prices[ticker_symbol] = closing_price  # Update cache
                        else:
                            closing_price = None

                # Use the determined closing price for previous business day or current day
                previous_day_close_price = closing_price
            
                total_value_yesterday = previous_day_close_price * self.shares[ticker_symbol]
                total_portfolio_value_yesterday += total_value_yesterday
                

                # Fetch the current price including extended hours trading
                current_day_data = ticker.history(period='1d', interval='1m', prepost=True).tail(1)
                print(f"Current Day Data for {ticker_symbol}: {current_day_data}")
                current_price = current_day_data['Close'].iloc[-1]
                

                total_value_today = current_price * self.shares[ticker_symbol]
                total_portfolio_value_today += total_value_today
                
            except Exception as e:
                print(f"Error occurred for {ticker_symbol}: {e}")
        print(closing_price)
        print(current_price)
        change_in_value = total_portfolio_value_today - total_portfolio_value_yesterday
        formatted_change_in_value = "${:,.2f}".format(change_in_value)
        print(formatted_change_in_value)
        
        self.update_status_bar(formatted_change_in_value, change_in_value)

    def update_status_bar(self, text, value):
        print(f"Updating status bar with: {text}")
        # Create a custom view for the status bar
        view = NSView.alloc().initWithFrame_(NSMakeRect(0, 0, 100, 22))
        label = NSTextField.alloc().initWithFrame_(NSMakeRect(0, 0, 100, 22))

        # Set the label properties
        label.setStringValue_(text)
        label.setBezeled_(False)
        label.setDrawsBackground_(False)
        label.setEditable_(False)
        label.setSelectable_(False)
        label.setAlignment_(2)  # Center alignment

        # Set the text color based on the value
        if value >= 0:
            label.setTextColor_(NSColor.blueColor())
        else:
            label.setTextColor_(NSColor.redColor())

        # Add the label to the view
        view.addSubview_(label)
        
        # Set the custom view to the status item
        self.status_item.setView_(view)

    @rumps.clicked("Update")
    def on_update(self, _):
        self.on_tick(None)

if __name__ == "__main__":
    # Initialize NSApplication
    app = NSApplication.sharedApplication()
    stock_app = StockApp()
    stock_app.run()
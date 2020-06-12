#!/usr/bin/env python
# coding: utf-8

# In[2]:


# Google Sheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Selenium
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time 

# Essentials
import re
import numpy as py
import pandas as pd
from pandas.tseries.offsets import BDay
from datetime import datetime, date, timedelta

# Stock price
from yahoo_fin import stock_info as si

# Pickle
import pickle

# Data Analysis
import scipy.stats
import seaborn as sns


# In[3]:


import requests
class APIError(Exception):
    """An API Error Exception"""

    def __init__(self, status):
        self.status = status

    def __str__(self):
        return "APIError: status={}".format(self.status)


# In[4]:


# Gets the original data needed to start the analysis. The sheet has a record of all my past trades 

def scrape_sheets():
    # Get access to the sheet
    scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('Options_Scrapper.json',scope)
    client = gspread.authorize(creds)
    sh = client.open('Account')
    worksheet_history = sh.worksheet("History")

    # Essential variables
    total_quant_list = []
    df_history = pd.DataFrame()

    # All the information
    security_list = worksheet_history.col_values(4)[1:]
    cost_basis_list = worksheet_history.col_values(5)[1:]
    quantity_list = [int(quantity) for quantity in worksheet_history.col_values(3)[1:]]
    date_list = worksheet_history.col_values(1)[1:]
    action_list = [-1 if action=="Sold" else 1 for action in worksheet_history.col_values(2)[1:]]
    
    # Multiplies the bought(1)/sold(-1) by the amount
    for i in range(0, len(quantity_list)):
        total_quant_list.append(action_list[i] * quantity_list[i])

    df_history[worksheet_history.col_values(4)[0]] = security_list
    df_history[worksheet_history.col_values(5)[0]] = cost_basis_list
    df_history[worksheet_history.col_values(1)[0]] = date_list
    df_history["Quantity"] = total_quant_list
    return(df_history)


# In[5]:


# Finds all the options needed to be scraped to find current price. Returns the expiration and security

def match_tickers(df_history):
    ref_dict = {}
    
    # goes through each row of the dataframe 
    for index, row in df_history.iterrows():
        security = row["Security"]
        option_info = security.split(' ')
        ticker = option_info[0]
        # try becase an equity will only have the ticker
        try: 
            expiration_seek = option_info[1]
            compare_date_seek = datetime.strptime(expiration_seek, '%m/%d/%Y')
            append_info = [expiration_seek, security]
            
            # if the option has expired, don't find the price
            if compare_date_seek > datetime.today():
                
                # If its a new ticker, put the info into a new key
                if ticker not in ref_dict:
                    ref_dict[ticker] = [append_info]
        
                else:
                    # if it matches a ticker, add the entry to it
                    if append_info not in ref_dict[ticker]:
                        ref_dict[ticker].append(append_info)
        except:
            pass
    return(ref_dict)


# In[6]:


# Creates the month and year string to find the match on the Cboe Website

def year_month(date):
    broken = date.split("/")
    month_number = int(broken[0])
    day = int(broken[1])
    year = int(broken[2])
    datetime_object = datetime(year, month_number, day)
    month_abre = datetime_object.strftime('%b')
    complete_string = month_abre + " " + str(year) 
    return(complete_string)


# In[7]:


# Using the match_ticker function, we have the security and expiration
# We find the price using Cboe 

def get_all_option_quotes(match_ticker_dict):
    option_value_dict = {}
    driver = webdriver.Chrome(r'C:\Users\ericn\.wdm\drivers\chromedriver\83.0.4103.39\win32\chromedriver.exe')
    
    # Sometimes the website can be laggy, so try up to 5 times
    x = 0
    while x < 5:
        try:
            
            # Get all the tickers in the dictionary, iterate through them
            all_tickers = list(match_ticker_dict.keys())
            for i in range(len(all_tickers)):
                ticker = all_tickers[i]
                base_url = "http://www.cboe.com/delayedquote/quote-table"
                driver.get(base_url)

                # Put in the ticker and search
                loginInput = driver.find_element_by_xpath("""//*[@id="txtSymbol"]""")
                loginInput.send_keys(ticker)
                driver.find_element_by_xpath("""//*[@id="btnSubmit"]""").click()

                time.sleep(2)

                # Change criteria to all strikes
                driver.find_element_by_xpath("""//*[@id="ddlRange"]""").click()
                driver.find_element_by_xpath("""//*[@id="ddlRange"]/option[1]""").click()

                # Finds the expiration associated with the ticker
                expiration_security = match_ticker_dict[ticker]
                for i in range(len(expiration_security)):
                    current_expiration = expiration_security[i][0]
                    current_security = expiration_security[i][1]
#                     print(current_security)
                    info = current_security.split(' ')
                    ticker = info[0]
                    strike = info[2]
                    option_type = info[3]

                    # Change criteria to specific date by clicking on the bottom 
                    # We need to go by month first because there might be too many quotes

                    # Finds the list of dates
                    exp = driver.find_element_by_xpath("""//*[@id="divpgen-cntr"]""").text
                    all_expirations_list = exp.splitlines()

                    # clicks on the one that matches the month and year we are looking for
                    year_month_expiration = year_month(current_expiration)
                    xpath_expiration_index = all_expirations_list.index(year_month_expiration) + 2
                    xpath_expiration = """//*[@id="ddlMonth"]/option[""" + str(xpath_expiration_index) + "]"
                    driver.find_element_by_xpath(xpath_expiration).click()

                    # Apply Criteria
                    driver.find_element_by_xpath("""//*[@id="btnFilter"]""").click()
                    time.sleep(3)

                    # Finds all expiration days within the month
                    expiration_list = []
                    expiration_object = driver.find_elements_by_id("expiration_date")
                    for expiration in expiration_object:
                        expiration_list.append(expiration.text)
                    
                    # Find the strikes within expiration within the month that we are looking for
                    expiration_location = expiration_list.index(current_expiration)
                    index_string = str(expiration_location+1) + '"'
                    strike_chain = driver.find_element_by_xpath("""//*[@id="temp-default""" + index_string + """]/div[2]""")
                    strike_chain_array = strike_chain.text.splitlines()
                    
                    # If the option is a call
                    if option_type == "C":

                        call_xpath_index = """//*[@id="temp-default""" + str(expiration_location+1) +""""]/div[1]"""
                        call_chain = driver.find_element_by_xpath(call_xpath_index)
                        call_chain_array = call_chain.text.splitlines()[2:]

                        strike_find = ticker + " " + str(strike) + "0"
                        strike_chain_index = strike_chain_array.index(strike_find) - 2


                        #['Last', 'Net', 'Bid', 'Ask', 'Vol', 'IV', 'Delta', 'Gamma', 'Int']
                        call_info = call_chain_array[strike_chain_index].split(' ')
                        bid_price = float(call_info[2])
                        ask_price = float(call_info[3])
                        delta = float(call_info[6])

                        # Price is the average of the bid and ask
                        call_value = (bid_price + ask_price)/2
                        final_value = call_value

                    # if its a put
                    else:
                        
                        put_xpath_index = """//*[@id="temp-default""" + str(expiration_location+1) +""""]/div[3]"""
                        put_chain = driver.find_element_by_xpath(put_xpath_index)
                        put_chain_array = call_chain.text.splitlines()[2:]

                        # Gets the strike index, includes the strike header, so -2 vs -1
                        strike_find = ticker + " " + str(strike) + "0"
                        strike_chain_index = strike_chain_array.index(strike_find) - 2

                        # Uses the index to get the exact call values
                        #['Last', 'Net', 'Bid', 'Ask', 'Vol', 'IV', 'Delta', 'Gamma', 'Int']
                        put_info = call_chain_array[strike_chain_index].split(' ')
                        bid_price = float(put_info[2])
                        ask_price = float(put_info[3])
                        delta = float(put_info[6])

                        # Price is the average of the bid and ask
                        put_value = (bid_price + ask_price)/2
                        final_value = put_value

                    option_value_dict[current_security] = [round(final_value*100,3), delta]
                # Wait for data to refresh
                    # break the while loop
                x += 6
            print("x")
                
        except:
            x += 1
    driver.close()
    return(option_value_dict)


# In[8]:


# Finds live stock price of an equity
def find_stock_price(security):
    price = si.get_live_price(security)
    return(round(price,3))


# In[9]:


# Returns the security price of either an equity or options
# Uses the option_values_dict which holds all the values needed for options

def find_security_price(security, option_values_dict):
    # If it is just an equity
    if len(security) > 5:
        # Find if the option has not expired yet
        option_info = security.split(' ')
        expiration_seek = option_info[1]
        compare_date_seek = datetime.strptime(expiration_seek, '%m/%d/%Y')
#         print(expiration_seek)
        # if it has expired, we can't hold it so return a value of 0, it will be deleted later
        if compare_date_seek < datetime.today():
            return([0,0])
        # find the price from the dictionary
        return_price = option_values_dict[security]
        
    # If it is an equity
    else:
        final = [find_stock_price(security)]
        return_price = final
    return(return_price)


# In[16]:


# Returns dictionaries for active and sold securities to be put in google sheets

active_dict = {}
sold_dict = {}

def sold_active(df_history):
    match_ticker = match_tickers(df_history)
    option_values_dict = get_all_option_quotes(match_ticker)

    for i in range(len(df_history)):
        item = df_history.loc[i]
        security = item["Security"]
        date = item["Date"]
        # cost will be sold_price if sold or cost_basis if bought
        cost = float(item["Cost Basis"][1:].replace(',',''))
        quantity = item["Quantity"]
        # Active
        if quantity > 0:
                
            # if it has been bought already, add the new quantity and update the price and unrealized gain (loss)
            if security in active_dict.keys():
                # Adds new quantity (Q1 + Q2)
                previous_quantity = active_dict[security][1]
                new_quantity = previous_quantity + quantity
                
                # Adds new cost_basis (Q1 + Q2)/(C1 + C2)
                previous_cost_basis = active_dict[security][2]
                new_cost_basis = (previous_cost_basis + cost) / (new_quantity) 
                
                # Finds stock or option price, delta was already found the first time
                option_stock = find_security_price(security, option_values_dict)
                if len(option_stock) > 1:
                    security_price = option_stock[0] * new_quantity
                else:
                    security_price = option_stock[0] * new_quantity
                
                unrealized_profit = security_price - new_cost_basis
                
                # Finds the new risk with the additional security
                new_risk = risk_assessment(security, active_dict[security][5], new_quantity)
                
                # Puts everything in the dictionary (except the date & delta, which stays the same)
                active_dict[security][1] = new_quantity
                active_dict[security][2] = new_cost_basis
                active_dict[security][3] = security_price
                active_dict[security][4] = unrealized_profit
                active_dict[security][6] = new_risk
                
            # if it hasn't been bought yet, add a new entry
            else:
                
                option_stock = find_security_price(security, option_values_dict)
#                 print("option_chain_dict", option_chain_dict)
                if len(option_stock) > 1:
                    security_price = option_stock[0] * quantity
                    security_delta = option_stock[1]
                else:
                    security_price = option_stock[0] * quantity
                    security_delta = ""

                unrealized_profit = security_price - cost
                risk = risk_assessment(security, security_delta, quantity)
                active_dict[security] = [date, quantity, cost, security_price, unrealized_profit, security_delta, risk]
        
        # if the quantity is less than 0, it has been sold
        if quantity < 0:
            print(security)
            # finds cost basis for the quantity
            cost_basis = active_dict[security][2]/active_dict[security][1]*quantity
            
            # if its already been sold before, add to the list
            if security in sold_dict.keys():
                
                # Finds new quantity 
                previous_quantity = sold_dict[security][1]
                new_quantity = previous_quantity + quantity
                
                # If we had already sold it before, find the average price
                previous_sold_price = sold_dict[security][2]
                new_sold_price = (previous_sold_price + cost) / new_quantity
                
                # Finds realized profit
                realized_profit = new_sold_price - cost_basis
                
                # Puts it in the dictionary
                sold_dict[security] = [date, quantity, new_sold_price, realized_profit]
                
                # Changes the active quantity by the amount sold
                active_quantity = active_dict[security][1] + quantity
                
                # if you sold the last security, it is no longer active, take it out of the active dict
                if active_quantity == 0:
                    active_dict.pop(security)
            
            # if its not been sold before
            else:
                
                realized_profit = cost + cost_basis 
                sold_dict[security] = [date, quantity, cost, cost_basis, realized_profit]
            
                # update how many are still active
                # if you sold the last security, it is no longer active, take it out of the active dict
                active_quantity = active_dict[security][1] + quantity
                if active_quantity == 0:
                    active_dict.pop(security)
                    
    all_records = [active_dict,sold_dict]
    print(sold_dict)
    return(all_records)


# In[11]:


def update_worksheet(all_records):
    scope = ['https://spreadsheets.google.com/feeds',
     'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('Options_Scrapper.json',scope)
    client = gspread.authorize(creds)
    sh = client.open('Account')
    worksheet_account = sh.worksheet("Account")
    
    active_dict = all_records[0]
    solddict = all_records[1]
    
    # [date, quantity, cost, security_price, unrealized_profit, delta]
    active_sec = list(active_dict.keys())
    # time sleep because of google drive API restrictions
    x = 0
    while x < len(active_sec):
        for i in range(len(active_dict)):
            try:
                
                security = active_sec[i]
                date = active_dict[security][0]
                quantity = int(active_dict[security][1])
                cost = active_dict[security][2]
                current_price = active_dict[security][3]
                unrealized = active_dict[security][4]
                delta = active_dict[security][5]
                risk = active_dict[security][6]

                # (i + 2) vs (i + 1) because of the header
                worksheet_account.update_cell(i+2,1,security)
                worksheet_account.update_cell(i+2,2,date)
                worksheet_account.update_cell(i+2,3,quantity)
                worksheet_account.update_cell(i+2,4,cost)
                worksheet_account.update_cell(i+2,5,current_price)
                worksheet_account.update_cell(i+2,6,unrealized)
                worksheet_account.update_cell(i+2,7,delta)
                worksheet_account.update_cell(i+2,8,risk)
                
                x += 1
            except Exception as err:
                print(err)
                time.sleep(100)

        # Actual gap is number - 1
        gap_inbetween = 6
        total_spaces = len(active_dict) + gap_inbetween + 1
        worksheet_account.update_cell(total_spaces-1,1,"Security")
        worksheet_account.update_cell(total_spaces-1,2,"Date")
        worksheet_account.update_cell(total_spaces-1,3,"Quantity")
        worksheet_account.update_cell(total_spaces-1,4,"Cost Basis")
        worksheet_account.update_cell(total_spaces-1,5,"Price Sold")
        worksheet_account.update_cell(total_spaces-1,6,"Realized Profit")



#     [date, quantity, cost, cost_basis, realized_profit], cost is price sold
    sold_sec = list(sold_dict.keys())
    
    # time sleep because of google drive API restrictions
    i = 0
    while i < len(sold_dict):
        try:
            security_sold = sold_sec[i]
            date_sold = sold_dict[security_sold][0]
            quantity_sold = int(sold_dict[security_sold][1])
            price_sold = sold_dict[security_sold][2]
            cost_basis_sold = sold_dict[security_sold][3]
            realized_sold = sold_dict[security_sold][4]

            worksheet_account.update_cell(i+total_spaces,1,security_sold)
            worksheet_account.update_cell(i+total_spaces,2,date_sold)
            worksheet_account.update_cell(i+total_spaces,3,quantity_sold)
            worksheet_account.update_cell(i+total_spaces,4,cost_basis_sold)
            worksheet_account.update_cell(i+total_spaces,5,price_sold)
            worksheet_account.update_cell(i+total_spaces,6,realized_sold)
            
            i += 1
        except Exception as err:
            print(err)
            time.sleep(100)
    print("COMPLETED")


# In[12]:


# Gets the lowest 5% of returns
def lower_confidence_interval(data, confidence):
    
    k = int(round(len(data) * (1-confidence),0))
    idx = py.argpartition(data, k)
    lowest_performance = data[idx[:k]]
    return(max(lowest_performance))
#     return(max(idx))
#     a = 1.0 * py.array(data)
#     n = len(a)
#     m, se = py.mean(a), scipy.stats.sem(a)
#     print("se",se)
#     h = se * scipy.stats.t.ppf((1 + confidence) / 2., n-1)
#     print(2*h)
#     return  m-(h)


# In[13]:


def risk_assessment(security, delta, quantity):
    
    # security info needed
    option_info = security.split(' ')
    ticker = option_info[0]
    current_price = si.get_live_price(ticker)
    
    # Time info needed
    length_of_time = 365
    current_date = date.today().isoformat()  
    end_date = (date.today()-timedelta(days=1)).isoformat()
    start_date = (date.today()-timedelta(days=length_of_time)).isoformat()
    
    # Analysis
    #get_data(ticker, start_date = None, end_date = None, index_as_date = True, interval = “1d”)
    data = si.get_data(ticker, start_date, end_date)
    closing_prices = data["close"].tolist()
    day_chg = (py.array(closing_prices[:-1]) - py.array(closing_prices[1:])) / py.array(closing_prices[:-1])
    lower_confidence = lower_confidence_interval(day_chg, 0.95)
    lowest_day_change = current_price * lower_confidence
    total_position_affected = lowest_day_change * quantity
    
    # If its an option
    if len(security) > 5:

        position_affected = total_position_affected * delta * 100
        return(position_affected)
    return(total_position_affected)


# In[26]:


# # Scrape the history of trades on a websheet
df_history = scrape_sheets()

# organize the trades to those that are still sold vs active
# Also analyze the data to find the delta and risk of the portfolio
all_records = sold_active(df_history) 

# Put the information on the google sheet
update_worksheet(all_records)


# In[ ]:





# In[15]:


# test_active = {'AAL 01/15/2021 9.00 C': ['04/24/2020', 1, 402.0, 730.0000000000001, 328.0000000000001], 'GS 01/15/2021 185.00 C': ['04/27/2020', 1, 2208.0, 4132.5, 1924.5], 'SBUX 07/17/2020 75.00 C': ['04/29/2020', 1, 690.0, 597.5, -92.5], 'QQQ 07/17/2020 212.00 C': ['05/01/2020', 1, 1342.0, 2682.0, 1340.0], 'AAL 01/21/2022 10.00 C': ['05/04/2020', 1, 520.0, 792.5000000000001, 272.5000000000001], 'AAPL': ['05/08/2020', 2, 618.77, 368.04, -250.72999999999996], 'QQQ 08/21/2020 225.00 C': ['05/13/2020', 1, 1260.0, 1876.5, 616.5], 'GIS 06/19/2020 60.00 C': ['05/19/2020', 1, 255.0, 140.5, -114.5], 'ITB 10/16/2020 45.00 C': ['05/27/2020', 1, 430.0, 480.0, 50.0]}
# test_sold = {'TZA 05/08/2020 37.00 C': ['05/01/2020', -1, 371.0, 771.0], 'TZA 05/01/2020 35.00 C': ['05/01/2020', -1, 375.0, 760.0], 'SPXS 05/15/2020 10.00 C': ['05/05/2020', -1, 84.0, 224.0], 'SQQQ 05/08/2020 12.00 C': ['05/11/2020', -1, 0.0, 109.0], 'QQQ 05/29/2020 213.00 C': ['05/11/2020', -1, 1481.0, 2446.0], 'GIS 05/15/2020 60.00 C': ['05/13/2020', -1, 351.0, 926.0], 'GIS 06/19/2020 60.00 C': ['05/14/2020', -1, 364.0, 874.0], 'QQQ 06/05/2020 220.00 C': ['05/26/2020', -1, 1275.0, 2025.0], 'QQQ 06/19/2020 231.00 C': ['06/03/2020', -1, 815.0, 1466.0]}
# [date, quantity, cost_basis, security_price, unrealized]


# [date, quantity, cost_basis, realized_price]


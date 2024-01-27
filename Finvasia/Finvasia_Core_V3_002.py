# -*- coding: utf-8 -*-
"""
Created on Fri Oct 14 09:13:23 2022

This tool is to fetch two option chain (of any segment/symbol) data and trade terminal sheet, where all input comes using web socket data which leads to low lag and no api restriction. 

Contact details :
Telegram Channel:  https://t.me/pythontrader
Developer Telegram ID : https://t.me/pythontrader_admin
Gmail ID:   mnkumar2020@gmail.com 
Whatsapp : 9470669031 
Youtube : youtube.com/@pythontraders

Disclaimer: The information provided by the Python Traders channel is for educational purposes only, so please contact your financial adviser before placing your trade. Developer is not responsible for any profit/loss happened due to coding, logical or any type of error.


Toots Feature :
1. Trade terminal where you can take paper/real trade.
2. You can set target/sl together using tool.
3. You can also trail your sl.
4. Tool has voice and telegram notification/alert feature.
5. Tools have orderbook sheet where you can cancel any open orders.
6. Tools have openposition sheet where you can exit any/all running trade.
7. Tools has two option chain and each option chain contains live data with option greeks.
"""

Market_Safety = 1

import warnings
warnings.filterwarnings("ignore")

#from NorenRestApiPy.NorenApi import NorenApi
from Noren import NorenApi
import os
import time
import json
import sys
import platform
from datetime import datetime as dt, timedelta, time, date
from time import sleep
import logging
from threading import Thread
import numpy as np

try:
    import pandas as pd
except (ModuleNotFoundError, ImportError):
    print("pandas module not found")
    os.system(f"{sys.executable} -m pip install -U pandas")
finally:
    import pandas as pd
        
try:
    import pyotp
except (ModuleNotFoundError, ImportError):
    print("pyotp module not found")
    os.system(f"{sys.executable} -m pip install -U pyotp")
finally:
    import pyotp
    
try:
    import xlwings as xw
except (ModuleNotFoundError, ImportError):
    print("xlwings module not found")
    os.system(f"{sys.executable} -m pip install -U xlwings")
finally:
    import xlwings as xw

try:
    import pyttsx3
except (ModuleNotFoundError, ImportError):
    #print("pyttsx3 module not found")
    os.system(f"{sys.executable} -m pip install -U pyttsx3")
finally:
    import pyttsx3

try:
    import requests
except (ModuleNotFoundError, ImportError):
    print("requests module not found")
    os.system(f"{sys.executable} -m pip install -U requests")
finally:
    import requests
    
try:
    from GetIVGreeks import DayCountType, ExpType, TryMatchWith, CalcIvGreeks
except ImportError:
    exec(
        __import__("requests")
        .get(
            "https://gist.githubusercontent.com/ShabbirHasan1"
            + "/7695687d87053c7e3df46810ab2e3046"
            + "/raw/60b2baf44b4801dd91ba583c85076f2605783a4b"
            + "/GetIVGreeks.py"
        )
        .text
    )


print("HI welcome")
    
#logging.basicConfig(level=logging.DEBUG)

df_ins_NFO = pd.DataFrame()

OptionChain_template = []
subs_lst = []
subs_pending_lst = []

Indices_To_check_instrument_Sheet = ['NIFTY','BANKNIFTY',"SENSEX" "SENSEX50", "BANKEX"]
IndexList = ["NIFTY", "BANKNIFTY", "FINNIFTY",'INDIAVIX','MIDCPNIFTY', "SENSEX", "SENSEX50", "BANKEX"]
Token_list = {'NIFTY':26000,'MIDCPNIFTY':26074,'BANKNIFTY':26009,'FINNIFTY':26037,'INDIAVIX':26017, "SENSEX":1, "SENSEX50":47, "BANKEX":12}    

try:
    TerminalSheetName = sys.argv[1]
    if not os.path.exists(TerminalSheetName):
        TerminalSheetName = "Finvasia_Trade_Terminal_v3.xlsm"
except Exception as e:
    TerminalSheetName = "Finvasia_Trade_Terminal_v3.xlsm"
    pass


def Text2Speech(Text):
    #print(Text)
    
    global engine
    try:
        engine = pyttsx3.init()
        voices = engine.getProperty('voices')
        try:
           engine.setProperty('voice', voices[1].id)
        except Exception as e:
            pass
        engine.setProperty('rate', 160)
        engine.say(Text)
        engine.runAndWait()
        del engine
    except Exception as e:
        #print(f"Issue in voice module : {e}")
        pass
    
Text2Speech("Welcome to python trader trade terminal tool, have a nice day. Hope u r enjoying our work")


def Shoonya_login():
    global api
    global TelegramBotCredential, ReceiverTelegramID
    global logger
    global userid, client_name
    client_name = ""
    isConnected = 0
    excel_name = xw.Book(TerminalSheetName)
    Credential_sheet = excel_name.sheets["User_Credential"]
    
    Credential_sheet.range('a1').value = 'Welcome To Python Trader'
    Credential_sheet.range('a1').color = (46,132,198)
    Credential_sheet.range('a14').value = 'Tool developed by PythonTrader, Please follow us on social media site for more tools or freelancing work'
    Credential_sheet.range('a14').color = (46,132,198)
    Credential_sheet.range('b15').value = 'https://www.youtube.com/@pythontraders'
    Credential_sheet.range('b16').value =  'https://www.t.me/pythontrader'
    Credential_sheet.range('b15:b16').color = (220,214,32)
        
    try:
        userid = Credential_sheet.range("B2").value
        
        Timestamp = dt.now().strftime("%d%m%Y_%H%M%S")
        
        subdir = "Logs"

        if not os.path.exists(subdir):
            os.makedirs(subdir)
            
        os_name = platform.system()
        if(os_name == 'Windows'):
            LogFolder = "Logs\\"
        else:
            LogFolder = "Logs/"
            
        LogFile  =  LogFolder + "Finvasia_TT_" + str(userid) + "_" + str(Timestamp)  + str('.log') 

        logging.basicConfig(filename=LogFile, format='%(asctime)s %(message)s', filemode='w')
        logger = logging.getLogger() 
        logger.setLevel(logging.INFO)

        try:
            class ShoonyaApiPy(NorenApi):
                def __init__(self):
                    NorenApi.__init__(self, host='https://api.shoonya.com/NorenWClientTP/', websocket='wss://api.shoonya.com/NorenWSTP/', eodhost='https://api.shoonya.com/chartApi/getdata/')
             
            api = ShoonyaApiPy()
        except Exception as e:
            class ShoonyaApiPy(NorenApi):
                def __init__(self):
                    NorenApi.__init__(self, host='https://api.shoonya.com/NorenWClientTP/', websocket='wss://api.shoonya.com/NorenWSTP/')
            
            api = ShoonyaApiPy()
            pass
                
        
        TelegramBotCredential = str(Credential_sheet.range("b10").value)
        ReceiverTelegramID = str(Credential_sheet.range("b11").value)
        index = ReceiverTelegramID.find(".")
        if index != -1:
            ReceiverTelegramID = ReceiverTelegramID[:len(ReceiverTelegramID)-2]
        
        print(f"TelegramBotCredential= ({TelegramBotCredential}), ReceiverTelegramID=({ReceiverTelegramID})")
        

        password = str(Credential_sheet.range("B3").value)
            
        index = password.find(".")
        if index != -1:
            password = password[:len(password)-2]
        
        LoginMethod = str(Credential_sheet.range("B4").value)
        if (LoginMethod == "New_Session"):
            TotpKey = str(Credential_sheet.range('B5').value)
            index = TotpKey.find(".")
            if index != -1:
                twoFA = int(TotpKey[:6])
            else:
                pin = pyotp.TOTP(TotpKey).now()
                twoFA = f"{int(pin):06d}" if len(pin) <=5 else pin    
                
            vendor_code = Credential_sheet.range("B6").value
            api_secret = Credential_sheet.range("B7").value
            imei = "abcd1234"

            print(
                f"userid={userid},password={password},twoFA={twoFA},vendor_code={vendor_code},api_secret={api_secret}, imei={imei}"
            )
            login_status = api.login(
                userid=userid,
                password=str(password),
                twoFA=str(twoFA),
                vendor_code=vendor_code,
                api_secret=api_secret,
                imei=imei,
            )

            client_name = login_status.get("uname")
            token = login_status.get('susertoken')
            print(login_status)
            Credential_sheet.range("c2").value = "Login Successful, Welcome " + client_name + "\nTool Validity : Demo" + "\nGenerated Token = (" + str(token) + ")"
            isConnected = 1
            Text2Speech("Login Successful, Welcome " + str(client_name) )
            Credential_sheet.range('c2').color = (118,224,280)
        else:
            try:
                ExistingToken = Credential_sheet.range("B8").value
                login_status = api.set_session(userid=userid, password=password,usertoken=ExistingToken)
                get_limits = api.get_limits()
                print(get_limits)
                if(get_limits['stat'] == 'Ok'):
                    Credential_sheet.range("c2").value = "Login Successful \nTool Validity : Demo" + "\nLoggedin using Token = (" + str(ExistingToken) + ")"
                    isConnected = 1
                    Credential_sheet.range('c2').color = (118,224,280)
                    Text2Speech("Login Successful")
                else:
                    Credential_sheet.range("c2").value = "Wrong credential \n" + str(get_limits['emsg']) +"\nPlease login using New_Session"
                    Text2Speech("Login unsuccessful")
                    Credential_sheet.range('c2').color = (255, 0, 0)
            except Exception as e:
                Credential_sheet.range("c2").value = "Wrong credential"
                Text2Speech("Login unsuccessful")
                Credential_sheet.range('c2').color = (255, 0, 0)
                
    except Exception as e:
        print(f"Error : {e}")
        Credential_sheet.range("c2").value = "Wrong credential"
        Text2Speech("Login unsuccessful")
        Credential_sheet.range('c2').color = (255, 0, 0)

    return isConnected

feed_opened = False
SYMBOLDICT = {}
live_data = {}

Telegram_Message = ["Welcome to Python Trader excel based trade terminal","Have a Good Day"]
print(Telegram_Message)
Voice_Message = []
TelegramBotCredential = None
ReceiverTelegramID = None

def convert_to_float(item):
    try:
        return float(item)
    except (ValueError, TypeError):
        return 0

def SendMessageToTelegram(Message):
    #print(f"SendMessageToTelegram = {Message}")
    global TelegramBotCredential, ReceiverTelegramID
    try:
        Url = "https://api.telegram.org/bot" + str(TelegramBotCredential) +  "/sendMessage?chat_id=" + str(ReceiverTelegramID)
        
        textdata ={ "text":Message}
        response = requests.request("POST",Url,params=textdata)
    except Exception as e:
        #print(f"SendMessageToTelegram exception occur : {e}")
        pass

def event_handler_quote_update(inmessage):
    global live_data

    global SYMBOLDICT
    # e   Exchange
    # tk  Token
    # lp  LTP
    # pc  Percentage change
    # v   volume
    # o   Open price
    # h   High price
    # l   Low price
    # c   Close price
    # ap  Average trade price
    # oi  Open interest

    fields = [
        "ts",
        "lp",
        "pc",
        "c",
        "o",
        "h",
        "l",
        "v",
        "ltq",
        "ltp",
        "bp1",
        "sp1",
        "ap",
        "oi",
        "ap",
        "poi",
        "toi",
    ]

    message = {field: inmessage[field] for field in set(fields) & set(inmessage.keys())}
    # print(message)
    key = inmessage["e"] + "|" + inmessage["tk"]

    if key in SYMBOLDICT:
        symbol_info = SYMBOLDICT[key]
        symbol_info.update(message)
        SYMBOLDICT[key] = symbol_info
        live_data[key] = symbol_info
    else:
        SYMBOLDICT[key] = message
        live_data[key] = message

def event_handler_order_update(tick_data):
    print(f"Order update {tick_data}")

def open_callback():
    global feed_opened
    feed_opened = True
    
    #add below lines if issue facing while websocket distruption
    global subs_lst
    subs_lst = []
    
def event_handler_socket_closed():
    print(f"socket closed, so again trying to reconnect")
    sleep(2)
    
def order_status(orderid):
    #print(f"Checking order_status for ({orderid})")
    AverageExecutedPrice = 0
    status = ''
    try:
        order_book = get_order_book()
        order_book = order_book[order_book['Order No'] == str(orderid)]

        status = order_book.iloc[0]["Status"]
        if status == "COMPLETE":
            AverageExecutedPrice = order_book.iloc[0]["Executed Price"]
    except Exception as e:
        Message = str(e) + " : Exception occur in order_status"
        print(Message)
    return status, AverageExecutedPrice

def place_trade(symbol, quantity, buy_or_sell, order_type = None, price = None):
    global api
    global Product_type
    global logger
    tradingsymbol = symbol[4:]
    exchange = symbol[:3]
    
    if order_type == 'MARKET':
        price = 0
        trigger_price = None
        price_type = "MKT"
    elif order_type == 'LIMIT':
        trigger_price = None
        price_type = "LMT"
    elif order_type == 'SL-M':
        price_type = "SL-LMT"
        trigger_price = price
        
        if buy_or_sell == 'BUY':
            price = price * (100 + Market_Safety) / 100
            if exchange == 'CDS':
                round_to = .0025
                result = ( int(price / round_to) + 1) * round_to
                price = float( "{:.4f}".format(result) )
            else:
                round_to = .1
                result = ( int(price / round_to) + 1) * round_to
                price = float( "{:.1f}".format(result) )
        else:
            price = price * (100 - Market_Safety) / 100
            if exchange == 'CDS':
                round_to = .0025
                result = ( int(price / round_to) ) * round_to
                price = float( "{:.4f}".format(result) )
            else:
                round_to = .1
                result = ( int(price / round_to) ) * round_to
                price = float( "{:.1f}".format(result) )
        
        
    if(Product_type == 'MIS'):
            product = "I"
    else:
    
        if exchange in ["NSE","BSE"]:
            product = "C"
        else:
            product = "M"
    Message = "Order Details : buy_or_sell = " + str(buy_or_sell[0]) + ", product_type = " + str(product)+ ", exchange = " + str(exchange)+ ", tradingsymbol = " + str(tradingsymbol)+ ", quantity = " + str(quantity)+ ", price_type = " + str(price_type)+ ", price = " + str(price)+ ", trigger_price = " + str(trigger_price)
    logger.info(Message)
    
    order_id = api.place_order(
        buy_or_sell=buy_or_sell[0],
        product_type=product,
        exchange=exchange,
        tradingsymbol=tradingsymbol,
        quantity=quantity,
        discloseqty=0,
        price_type=price_type,
        price=price,
        trigger_price=trigger_price,
        retention="DAY",
        remarks="Python_Trader",
    ).get("norenordno")

    Message = "Order placed for " + str (tradingsymbol) + " " + str(quantity) + " quantity " + str(buy_or_sell) + ", ur order id :" + str(order_id)
    print(Message)
    logger.info(Message)
    Telegram_Message.append(Message)
    Voice_Message.append(Message)
    
    print(f"Order id = {order_id}")

    return order_id

def get_order_book():
    try:
        global api
        order_book = api.get_order_book()
        df_order_book = pd.DataFrame(order_book)
        if(len(df_order_book) > 0):
            df_order_book = df_order_book.sort_values(by=["norenordno"]).reset_index(drop=True)
            try:
                df_order_book = df_order_book[['norenordno' ,'status','exch','tsym','prctyp','trantype','qty','fillshares','prc','avgprc','prd','token','ls','remarks','rejreason','trgprc']]
                df_order_book = df_order_book.rename(columns={'norenordno' : 'Order No','exch' : 'Exchange','tsym' : 'Trading Symbol','qty' : 'Quantity','trantype' : 'Transaction Type','prctyp' : 'Product Type','trgprc':'Trigger Price','token' : 'Token','ls' : 'Lot Size','status' : 'Status','remarks' : 'Remarks','rejreason' : 'Order rejection reason','prd':'Order Type','avgprc':'Executed Price','fillshares':'Filled Shares','prc':'Price'})
            except Exception as e:
                try:
                    df_order_book = df_order_book[['norenordno' ,'status','exch','tsym','prctyp','trantype','qty','fillshares','prc','avgprc','prd','token','ls','remarks','trgprc']]
                    df_order_book = df_order_book.rename(columns={'norenordno' : 'Order No','exch' : 'Exchange','tsym' : 'Trading Symbol','qty' : 'Quantity','trantype' : 'Transaction Type','prctyp' : 'Product Type','trgprc':'Trigger Price','token' : 'Token','ls' : 'Lot Size','status' : 'Status','remarks' : 'Remarks','prd':'Order Type','avgprc':'Executed Price','fillshares':'Filled Shares','prc':'Price'})
                except Exception as e:
                    try:
                        df_order_book = df_order_book[['norenordno' ,'status','exch','tsym','prctyp','trantype','qty','prc','prd','token','ls','remarks','avgprc']]
                        df_order_book = df_order_book.rename(columns={'norenordno' : 'Order No','exch' : 'Exchange','tsym' : 'Trading Symbol','qty' : 'Quantity','trantype' : 'Transaction Type','prctyp' : 'Product Type','token' : 'Token','ls' : 'Lot Size','status' : 'Status','remarks' : 'Remarks','prd':'Order Type','prc':'Price','avgprc':'Executed Price'})
                    except Exception as e:
                        df_order_book = df_order_book[['norenordno' ,'status','exch','tsym','prctyp','trantype','qty','prc','prd','token','ls','remarks']]
                        df_order_book = df_order_book.rename(columns={'norenordno' : 'Order No','exch' : 'Exchange','tsym' : 'Trading Symbol','qty' : 'Quantity','trantype' : 'Transaction Type','prctyp' : 'Product Type','token' : 'Token','ls' : 'Lot Size','status' : 'Status','remarks' : 'Remarks','prd':'Order Type','prc':'Price'})
        #print(df_order_book)
    except Exception as e:
        print(f"Exception occur in get_order_book : {e}")

    return df_order_book

def GetToken_UsingSymbol(exchange, tradingsymbol):
    global df_ins_NSE, df_ins_BSE, df_ins_NFO, df_ins_CDS, df_ins_MCX, df_ins_BFO
    global api
    try:
        if exchange == 'NSE':
            df_ins_temp = df_ins_NSE[df_ins_NSE.TradingSymbol == tradingsymbol]
            Token = df_ins_temp.iloc[0]['Token']
        elif exchange == 'BSE':
            df_ins_temp = df_ins_BSE[df_ins_BSE.TradingSymbol == tradingsymbol]
            Token = df_ins_temp.iloc[0]['Token']
        elif exchange == 'NFO':
            df_ins_temp = df_ins_NFO[df_ins_NFO.TradingSymbol == tradingsymbol]
            Token = df_ins_temp.iloc[0]['Token']
        elif exchange == 'BFO':
            df_ins_temp = df_ins_BFO[df_ins_BFO.TradingSymbol == tradingsymbol]
            Token = df_ins_temp.iloc[0]['Token']
        elif exchange == 'CDS':
            df_ins_temp = df_ins_CDS[df_ins_CDS.TradingSymbol == tradingsymbol]
            Token = df_ins_temp.iloc[0]['Token']
        elif exchange == 'MCX':
            df_ins_temp = df_ins_MCX[df_ins_MCX.TradingSymbol == tradingsymbol]
            Token = df_ins_temp.iloc[0]['Token']
        #print(f"Token = {Token}")
    except Exception as e:
        #print(f"Exception : {e}")
        Token = api.searchscrip(exchange=exchange, searchtext=tradingsymbol).get("values")[0].get("token")
        #print(f"latest Token = {Token}")
        pass
    return Token

def subscribe_new_token(exchange, Token):
    global api
    symbol = []
    symbol.append(f"{exchange}|{Token}")
    api.subscribe(symbol)

LimitOrderBook = {}
def start_Trade_Terminal():
    #print("I am inside start_Trade_Terminal")
    global api, live_data
    global SYMBOLDICT
    global Product_type
    global Trade_Mode
    global Telegram_Message, Voice_Message    
    global subs_lst
    excel_TT = xw.Book(TerminalSheetName)
    tt = excel_TT.sheets("Trade_Terminal")
    tt.range("a2:d2").value  = 0
    tt.range("q4:s1000").value  = None
    tt.range("u4:w1000").value  = None
    tt.range("aa4:ac1000").value  = None
    tt.range(f"a3:ac3").value = [ "Symbol", "Open", "High", "Low", "Close", "VWAP", "Best Buy Price",
                                "Best Sell Price","Volume","OI", "LTP","Percentage change", "Qty", "BUY/SELL", "Entry Signal","Entry Limit Price", "Entry Done @","Entry Order ID", "Entry Remarks","Exit Signal","Exit Done @","Exit Order ID","Exit Remarks", "Target","SL" ,"Trail Enable",    "Latest SL","Trade Status","PnL"]
                                
    tt.range('k1').value =  'PYTHON TRADER'
    tt.range('k1').color = (46,132,198)  
    Trade_Mode = str(tt.range('s2').value).upper()
    
    AlertMessage = Trade_Mode + " trade mode enabled"
    Telegram_Message.append(AlertMessage)
    Voice_Message.append(AlertMessage)    
    
    Symbol_Token = {}
    global LimitOrderBook
    #run a parallel thread to update the status
    while True:
        try:
            Product_type = tt.range(f"P2").value
            symbols = tt.range(f"a{4}:a{1000}").value
            trading_info = tt.range(f"m{4}:ac{1000}").value
            main_list = []

            idx = 0
            for i in symbols:
                lst = [None, None, None, None,None, None, None, None, None,None,None]
                if i:
                    if i not in subs_lst:
                        subs_lst.append(i)
                        try:
                            exchange = i[:3]
                            tradingsymbol = i[4:]
                            Token = GetToken_UsingSymbol(exchange, tradingsymbol)
                            Symbol_Token[i] = exchange + "|" + str(Token)
                            subscribe_new_token(exchange, Token)
                            print(f"Symbol = {i}, Token={Token} subscribed")

                        except Exception as e:
                            print(f"Subscribe error {i} : {e}")
                    if i in subs_lst:
                        try:
                            TokenKey = Symbol_Token[i]

                            lst = [
                                live_data[TokenKey].get("o", "-"),
                                live_data[TokenKey].get("h", "-"),
                                live_data[TokenKey].get("l", "-"),
                                live_data[TokenKey].get("c", "-"),
                                live_data[TokenKey].get("ap", "-"),
                                live_data[TokenKey].get("bp1", "-"),
                                live_data[TokenKey].get("sp1", "-"),
                                live_data[TokenKey].get("v", "-"),
                                live_data[TokenKey].get("oi", "-"),
                                live_data[TokenKey].get("lp", "-"),
                                live_data[TokenKey].get("pc", "-"),
                            ]

                            try:
                                trade_info = trading_info[idx]
                                idx_location = idx + 2
                                #print(f" {i} : {trade_info}")
                                if trade_info[0] is not None and trade_info[1] is not None:
                                    if type(trade_info[0]) is float and type(trade_info[1]) is str:
                                        
                                        LTP = float(live_data[TokenKey].get("lp", 0))
                                        
                                        if Trade_Mode == 'REAL':
                                            #Real trade mode handling will handle here
                                            if trade_info[1].upper() == "BUY" and LTP != 0:
                                                if trade_info[2] in ['True_Market' ,'True_Limit_LTP', 'Limit_Below', 'Limit_Above']:
                                                    
                                                    
                                                    #0:Qty    1:BUY/SELL    2:Entry_Signal  3:Entry_Limit_Price  4:Entry_Done@  5:Entry_Order_ID 6: Entry_Remarks  7:Exit_Signal 8:Exit_Done @ 9:Exit_Order_ID  10: Exit_Remarks 11:Target    12:SL  13:Trail Enable 14: Latest_SL 15:Trade_status 16: PnL
                                                    
            
                                                    if trade_info[15] != 'Active' and trade_info[15] != 'Entry_Pending' and trade_info[15] != 'Exit_Pending' and trade_info[15] != 'Closed' and (trade_info[15] is None or trade_info[15] == ''):
                                                        
                                                        
                                                        if trade_info[2] == 'True_Market':
                                                            #Entry buy trade immediately
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[1]) + ", " + str(trade_info[2])+ ", " + str(trade_info[3])+ " triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","MARKET")
                                                            
                                                            if order_id is None:
                                                                tt.range(f"O{idx_location + 2}").value = None
                                                            else:    
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"q{idx_location + 2}").value = None
                                                                tt.range(f"r{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Entry_Pending'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None
                                                            
                                                        elif trade_info[2] == 'True_Limit_LTP':    
                                                            #Entry buy trade immediately at ltp price
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[1]) + ", " + str(trade_info[2])+ ", " + str(trade_info[3])+ " triggered"
                                                            logger.info(Message)
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","LIMIT",LTP)
                                                            if order_id is None:
                                                                tt.range(f"O{idx_location + 2}").value = None
                                                            else: 
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"q{idx_location + 2}").value = None
                                                                tt.range(f"r{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Entry_Pending'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None
                                                            
                                                        elif trade_info[2] == 'Limit_Below' and trade_info[3] is not None:
                                                            #print(f"LTP={LTP}, trade_info[3] ={trade_info[3]}")
                                                            
                                                            #handling real Limit_Below based trade
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[1]) + ", " + str(trade_info[2])+ ", " + str(trade_info[3])+ " triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","LIMIT",trade_info[3])
                                                            if order_id is None:
                                                                tt.range(f"O{idx_location + 2}").value = None
                                                            else: 
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"q{idx_location + 2}").value = None
                                                                tt.range(f"r{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Entry_Pending'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None
                                                           
                                                            
                                                        elif trade_info[2] == 'Limit_Above' and trade_info[3] is not None:
                                                            #print(f"LTP={LTP}, trade_info[3] ={trade_info[3]}")
                                                            
                                                            #handling real Limit_Above based trade
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[1]) + ", " + str(trade_info[2])+ ", " + str(trade_info[3])+ " triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","SL-M",trade_info[3])
                                                            if order_id is None:
                                                                tt.range(f"O{idx_location + 2}").value = None
                                                            else: 
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"q{idx_location + 2}").value = None
                                                                tt.range(f"r{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Entry_Pending'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None
                                                            
                                                    elif trade_info[15] == 'Entry_Pending' :
                                                        Order_id = str(int(trade_info[5]))
                                                        #print(f"Order_id={Order_id}")
                                                        status = LimitOrderBook[Order_id]['status']
                                                        if status == 'COMPLETE':
                                                            Executed_price = LimitOrderBook[Order_id]['Executed_price']
                                                            tt.range(f"q{idx_location + 2}").value = Executed_price
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Active'
                                                        
                                                            Message = str(int(trade_info[0])) + " quantity " + str(i[4:]) + " Bought at " + str(Executed_price)
                                                            
                                                            Telegram_Message.append(Message)
                                                            Voice_Message.append(Message)
                        
                        
                                                        tt.range(f"S{idx_location + 2}").value = LimitOrderBook[Order_id]['Remarks']
                                                        
                                                    elif trade_info[15] == 'Active':
                                                        
                                                        
                                                        PnL = (float(LTP) - float(trade_info[4])) * int(trade_info[0])
                                                        tt.range(f"AC{idx_location + 2}").value = PnL
                                                        #print(f"PnL = {PnL}")
                                                        
                                                        TSL = 0
                                                        if type(trade_info[12]) is float:
                                                            if trade_info[13] == True:
                                                                if(LTP >= trade_info[4]):
                                                                    CUR_TSL = LTP - trade_info[4] + trade_info[12]
                                                                else:
                                                                    CUR_TSL = trade_info[12]
                                                                
                                                                OLD_TSL = trade_info[14]
                                                                if OLD_TSL is not None:
                                                                    TSL = max(CUR_TSL,OLD_TSL)
                                                                else:
                                                                    TSL = CUR_TSL 
                                                            else:
                                                                TSL = trade_info[12]
                                                            
                                                            tt.range(f"AA{idx_location + 2}").value = TSL
                                                        
                                                        if trade_info[7] == 'True_Market':
                                                            #exit buy order immediately
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[7]) +  " exit triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","MARKET")
                                                            if order_id is None:
                                                                tt.range(f"T{idx_location + 2}").value = None
                                                            else:  
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"v{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Exit_Pending'
                                                                
                                                        elif trade_info[7] == 'True_Limit_LTP':
                                                            #exit buy at ltp limit 
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[7]) +  " exit triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",LTP)
                                                            if order_id is None:
                                                                tt.range(f"T{idx_location + 2}").value = None
                                                            else:  
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"v{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Exit_Pending'
                                                            
                                                        elif type(trade_info[11]) is float and trade_info[11] <= LTP:
                                                            #target meets, so exit the buy order
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", target meets so exit triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","MARKET")
                                                            
                                                            LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                            tt.range(f"v{idx_location + 2}").value = order_id
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Exit_Pending'
                                                        
                                                        elif TSL >= LTP and type(trade_info[12]) is float:
                                                            #sl meets, so exiting the buy order
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", TSL meets so exit triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","MARKET")
                                                            
                                                            LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                            
                                                            tt.range(f"v{idx_location + 2}").value = order_id
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Exit_Pending'
                                                            
                                                    
                                                    elif trade_info[15] ==  'Exit_Pending':
                                                        Order_id = str(int(trade_info[9]))
                                                        #print(f"Order_id={Order_id}")
                                                        status = LimitOrderBook[Order_id]['status']
                                                        if status == 'COMPLETE':
                                                            Executed_price = LimitOrderBook[Order_id]['Executed_price']
                                                            tt.range(f"u{idx_location + 2}").value = Executed_price
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Closed'
                                                            
                                                            PnL = (float(Executed_price) - float(trade_info[4])) * int(trade_info[0])
                                                            tt.range(f"AC{idx_location + 2}").value = PnL
                                                            
                                                            Message = str(int(trade_info[0])) + " quantity " + str(i[4:]) + " Sold at " + str(Executed_price) + ", Trade profit or Loss = " + str(round(PnL,2))
                                                            
                                                            Telegram_Message.append(Message)
                                                            Voice_Message.append(Message)
                                                            
                                                        tt.range(f"W{idx_location + 2}").value = LimitOrderBook[Order_id]['Remarks']
                                                        
                                            if trade_info[1].upper() == "SELL" and LTP != 0:
                                                
                                                if trade_info[2] in ['True_Market' , 'True_Limit_LTP' , 'Limit_Below' , 'Limit_Above']:
                                                    #0:Qty    1:BUY/SELL    2:Entry_Signal  3:Entry_Limit_Price  4:Entry_Done@  5:Entry_Order_ID 6: Entry_Remarks  7:Exit_Signal 8:Exit_Done @ 9:Exit_Order_ID   10:Target    11:SL  12:Trail Enable 13: Latest_SL 14:Trade_status 15: PnL
                                                    
                                                    
                                                    if trade_info[15] != 'Active' and trade_info[15] != 'Entry_Pending' and trade_info[15] != 'Exit_Pending' and trade_info[15] != 'Closed' and (trade_info[15] is None or trade_info[15] == '' ):
                                                
                                                        if trade_info[2] == 'True_Market':
                                                            #entry sell order immediately
                                                            
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[1]) + ", " + str(trade_info[2])+ ", " + str(trade_info[3])+ " triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","MARKET")
                                                            if order_id is None:
                                                                tt.range(f"O{idx_location + 2}").value = None
                                                            else: 
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                            
                                                                tt.range(f"q{idx_location + 2}").value = None
                                                                tt.range(f"r{idx_location + 2}").value = order_id
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Entry_Pending'
                                                                tt.range(f"AC{idx_location + 2}").value = None
                                                                #print("b")
                                                            
                                                        elif trade_info[2] == 'True_Limit_LTP':    
                                                            #Entry buy trade immediately at ltp price
                                                            
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[1]) + ", " + str(trade_info[2])+ ", " + str(trade_info[3])+ " triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",LTP)
                                                            if order_id is None:
                                                                tt.range(f"O{idx_location + 2}").value = None
                                                            else: 
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"q{idx_location + 2}").value = None
                                                                tt.range(f"r{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Entry_Pending'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None
                                                                
                                                        elif trade_info[2] == 'Limit_Above' and trade_info[3] is not None:
                                                            
                                                            #handling real Limit_Above based trade
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[1]) + ", " + str(trade_info[2])+ ", " + str(trade_info[3])+ " triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",trade_info[3])
                                                            if order_id is None:
                                                                tt.range(f"O{idx_location + 2}").value = None
                                                            else: 
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"q{idx_location + 2}").value = None
                                                                tt.range(f"r{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Entry_Pending'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None
                                                            
                                                                    
                                                        elif trade_info[2] == 'Limit_Below' and trade_info[3] is not None:
                                                           
                                                            #handling real Limit_Above based trade
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[1]) + ", " + str(trade_info[2])+ ", " + str(trade_info[3])+ " triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","SL-M",trade_info[3])
                                                            
                                                            if order_id is None:
                                                                tt.range(f"O{idx_location + 2}").value = None
                                                            else: 
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"q{idx_location + 2}").value = None
                                                                tt.range(f"r{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Entry_Pending'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None
                                                            
                                                        
                                                    elif trade_info[15] == 'Entry_Pending' :
                                                        Order_id = str(int(trade_info[5]))
                                                        #print(f"Order_id={Order_id}")
                                                        status = LimitOrderBook[Order_id]['status']
                                                        if status == 'COMPLETE':
                                                            Executed_price = LimitOrderBook[Order_id]['Executed_price']
                                                            tt.range(f"q{idx_location + 2}").value = Executed_price
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Active'
                                                            
                                                            Message = str(int(trade_info[0])) + " quantity " + str(i[4:]) + " Sold at " + str(Executed_price) 
                                                            
                                                            Telegram_Message.append(Message)
                                                            Voice_Message.append(Message)
                                                            
                                                        tt.range(f"S{idx_location + 2}").value = LimitOrderBook[Order_id]['Remarks']
                                                        
                                                    elif trade_info[15] == 'Active':
                                                        
                                                        PnL = (float(trade_info[4]) - float(LTP) ) * int(trade_info[0])
                                                        tt.range(f"AC{idx_location + 2}").value = PnL
                                                        
                                                        TSL = 9999999
                                                        
                                                        if type(trade_info[12]) is float:
                                                            if trade_info[13] == True :
                                                                if(LTP <= trade_info[4]):
                                                                    CUR_TSL = trade_info[12] - (trade_info[4] - LTP )
                                                                else:
                                                                    CUR_TSL = trade_info[12]
                                                                
                                                                OLD_TSL = trade_info[14]
                                                                if OLD_TSL is not None:
                                                                    TSL = min(CUR_TSL,OLD_TSL)
                                                                else:
                                                                    TSL = CUR_TSL 
                                                            else:
                                                                TSL = trade_info[12]
                                                            
                                                            tt.range(f"AA{idx_location + 2}").value = TSL
                                                            
                                                        
                                                        if trade_info[7] == 'True_Market':
                                                            #exit sell order immediately
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[7]) +  " exit triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","MARKET")
                                                            if order_id is None:
                                                                tt.range(f"T{idx_location + 2}").value = None
                                                            else:  
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                    
                                                                tt.range(f"v{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = "Exit_Pending"
                                                                
                                                        elif trade_info[7] == 'True_Limit_LTP':
                                                            #exit SELL at ltp limit 
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", " + str(trade_info[7]) +  " exit triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","LIMIT",LTP)
                                                            
                                                            if order_id is None:
                                                                tt.range(f"T{idx_location + 2}").value = None
                                                            else:  
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"v{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Exit_Pending'    
                                                                
                                                        elif type(trade_info[11]) is float and trade_info[11] >= LTP:
                                                            #target meets, so exiting the sell order
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", target meets so exit triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","MARKET")
                                                            
                                                            LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                            
                                                            tt.range(f"v{idx_location + 2}").value = order_id
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Exit_Pending'
                                                            
                                                        
                                                        elif TSL <= LTP and type(trade_info[12]) is float:
                                                            #sl hit, so exiting the sell order
                                                            Message =  str(i) + " : " +  str(trade_info[0]) + ", TSL meets so exit triggered"
                                                            logger.info(Message)
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","MARKET")
                                                            
                                                            if order_id is None:
                                                                tt.range(f"O{idx_location + 2}").value = None
                                                            else: 
                                                            
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                
                                                                tt.range(f"v{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Exit_Pending'
                                                            
                                                    elif trade_info[15] ==  'Exit_Pending':
                                                        Order_id = str(int(trade_info[9]))
                                                        #print(f"Order_id={Order_id}")
                                                        status = LimitOrderBook[Order_id]['status']
                                                        if status == 'COMPLETE':
                                                            Executed_price = LimitOrderBook[Order_id]['Executed_price']
                                                            tt.range(f"u{idx_location + 2}").value = Executed_price
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Closed'
                                                            
                                                            PnL = (float(trade_info[4]) - float(Executed_price)) * int(trade_info[0])
                                                            tt.range(f"AC{idx_location + 2}").value = PnL
                                                            
                                                            Message = str(int(trade_info[0])) + " quantity " + str(i[4:]) + " buyback at " + str(Executed_price) + ", Trade profit or Loss = " + str(round(PnL,2))
                                                            
                                                            Telegram_Message.append(Message)
                                                            Voice_Message.append(Message)
                                                            
                                                        tt.range(f"W{idx_location + 2}").value = LimitOrderBook[Order_id]['Remarks']
                                        else:
                                            #paper trade mode handling will handle here
                                            if trade_info[1].upper() == "BUY" and LTP != 0:
                                                if trade_info[2]  in ['True_Market' ,'True_Limit_LTP' , 'Limit_Below','Limit_Above']:
                                                    
                                                   
                                                    
                                                    #0:Qty    1:BUY/SELL    2:Entry_Signal  3:Entry_Limit_Price  4:Entry_Done@  5:Entry_Order_ID 6: Entry_Remarks  7:Exit_Signal 8:Exit_Done @ 9:Exit_Order_ID   10:Target    11:SL  12:Trail Enable 13: Latest_SL 14:Trade_status 15: PnL
                                                    
            
                                                    if trade_info[15] != 'Active' and trade_info[15] != 'Exit_Pending' and trade_info[15] != 'Closed' and (trade_info[15] is None or trade_info[15] == ''or trade_info[15] == 'Entry_Pending' ):
                                                        
                                                        
                                                        if trade_info[2] in[ 'True_Market', 'True_Limit_LTP']:
                                                            #Entry buy trade immediately
                                                            order_id = 'PAPER'
                                                            Executed_price = LTP
                                                                
                                                            tt.range(f"q{idx_location + 2}").value = Executed_price
                                                            tt.range(f"r{idx_location + 2}").value = order_id
                                                            tt.range(f"u{idx_location + 2}").value = None
                                                            tt.range(f"v{idx_location + 2}").value = None
                                                            tt.range(f"AA{idx_location + 2}").value = None
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Active'
                                                            tt.range(f"AC{idx_location + 2}").value = None
                                                        
                                                        elif trade_info[2] == 'Limit_Below' and trade_info[3] is not None:
                                                            #print(f"LTP={LTP}, trade_info[3] ={trade_info[3]}")
                                                            
                                                            #handling paper Limit_Below based trade
                                                            if LTP <= trade_info[3]:
                                                                #Entry limit below buy
                                                                
                                                                order_id = 'PAPER'
                                                                Executed_price = LTP
                                                                    
                                                                tt.range(f"q{idx_location + 2}").value = Executed_price
                                                                tt.range(f"r{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Active'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None

                                                            elif trade_info[15] != 'Entry_Pending':
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Entry_Pending'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None
                                                        
                                                        elif trade_info[2] == 'Limit_Above' and trade_info[3] is not None:
                                                            #print(f"LTP={LTP}, trade_info[3] ={trade_info[3]}")
                                                            
                                                            if LTP >= trade_info[3]:
                                                                #Entry limit above buy
                                                                order_id = 'PAPER'
                                                                Executed_price = LTP
                                                                    
                                                                tt.range(f"q{idx_location + 2}").value = Executed_price
                                                                tt.range(f"r{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Active'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None

                                                            elif trade_info[15] != 'Entry_Pending':
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Entry_Pending'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None
                                                                                                            
                                                                
                                                    elif trade_info[15] == 'Active' or trade_info[15] == 'Exit_Pending':
                                                        
                                                        
                                                        PnL = (float(LTP) - float(trade_info[4])) * int(trade_info[0])
                                                        tt.range(f"AC{idx_location + 2}").value = PnL
                                                        #print(f"PnL = {PnL}")
                                                        
                                                        TSL = 0
                                                        if type(trade_info[12]) is float:
                                                            if trade_info[13] == True:
                                                                if(LTP >= trade_info[4]):
                                                                    CUR_TSL = LTP - trade_info[4] + trade_info[12]
                                                                else:
                                                                    CUR_TSL = trade_info[12]
                                                                
                                                                OLD_TSL = trade_info[14]
                                                                if OLD_TSL is not None:
                                                                    TSL = max(CUR_TSL,OLD_TSL)
                                                                else:
                                                                    TSL = CUR_TSL 
                                                            else:
                                                                TSL = trade_info[12]
                                                            
                                                            tt.range(f"AA{idx_location + 2}").value = TSL
                                                        
                                                        if trade_info[7] in  ['True_Market','True_Limit_LTP']:
                                                            #exit buy order immediately
                                                            order_id = 'PAPER'
                                                            Executed_price = LTP
                                                                
                                                            tt.range(f"u{idx_location + 2}").value = Executed_price
                                                            tt.range(f"v{idx_location + 2}").value = order_id
                                                            tt.range(f"v{idx_location + 2}").value = order_id
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Closed'
                                                            
                                                            PnL = (float(Executed_price) - float(trade_info[4])) * int(trade_info[0])
                                                            tt.range(f"AC{idx_location + 2}").value = PnL
                                                        
                                                        
                                                        elif type(trade_info[11]) is float and trade_info[11] <= LTP:
                                                            #target meets, so exit the buy order
                                                            order_id = 'PAPER'
                                                            Executed_price = LTP
                                                                
                                                            tt.range(f"u{idx_location + 2}").value = Executed_price
                                                            tt.range(f"v{idx_location + 2}").value = order_id
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Closed'
                                                            
                                                            PnL = (float(Executed_price) - float(trade_info[4])) * int(trade_info[0])
                                                            tt.range(f"AC{idx_location + 2}").value = PnL
                                                        
                                                        elif TSL >= LTP and type(trade_info[12]) is float:
                                                            #sl meets, so exiting the buy order
                                                            order_id = 'PAPER'
                                                            Executed_price = LTP
                                                                
                                                            tt.range(f"u{idx_location + 2}").value = Executed_price
                                                            tt.range(f"v{idx_location + 2}").value = order_id
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Closed'
                                                            
                                            if trade_info[1].upper() == "SELL" and LTP != 0:
                                                
                                                if trade_info[2] in ['True_Market','True_Limit_LTP',  'Limit_Below', 'Limit_Above']:
                                                    #0:Qty    1:BUY/SELL    2:Entry_Signal  3:Entry_Limit_Price  4:Entry_Done@  5:Entry_Order_ID 6: Entry_Remarks  7:Exit_Signal 8:Exit_Done @ 9:Exit_Order_ID   10:Target    11:SL  12:Trail Enable 13: Latest_SL 14:Trade_status 15: PnL
                                                    
                                                    
                                                    if trade_info[15] != 'Active' and trade_info[15] != 'Exit_Pending' and trade_info[15] != 'Closed' and (trade_info[15] is None or trade_info[15] == '' or trade_info[15] == 'Entry_Pending'):
                                                
                                                        if trade_info[2] in ['True_Market', 'True_Limit_LTP']:
                                                            #entry sell order immediately
                                                            order_id = 'PAPER'
                                                            Executed_price = LTP
                                                                
                                                            tt.range(f"q{idx_location + 2}").value = Executed_price
                                                            tt.range(f"r{idx_location + 2}").value = order_id
                                                            tt.range(f"u{idx_location + 2}").value = None
                                                            tt.range(f"v{idx_location + 2}").value = None
                                                            tt.range(f"AA{idx_location + 2}").value = None
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Active'
                                                            tt.range(f"AC{idx_location + 2}").value = None
                                                            #print("b")
                                                            
                                                        elif trade_info[2] == 'Limit_Above' and trade_info[3] is not None:
                                                            
                                                            
                                                            if LTP >= trade_info[3]:
                                                                #entry sell order with Limit_Above
                                                                order_id = 'PAPER'
                                                                Executed_price = LTP
                                                                    
                                                                tt.range(f"q{idx_location + 2}").value = Executed_price
                                                                tt.range(f"r{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Active'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None
                                                                
                                                            elif trade_info[15] != 'Entry_Pending':
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Entry_Pending'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None 
                                                                
                                                        elif trade_info[2] == 'Limit_Below' and trade_info[3] is not None:
                                                            
                                                            
                                                            if LTP <= trade_info[3]:
                                                                #entry sell order with Limit_Below
                                                                
                                                                order_id = 'PAPER'
                                                                Executed_price = LTP
                                                                    
                                                                tt.range(f"q{idx_location + 2}").value = Executed_price
                                                                tt.range(f"r{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Active'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None
                                                                
                                                            elif trade_info[15] != 'Entry_Pending':
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Entry_Pending'
                                                                tt.range(f"u{idx_location + 2}").value = None
                                                                tt.range(f"v{idx_location + 2}").value = None
                                                                tt.range(f"AA{idx_location + 2}").value = None
                                                                tt.range(f"AC{idx_location + 2}").value = None    
                                                        
                                                    elif trade_info[15] == 'Active' or trade_info[15] == 'Exit_Pending':
                                                        
                                                        PnL = (float(trade_info[4]) - float(LTP) ) * int(trade_info[0])
                                                        tt.range(f"AC{idx_location + 2}").value = PnL
                                                        
                                                        TSL = 9999999
                                                        
                                                        if type(trade_info[12]) is float:
                                                            if trade_info[13] == True :
                                                                if(LTP <= trade_info[4]):
                                                                    CUR_TSL = trade_info[12] - (trade_info[4] - LTP )
                                                                else:
                                                                    CUR_TSL = trade_info[12]
                                                                
                                                                OLD_TSL = trade_info[14]
                                                                if OLD_TSL is not None:
                                                                    TSL = min(CUR_TSL,OLD_TSL)
                                                                else:
                                                                    TSL = CUR_TSL 
                                                            else:
                                                                TSL = trade_info[12]
                                                            
                                                            tt.range(f"AA{idx_location + 2}").value = TSL
                                                            
                                                        
                                                        if trade_info[7] in  ['True_Market','True_Limit_LTP']:
                                                            #exit sell order immediately
                                                            order_id = 'PAPER'
                                                            Executed_price = LTP
                                                                
                                                            tt.range(f"u{idx_location + 2}").value = Executed_price
                                                            tt.range(f"v{idx_location + 2}").value = order_id
                                                            tt.range(f"Ab{idx_location + 2}").value = "Closed"
                                                            
                                                            PnL = (float(trade_info[4]) - float(Executed_price) ) * int(trade_info[0])
                                                            tt.range(f"AC{idx_location + 2}").value = PnL
                                                        
                                                            
                                                        elif type(trade_info[11]) is float and trade_info[11] >= LTP:
                                                            #target meets, so exiting the sell order
                                                            order_id = 'PAPER'
                                                            Executed_price = LTP
                                                                
                                                            tt.range(f"u{idx_location + 2}").value = Executed_price
                                                            tt.range(f"v{idx_location + 2}").value = order_id
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Closed'
                                                            
                                                            PnL = (float(trade_info[4]) - float(Executed_price) ) * int(trade_info[0])
                                                            tt.range(f"AC{idx_location + 2}").value = PnL
                                                        
                                                        elif TSL <= LTP and type(trade_info[12]) is float:
                                                            #sl hit, so exiting the sell order
                                                            order_id = 'PAPER'
                                                            Executed_price = LTP    
                                                                
                                                            tt.range(f"u{idx_location + 2}").value = Executed_price
                                                            tt.range(f"v{idx_location + 2}").value = order_id
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Closed'
                                                            
                                                            PnL = (float(trade_info[4]) - float(Executed_price) ) * int(trade_info[0])
                                                            tt.range(f"AC{idx_location + 2}").value = PnL
                                
                            except Exception as e:
                                Message = "Exception occur in core order book management:" + str(e)
                                #print(Message)
                                logger.info(Message)
                                pass
                        except Exception as e:
                            #print(f"exception : {str(e)}")
                            pass
                main_list.append(lst)

                idx += 1

            tt.range("b4:l1000").value = main_list
                
        except Exception as e:
            #print(f"Exception : {str(e)}")
            pass

def get_position():
    #print(f"I am inside get_position")
    global api
    try:
        positions = api.get_positions()
        df_positions = pd.DataFrame(positions)
        #df_positions.to_csv("positions.csv")
        #print(df_positions)
        mtm = 0
        pnl = 0
        day_m2m = 0
        if len(df_positions) > 0:
            df_positions = df_positions.sort_values(by=["tsym"]).reset_index(drop=True)
            for i in positions:
                mtm += float(i['urmtom'])
                pnl += float(i['rpnl'])
                day_m2m = mtm + pnl
            #print(f'{day_m2m} is your Daily MTM')
        
            df_positions = df_positions[['exch','tsym','prd','netqty','netavgprc','lp','daybuyqty','daysellqty','daybuyavgprc','daysellavgprc','openbuyqty','opensellqty']]
        
            df_positions = df_positions.rename(columns={'exch':'Exchange' ,'tsym':'Symbol' ,'prd':'Product' ,'netqty':'Net Quantity' ,'netavgprc':'Avg Price' ,'lp':'Last Price' ,'daybuyqty':'Buy Quantity' ,'daysellqty':'Sell Quantity' ,'daybuyavgprc':'Avg Buy Price' ,'daysellavgprc':'Avg. Sell Price' ,'openbuyqty':'Open Buy Quantity' ,'opensellqty':'Open Sell Quantity'})
        
            df_positions['Net Quantity'] = df_positions['Net Quantity'].astype('int')
        
    except Exception as e:
        print(f"Error in get_position: {e}")
    return df_positions,day_m2m 

def LoadInstrument_token(Token_4_Exchange = ['NSE','BSE','NFO', 'BFO','CDS','MCX']):
    global df_ins_NSE, df_ins_BSE, df_ins_NFO, df_ins_BFO,df_ins_CDS, df_ins_MCX
    global api

    try:
        subdir = "Instrument"

        if not os.path.exists(subdir):
            os.makedirs(subdir)
            
        print("Finvasia intrument token download started, may take upto 2-3 minutes ..")
           
        if 'NSE' in Token_4_Exchange:
            #reading nse instrument symbol
            zip_file = "NSE_symbols.txt.zip"
            url = f"https://api.shoonya.com/{zip_file}"
            r = requests.get(f"{url}", allow_redirects=True)
            open(zip_file, "wb").write(r.content)
            df_ins_NSE = pd.read_csv(zip_file)
            os.remove(zip_file)
            try:
                df_ins_NSE.to_csv(os.path.join(subdir,"NSE_symbols.csv"),index = False)
            except:
                pass
                
        if 'BSE' in Token_4_Exchange:        
            #reading BSE instrument symbol
            zip_file = "BSE_symbols.txt.zip"
            url = f"https://api.shoonya.com/{zip_file}"
            r = requests.get(f"{url}", allow_redirects=True)
            open(zip_file, "wb").write(r.content)
            df_ins_BSE = pd.read_csv(zip_file)
            os.remove(zip_file)
            try:
                df_ins_BSE.to_csv(os.path.join(subdir,"BSE_symbols.csv"),index = False)
            except:
                pass
                
        
        if 'NFO' in Token_4_Exchange:
            #reading nfo instrument symbol
            zip_file = "NFO_symbols.txt.zip"
            url = f"https://api.shoonya.com/{zip_file}"
            r = requests.get(f"{url}", allow_redirects=True)
            open(zip_file, "wb").write(r.content)
            df_ins_NFO = pd.read_csv(zip_file)
            df_ins_NFO['Expiry'] = pd.to_datetime(df_ins_NFO['Expiry']).apply(lambda x: x.date())
            df_ins_NFO = df_ins_NFO.sort_values(by=['Expiry',"Symbol",'StrikePrice'], ascending=[True,True,True])
            df_ins_NFO = df_ins_NFO.astype({"StrikePrice": str}) 
            os.remove(zip_file)
            try:
                df_ins_NFO.to_csv(os.path.join(subdir,"NFO_symbols.csv"),index = False)
            except:
                pass

        if 'BFO' in Token_4_Exchange:
            #reading nfo instrument symbol
            zip_file = "BFO_symbols.txt.zip"
            url = f"https://api.shoonya.com/{zip_file}"
            r = requests.get(f"{url}", allow_redirects=True)
            open(zip_file, "wb").write(r.content)
            df_ins_BFO = pd.read_csv(zip_file, usecols=["Exchange", "Token", "LotSize", "TradingSymbol", "Expiry", "Instrument", "OptionType", "StrikePrice", "TickSize"])            
            bfo_symbols = np.where(
                                df_ins_BFO['TradingSymbol'].str.contains('SENSEX50', regex=False), 'SENSEX50',
                                df_ins_BFO['TradingSymbol'].str.extract(r'(.*?)(?:\d)', expand=False)
                                )
            df_ins_BFO.insert(3, "Symbol", bfo_symbols)
            df_ins_BFO['Expiry'] = pd.to_datetime(df_ins_BFO['Expiry']).apply(lambda x: x.date())
            df_ins_BFO = df_ins_BFO.sort_values(by=['Expiry',"Symbol",'StrikePrice'], ascending=[True,True,True])
            df_ins_BFO = df_ins_BFO.astype({"StrikePrice": str}) 
            os.remove(zip_file)
            try:
                df_ins_BFO.to_csv(os.path.join(subdir,"BFO_symbols.csv"),index = False)
            except:
                pass
                
        if 'CDS' in Token_4_Exchange:
            #reading CDS instrument symbol
            zip_file = "CDS_symbols.txt.zip"
            url = f"https://api.shoonya.com/{zip_file}"
            r = requests.get(f"{url}", allow_redirects=True)
            open(zip_file, "wb").write(r.content)
            df_ins_CDS = pd.read_csv(zip_file)
            df_ins_CDS['Expiry'] = pd.to_datetime(df_ins_CDS['Expiry']).apply(lambda x: x.date())
            df_ins_CDS = df_ins_CDS.sort_values(by=['Instrument','Expiry',"Symbol",'StrikePrice'], ascending=[False,True,True,True])
            df_ins_CDS = df_ins_CDS.astype({"StrikePrice": str})
            os.remove(zip_file)
            try:
                df_ins_CDS.to_csv(os.path.join(subdir,"CDS_symbols.csv"),index = False)
            except:
                pass
                
        if 'MCX' in Token_4_Exchange:
            #reading MCX instrument symbol
            zip_file = "MCX_symbols.txt.zip"
            url = f"https://api.shoonya.com/{zip_file}"
            r = requests.get(f"{url}", allow_redirects=True)
            open(zip_file, "wb").write(r.content)
            df_ins_MCX = pd.read_csv(zip_file)
            df_ins_MCX['Expiry'] = pd.to_datetime(df_ins_MCX['Expiry']).apply(lambda x: x.date())
            df_ins_MCX = df_ins_MCX.sort_values(by=['Instrument','Expiry',"Symbol",'StrikePrice'], ascending=[False,True,True,True])
            df_ins_MCX = df_ins_MCX.astype({"StrikePrice": str})
            os.remove(zip_file)
            try:
                df_ins_MCX.to_csv(os.path.join(subdir,"MCX_symbols.csv"),index = False)
            except:
                pass
                
        print("Finvasia intrument token download completed")
        
    except Exception as e:
        print(f"Exception in LoadInstrument_token : {e}")
        pass
        
def start_optionchain():
    global subscribe_symbol
    global live_data
    global df_ins_NFO
    global df_ins_NSE, df_ins_BSE, df_ins_NFO, df_ins_CDS, df_ins_MCX, df_ins_BFO
    
    global subs_lst
    excel_name = xw.Book(TerminalSheetName)
    oci = excel_name.sheets("Option_Chain_Input")
    Option_Chain_Output = excel_name.sheets("Option_Chain_Output")
    

    oci.range("F3").value = None
    oci.range("F4").value = None
    oci.range("a2:c500").value = None
    
    Option_Chain_Output.range('a3:ae500').value = None

    IterationSleep = 1
    
    oci.range("d2").value = "Segment==>>"
    oci.range("d3").value, oci.range("d4").value = "Symbol==>>", "Expiry==>>",
    oci.range("d5").value, oci.range("d6").value = "RefreshRate==>>", "NoOfStrike==>>",
    oci.range("d7").value, oci.range("d8").value = "ExpiryType==>>" , "GreekMatch==>>"
    
    pre_selected_segment = pre_selected_symbol = pre_selected_expiry = "" 
    pre_selected_NoOfStrike = 0
    
    while True:
        try:
            selected_segment = oci.range("E2").value
            Exchange = selected_segment[:3]
            if Exchange not in ['NFO','CDS','MCX', 'BFO']:
                Exchange = 'NFO'
            
            oci.range("F2").value = None
            if Exchange == 'NFO':
                df_instrument = df_ins_NFO
            elif Exchange == 'CDS':
                df_instrument = df_ins_CDS
            elif Exchange == 'MCX':
                df_instrument = df_ins_MCX
            elif Exchange == 'BFO':
                df_instrument = df_ins_BFO
            else:
                oci.range("F2").value = "Please select correct segment"
            #print(f"pre : pre_selected_segment = {pre_selected_segment}, selected_segment={selected_segment}")
            if pre_selected_segment != selected_segment:
                if Exchange == 'NFO':
                    df_symbol = df_ins_NFO
                if Exchange == 'BFO':
                    df_symbol = df_ins_BFO                
                if Exchange == 'CDS':
                    df_symbol = df_ins_CDS[df_ins_CDS['Instrument'] == 'UNDCUR']
                if Exchange == 'MCX':
                    df_symbol = df_ins_MCX[df_ins_MCX['Instrument'] == 'OPTFUT']
                    
                df_symbol = df_symbol.drop_duplicates( "Symbol" , keep='first')
                df_symbol =df_symbol[['Symbol']]
                oci.range("a2:c500").value = None
                oci.range("a1").options(index=False, header=True).value = df_symbol
                    
            pre_selected_segment = selected_segment
            #print(f"post : pre_selected_segment = {pre_selected_segment}, selected_segment={selected_segment}")
            
            input_symbol = str(oci.range("E3").value).strip()
            if input_symbol != None:

                df_instrument_temp = df_instrument[(df_instrument.Symbol == input_symbol) & (df_instrument['OptionType'].isin(['CE','PE']))]
                
                if len(df_instrument_temp) > 0:
                    if pre_selected_symbol != input_symbol:
                        df_exp = df_instrument_temp.drop_duplicates( "Expiry" , keep='first')
                        
                        df_exp =df_exp[['Expiry','LotSize']]
                        df_exp.sort_values(by = 'Expiry',inplace = True)
                        oci.range("b2:C100").value = None
                        oci.range("b2").options(index=False, header=False).value = df_exp
                    
                    pre_selected_symbol = input_symbol                    
                    
                    oci.range("F3").value = None
                    
                    expiry_input = oci.range("E4").value
                    
                    expiry_input = expiry_input.date() 

                    #print(f"Option Chain Details: input_symbol={input_symbol},expiry_input={expiry_input}")
                    if expiry_input != None:

                        df_instrument_temp = df_instrument_temp[df_instrument_temp.Expiry == expiry_input]
                        if len(df_instrument_temp) > 0:

                            oci.range("F4").value = None
                            try:
                                IterationSleep = oci.range("E5").value
                                if IterationSleep != None:
                                    IterationSleep = int(IterationSleep)
                                    oci.range("F5").value = None
                            except:
                                IterationSleep = 5
                                Message = "Iteration should be number"
                                print(Message)
                                oci.range("F5").value = Message

                            symbol_in_template = [itr for itr in OptionChain_template if itr["symbol"] ==input_symbol ]
                            if len(symbol_in_template) == 0:
                                #print("Symbol Not found, so append the details")

                                ind_df_specific_symbol = df_instrument[(df_instrument.Symbol == input_symbol) & (df_instrument['OptionType'].isin(['CE','PE']))]
                                #print(ind_df_specific_symbol)

                                List_of_Expiry = ind_df_specific_symbol["Expiry"].unique()
                                #print(List_of_Expiry)

                                Expiry_strikelist_list = []

                                for expiry in List_of_Expiry:
                                    if str( expiry ) != 'NaT' :
                                        #print(expiry)
                                        ind_df_specific_symbol_expiry = ind_df_specific_symbol[(ind_df_specific_symbol.Expiry == expiry) ]
                                        #print(ind_df_specific_symbol_expiry)

                                        List_of_strikes = ind_df_specific_symbol_expiry["StrikePrice"].unique()

                                        LotSize = ind_df_specific_symbol_expiry.iloc[0]["LotSize"]
                                        #print(LotSize)
                                        #print(List_of_strikes)
                                        strike_pe_ce_list = []

                                        for strike in List_of_strikes:
                                            #print(f"\nsymbol:{input_symbol},Expiry:{expiry},strike:{strike}")
                                            ind_df_specific_symbol_expiry_pe = ind_df_specific_symbol_expiry[
                                                (ind_df_specific_symbol_expiry.StrikePrice == str(strike)) & (ind_df_specific_symbol_expiry.OptionType == "PE") ]
                                            if len(ind_df_specific_symbol_expiry_pe) > 0:
                                                pe_token = ind_df_specific_symbol_expiry_pe.iloc[0]["Token"]
                                            else:
                                                pe_token = "NA"

                                            ind_df_specific_symbol_expiry_ce = ind_df_specific_symbol_expiry[(            ind_df_specific_symbol_expiry.StrikePrice== str(strike))& (            ind_df_specific_symbol_expiry.OptionType == "CE")]
                                            if len(ind_df_specific_symbol_expiry_ce) > 0:
                                                ce_token = ind_df_specific_symbol_expiry_ce.iloc[0]["Token"]
                                            else:
                                                ce_token = "NA"

                                            # print(f"symbol:{input_symbol},Expiry:{expiry},strike:{strike},pe_token={pe_token},ce_token={ce_token}")

                                            strike_pe_ce_dictionary = dict(
                                                {
                                                    "strike": strike,
                                                    "PE_Token": pe_token,
                                                    "CE_Token": ce_token,
                                                }
                                            )

                                            strike_pe_ce_list.append(strike_pe_ce_dictionary)

                                        # print(strike_pe_ce_list)

                                        expiry_stikelist_dict = dict(
                                            {
                                                "Expiry": expiry,
                                                "LotSize": LotSize,
                                                "Strike_list": strike_pe_ce_list,
                                            }
                                        )

                                        Expiry_strikelist_list.append(expiry_stikelist_dict)
                                    else:
                                        print(f"Ignoring the wrong date : {expiry}")
                                # print("\n\n")
                                # print(Expiry_strikelist_list)

                                Final_dic = dict(
                                    {
                                        "symbol": input_symbol,
                                        "Expiry_Strike_token": Expiry_strikelist_list,
                                    }
                                )

                                OptionChain_template.append(Final_dic)
                                # print("\n\n")
                                #print(OptionChain_template)
                            else:
                                Message = (
                                    "symbol already available in option chain template"
                                )

                            List_of_Expiry_Strike_token = [
                                itr["Expiry_Strike_token"]
                                for itr in OptionChain_template
                                if itr["symbol"] == input_symbol
                            ][0]
                            #print(f"\n\n******\n\n{List_of_Expiry_Strike_token}")

                            for expiry_strike in List_of_Expiry_Strike_token:
                                
                               
                                if input_symbol not in subs_lst:

                                    List_of_particular_expiry_strike = [
                                        itr["Strike_list"]
                                        for itr in List_of_Expiry_Strike_token
                                        if itr["Expiry"] == expiry_strike["Expiry"]
                                    ][0]

                                    for strike_dict in List_of_particular_expiry_strike:
                                        print(f"Going to subscribe {input_symbol} strike {strike_dict.get('strike')}")
                                        PE_Token = strike_dict.get("PE_Token")
                                        # print(PE_Token)
                                        CE_Token = strike_dict.get("CE_Token")
                                        # print(CE_Token)
                                        if PE_Token != "NA":
                                            subscribe_new_token(Exchange, PE_Token)
                                        if CE_Token != "NA":
                                            subscribe_new_token(Exchange, CE_Token)

                            if input_symbol not in subs_lst:
                                subs_lst.append(input_symbol)
                                print(f"{input_symbol} subscription completed")

                            pd_oc = pd.DataFrame(columns=[ "CE_token", "CE_oi",  "CE_poi", "CE_toi", "CE_lp", "CE_pc", "CE_bq1","CE_bp1", "CE_sq1", "CE_sp1",  "strike", "PE_bq1", "PE_bp1", "PE_sq1", "PE_sp1", "PE_pc",  "PE_lp",  "PE_toi", "PE_poi", "PE_oi", "PE_token","CE_coi","PE_coi","CE_v","PE_v"] )

                            # prepare option chain
                            List_of_Expiry_Strike_token = [
                                itr["Expiry_Strike_token"]
                                for itr in OptionChain_template
                                if itr["symbol"] == input_symbol
                            ][0]
                            #print(f"List_of_Expiry_Strike_token = {List_of_Expiry_Strike_token}")
                            try:
                                Lot_Size = [
                                    itr["LotSize"]
                                    for itr in List_of_Expiry_Strike_token
                                    if itr["Expiry"] == expiry_input
                                ][0]
                                List_of_particular_expiry_strike = [
                                    itr["Strike_list"]
                                    for itr in List_of_Expiry_Strike_token
                                    if itr["Expiry"] == expiry_input
                                ][0]
                                #print(f"****{live_data}")
                                #print(f"List_of_particular_expiry_strike= {List_of_particular_expiry_strike}")
                                
                                isFound, Fut_Token = GetToken(Exchange,input_symbol)
                                #print(f"isFound {isFound} Fut_Token {Fut_Token}")
                                if Exchange == 'NFO':
                                    isFound, Spot_Token = GetToken('NSE',input_symbol)
                                    spot_ltp = convert_to_float(api.get_quotes("NSE", str(Spot_Token)).get("lp"))
                                elif Exchange == 'BFO':
                                    isFound, Spot_Token = GetToken('BSE',input_symbol)
                                    spot_ltp = convert_to_float(api.get_quotes("BSE", str(Spot_Token)).get("lp"))
                                else:    
                                    Spot_Token = Fut_Token
                                    spot_ltp = convert_to_float(api.get_quotes(Exchange, str(Spot_Token)).get("lp"))
                                #print(f"Spot_Token {Spot_Token} spot_ltp {spot_ltp}")
                                future_ltp = convert_to_float(api.get_quotes(Exchange, str(Fut_Token)).get("lp"))
                                
                                #print(f"{input_symbol} spot ltp = {spot_ltp} future ltp = {future_ltp}")

                                for strike_dict in List_of_particular_expiry_strike:
                                    
                                    Strike = convert_to_float(strike_dict.get("strike"))
                                    #print(Strike)
                                    PE_Token = strike_dict.get("PE_Token")
                                    PE_Token = str(Exchange)+ "|"  + str(PE_Token)
                                    #print(PE_Token)
                                    CE_Token = strike_dict.get("CE_Token")
                                    CE_Token =  str(Exchange)+ "|" + str(CE_Token)
                                    #print(CE_Token)
                                    
                                    
                                    try:
                                        CE_oi = live_data[str(CE_Token)].get("oi", 0)
                                    except:
                                        CE_oi = 0
                                    try:
                                        CE_poi = live_data[str(CE_Token)].get("poi", 0)
                                    except:
                                        CE_poi = 0
                                    
                                    CE_coi = int(CE_oi) - int(CE_poi)
                                    
                                    try:
                                        CE_toi = live_data[str(CE_Token)].get("toi", "-")
                                    except:
                                        CE_toi = "-"
                                    try:
                                        CE_lp = live_data[str(CE_Token)].get("lp", 0)
                                    except:
                                        CE_lp = 0
                                    try:
                                        CE_pc = live_data[str(CE_Token)].get("pc", "-")
                                    except:
                                        CE_pc = "-"
                                    try:
                                        CE_bq1 = live_data[str(CE_Token)].get("bq1", "-")
                                    except:
                                        CE_bq1 = "-"
                                    try:
                                        CE_bp1 = live_data[str(CE_Token)].get("bp1", "-")
                                    except:
                                        CE_bp1 = "-"
                                    try:
                                        CE_sq1 = live_data[str(CE_Token)].get("sq1", "-")
                                    except:
                                        CE_sq1 = "-"
                                    try:
                                        CE_sp1 = live_data[str(CE_Token)].get("sp1", "-")
                                    except:
                                        CE_sp1 = "-"
                                    
                                    try:
                                        PE_oi = live_data[str(PE_Token)].get("oi", 0)
                                    except:
                                        PE_oi = 0
                                    try:
                                        PE_poi = live_data[str(PE_Token)].get("poi", 0)
                                    except:
                                        PE_poi = 0
                                    
                                    PE_coi = int(PE_oi) - int(PE_poi)
                                    
                                    #print(f"CE COI : {CE_oi} {CE_poi} {CE_coi}")
                                    #print(f"PE COI : {PE_oi} {PE_poi} {PE_coi}")
                                    try:
                                        PE_toi = live_data[str(PE_Token)].get("toi", "-")
                                    except:
                                        PE_toi = "-"
                                    try:
                                        PE_lp = live_data[str(PE_Token)].get("lp", 0)
                                    except:
                                        PE_lp = 0
                                    try:
                                        PE_pc = live_data[str(PE_Token)].get("pc", "-")
                                    except:
                                        PE_pc = "-"
                                    try:
                                        PE_bq1 = live_data[str(PE_Token)].get("bq1", "-")
                                    except:
                                        PE_bq1 = "-"
                                    try:
                                        PE_bp1 = live_data[str(PE_Token)].get("bp1", "-")
                                    except:
                                        PE_bp1 = "-"
                                    try:
                                        PE_sq1 = live_data[str(PE_Token)].get("sq1", "-")
                                    except:
                                        PE_sq1 = "-"
                                    try:
                                        PE_sp1 = live_data[str(PE_Token)].get("sp1", "-")
                                    except:
                                        PE_sp1 = "-"
                                    
                                    try:
                                        CE_v = live_data[str(CE_Token)].get("v", 0)
                                    except:
                                        CE_v = 0
                                        
                                    try:
                                        PE_v = live_data[str(PE_Token)].get("v", 0)
                                    except:
                                        PE_v = 0
                                    
                                    dic_data = {
                                            "CE_token": CE_Token,
                                            "CE_oi": CE_oi,
                                            "CE_poi": CE_poi,
                                            "CE_toi": CE_toi,
                                            "CE_lp": CE_lp,
                                            "CE_pc": CE_pc,
                                            "CE_bq1": CE_bq1,
                                            "CE_bp1": CE_bp1,
                                            "CE_sq1": CE_sq1,
                                            "CE_sp1": CE_sp1,
                                            "strike": Strike,
                                            "PE_bq1": PE_bq1,
                                            "PE_bp1": PE_bp1,
                                            "PE_sq1": PE_sq1,
                                            "PE_sp1": PE_sp1,
                                            "PE_pc": PE_pc,
                                            "PE_lp": PE_lp,
                                            "PE_toi": PE_toi,
                                            "PE_poi": PE_poi,
                                            "PE_oi": PE_oi,
                                            "PE_token": PE_Token,
                                            "CE_coi":CE_coi,
                                            "PE_coi":PE_coi,
                                            "CE_v":CE_v,
                                            "PE_v":PE_v,
                                        }
                                    pd_oc = pd.concat([pd_oc, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                    
                                #pd_oc.to_csv("pd_oc.csv")                            
                                #print(pd_oc)
                                #pd_oc = pd_oc.astype({"strike": float})
                                pd_oc['strike'] = pd_oc['strike'].apply(convert_to_float)
                                pd_oc = pd_oc.sort_values(by="strike", ascending=True)

                                pd_oc = pd_oc.fillna(0)
                                
                                pd_oc = pd_oc.astype({"CE_oi": int})
                                pd_oc = pd_oc.astype({"PE_oi": int})
                                pd_oc = pd_oc.astype({"CE_poi": int})
                                pd_oc = pd_oc.astype({"PE_poi": int})
                                pd_oc = pd_oc.astype({"CE_coi": int})
                                pd_oc = pd_oc.astype({"PE_coi": int})
                                
                                pd_oc = pd_oc.astype({"CE_v": int})
                                pd_oc = pd_oc.astype({"PE_v": int})
                                
                                pd_oc["CE_v"] = pd_oc["CE_v"] / int(Lot_Size)
                                pd_oc["PE_v"] = pd_oc["PE_v"] / int(Lot_Size)
                                
                                pd_oc["CE_oi"] = pd_oc["CE_oi"] / int(Lot_Size)
                                pd_oc["PE_oi"] = pd_oc["PE_oi"] / int(Lot_Size)
                                pd_oc["CE_poi"] = pd_oc["CE_poi"] / int(Lot_Size)
                                pd_oc["PE_poi"] = pd_oc["PE_poi"] / int(Lot_Size)
                                pd_oc["CE_coi"] = pd_oc["CE_coi"] / int(Lot_Size)
                                pd_oc["PE_coi"] = pd_oc["PE_coi"] / int(Lot_Size)
                                pd_oc["OI_SUM"] = pd_oc["CE_oi"] + pd_oc["PE_oi"]

                                #print(pd_oc)
                                
                                df_oc_pro = pd_oc
                                
                                try:
                                    NoOfStrike = int(oci.range("E6").value)
                                except Exception as e:
                                    NoOfStrike = 100
                                
                                df_oc_pro['strike_diff'] = abs(df_oc_pro['strike'] - spot_ltp)
            
                                df_oc_pro.sort_values(by = 'strike_diff',inplace = True)
                                
                                #print(f"\n\n***\n\n{df_oc_pro}")
                                AtmStrike = convert_to_float(df_oc_pro.iloc[0]['strike'])
                                AtmStrikeCallPrice = convert_to_float(df_oc_pro.iloc[0]['CE_lp'])
                                AtmStrikePutPrice = convert_to_float(df_oc_pro.iloc[0]['PE_lp'])
                                
                                #additional detail related to dump on input page
                                Future_LTP = future_ltp
                                Max_Pain_at_Strike = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['strike']
                                Ltp_at_Max_Pain_CE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['CE_lp']
                                Ltp_at_Max_Pain_PE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['PE_lp']
                                #ATM_Strike = 
                                LTP_at_ATM_CE = df_oc_pro[df_oc_pro.strike == AtmStrike].iloc[0]['CE_lp']
                                LTP_at_ATM_PE = df_oc_pro[df_oc_pro.strike == AtmStrike].iloc[0]['PE_lp']
                                Total_OI_CE = df_oc_pro["CE_oi"].sum()
                                Total_OI_PE = df_oc_pro["PE_oi"].sum()
                                
                                Max_OI_CE = df_oc_pro["CE_oi"].max()
                                Max_OI_PE = df_oc_pro["PE_oi"].max()
                                Max_OI_at_Strike_CE = df_oc_pro[df_oc_pro.CE_oi == df_oc_pro["CE_oi"].max()].iloc[0]['strike']
                                Max_OI_at_Strike_PE = df_oc_pro[df_oc_pro.PE_oi == df_oc_pro["PE_oi"].max()].iloc[0]['strike']
                                LTP_of_Max_OI_Strike_CE = df_oc_pro[df_oc_pro.CE_oi == df_oc_pro["CE_oi"].max()].iloc[0]['CE_lp']
                                LTP_of_Max_OI_Strike_PE = df_oc_pro[df_oc_pro.PE_oi == df_oc_pro["PE_oi"].max()].iloc[0]['PE_lp']

                                Total_OI_Change_CE = df_oc_pro["CE_coi"].sum()
                                Total_OI_Change_PE = df_oc_pro["PE_coi"].sum()
                                
                                Max_Change_in_OI_addition_CE = df_oc_pro["CE_coi"].max()
                                Max_Change_in_OI_addition_PE = df_oc_pro["PE_coi"].max()
                                Max_OI_addition_at_Srike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].max()].iloc[0]['strike']
                                Max_OI_addition_at_Srike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].max()].iloc[0]['strike']
                                LTP_of_Max_OI_addition_Strike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].max()].iloc[0]['CE_lp']
                                LTP_of_Max_OI_addition_Strike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].max()].iloc[0]['PE_lp']
                                Max_Change_in_OI_unwinding_CE = -1 * int(df_oc_pro["CE_coi"].min())
                                Max_Change_in_OI_unwinding_PE = -1 * int(df_oc_pro["PE_coi"].min())
                                Max_OI_unwinding_at_Srike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].min()].iloc[0]['strike']
                                Max_OI_unwinding_at_Srike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].min()].iloc[0]['strike']
                                LTP_of_Max_OI_unwinding_Strike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].min()].iloc[0]['CE_lp']
                                LTP_of_Max_OI_unwinding_Strike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].min()].iloc[0]['PE_lp']
                            
                                df_additional_detail = pd.DataFrame(columns = ['CE','PE'])
                                dic_data = {'CE':Future_LTP}

                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_Pain_at_Strike}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Ltp_at_Max_Pain_CE,'PE': Ltp_at_Max_Pain_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':AtmStrike}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':LTP_at_ATM_CE,'PE':LTP_at_ATM_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Total_OI_CE, 'PE':Total_OI_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Total_OI_Change_CE, 'PE':Total_OI_Change_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_OI_CE,  'PE':Max_OI_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_OI_at_Strike_CE,'PE':Max_OI_at_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':LTP_of_Max_OI_Strike_CE,'PE':LTP_of_Max_OI_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_Change_in_OI_addition_CE,'PE':Max_Change_in_OI_addition_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_OI_addition_at_Srike_CE,'PE':Max_OI_addition_at_Srike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':LTP_of_Max_OI_addition_Strike_CE, 'PE':LTP_of_Max_OI_addition_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_Change_in_OI_unwinding_CE, 'PE':Max_Change_in_OI_unwinding_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_OI_unwinding_at_Srike_CE, 'PE':Max_OI_unwinding_at_Srike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':LTP_of_Max_OI_unwinding_Strike_CE,'PE':LTP_of_Max_OI_unwinding_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                

                                oci.range("i3").options(index=False, header=False).value = df_additional_detail
                                
                                
                                
                                df_oc_pro = df_oc_pro[:2*int(NoOfStrike)+1]
                                
                                df_oc_pro.sort_values(by ='strike',inplace = True)
                                
                                df_oc_pro['CE_Delta'] = None
                                df_oc_pro['CE_Gamma'] = None
                                df_oc_pro['CE_Theta'] = None
                                df_oc_pro['CE_Vega'] = None
                                df_oc_pro['CE_Rho'] = None
                                df_oc_pro['CE_IV'] = None
                                
                                df_oc_pro['PE_Delta'] = None
                                df_oc_pro['PE_Gamma'] = None
                                df_oc_pro['PE_Theta'] = None
                                df_oc_pro['PE_Vega'] = None
                                df_oc_pro['PE_Rho'] = None
                                df_oc_pro['PE_IV'] = None
                    
                                df_oc_pro = df_oc_pro.reindex(['CE_Delta','CE_Gamma','CE_Theta','CE_Vega','CE_Rho','CE_oi','CE_coi','CE_v','CE_IV','CE_lp','CE_pc','CE_bq1','CE_bp1','CE_sp1','CE_sq1','strike','PE_bq1','PE_bp1','PE_sp1','PE_sq1','PE_pc','PE_lp','PE_IV','PE_v','PE_coi','PE_oi','PE_Rho','PE_Vega','PE_Theta','PE_Gamma','PE_Delta'], axis=1)
                                
                                SpotPrice = convert_to_float(spot_ltp)
                                FuturePrice = convert_to_float(future_ltp)
                                ExpiryDateTime = dt(expiry_input.year, expiry_input.month, expiry_input.day, 0, 0, 0)
                                
                                #print(f"SpotPrice = {SpotPrice}, FuturePrice={FuturePrice}, AtmStrike={AtmStrike}, AtmStrikeCallPrice={AtmStrikeCallPrice}, AtmStrikePutPrice={AtmStrikePutPrice}, ExpiryDateTime={ExpiryDateTime}")
                                
                                ExpiryType = oci.range("F7").value
                                GreekMatch = oci.range("F8").value
                                
                                if ExpiryType == 'WEEKLY':
                                    ExpiryDateType = ExpType.WEEKLY
                                else:
                                    ExpiryDateType = ExpType.MONTHLY
                                
                                FromDateTime = dt.now() 
                                if Exchange == 'NFO':
                                    if dt.now().time() > time(15, 30, 0):
                                        FromDateTime = dt(dt.now().year, dt.now().month,dt.now().day, 15, 30, 0)
                                    
                                    
                                if GreekMatch == "SENSIBULL":
                                    tryMatchWith=TryMatchWith.SENSIBULL
                                else:
                                    tryMatchWith=TryMatchWith.NSE
                                
                                dayCountType = DayCountType.CALENDARDAYS
                                
                                IvGreeks = CalcIvGreeks( SpotPrice = SpotPrice,  FuturePrice = FuturePrice, AtmStrike = AtmStrike, AtmStrikeCallPrice = AtmStrikeCallPrice, AtmStrikePutPrice = AtmStrikePutPrice, ExpiryDateTime = ExpiryDateTime, ExpiryDateType = ExpiryDateType, FromDateTime = FromDateTime, tryMatchWith = tryMatchWith, dayCountType = dayCountType)
                
                                #print(f"SpotPrice={SpotPrice}, FuturePrice={FuturePrice},  AtmStrike={AtmStrike}, AtmStrikeCallPrice={AtmStrikeCallPrice}, AtmStrikePutPrice={AtmStrikePutPrice}, ExpiryDateTime={ExpiryDateTime},  ExpiryDateType={ExpiryDateType}, FromDateTime={FromDateTime}, tryMatchWith={tryMatchWith}")
                                
                                for ind in df_oc_pro.index:
                                    
                                    StrikePrice= convert_to_float(df_oc_pro['strike'][ind])
                                    StrikeCallPrice= convert_to_float(df_oc_pro['CE_lp'][ind])
                                    StrikePutPrice= convert_to_float(df_oc_pro['PE_lp'][ind])
                                    #print(f"StrikePrice={StrikePrice}, StrikeCallPrice={StrikeCallPrice}, StrikePutPrice={StrikePutPrice}")
                                    Greeks = IvGreeks.GetImpVolAndGreeks( StrikePrice = StrikePrice, StrikeCallPrice = StrikeCallPrice, StrikePutPrice = StrikePutPrice)
                                    #print(Greeks)
                                    
                                    df_oc_pro['CE_Delta'][ind] = round(Greeks["CallDelta"],2)
                                    df_oc_pro['CE_Gamma'][ind] = round(Greeks["Gamma"],4)
                                    df_oc_pro['CE_Theta'][ind] = round(Greeks["Theta"],2)
                                    df_oc_pro['CE_Vega'][ind] = round(Greeks["Vega"],2)
                                    df_oc_pro['CE_Rho'][ind] = round(Greeks["RhoCall"],4)
                                    
                                    
                                    df_oc_pro['PE_Delta'][ind] = round(Greeks["PutDelta"],2)
                                    df_oc_pro['PE_Gamma'][ind] = round(Greeks["Gamma"],4)
                                    df_oc_pro['PE_Theta'][ind] = round(Greeks["Theta"],2)
                                    df_oc_pro['PE_Vega'][ind] = round(Greeks["Vega"],2)
                                    df_oc_pro['PE_Rho'][ind] = round(Greeks["RhoPut"],4)
                                    
                                    if GreekMatch == "NSE":
                                        df_oc_pro['CE_IV'][ind] = round(Greeks["CallIV"],2)
                                        df_oc_pro['PE_IV'][ind] = round(Greeks["PutIV"],2)
                                    else:
                                        df_oc_pro['CE_IV'][ind] = round(Greeks["ImplVol"],2)
                                        df_oc_pro['PE_IV'][ind] = round(Greeks["ImplVol"],2)
                                
                                del IvGreeks
                                df_oc_pro.round({"CE_lp":2, 'PE_lp':2})
                                #print(df_oc_pro)
                                
                                if pre_selected_NoOfStrike != NoOfStrike:
                                    pre_selected_NoOfStrike = NoOfStrike
                                    Option_Chain_Output.range('a3:ae500').value = None
                                    Option_Chain_Output.range(f"a3:AE500").color = (255,255,255)
                                    Option_Chain_Output.range(f"p3:p{2 * int(NoOfStrike) + 3}").color = (46,132,198) 
                                
                                df_oc_pro = df_oc_pro.reset_index(drop = True)
                                ATM_pos = df_oc_pro[df_oc_pro.strike == AtmStrike].index.values[0]
                                if NoOfStrike * 2 < len(df_oc_pro):
                                    df_oc_pro = df_oc_pro.iloc[ATM_pos - int(NoOfStrike) : ATM_pos + int(NoOfStrike) + 1]
                                    ATM_Row = int(NoOfStrike) + 3
                                    Option_Chain_Output.range(f"a{ATM_Row}:AE{ATM_Row}").color = (46,132,198)
                        
                                Option_Chain_Output.range('a3').options(index=False,header=False).value = df_oc_pro
                                
                            except Exception as e:
                                Option_Chain_Output.range('a3:ae500').value = None
                                oci.range('I3:J18').value = None
                                Message = "Please check the all provided detail to load option chain:" + str(e)
                                print(Message)
                                oci.range("F4").value = Message
                                
                        else:
                            Option_Chain_Output.range('a3:ae500').value = None
                            oci.range('I3:J18').value = None
                            Message = "Please enter correct expiry in dd-mm-YYYY (date format)"
                            print(Message)
                            oci.range("F4").value = Message
                            
                    else:
                        Option_Chain_Output.range('a3:ae500').value = None
                        oci.range('I3:J18').value = None
                        Message = "Please enter the expiry"
                        print(Message)
                        oci.range("F4").value = Message
                        
                else:
                    Option_Chain_Output.range('a3:ae500').value = None
                    oci.range('I3:J18').value = None
                    Message = "Please enter correct symbol"
                    print(Message)
                    oci.range("F3").value = Message
                    oci.range("F4").value = None
                    oci.range("b2:c100").value = None
                    
            else:
                Option_Chain_Output.range('a3:ae500').value = None
                oci.range('I3:J18').value = None
                Message = "Please enter the symbol"
                print(Message)
                oci.range("F3").value = Message
                oci.range("F4").value = None
                oci.range("b2:c100").value = None
                
        except Exception as e:
            print(f"Excption : {e}")
            pass
        sleep(int(IterationSleep))
        
def GetToken(Exchange,Symbol, Type = 'FUT' ,Expiry=None,Strike = None):
    #print(f"GetToken called with parametrer \nExchange = {Exchange}, Symbol = {Symbol}")
    global df_ins_NSE, df_ins_BSE, df_ins_NFO, df_ins_CDS, df_ins_MCX, df_ins_BFO
    try:
        Token = None
        isTokenFound = False
        Symbol = Symbol.upper()
        if Exchange == 'NSE':
            if Symbol in IndexList:    
                Token = Token_list[Symbol]
                isTokenFound = True
            else:
                df_ins_temp = df_ins_NSE[df_ins_NSE.Symbol == Symbol]
                if len(df_ins_temp) > 0:
                    Token = df_ins_temp.iloc[0]['Token']
                    isTokenFound = True
        
        elif Exchange == 'BSE':
            if Symbol in IndexList:    
                Token = Token_list[Symbol]
                isTokenFound = True
            else:
                df_ins_temp = df_ins_BSE[df_ins_BSE.Symbol == Symbol]
                if len(df_ins_temp) > 0:
                    Token = df_ins_temp.iloc[0]['Token']
                    isTokenFound = True
        elif Exchange == 'NFO':
            df_ins_temp = df_ins_NFO[(df_ins_NFO.Symbol == Symbol) & (df_ins_NFO['Instrument'].isin(['FUTIDX','FUTSTK']))]
            if len(df_ins_temp) > 0:
                df_ins_temp.sort_values(by = 'Expiry',inplace = True)
                Token = df_ins_temp.iloc[0]['Token']
                isTokenFound = True
        elif Exchange == 'BFO':
            df_ins_temp = df_ins_BFO[(df_ins_BFO.Symbol == Symbol) & (df_ins_BFO['Instrument'].isin(['FUTIDX','FUTSTK']))]
            if len(df_ins_temp) > 0:
                df_ins_temp.sort_values(by = 'Expiry',inplace = True)
                Token = df_ins_temp.iloc[0]['Token']
                isTokenFound = True
        elif Exchange == 'CDS':
            df_ins_temp = df_ins_CDS[(df_ins_CDS.Symbol == Symbol) & (df_ins_CDS.Instrument == 'FUTCUR')]
            if len(df_ins_temp) > 0:
                df_ins_temp.sort_values(by = 'Expiry',inplace = True)
                Token = df_ins_temp.iloc[0]['Token']
                isTokenFound = True
        elif Exchange == 'MCX':
            df_ins_temp = df_ins_MCX[(df_ins_MCX.Symbol == Symbol) & (df_ins_MCX.Instrument == 'FUTCOM')]
            if len(df_ins_temp) > 0:
                df_ins_temp.sort_values(by = 'Expiry',inplace = True)
                Token = df_ins_temp.iloc[0]['Token']
                isTokenFound = True
    except Exception as e:
        print(f"Exception occur in GetToken : {e}")
    
    #print(f"Returning value isTokenFound = {isTokenFound}, Token = {Token} ")
    return isTokenFound, Token
    
def start_optionchain_Pro():
    global subscribe_symbol
    global live_data
    global df_ins_NFO
    global df_ins_NSE, df_ins_BSE, df_ins_NFO, df_ins_CDS, df_ins_MCX, df_ins_BFO
    
    global subs_lst
    excel_name = xw.Book(TerminalSheetName)
    oci_pro = excel_name.sheets("Option_Chain_Pro_Input")
    Option_Chain_Pro_Output = excel_name.sheets("Option_Chain_Pro_Output")
    

    oci_pro.range("F3").value = None
    oci_pro.range("F4").value = None
    oci_pro.range("a2:c500").value = None
    
    Option_Chain_Pro_Output.range('a3:ae500').value = None

    IterationSleep = 1
    
    oci_pro.range("d2").value = "Segment==>>"
    oci_pro.range("d3").value, oci_pro.range("d4").value = "Symbol==>>", "Expiry==>>",
    oci_pro.range("d5").value, oci_pro.range("d6").value = "RefreshRate==>>", "NoOfStrike==>>",
    oci_pro.range("d7").value, oci_pro.range("d8").value = "ExpiryType==>>" , "GreekMatch==>>"
    
    pre_selected_segment = pre_selected_symbol = pre_selected_expiry = "" 
    pre_selected_NoOfStrike = 0
    
    while True:
        try:
            selected_segment = oci_pro.range("E2").value
            Exchange = selected_segment[:3]
            if Exchange not in ['NFO','CDS','MCX','BFO']:
                Exchange = 'NFO'
            
            oci_pro.range("F2").value = None
            if Exchange == 'NFO':
                df_instrument = df_ins_NFO
            elif Exchange == 'CDS':
                df_instrument = df_ins_CDS
            elif Exchange == 'MCX':
                df_instrument = df_ins_MCX
            elif Exchange == 'BFO':
                df_instrument = df_ins_BFO
            else:
                oci_pro.range("F2").value = "Please select correct segment"
            #print(f"pre : pre_selected_segment = {pre_selected_segment}, selected_segment={selected_segment}")
            if pre_selected_segment != selected_segment:
                if Exchange == 'NFO':
                    df_symbol = df_ins_NFO
                if Exchange == 'BFO':
                    df_symbol = df_ins_BFO
                if Exchange == 'CDS':
                    df_symbol = df_ins_CDS[df_ins_CDS['Instrument'] == 'UNDCUR']
                if Exchange == 'MCX':
                    df_symbol = df_ins_MCX[df_ins_MCX['Instrument'] == 'OPTFUT']
                    
                df_symbol = df_symbol.drop_duplicates( "Symbol" , keep='first')
                df_symbol =df_symbol[['Symbol']]
                oci_pro.range("a2:c500").value = None
                oci_pro.range("a1").options(index=False, header=True).value = df_symbol
                    
            pre_selected_segment = selected_segment
            #print(f"post : pre_selected_segment = {pre_selected_segment}, selected_segment={selected_segment}")
            
            input_symbol = str(oci_pro.range("E3").value).strip()
            if input_symbol != None:

                df_instrument_temp = df_instrument[(df_instrument.Symbol == input_symbol) & (df_instrument['OptionType'].isin(['CE','PE']))]
                
                if len(df_instrument_temp) > 0:
                    if pre_selected_symbol != input_symbol:
                        df_exp = df_instrument_temp.drop_duplicates( "Expiry" , keep='first')
                        
                        df_exp =df_exp[['Expiry','LotSize']]
                        df_exp.sort_values(by = 'Expiry',inplace = True)
                        oci_pro.range("b2:C100").value = None
                        oci_pro.range("b2").options(index=False, header=False).value = df_exp
                    
                    pre_selected_symbol = input_symbol                    
                    
                    oci_pro.range("F3").value = None
                    
                    expiry_input = oci_pro.range("E4").value
                    
                    expiry_input = expiry_input.date() 

                    #print(f"Option Chain Details: input_symbol={input_symbol},expiry_input={expiry_input}")
                    if expiry_input != None:

                        df_instrument_temp = df_instrument_temp[df_instrument_temp.Expiry == expiry_input]
                        if len(df_instrument_temp) > 0:

                            oci_pro.range("F4").value = None
                            try:
                                IterationSleep = oci_pro.range("E5").value
                                if IterationSleep != None:
                                    IterationSleep = int(IterationSleep)
                                    oci_pro.range("F5").value = None
                            except:
                                IterationSleep = 5
                                Message = "Iteration should be number"
                                print(Message)
                                oci_pro.range("F5").value = Message

                            symbol_in_template = [itr for itr in OptionChain_template if itr["symbol"] ==input_symbol ]
                            if len(symbol_in_template) == 0:
                                #print("Symbol Not found, so append the details")

                                ind_df_specific_symbol = df_instrument[(df_instrument.Symbol == input_symbol) & (df_instrument['OptionType'].isin(['CE','PE']))]
                                #print(ind_df_specific_symbol)

                                List_of_Expiry = ind_df_specific_symbol["Expiry"].unique()
                                #print(List_of_Expiry)

                                Expiry_strikelist_list = []

                                for expiry in List_of_Expiry:
                                    if str( expiry ) != 'NaT' :
                                        #print(expiry)
                                        ind_df_specific_symbol_expiry = ind_df_specific_symbol[(ind_df_specific_symbol.Expiry == expiry) ]
                                        #print(ind_df_specific_symbol_expiry)

                                        List_of_strikes = ind_df_specific_symbol_expiry["StrikePrice"].unique()

                                        LotSize = ind_df_specific_symbol_expiry.iloc[0]["LotSize"]
                                        #print(LotSize)
                                        #print(List_of_strikes)
                                        strike_pe_ce_list = []

                                        for strike in List_of_strikes:
                                            #print(f"\nsymbol:{input_symbol},Expiry:{expiry},strike:{strike}")
                                            ind_df_specific_symbol_expiry_pe = ind_df_specific_symbol_expiry[
                                                (ind_df_specific_symbol_expiry.StrikePrice == str(strike)) & (ind_df_specific_symbol_expiry.OptionType == "PE") ]
                                            if len(ind_df_specific_symbol_expiry_pe) > 0:
                                                pe_token = ind_df_specific_symbol_expiry_pe.iloc[0]["Token"]
                                            else:
                                                pe_token = "NA"

                                            ind_df_specific_symbol_expiry_ce = ind_df_specific_symbol_expiry[(            ind_df_specific_symbol_expiry.StrikePrice== str(strike))& (            ind_df_specific_symbol_expiry.OptionType == "CE")]
                                            if len(ind_df_specific_symbol_expiry_ce) > 0:
                                                ce_token = ind_df_specific_symbol_expiry_ce.iloc[0]["Token"]
                                            else:
                                                ce_token = "NA"

                                            # print(f"symbol:{input_symbol},Expiry:{expiry},strike:{strike},pe_token={pe_token},ce_token={ce_token}")

                                            strike_pe_ce_dictionary = dict(
                                                {
                                                    "strike": strike,
                                                    "PE_Token": pe_token,
                                                    "CE_Token": ce_token,
                                                }
                                            )

                                            strike_pe_ce_list.append(strike_pe_ce_dictionary)

                                        # print(strike_pe_ce_list)

                                        expiry_stikelist_dict = dict(
                                            {
                                                "Expiry": expiry,
                                                "LotSize": LotSize,
                                                "Strike_list": strike_pe_ce_list,
                                            }
                                        )

                                        Expiry_strikelist_list.append(expiry_stikelist_dict)
                                    else:
                                        print(f"Ignoring the wrong date : {expiry}")
                                # print("\n\n")
                                # print(Expiry_strikelist_list)

                                Final_dic = dict(
                                    {
                                        "symbol": input_symbol,
                                        "Expiry_Strike_token": Expiry_strikelist_list,
                                    }
                                )

                                OptionChain_template.append(Final_dic)
                                # print("\n\n")
                                #print(OptionChain_template)
                            else:
                                Message = (
                                    "symbol already available in option chain template"
                                )

                            List_of_Expiry_Strike_token = [
                                itr["Expiry_Strike_token"]
                                for itr in OptionChain_template
                                if itr["symbol"] == input_symbol
                            ][0]
                            #print(f"\n\n******\n\n{List_of_Expiry_Strike_token}")

                            for expiry_strike in List_of_Expiry_Strike_token:
                                
                               
                                if input_symbol not in subs_lst:

                                    List_of_particular_expiry_strike = [
                                        itr["Strike_list"]
                                        for itr in List_of_Expiry_Strike_token
                                        if itr["Expiry"] == expiry_strike["Expiry"]
                                    ][0]

                                    for strike_dict in List_of_particular_expiry_strike:
                                        print(f"Going to subscribe {input_symbol} strike {strike_dict.get('strike')}")
                                        PE_Token = strike_dict.get("PE_Token")
                                        # print(PE_Token)
                                        CE_Token = strike_dict.get("CE_Token")
                                        # print(CE_Token)
                                        if PE_Token != "NA":
                                            subscribe_new_token(Exchange, PE_Token)
                                        if CE_Token != "NA":
                                            subscribe_new_token(Exchange, CE_Token)

                            if input_symbol not in subs_lst:
                                subs_lst.append(input_symbol)
                                print(f"{input_symbol} subscription completed")

                            pd_oc = pd.DataFrame(columns=[ "CE_token", "CE_oi",  "CE_poi", "CE_toi", "CE_lp", "CE_pc", "CE_bq1","CE_bp1", "CE_sq1", "CE_sp1",  "strike", "PE_bq1", "PE_bp1", "PE_sq1", "PE_sp1", "PE_pc",  "PE_lp",  "PE_toi", "PE_poi", "PE_oi", "PE_token","CE_coi","PE_coi","CE_v","PE_v"] )

                            # prepare option chain
                            List_of_Expiry_Strike_token = [
                                itr["Expiry_Strike_token"]
                                for itr in OptionChain_template
                                if itr["symbol"] == input_symbol
                            ][0]
                            #print(f"List_of_Expiry_Strike_token = {List_of_Expiry_Strike_token}")
                            try:
                                Lot_Size = [
                                    itr["LotSize"]
                                    for itr in List_of_Expiry_Strike_token
                                    if itr["Expiry"] == expiry_input
                                ][0]
                                List_of_particular_expiry_strike = [
                                    itr["Strike_list"]
                                    for itr in List_of_Expiry_Strike_token
                                    if itr["Expiry"] == expiry_input
                                ][0]
                                #print(f"****{live_data}")
                                #print(f"List_of_particular_expiry_strike= {List_of_particular_expiry_strike}")
                                
                                isFound, Fut_Token = GetToken(Exchange,input_symbol)
                                #print(f"isFound {isFound} Fut_Token {Fut_Token}")
                                if Exchange == 'NFO':
                                    isFound, Spot_Token = GetToken('NSE',input_symbol)
                                    spot_ltp = convert_to_float(api.get_quotes("NSE", str(Spot_Token)).get("lp"))
                                if Exchange == 'BFO':
                                    isFound, Spot_Token = GetToken('BSE',input_symbol)
                                    spot_ltp = convert_to_float(api.get_quotes("BSE", str(Spot_Token)).get("lp"))
                                else:    
                                    Spot_Token = Fut_Token
                                    spot_ltp = convert_to_float(api.get_quotes(Exchange, str(Spot_Token)).get("lp"))
                                #print(f"Spot_Token {Spot_Token} spot_ltp {spot_ltp}")
                                future_ltp = convert_to_float(api.get_quotes(Exchange, str(Fut_Token)).get("lp"))
                                
                                #print(f"{input_symbol} spot ltp = {spot_ltp} future ltp = {future_ltp}")

                                for strike_dict in List_of_particular_expiry_strike:
                                    
                                    Strike = convert_to_float(strike_dict.get("strike"))
                                    #print(Strike)
                                    PE_Token = strike_dict.get("PE_Token")
                                    PE_Token = str(Exchange)+ "|"  + str(PE_Token)
                                    #print(PE_Token)
                                    CE_Token = strike_dict.get("CE_Token")
                                    CE_Token =  str(Exchange)+ "|" + str(CE_Token)
                                    #print(CE_Token)
                                    
                                    
                                    try:
                                        CE_oi = live_data[str(CE_Token)].get("oi", 0)
                                    except:
                                        CE_oi = 0
                                    try:
                                        CE_poi = live_data[str(CE_Token)].get("poi", 0)
                                    except:
                                        CE_poi = 0
                                    
                                    CE_coi = int(CE_oi) - int(CE_poi)
                                    
                                    try:
                                        CE_toi = live_data[str(CE_Token)].get("toi", "-")
                                    except:
                                        CE_toi = "-"
                                    try:
                                        CE_lp = live_data[str(CE_Token)].get("lp", 0)
                                    except:
                                        CE_lp = 0
                                    try:
                                        CE_pc = live_data[str(CE_Token)].get("pc", "-")
                                    except:
                                        CE_pc = "-"
                                    try:
                                        CE_bq1 = live_data[str(CE_Token)].get("bq1", "-")
                                    except:
                                        CE_bq1 = "-"
                                    try:
                                        CE_bp1 = live_data[str(CE_Token)].get("bp1", "-")
                                    except:
                                        CE_bp1 = "-"
                                    try:
                                        CE_sq1 = live_data[str(CE_Token)].get("sq1", "-")
                                    except:
                                        CE_sq1 = "-"
                                    try:
                                        CE_sp1 = live_data[str(CE_Token)].get("sp1", "-")
                                    except:
                                        CE_sp1 = "-"
                                    
                                    try:
                                        PE_oi = live_data[str(PE_Token)].get("oi", 0)
                                    except:
                                        PE_oi = 0
                                    try:
                                        PE_poi = live_data[str(PE_Token)].get("poi", 0)
                                    except:
                                        PE_poi = 0
                                    
                                    PE_coi = int(PE_oi) - int(PE_poi)
                                    
                                    #print(f"CE COI : {CE_oi} {CE_poi} {CE_coi}")
                                    #print(f"PE COI : {PE_oi} {PE_poi} {PE_coi}")
                                    try:
                                        PE_toi = live_data[str(PE_Token)].get("toi", "-")
                                    except:
                                        PE_toi = "-"
                                    try:
                                        PE_lp = live_data[str(PE_Token)].get("lp", 0)
                                    except:
                                        PE_lp = 0
                                    try:
                                        PE_pc = live_data[str(PE_Token)].get("pc", "-")
                                    except:
                                        PE_pc = "-"
                                    try:
                                        PE_bq1 = live_data[str(PE_Token)].get("bq1", "-")
                                    except:
                                        PE_bq1 = "-"
                                    try:
                                        PE_bp1 = live_data[str(PE_Token)].get("bp1", "-")
                                    except:
                                        PE_bp1 = "-"
                                    try:
                                        PE_sq1 = live_data[str(PE_Token)].get("sq1", "-")
                                    except:
                                        PE_sq1 = "-"
                                    try:
                                        PE_sp1 = live_data[str(PE_Token)].get("sp1", "-")
                                    except:
                                        PE_sp1 = "-"
                                    
                                    try:
                                        CE_v = live_data[str(CE_Token)].get("v", 0)
                                    except:
                                        CE_v = 0
                                        
                                    try:
                                        PE_v = live_data[str(PE_Token)].get("v", 0)
                                    except:
                                        PE_v = 0
                                    
                                    dic_data = {
                                            "CE_token": CE_Token,
                                            "CE_oi": CE_oi,
                                            "CE_poi": CE_poi,
                                            "CE_toi": CE_toi,
                                            "CE_lp": CE_lp,
                                            "CE_pc": CE_pc,
                                            "CE_bq1": CE_bq1,
                                            "CE_bp1": CE_bp1,
                                            "CE_sq1": CE_sq1,
                                            "CE_sp1": CE_sp1,
                                            "strike": Strike,
                                            "PE_bq1": PE_bq1,
                                            "PE_bp1": PE_bp1,
                                            "PE_sq1": PE_sq1,
                                            "PE_sp1": PE_sp1,
                                            "PE_pc": PE_pc,
                                            "PE_lp": PE_lp,
                                            "PE_toi": PE_toi,
                                            "PE_poi": PE_poi,
                                            "PE_oi": PE_oi,
                                            "PE_token": PE_Token,
                                            "CE_coi":CE_coi,
                                            "PE_coi":PE_coi,
                                            "CE_v":CE_v,
                                            "PE_v":PE_v,
                                        }
                                    pd_oc = pd.concat([pd_oc, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                    
                                #pd_oc.to_csv("pd_oc.csv")                            
                                #print(pd_oc)
                                #pd_oc = pd_oc.astype({"strike": float})
                                pd_oc['strike'] = pd_oc['strike'].apply(convert_to_float)
                                pd_oc = pd_oc.sort_values(by="strike", ascending=True)

                                pd_oc = pd_oc.fillna(0)
                                
                                pd_oc = pd_oc.astype({"CE_oi": int})
                                pd_oc = pd_oc.astype({"PE_oi": int})
                                pd_oc = pd_oc.astype({"CE_poi": int})
                                pd_oc = pd_oc.astype({"PE_poi": int})
                                pd_oc = pd_oc.astype({"CE_coi": int})
                                pd_oc = pd_oc.astype({"PE_coi": int})
                                
                                pd_oc = pd_oc.astype({"CE_v": int})
                                pd_oc = pd_oc.astype({"PE_v": int})
                                
                                pd_oc["CE_v"] = pd_oc["CE_v"] / int(Lot_Size)
                                pd_oc["PE_v"] = pd_oc["PE_v"] / int(Lot_Size)
                                
                                pd_oc["CE_oi"] = pd_oc["CE_oi"] / int(Lot_Size)
                                pd_oc["PE_oi"] = pd_oc["PE_oi"] / int(Lot_Size)
                                pd_oc["CE_poi"] = pd_oc["CE_poi"] / int(Lot_Size)
                                pd_oc["PE_poi"] = pd_oc["PE_poi"] / int(Lot_Size)
                                pd_oc["CE_coi"] = pd_oc["CE_coi"] / int(Lot_Size)
                                pd_oc["PE_coi"] = pd_oc["PE_coi"] / int(Lot_Size)
                                pd_oc["OI_SUM"] = pd_oc["CE_oi"] + pd_oc["PE_oi"]

                                #print(pd_oc)
                                
                                df_oc_pro = pd_oc
                                
                                try:
                                    NoOfStrike = int(oci_pro.range("E6").value)
                                except Exception as e:
                                    NoOfStrike = 100
                                
                                df_oc_pro['strike_diff'] = abs(df_oc_pro['strike'] - spot_ltp)
            
                                df_oc_pro.sort_values(by = 'strike_diff',inplace = True)
                                
                                #print(f"\n\n***\n\n{df_oc_pro}")
                                AtmStrike = convert_to_float(df_oc_pro.iloc[0]['strike'])
                                AtmStrikeCallPrice = convert_to_float(df_oc_pro.iloc[0]['CE_lp'])
                                AtmStrikePutPrice = convert_to_float(df_oc_pro.iloc[0]['PE_lp'])
                                
                                #additional detail related to dump on input page
                                Future_LTP = future_ltp
                                Max_Pain_at_Strike = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['strike']
                                Ltp_at_Max_Pain_CE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['CE_lp']
                                Ltp_at_Max_Pain_PE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['PE_lp']
                                #ATM_Strike = 
                                LTP_at_ATM_CE = df_oc_pro[df_oc_pro.strike == AtmStrike].iloc[0]['CE_lp']
                                LTP_at_ATM_PE = df_oc_pro[df_oc_pro.strike == AtmStrike].iloc[0]['PE_lp']
                                Total_OI_CE = df_oc_pro["CE_oi"].sum()
                                Total_OI_PE = df_oc_pro["PE_oi"].sum()
                                
                                Max_OI_CE = df_oc_pro["CE_oi"].max()
                                Max_OI_PE = df_oc_pro["PE_oi"].max()
                                Max_OI_at_Strike_CE = df_oc_pro[df_oc_pro.CE_oi == df_oc_pro["CE_oi"].max()].iloc[0]['strike']
                                Max_OI_at_Strike_PE = df_oc_pro[df_oc_pro.PE_oi == df_oc_pro["PE_oi"].max()].iloc[0]['strike']
                                LTP_of_Max_OI_Strike_CE = df_oc_pro[df_oc_pro.CE_oi == df_oc_pro["CE_oi"].max()].iloc[0]['CE_lp']
                                LTP_of_Max_OI_Strike_PE = df_oc_pro[df_oc_pro.PE_oi == df_oc_pro["PE_oi"].max()].iloc[0]['PE_lp']

                                Total_OI_Change_CE = df_oc_pro["CE_coi"].sum()
                                Total_OI_Change_PE = df_oc_pro["PE_coi"].sum()
                                
                                Max_Change_in_OI_addition_CE = df_oc_pro["CE_coi"].max()
                                Max_Change_in_OI_addition_PE = df_oc_pro["PE_coi"].max()
                                Max_OI_addition_at_Srike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].max()].iloc[0]['strike']
                                Max_OI_addition_at_Srike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].max()].iloc[0]['strike']
                                LTP_of_Max_OI_addition_Strike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].max()].iloc[0]['CE_lp']
                                LTP_of_Max_OI_addition_Strike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].max()].iloc[0]['PE_lp']
                                Max_Change_in_OI_unwinding_CE = -1 * int(df_oc_pro["CE_coi"].min())
                                Max_Change_in_OI_unwinding_PE = -1 * int(df_oc_pro["PE_coi"].min())
                                Max_OI_unwinding_at_Srike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].min()].iloc[0]['strike']
                                Max_OI_unwinding_at_Srike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].min()].iloc[0]['strike']
                                LTP_of_Max_OI_unwinding_Strike_CE = df_oc_pro[df_oc_pro["CE_coi"] == df_oc_pro["CE_coi"].min()].iloc[0]['CE_lp']
                                LTP_of_Max_OI_unwinding_Strike_PE = df_oc_pro[df_oc_pro["PE_coi"] == df_oc_pro["PE_coi"].min()].iloc[0]['PE_lp']
                            
                                df_additional_detail = pd.DataFrame(columns = ['CE','PE'])
                                dic_data = {'CE':Future_LTP}

                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_Pain_at_Strike}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Ltp_at_Max_Pain_CE,'PE': Ltp_at_Max_Pain_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':AtmStrike}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':LTP_at_ATM_CE,'PE':LTP_at_ATM_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Total_OI_CE, 'PE':Total_OI_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Total_OI_Change_CE, 'PE':Total_OI_Change_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_OI_CE,  'PE':Max_OI_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_OI_at_Strike_CE,'PE':Max_OI_at_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':LTP_of_Max_OI_Strike_CE,'PE':LTP_of_Max_OI_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_Change_in_OI_addition_CE,'PE':Max_Change_in_OI_addition_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_OI_addition_at_Srike_CE,'PE':Max_OI_addition_at_Srike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':LTP_of_Max_OI_addition_Strike_CE, 'PE':LTP_of_Max_OI_addition_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_Change_in_OI_unwinding_CE, 'PE':Max_Change_in_OI_unwinding_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':Max_OI_unwinding_at_Srike_CE, 'PE':Max_OI_unwinding_at_Srike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                
                                dic_data = {'CE':LTP_of_Max_OI_unwinding_Strike_CE,'PE':LTP_of_Max_OI_unwinding_Strike_PE}
                                df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                                

                                oci_pro.range("i3").options(index=False, header=False).value = df_additional_detail
                                
                                
                                
                                df_oc_pro = df_oc_pro[:2*int(NoOfStrike)+1]
                                
                                df_oc_pro.sort_values(by ='strike',inplace = True)
                                
                                df_oc_pro['CE_Delta'] = None
                                df_oc_pro['CE_Gamma'] = None
                                df_oc_pro['CE_Theta'] = None
                                df_oc_pro['CE_Vega'] = None
                                df_oc_pro['CE_Rho'] = None
                                df_oc_pro['CE_IV'] = None
                                
                                df_oc_pro['PE_Delta'] = None
                                df_oc_pro['PE_Gamma'] = None
                                df_oc_pro['PE_Theta'] = None
                                df_oc_pro['PE_Vega'] = None
                                df_oc_pro['PE_Rho'] = None
                                df_oc_pro['PE_IV'] = None
                    
                                df_oc_pro = df_oc_pro.reindex(['CE_Delta','CE_Gamma','CE_Theta','CE_Vega','CE_Rho','CE_oi','CE_coi','CE_v','CE_IV','CE_lp','CE_pc','CE_bq1','CE_bp1','CE_sp1','CE_sq1','strike','PE_bq1','PE_bp1','PE_sp1','PE_sq1','PE_pc','PE_lp','PE_IV','PE_v','PE_coi','PE_oi','PE_Rho','PE_Vega','PE_Theta','PE_Gamma','PE_Delta'], axis=1)
                                
                                SpotPrice = convert_to_float(spot_ltp)
                                FuturePrice = convert_to_float(future_ltp)
                                ExpiryDateTime = dt(expiry_input.year, expiry_input.month, expiry_input.day, 0, 0, 0)
                                
                                #print(f"SpotPrice = {SpotPrice}, FuturePrice={FuturePrice}, AtmStrike={AtmStrike}, AtmStrikeCallPrice={AtmStrikeCallPrice}, AtmStrikePutPrice={AtmStrikePutPrice}, ExpiryDateTime={ExpiryDateTime}")
                                
                                ExpiryType = oci_pro.range("F7").value
                                GreekMatch = oci_pro.range("F8").value
                                
                                if ExpiryType == 'WEEKLY':
                                    ExpiryDateType = ExpType.WEEKLY
                                else:
                                    ExpiryDateType = ExpType.MONTHLY
                                
                                FromDateTime = dt.now() 
                                if Exchange == 'NFO':
                                    if dt.now().time() > time(15, 30, 0):
                                        FromDateTime = dt(dt.now().year, dt.now().month,dt.now().day, 15, 30, 0)
                                    
                                    
                                if GreekMatch == "SENSIBULL":
                                    tryMatchWith=TryMatchWith.SENSIBULL
                                else:
                                    tryMatchWith=TryMatchWith.NSE
                                
                                dayCountType = DayCountType.CALENDARDAYS
                                
                                IvGreeks = CalcIvGreeks( SpotPrice = SpotPrice,  FuturePrice = FuturePrice, AtmStrike = AtmStrike, AtmStrikeCallPrice = AtmStrikeCallPrice, AtmStrikePutPrice = AtmStrikePutPrice, ExpiryDateTime = ExpiryDateTime, ExpiryDateType = ExpiryDateType, FromDateTime = FromDateTime, tryMatchWith = tryMatchWith, dayCountType = dayCountType)
                
                                #print(f"SpotPrice={SpotPrice}, FuturePrice={FuturePrice},  AtmStrike={AtmStrike}, AtmStrikeCallPrice={AtmStrikeCallPrice}, AtmStrikePutPrice={AtmStrikePutPrice}, ExpiryDateTime={ExpiryDateTime},  ExpiryDateType={ExpiryDateType}, FromDateTime={FromDateTime}, tryMatchWith={tryMatchWith}")
                                
                                for ind in df_oc_pro.index:
                                    
                                    StrikePrice= convert_to_float(df_oc_pro['strike'][ind])
                                    StrikeCallPrice= convert_to_float(df_oc_pro['CE_lp'][ind])
                                    StrikePutPrice= convert_to_float(df_oc_pro['PE_lp'][ind])
                                    #print(f"StrikePrice={StrikePrice}, StrikeCallPrice={StrikeCallPrice}, StrikePutPrice={StrikePutPrice}")
                                    Greeks = IvGreeks.GetImpVolAndGreeks( StrikePrice = StrikePrice, StrikeCallPrice = StrikeCallPrice, StrikePutPrice = StrikePutPrice)
                                    #print(Greeks)
                                    
                                    df_oc_pro['CE_Delta'][ind] = round(Greeks["CallDelta"],2)
                                    df_oc_pro['CE_Gamma'][ind] = round(Greeks["Gamma"],4)
                                    df_oc_pro['CE_Theta'][ind] = round(Greeks["Theta"],2)
                                    df_oc_pro['CE_Vega'][ind] = round(Greeks["Vega"],2)
                                    df_oc_pro['CE_Rho'][ind] = round(Greeks["RhoCall"],4)
                                    
                                    
                                    df_oc_pro['PE_Delta'][ind] = round(Greeks["PutDelta"],2)
                                    df_oc_pro['PE_Gamma'][ind] = round(Greeks["Gamma"],4)
                                    df_oc_pro['PE_Theta'][ind] = round(Greeks["Theta"],2)
                                    df_oc_pro['PE_Vega'][ind] = round(Greeks["Vega"],2)
                                    df_oc_pro['PE_Rho'][ind] = round(Greeks["RhoPut"],4)
                                    
                                    if GreekMatch == "NSE":
                                        df_oc_pro['CE_IV'][ind] = round(Greeks["CallIV"],2)
                                        df_oc_pro['PE_IV'][ind] = round(Greeks["PutIV"],2)
                                    else:
                                        df_oc_pro['CE_IV'][ind] = round(Greeks["ImplVol"],2)
                                        df_oc_pro['PE_IV'][ind] = round(Greeks["ImplVol"],2)
                                
                                del IvGreeks
                                df_oc_pro.round({"CE_lp":2, 'PE_lp':2})
                                #print(df_oc_pro)
                                
                                if pre_selected_NoOfStrike != NoOfStrike:
                                    pre_selected_NoOfStrike = NoOfStrike
                                    Option_Chain_Pro_Output.range('a3:ae500').value = None
                                    Option_Chain_Pro_Output.range(f"a3:AE500").color = (255,255,255)
                                    Option_Chain_Pro_Output.range(f"p3:p{2 * int(NoOfStrike) + 3}").color = (46,132,198) 
                                
                                df_oc_pro = df_oc_pro.reset_index(drop = True)
                                ATM_pos = df_oc_pro[df_oc_pro.strike == AtmStrike].index.values[0]
                                if NoOfStrike * 2 < len(df_oc_pro):
                                    df_oc_pro = df_oc_pro.iloc[ATM_pos - int(NoOfStrike) : ATM_pos + int(NoOfStrike) + 1]
                                    ATM_Row = int(NoOfStrike) + 3
                                    Option_Chain_Pro_Output.range(f"a{ATM_Row}:AE{ATM_Row}").color = (46,132,198)
                        
                                Option_Chain_Pro_Output.range('a3').options(index=False,header=False).value = df_oc_pro
                                
                            except Exception as e:
                                Option_Chain_Pro_Output.range('a3:ae500').value = None
                                oci_pro.range('I3:J18').value = None
                                Message = "Please check the all provided detail to load option chain:" + str(e)
                                print(Message)
                                oci_pro.range("F4").value = Message
                                
                        else:
                            Option_Chain_Pro_Output.range('a3:ae500').value = None
                            oci_pro.range('I3:J18').value = None
                            Message = "Please enter correct expiry in dd-mm-YYYY (date format)"
                            print(Message)
                            oci_pro.range("F4").value = Message
                            
                    else:
                        Option_Chain_Pro_Output.range('a3:ae500').value = None
                        oci_pro.range('I3:J18').value = None
                        Message = "Please enter the expiry"
                        print(Message)
                        oci_pro.range("F4").value = Message
                        
                else:
                    Option_Chain_Pro_Output.range('a3:ae500').value = None
                    oci_pro.range('I3:J18').value = None
                    Message = "Please enter correct symbol"
                    print(Message)
                    oci_pro.range("F3").value = Message
                    oci_pro.range("F4").value = None
                    oci_pro.range("b2:c100").value = None
                    
            else:
                Option_Chain_Pro_Output.range('a3:ae500').value = None
                oci_pro.range('I3:J18').value = None
                Message = "Please enter the symbol"
                print(Message)
                oci_pro.range("F3").value = Message
                oci_pro.range("F4").value = None
                oci_pro.range("b2:c100").value = None
                
        except Exception as e:
            print(f"Excption : {e}")
            pass
        sleep(int(IterationSleep))

def CloseTrade():
    global api
    print("I am inside CloseTrade")
    try:
        df_open_position_net,OverAllPnL  = get_position()

        price_type = "MKT"
        price = 0.0
        
        df_open_position_net['Net Quantity'] = df_open_position_net['Net Quantity'].astype('int')
        df_open_position_net = df_open_position_net[df_open_position_net['Net Quantity']!= 0 ] #make it 0 
        
        if(len(df_open_position_net) > 0):
            
            df_position_net_short = df_open_position_net[df_open_position_net['Net Quantity'] < 0]
            
            if(len(df_position_net_short) > 0 ):
                for ind in df_position_net_short.index:

                    try:
                    
                        api.place_order(
                            buy_or_sell='B',
                            product_type=df_position_net_short['Product'][ind],
                            exchange=df_position_net_short['Exchange'][ind],
                            tradingsymbol=df_position_net_short['Symbol'][ind],
                            quantity=abs(df_position_net_short['Net Quantity'][ind]),
                            discloseqty=0,
                            price_type=price_type,
                            price=price,
                            trigger_price=None,
                            retention="DAY",
                            remarks="Python_Trader_UserSelected_squareoff",
                        )
                    
                    except Exception as e:
                        Message =  str(e) + " : Exception occur in Finvasia order placement"
                        print(Message) 
            
            df_position_net_long = df_open_position_net[df_open_position_net['Net Quantity'] > 0]
            
            if(len(df_position_net_long) > 0 ):
                for ind in df_position_net_long.index:
                    
                    try:
                        api.place_order(
                            buy_or_sell='S',
                            product_type=df_position_net_long['Product'][ind],
                            exchange=df_position_net_long['Exchange'][ind],
                            tradingsymbol=df_position_net_long['Symbol'][ind],
                            quantity=abs(df_position_net_long['Net Quantity'][ind]),
                            discloseqty=0,
                            price_type=price_type,
                            price=price,
                            trigger_price=None,
                            retention="DAY",
                            remarks="Python_Trader_UserSelected_squareoff",)
                        
                    except Exception as e:
                        Message =  str(e) + " : Exception occur in finvasia order placement"
                        print(Message)

        else:
            Message = "No Open position found"
            print(Message)
    except Exception as e:
        print(f"Exception occur in CloseTrade : {e}")

df_orderbook = pd.DataFrame()
df_openPosition = pd.DataFrame()

def start_Open_Position():
    
    excel_op = xw.Book(TerminalSheetName)
    op_op = excel_op.sheets("OpenPosition")
    op_tt = excel_op.sheets("Trade_Terminal")
    op_hold = excel_op.sheets("Holdings")
    op_config = excel_op.sheets("Config")
    op_ob = excel_op.sheets("OrderBook")
    
    isTelegramEnable = False
    if op_config.range("b3").value == True:
        isTelegramEnable = True
    
    isVoiceEnable = False
    if op_config.range("b6").value == True:
        isVoiceEnable = True
        
    op_op.range(f"d2").value = False
    op_op.range(f"e2").value = 0
    op_op.range("b3:az1000").value = None
    op_hold.range("a1:w500").value  = None
        
    op_ob.range("b1:az100").value = None
    
    global LimitOrderBook
    global df_orderbook, df_openPosition
    global api
    global Telegram_Message, Voice_Message
    
    while True:
        try:
            get_limits = api.get_limits()
            
            op_tt.range("a2").value = get_limits['cash']
            try:
                op_tt.range("b2").value = get_limits['marginused']
                op_tt.range("c2").value = get_limits['expo']
                op_tt.range("d2").value = get_limits['span']
            except Exception as e:
                op_tt.range("b2").value = 0
                op_tt.range("c2").value = 0
                op_tt.range("d2").value = 0
             
             
            df_openPosition, OverAllPnL  = get_position()
            if(len(df_openPosition) > 0):
                op_op.range(f"a2").value = OverAllPnL
                op_tt.range(f"f2").value = OverAllPnL
                if excel_op.sheets.active.name == "OpenPosition":
                    op_op.range("b3").options(index=False,header=True).value = df_openPosition
                    
                    KillSwitch = op_op.range(f"d2").value
                    Reconfirm = int(op_op.range(f"e2").value)
                    #print(f"KillSwitch={KillSwitch} Reconfirm={Reconfirm}")
                    if(KillSwitch == 'Execute' and Reconfirm == 1):
                        CloseTrade()        
                        op_op.range(f"d2").value = False
                        op_op.range(f"e2").value = 0
                    else:
                        
                        #print("check if any open order needs to be cancel")
                        UserAction = op_op.range(f"a{4}:a{3 + len(df_openPosition)}").value
                        #print(f"UserAction = {UserAction}")
                        if UserAction != None:
                            for i in range (0,len(df_openPosition)):
                                #print(f"{i} : {UserAction[i]}")
                                if UserAction[i] == 'Square_Off' or UserAction[i] == 'S':
                                    #print(f"*{df_openPosition}**")
                                    exchange = df_openPosition['Exchange'][i]
                                    tradingsymbol = df_openPosition['Symbol'][i]
                                    if df_openPosition['Net Quantity'][i] < 0:
                                        buy_or_sell = 'B'
                                    else:
                                        buy_or_sell = 'S'
                                    quantity = abs(df_openPosition['Net Quantity'][i])
                                    product_type = df_openPosition['Product'][i]
                                    tag = "PT individual Exit"
                                    price_type = "MKT"
                                    price = 0.0
                                    #print(f"squareoff {tradingsymbol}  {exchange} {product} {quantity}{transaction_type}")
                                    try:
                                        api.place_order(buy_or_sell=buy_or_sell,product_type=product_type,exchange=exchange,tradingsymbol=tradingsymbol,quantity=quantity,discloseqty=0,price_type=price_type,price=price,trigger_price=None,retention="DAY",remarks=tag)
                                    except Exception as e:
                                        print(f"squareoff can't be done for {df_openPosition['Symbol'][i]} with reason : {e}")
                                        pass
                                    op_op.range(f"a{4+i}").value = None
        except Exception as e:
            print(f"Exception occur in OpenPosition : {e}")
            pass
        
        try:
            if excel_op.sheets.active.name == "Holdings":
                df_holding = getholdings()
                if(len(df_holding)> 0 ):
                    op_hold.range("a1").options(index=False,header=True).value = df_holding
        except Exception as e:
            print(f"Exception occur in Holdings : {e}")
            pass
        
        try:
            if excel_op.sheets.active.name == "OrderBook":
                df_orderbook = get_order_book()
                if (len(df_orderbook) > 0):
                    op_ob.range("b1").options(index=False,header=True).value = df_orderbook
                    
                    
                    #check if any open order needs to be cancel
                    UserAction = op_ob.range(f"a{2}:a{1 + len(df_orderbook)}").value
                    #print(f"UserAction = {UserAction}")
                    if UserAction is not None:
                        for i in range (0,len(df_orderbook)):
                            #print(f"{i} : {UserAction[i]}")
                            if UserAction[i] == 'CANCEL' or UserAction[i] == 'C':
                                #print(f"Cancel order id {df_orderbook['order_id'][i]} variety {df_orderbook['variety'][i]}")
                                try:
                                    api.cancel_order(orderno=df_orderbook['Order No'][i])
                                except Exception as e:
                                    print(f"Order can't be cancelled with reason : {e}")
                                    pass
                                op_ob.range(f"a{2+i}").value = None
                
        except Exception as e:
            print(f"Exception occur in OrderBook : {e}")
            pass
        
        try:
            
            for key, value in LimitOrderBook.items():
                #print(f"check status of {key} , having current status {value['status']}")
                if value['status'] == 'PENDING':
                    status, Executed_price = order_status (key)
                    #print(f"filled_quantity = {filled_quantity},Executed_price={Executed_price}, status={status} ")
                    if status == 'COMPLETE':
                        value['status'] = 'COMPLETE'
                        value['Executed_price'] = Executed_price
                    elif status in ['CANCELLED','CANCELED']:
                        value['status'] = 'CANCELLED'
                    elif status in ['REJECTED']:
                        value['status'] = 'REJECTED'
                    value['Remarks'] = status
                        
                #print(f"current status of {key} , with updated values {value['status']}")        
                
        except Exception as e:
            #print(f"Exception in open position during LimitOrderBook checking : {e}")
            pass
        
        #print(f"Telegram_Message queue = {Telegram_Message}")
        try:
            if isTelegramEnable:
                if len(Telegram_Message) > 0:
                    SendMessageToTelegram(Telegram_Message[0])
                    del Telegram_Message[0]
        
        
        except Exception as e:
            #print(f"Exception in telegram Send Message : {e}")
            pass
            
        try:
            if isVoiceEnable:
                if len(Voice_Message) > 0:
                    Text2Speech(Voice_Message[0])
                    del Voice_Message[0]
        
        
        except Exception as e:
            #print(f"Exception in voice Message : {e}")
            pass
                
def getholdings():  
    #print("I am inside getholdings")
    global api

    df_holding = pd.DataFrame(columns=['Exchange','Name','Holding quantity','Non Poa display quantity','Average price'])
    holdings = api.get_holdings()
    for holding in holdings:
        exch = holding['exch_tsym'][0]['exch']
        tsym = holding['exch_tsym'][0]['tsym']
        holdqty = holding['holdqty']
        try:
            npoadqty = holding['npoadqty']
        except Exception as e:
            npoadqty = 0
        upldprc = holding['upldprc']
        
        #print(f"{exch} {tsym} {holdqty} {npoadqty} {upldprc}")
        dic_data = {'Exchange':exch,'Name':tsym,'Holding quantity':holdqty,'Non Poa display quantity':npoadqty,'Average price':upldprc}
        df_holding = pd.concat([df_holding, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
    return df_holding
    
def StartThread():
    excel_name = xw.Book(TerminalSheetName)
    Config_sheet = excel_name.sheets['Config']
    try:
        
        # Define the threads and put them in an array
        threads = []
        
        if Config_sheet.range("b2").value == True:
            threads.append(Thread(target=start_Trade_Terminal))
            threads.append(Thread(target=start_Open_Position))      
           
        if Config_sheet.range("b4").value == True:
            threads.append(Thread(target=start_optionchain))
        
        if Config_sheet.range("b5").value == True:
            threads.append(Thread(target=start_optionchain_Pro))
            
        if len(threads) != 0 :    
            # Func1 and Func2 run in separate threads
            for thread in threads:
                thread.start()

            # Wait until both Func1 and Func2 have finished
            for thread in threads:
                thread.join()
        else:
            print("Please select atlease one feature in config and restart the algo")
        
    except Exception as e:
        Message = str(e) + " : Exception occur"
        print(Message)

print(f"Python Trader Excel Based Terminal program initialised")
if Shoonya_login() == 1:
    LoadInstrument_token()
    
    
    api.start_websocket(
        order_update_callback=event_handler_order_update,
        subscribe_callback=event_handler_quote_update,
        socket_open_callback=open_callback,
        socket_close_callback=event_handler_socket_closed,
    )

    while feed_opened == False:
        print("Trying to connect WebSocket...")
        pass

    print("Connected to WebSocket...")

    StartThread()
    print("Enjoy the automation...")
else:
    print("\n\nAlgo is not able to login using your given credential. Please follow below steps in matrix wise.\n\n1. Check entered userid/password/apikey/otp is correct or not. Try to login your finvasia account using same credential.\n2. If issue still exist please regenerate your password and api key and update the sheet.\n3. If you are using algo first time, please wait for 24 hours to activate your api.\n4. If issue still exist please contact to finvasia support team.")

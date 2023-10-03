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
from kiteext import KiteExt
import os, json, sys
from datetime import datetime as dt, timedelta, time
from time import sleep
import logging
import copy 
import threading
from threading import Thread
import platform

#print(f"001 :{dt.now()}")

try:
    from tzlocal import get_localzone
except (ModuleNotFoundError, ImportError):
    print("tzlocal module not found")
    os.system(f"{sys.executable} -m pip install -U tzlocal")
finally:
    from tzlocal import get_localzone
    
try:
    import numpy as np 
except (ModuleNotFoundError, ImportError):
    print("numpy module not found")
    os.system(f"{sys.executable} -m pip install -U numpy")
finally:
    import numpy as np 
    
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
    import psutil
except (ModuleNotFoundError, ImportError):
    print("psutil module not found")
    os.system(f"{sys.executable} -m pip install -U psutil")
finally:
    import psutil
    
try:
    import pandas as pd
except (ModuleNotFoundError, ImportError):
    print("pandas module not found")
    os.system(f"{sys.executable} -m pip install -U pandas")
finally:
    import pandas as pd
    

print("HI welcome")

def Text2Speech(Text):
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

TradeTerminalFileName = 'Zerodha_Trade_Terminal_V3.xlsm'
#logging.basicConfig(level=logging.DEBUG)

def Zerodha_login():
    print(f"I am inside Zerodha_login")
    global kite, excel_master
    global TelegramBotCredential, ReceiverTelegramID
    global UserID, client_name
    
    isConnected = 0
    Credential_sheet = excel_master.sheets['User_Credential']
    try:
        TelegramBotCredential = str(Credential_sheet.range("b8").value)
        ReceiverTelegramID = str(Credential_sheet.range("b9").value)
        index = ReceiverTelegramID.find(".")
        if index != -1:
            ReceiverTelegramID = ReceiverTelegramID[:len(ReceiverTelegramID)-2]
        
        kite = KiteExt()

        Credential_sheet.range('a1').value = 'Welcome To Python Trader'
        Credential_sheet.range('a1').color = (46,132,198)
        Credential_sheet.range('a11').value = 'Tool developed by PythonTrader, Please follow us on social media site for more tools or freelancing work'
        Credential_sheet.range('a11').color = (46,132,198)
        Credential_sheet.range('b12').value = 'https://www.youtube.com/@pythontraders'
        Credential_sheet.range('b13').value =  'https://www.t.me/pythontrader'
        Credential_sheet.range('b12:b13').color = (220,214,32)
        
        UserID = Credential_sheet.range('B2').value.strip()

        LoginMethod = str(Credential_sheet.range("B3").value)
        if (LoginMethod == "New_Session"):
            Password = str(Credential_sheet.range('B4').value)
            
            index = Password.find(".")
            if index != -1:
                Password = Password[:len(Password)-2]

            Totp = str(Credential_sheet.range('B5').value)
            index = Totp.find(".")
            if index != -1:
                twoFA = int(Totp[:6])
            else:
                pin = pyotp.TOTP(Totp).now()
                twoFA = f"{int(pin):06d}" if len(pin) <=5 else pin    
            
            print(f"UserID={UserID},Password={Password},Totp={Totp},twoFA={twoFA}")
            
            enctoken, public_token = kite.login_with_credentials(userid=UserID, password=Password, twofa=twoFA)
            
            #print(f"enctoken = ({enctoken}), public_token = ({public_token}) ")

        else:
        
            enctoken = str(Credential_sheet.range('B6').value)
            #print(f"UserID=({UserID}), enctoken=({enctoken})")
            kite.login_using_enctoken(userid=UserID, enctoken=enctoken, public_token=None) 
        
        Profile = kite.profile()
        print(Profile)
        client_name = Profile.get('user_name')
        Welcome_Message = "Login Successful, Welcome " + client_name + "\nTool Validity : Demo" + "\nToken = (" + str(enctoken) + ")"
        
        UserID = Profile['user_id']
        Text2Speech("Login Successful, Welcome " + str(client_name) )
        
        Credential_sheet.range('c2').value = Welcome_Message
        Credential_sheet.range('c2').color = (118,224,280)
        isConnected = 1
        
    except Exception as e:
        print(f"Error : {e}")
        Credential_sheet.range('c2').value = 'Wrong credential'
        Text2Speech("Login Unsuccessful due to wrong or expired credential/token , kindly check your provided details")
        Credential_sheet.range('c2').color = (255, 0, 0)

    return isConnected
    
SYMBOLDICT = {}
live_data = {}
Token_yet_to_subscribe = []
Telegram_Message = ["Welcome to Python Trader excel based trade terminal","Have a Good Day"]
Voice_Message = []
TelegramBotCredential = None
ReceiverTelegramID = None

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
            
def on_ticks(ws, ticks):
    # Callback to receive ticks.
    global Token_yet_to_subscribe, live_data
    #print("Ticks data: {}".format(ticks))
    
    for stock in ticks:
        #print(f"inside tick for loop : {stock}")
        try:
            volume = stock["volume_traded"]
        except:
            volume = 0
        try:
            OI = stock["oi"]
        except:
            OI = 0
        try:
            Vwap = stock["average_traded_price"]
        except:
            Vwap = 0
        try:
            bp1 = stock["depth"]["buy"][0]["price"]
        except:
            bp1 = 0
        try:
            sp1 = stock["depth"]["sell"][0]["price"]
        except:
            sp1 = 0        
        live_data[stock['instrument_token']] = {"Open": stock["ohlc"]["open"],
                                                                      "High": stock["ohlc"]["high"],
                                                                      "Low": stock["ohlc"]["low"],
                                                                      "Close": stock["ohlc"]["close"],
                                                                      "LTP": stock["last_price"],
                                                                      "Volume": volume,
                                                                      "OI": OI,
                                                                      "Vwap": Vwap,
                                                                      "change" : stock["change"],
                                                                      "bp1":bp1,
                                                                      "sp1":sp1}
        #print(f"after appending live data became : {live_data}")
    if(len(Token_yet_to_subscribe) > 0):
        try:
            #print(f"New symbol found, so going to subscribe, token= {Token_yet_to_subscribe}")
            kws.subscribe(Token_yet_to_subscribe)
            kws.set_mode(ws.MODE_FULL, Token_yet_to_subscribe)
            Token_yet_to_subscribe = []
        except Exception as e:
            #print(str(e)+":Exception in on_ticks")
            pass
            
def on_connect(ws, response):
    global Initial_Subscribr_TokenList
    # Callback on successful connect.
    #mode available MODE_FULL,MODE_LTP,MODE_QUOTE
    ws.subscribe(Initial_Subscribr_TokenList)
    ws.set_mode(ws.MODE_FULL, Initial_Subscribr_TokenList)

def on_error(ws, code, reason):
    logging.error('Ticker errored out. code = %d, reason = %s', code, reason)

def on_close(ws, code, reason):
    # On connection close stop the event loop.
    # Reconnection will not happen after executing `ws.stop()`
    ws.stop()

def on_order_update(ws, data):
    logging.info('Ticker: order update %s', data)

# Callback when connection closed with error.
def on_error(ws, code, reason):
    logging.info("Connection error: {code} - {reason}".format(code=code, reason=reason))


# Callback when reconnect is on progress
def on_reconnect(ws, attempts_count):
    logging.info("Reconnecting: {}".format(attempts_count))


# Callback when all reconnect failed (exhausted max retries)
def on_noreconnect(ws):
    logging.info("Reconnect failed.")
    
def stop_ticker():
    logging.info('Ticker: stopping..')
    kws.close(1000, "Manual close")

def on_max_reconnect_attempts(ws):
    logging.error('Ticker max auto reconnects attempted and giving up..')

def GetToken(exchange,tradingsymbol):
    global df_instrument
    
    Token = df_instrument[(df_instrument.tradingsymbol == tradingsymbol) & (df_instrument.exchange == exchange)].iloc[0]['instrument_token']
    
    return Token

def order_status (orderid):
    #print(f"I am at order_status function with order id {orderid}")
    filled_quantity = 0
    AverageExecutedPrice = 0
    global kite
    try:
        order_history = kite.order_history(orderid)
        filled_quantity = order_history[-1].get('filled_quantity')
        AverageExecutedPrice = order_history[-1].get('average_price')
        status = order_history[-1].get('status')
    except Exception as e:
        Message = str(e) + " : Exception occur in order_status"
        #print(Message)
        pass
    return filled_quantity, AverageExecutedPrice, status
    
def place_trade(tradingsymbol_exchange,quantity,transaction_type , order_type = None, price = None):
    
    global kite
    global Product_type
    try:
        exchange = tradingsymbol_exchange[:3]
        tradingsymbol = tradingsymbol_exchange[4:]
              
        if(Product_type == 'MIS'):
            product = kite.PRODUCT_MIS
        else:
            if(exchange == 'NSE'):
                product = kite.PRODUCT_CNC
            else:
                product = kite.PRODUCT_NRML
        
            
        variety =  kite.VARIETY_REGULAR
        
        
        if order_type == 'MARKET':
            price = 0
            trigger_price = None
        elif order_type == 'LIMIT':
            trigger_price = None
        elif order_type == 'SL-M':
            order_type = 'SL'
            trigger_price = price
                
            if transaction_type == 'BUY':
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
            
        
        Message = "Order placed for " + str (tradingsymbol) + " " + str(quantity) + " quantity " + str(transaction_type)
        print(Message)
        Telegram_Message.append(Message)
        Voice_Message.append(Message)
    
        order_id = None
        order_id = kite.place_order(tradingsymbol = tradingsymbol,
                                exchange = exchange,
                                transaction_type = transaction_type,
                                quantity = int(quantity),
                                price = price,
                                trigger_price=trigger_price,
                                variety = variety,
                                order_type = order_type,
                                product = product)
        print(f"Order id = {order_id}")
        
       
    except Exception as e:
        Message = "Order rejected with error : " + str(e)
        print(Message)
        Telegram_Message.append(Message)
    return order_id

def GetMarginDetail(segment  = 'equity'):
    global kite
    Fund_Detail = kite.margins()
    Available_Fund = Fund_Detail[segment]['net']
    Exposure_Margin = Fund_Detail[segment]['utilised']['exposure']
    SPAN_Margin = Fund_Detail[segment]['utilised']['span']
    
    return float(Available_Fund),float(Exposure_Margin),float(SPAN_Margin)

df_orderbook = pd.DataFrame()
df_openPosition = pd.DataFrame()
def start_Open_Position():
    
    excel_op = xw.Book(TradeTerminalFileName)
    op_op = excel_op.sheets("OpenPosition")
    op_tt = excel_op.sheets("Trade_Terminal")
    op_config = excel_op.sheets("Config")
    
    isTelegramEnable = False
    if op_config.range("b3").value == True:
        isTelegramEnable = True
    
    isVoiceEnable = False
    if op_config.range("b6").value == True:
        isVoiceEnable = True
        
    op_tt.range(f"r2").value = 0
    op_op.range(f"a2").value = 0
    op_hold = excel_op.sheets("Holdings")
    op_op.range("b3:k500").value  = None
    op_hold.range("a1:w500").value  = None
    op_ob = excel_op.sheets("OrderBook")
    op_ob.range("b1:p500").value = None
    op_ob.range("a2:a500").value = None
    op_op.range("a4:a100").value = None
    global LimitOrderBook
    global df_orderbook, df_openPosition
    global kite
    global Telegram_Message, Voice_Message
    
    while True:
        try:
            #print("start_Open_Position thread running")
            Available_Fund,Exposure_Margin,SPAN_Margin = GetMarginDetail()
                
            op_tt.range("a2").value = Available_Fund
            op_tt.range("b2").value = Exposure_Margin
            op_tt.range("c2").value = SPAN_Margin
            
            Available_Fund,Exposure_Margin,SPAN_Margin = GetMarginDetail('commodity')
            
            op_tt.range("e2").value = Available_Fund
            op_tt.range("f2").value = Exposure_Margin
            op_tt.range("g2").value = SPAN_Margin
                
                
            df_openPosition = get_position()
            #print(df_openPosition)
            if(len(df_openPosition) > 0):
                OverAllPnL = GetOverAllPnL()
                op_tt.range(f"r2").value = OverAllPnL
                op_op.range(f"a2").value = OverAllPnL
                if excel_op.sheets.active.name == "OpenPosition":
                    df_openPosition = RemoveUnwantedColumn(df_openPosition)
                    op_op.range("b3").options(index=False,header=True).value = df_openPosition
                    
                    KillSwitch = op_op.range(f"d2").value
                    Reconfirm = int(op_op.range(f"e2").value)
                    if(KillSwitch == 'Execute' and Reconfirm == 1):
                        isClosed = CloseTrade()        
                        op_op.range(f"d2").value = False
                        op_op.range(f"e2").value = 0
                    else:
                        #print("check if any open order needs to be cancel")
                        UserAction = op_op.range(f"a{4}:a{3 + len(df_openPosition)}").value
                        #print(f"UserAction = {UserAction}")
                        for i in range (0,len(df_openPosition)):
                            #print(f"{i} : {UserAction[i]}")
                            if UserAction[i] == 'Square_Off' or UserAction[i] == 'S':
                                exchange = df_openPosition['exchange'][i]
                                tradingsymbol = df_openPosition['tradingsymbol'][i]
                                product = df_openPosition['product'][i]
                                if df_openPosition['quantity'][i] < 0:
                                    transaction_type = 'BUY'
                                else:
                                    transaction_type = 'SELL'
                                quantity = abs(df_openPosition['quantity'][i])
                                tag = "PT user Exit"
                                #print(f"squareoff {tradingsymbol}  {exchange} {product} {quantity}{transaction_type}")
                                try:
                                    kite.place_order(variety = 'regular',
                                        exchange = exchange,
                                        tradingsymbol = tradingsymbol,
                                        transaction_type = transaction_type,
                                        quantity = quantity,
                                        product = product,
                                        order_type = 'MARKET',
                                        tag = tag)
                                except Exception as e:
                                    #print(f"squareoff can't be done for {df_openPosition['tradingsymbol'][i]} with reason : {e}")
                                    pass
                                op_op.range(f"a{4+i}").value = None
                                
            if excel_op.sheets.active.name == "Holdings":
                df_holdings = getholdings()
                if(len(df_holdings) > 0):
                    op_hold.range("a1").options(index=False,header=True).value = df_holdings
            
            
            if excel_op.sheets.active.name == "OrderBook":
                df_orderbook = get_order_book()
                #print(df_orderbook)
                if(len(df_orderbook) > 0):
                    op_ob.range("b1").options(index=False).value = df_orderbook
                    #check if any open order needs to be cancel
                    UserAction = op_ob.range(f"a{2}:a{1 + len(df_orderbook)}").value
                    #print(f"UserAction = {UserAction}")
                    if UserAction is not None:
                        for i in range (0,len(df_orderbook)):
                            #print(f"{i} : {UserAction[i]}")
                            if UserAction[i] == 'CANCEL' or UserAction[i] == 'C':
                                #print(f"Cancel order id {df_orderbook['order_id'][i]} variety {df_orderbook['variety'][i]}")
                                try:
                                    kite.cancel_order(variety=df_orderbook['variety'][i], order_id=df_orderbook['order_id'][i])
                                except Exception as e:
                                    print(f"Order can't be cancelled with reason : {e}")
                                    pass
                                op_ob.range(f"a{2+i}").value = None

        except Exception as e:
            #print(f"Exception in open position {e}")
            pass
        #sleep(5)
        try:
            
            for key, value in LimitOrderBook.items():
                #print(f"check status of {key} , having current status {value['status']}")
                if value['status'] == 'PENDING':
                    filled_quantity, Executed_price, status = order_status (key)
                    #print(f"filled_quantity = {filled_quantity},Executed_price={Executed_price}, status={status} ")
                    if status == 'COMPLETE':
                        value['status'] = 'COMPLETE'
                        value['Executed_price'] = Executed_price
                    elif status == 'CANCELLED':
                        value['status'] = 'CANCELLED'
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
    global kite
    df_holdings = pd.DataFrame()
    try:
        holding = kite.holdings()
        
        df_holdings = pd.DataFrame(holding)
        df_holdings = df_holdings.drop(['authorisation'],axis=1)
    except Exception as e:
        #print(f"Exception in getholdings : {e}")
        pass
    return df_holdings
    
def CloseTrade():
    global kite

    position = kite.positions()
    df_position_net = pd.DataFrame(position['net'])
    
    df_open_position_net = df_position_net[df_position_net.quantity != 0 ]
    
    if(len(df_open_position_net) > 0):
        
        df_position_net_short = df_position_net[df_position_net.quantity < 0]
        
        if(len(df_position_net_short) > 0 ):
            for ind in df_position_net_short.index:

                try:
                
                    kite.place_order(variety = 'regular',
                    exchange = df_position_net_short['exchange'][ind],
                    tradingsymbol = df_position_net_short['tradingsymbol'][ind],
                    transaction_type = 'BUY',
                    quantity = abs(df_position_net_short['quantity'][ind]),
                    product = df_position_net_short['product'][ind],
                    order_type = 'MARKET',
                    tag = "short_exit")
                
                except Exception as e:
                    Message =  str(e) + " : Exception occur in zerodha order placement"
                    print(Message) 
        
        df_position_net_long = df_position_net[df_position_net.quantity > 0]
        
        if(len(df_position_net_long) > 0 ):
            for ind in df_position_net_long.index:
                
                try:
                    kite.place_order(variety = 'regular',
                    exchange = df_position_net_long['exchange'][ind],
                    tradingsymbol = df_position_net_long['tradingsymbol'][ind],
                    transaction_type = 'SELL',
                    quantity = abs(df_position_net_long['quantity'][ind]),
                    product = df_position_net_long['product'][ind],
                    order_type = 'MARKET',
                    tag = "long_exit")
                    
                except Exception as e:
                    Message =  str(e) + " : Exception occur in zerodha order placement"
                    print(Message)

    else:
        Message = "No Open position found"

LimitOrderBook = {}
def start_Trade_Terminal():
    global excel_TT
    excel_TT = xw.Book(TradeTerminalFileName)
    global live_data , kws
    global SYMBOLDICT
    global Token_yet_to_subscribe
    global Product_type
    global Trade_Mode
    global Telegram_Message, Voice_Message
    tt = excel_TT.sheets("Trade_Terminal")
    tt.range("a2:c2").value  = 0
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
    
    subs_lst = []
    Symbol_Token = {}
    global LimitOrderBook
    #run a parallel thread to update the status
    while True:
        try:
            #sleep(.5)
            
            Product_type = tt.range(f"P2").value
            
            
            #print(live_data)
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
                            Token = GetToken(exchange,tradingsymbol)
                            Symbol_Token[i] =  int(Token)
                            Token_yet_to_subscribe.append(int(Token))
                            print(f"Symbol = {i}, Token={Token} subscribed")
                            
                        except Exception as e:
                            print(f"Subscribe error {i} : {e}")
                    if i in subs_lst:
                        try:
                            TokenKey = Symbol_Token[i]
                            #print(live_data)
                            lst = [live_data[TokenKey].get("Open", "-"),
                                   live_data[TokenKey].get("High", "-"),
                                   live_data[TokenKey].get("Low", "-"),
                                   live_data[TokenKey].get("Close", "-"),
                                   live_data[TokenKey].get("Vwap", "-"),
                                   live_data[TokenKey].get("bp1", "-"),
                                   live_data[TokenKey].get("sp1", "-"),
                                   live_data[TokenKey].get("Volume", "-"),
                                   live_data[TokenKey].get("OI", "-"),
                                   live_data[TokenKey].get("LTP", "-"),
                                   round(live_data[TokenKey].get("change", "0"),2)]
                            try:
                                trade_info = trading_info[idx]
                                idx_location = idx + 2
                                #print(f" {i} : {trade_info}")
                                if trade_info[0] is not None and trade_info[1] is not None:
                                    if type(trade_info[0]) is float and type(trade_info[1]) is str:
                                        
                                        LTP = live_data[TokenKey].get("LTP", 0)
                                        
                                        if Trade_Mode == 'REAL':
                                            #Real trade mode handling will handle here
                                            if trade_info[1].upper() == "BUY" and LTP != 0:
                                                if trade_info[2] in ['True_Market' ,'True_Limit_LTP', 'Limit_Below', 'Limit_Above']:
                                                    
                                                   
                                                    
                                                    #0:Qty    1:BUY/SELL    2:Entry_Signal  3:Entry_Limit_Price  4:Entry_Done@  5:Entry_Order_ID 6: Entry_Remarks  7:Exit_Signal 8:Exit_Done @ 9:Exit_Order_ID   10:Target    11:SL  12:Trail Enable 13: Latest_SL 14:Trade_status 15: PnL
                                                    
                                                    #0:Qty    1:BUY/SELL    2:Entry_Signal  3:Entry_Limit_Price  4:Entry_Done@  5:Entry_Order_ID 6: Entry_Remarks  7:Exit_Signal 8:Exit_Done @ 9:Exit_Order_ID  10: Exit_Remarks 11:Target    12:SL  13:Trail Enable 14: Latest_SL 15:Trade_status 16: PnL
                                                    
            
                                                    if trade_info[15] != 'Active' and trade_info[15] != 'Entry_Pending' and trade_info[15] != 'Exit_Pending' and trade_info[15] != 'Closed' and (trade_info[15] is None or trade_info[15] == ''):
                                                        
                                                        
                                                        if trade_info[2] == 'True_Market':
                                                            #Entry buy trade immediately
                                                            
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
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","MARKET")
                                                            if order_id is None:
                                                                tt.range(f"T{idx_location + 2}").value = None
                                                            else:  
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"v{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Exit_Pending'
                                                                
                                                        elif trade_info[7] == 'True_Limit_LTP':
                                                            #exit buy at ltp limit 
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","LIMIT",LTP)
                                                            if order_id is None:
                                                                tt.range(f"T{idx_location + 2}").value = None
                                                            else:  
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"v{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Exit_Pending'
                                                            
                                                        elif type(trade_info[11]) is float and trade_info[11] <= LTP:
                                                            #target meets, so exit the buy order
                                                            
                                                            order_id = place_trade(i, int(trade_info[0]), "SELL","MARKET")
                                                            
                                                            LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                            tt.range(f"v{idx_location + 2}").value = order_id
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Exit_Pending'
                                                        
                                                        elif TSL >= LTP and type(trade_info[12]) is float:
                                                            #sl meets, so exiting the buy order
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
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","MARKET")
                                                            if order_id is None:
                                                                tt.range(f"T{idx_location + 2}").value = None
                                                            else:  
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                    
                                                                tt.range(f"v{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = "Exit_Pending"
                                                                
                                                        elif trade_info[7] == 'True_Limit_LTP':
                                                            #exit SELL at ltp limit 
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","LIMIT",LTP)
                                                            
                                                            if order_id is None:
                                                                tt.range(f"T{idx_location + 2}").value = None
                                                            else:  
                                                                LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                                
                                                                tt.range(f"v{idx_location + 2}").value = order_id
                                                                tt.range(f"Ab{idx_location + 2}").value = 'Exit_Pending'    
                                                                
                                                        elif type(trade_info[11]) is float and trade_info[11] >= LTP:
                                                            #target meets, so exiting the sell order
                                                            order_id = place_trade(i, int(trade_info[0]), "BUY","MARKET")
                                                            
                                                            LimitOrderBook.update({str(order_id): {'status': 'PENDING', 'Remarks': None, 'Executed_price': None}})
                                                            
                                                            tt.range(f"v{idx_location + 2}").value = order_id
                                                            tt.range(f"Ab{idx_location + 2}").value = 'Exit_Pending'
                                                            
                                                        
                                                        elif TSL <= LTP and type(trade_info[12]) is float:
                                                            #sl hit, so exiting the sell order
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
                                #print('Exception occur in core order book management:' + str(e))
                                pass
                        except Exception as e:
                            #print('Exception occur in trade terminal:' + str(e))
                            pass
                main_list.append(lst)
                
                idx += 1

            tt.range("b4:l1000").value = main_list
                
        except Exception as e:
            #print('Exception occur in while :' + str(e))
            pass

def GetOverAllPnL():
    global kite
    position = kite.positions()
            
    df_position_net = pd.DataFrame(position['net'])
    Running_Algo_M2M = 0
    if(len(df_position_net) > 0 ):

        df_position_net['LTP'] = None 
        df_position_net['Algo_realised'] = None 
        df_position_net['Algo_non_realised'] = None 
        df_position_net['Algo_M2M'] = None

        for ind in df_position_net.index: 
            Exchange_Symbol = str(df_position_net['exchange'][ind]) + ':' + str(df_position_net['tradingsymbol'][ind]) 
            LTP = float(kite.quote([Exchange_Symbol])[Exchange_Symbol]['last_price']) 
            df_position_net['LTP'][ind] = LTP

            quantity = int(df_position_net['quantity'][ind])
            buy_quantity = int(df_position_net['buy_quantity'][ind])
            buy_price = float(df_position_net['buy_price'][ind])
            Sell_quantity = int(df_position_net['sell_quantity'][ind]) 
            sell_price = float(df_position_net['sell_price'][ind])
            multiplier = int(df_position_net['multiplier'][ind])

            if(quantity == 0 ): #no open position found
                df_position_net['Algo_realised'][ind]  = (sell_price - buy_price) *  buy_quantity * multiplier
                df_position_net['Algo_non_realised'][ind] = 0
            elif(quantity > 0): #few longs position r available
                if(Sell_quantity == 0):
                    df_position_net['Algo_realised'][ind]  = 0
                    df_position_net['Algo_non_realised'][ind] = (LTP - buy_price) *  buy_quantity * multiplier
                else:
                    df_position_net['Algo_realised'][ind]  = (sell_price - buy_price) *  Sell_quantity * multiplier
                    df_position_net['Algo_non_realised'][ind] = (LTP - buy_price) *  quantity * multiplier
            elif(quantity < 0): #few shorts position r available
                if(buy_quantity == 0):
                    df_position_net['Algo_realised'][ind]  = 0
                    df_position_net['Algo_non_realised'][ind] = (sell_price - LTP) *  Sell_quantity * multiplier
                else:
                    df_position_net['Algo_realised'][ind]  = (sell_price - buy_price) *  buy_quantity * multiplier
                    df_position_net['Algo_non_realised'][ind] = (sell_price - LTP) *  abs(quantity) * multiplier

            df_position_net['Algo_M2M'][ind] = df_position_net['Algo_realised'][ind]  + df_position_net['Algo_non_realised'][ind] 
        
        Running_Algo_M2M = df_position_net['Algo_M2M'].sum()
        
    
    return Running_Algo_M2M
       
def RemoveUnwantedColumn(df_openPosition):
    try:
        df_openPosition = df_openPosition.drop(columns =['overnight_quantity','multiplier','close_price','value','m2m','buy_quantity','buy_price','buy_value','buy_m2m','sell_quantity','sell_price','sell_value','sell_m2m','day_buy_quantity','day_buy_price','day_buy_value','day_sell_quantity','day_sell_price','day_sell_value'], axis=1)
    except Exception as e:
        pass
    return df_openPosition
    
def get_position():
    #print(f"I am inside get_position")
    global kite
    position = kite.positions()
    df_position_net = pd.DataFrame(position['net'])
    #print(df_position_net)
    return df_position_net
    
def get_order_book():
    #print(f"I am inside get_order_book")
    global kite
    
    order_book = pd.DataFrame()
    order_book = kite.orders()
    try:
        if(len(order_book) > 0):
            order_book = pd.DataFrame(order_book)
            
            order_book = order_book.drop(['meta','placed_by','validity','instrument_token','disclosed_quantity','modified','exchange_order_id','parent_order_id','status_message','status_message_raw','exchange_update_timestamp','exchange_timestamp','validity_ttl','market_protection','tag','guid'], axis=1)
            order_book.sort_values(by = 'order_id',inplace = True)
            order_book = order_book.reset_index(drop = True)
            #print(order_book)
            #order_book.to_csv("abc.csv")
    except Exception as e:
        #print(f"Exception in get_order_book: {e}")
        pass
    return order_book
      
prev_day_oi = {} 
prev_day_oi_pro = {} 
stop_get_oi_thread = False
stop_get_oi_pro_thread = False

Symbol_spot = {"NIFTY":{"Exchange":"NSE","Name":"NIFTY 50"},"BANKNIFTY":{"Exchange":"NSE","Name":"NIFTY BANK"},"FINNIFTY":{"Exchange":"NSE","Name":"NIFTY FIN SERVICE"},"MIDCPNIFTY":{"Exchange":"NSE","Name":"NIFTY MID SELECT"},"SENSEX":{"Exchange":"BSE","Name":"SENSEX"},"BANKEX":{"Exchange":"BSE","Name":"BANKEX"}}
        
def get_oi(data):
    global prev_day_oi, kite, stop_get_oi_thread 
    
    for symbol, v in data.items(): 
        if stop_get_oi_thread:
            break 
        while True: 
            try:
                prev_day_oi[symbol]
                break 
            except Exception as e: 
                try: 
                    pre_day_data = kite.historical_data(v["token"], (dt.now() - timedelta(days=5)).date(),(dt.now() - timedelta(days=1)).date(), "day", oi=True) 
                    try:
                        prev_day_oi[symbol] = pre_day_data[-1]["oi"] 
                    except Exception as e:
                        #print("previous day oi data not found, exception :" + str(e))
                        prev_day_oi[symbol] = 0 
                    sleep(0.5)
                    break 
                except Exception as e:
                    #print("Exception occur while downloading previous day oi :" + str(e))
                    sleep(0.5)
    stop_get_oi_thread = True
    
def start_optionchain():
    global excel_OC, TradeTerminalFileName
    excel_OC = xw.Book(TradeTerminalFileName)
    global df_instrument
    global prev_day_oi
    pre_selected_segment = pre_selected_symbol = pre_selected_expiry = "" 
    pre_selected_NoOfStrike = 1000
    instrument_dict = {}
    global stop_get_oi_thread
    
    oci = excel_OC.sheets("Option_Chain_Input")
    oco = excel_OC.sheets("Option_Chain_Output") 
    
    
    oci.range("d2").value = "Segment==>>"
    oci.range("d3").value, oci.range("d4").value = "Symbol==>>", "Expiry==>>",
    oci.range("d5").value, oci.range("d6").value = "RefreshRate==>>", "NoOfStrike==>>",
    oci.range("d7").value, oci.range("d8").value = "ExpiryType==>>" , "GreekMatch==>>"
    
    print("Excel option chain : Started") 
    #print(f"pro 003:{dt.now()}")
    while True:
        while True: #excel_OC_pro.sheets.active.name in ["Option_Chain_Pro_Input","Option_Chain_Pro_Output"]:
            
            UserInput = oci.range(f"e{2}:e{8}").value
            selected_segment , selected_symbol, selected_expiry, selected_refreshrate, NoOfStrike, ExpiryType, GreekMatch = UserInput[0],UserInput[1], UserInput[2] , UserInput[3] , UserInput[4], UserInput[5] , UserInput[6] 
            
            if selected_refreshrate is None:
                selected_refreshrate = 5
            
            if NoOfStrike is None:
                NoOfStrike = 100
            Exchange = selected_segment[:3]
            if Exchange not in ['NFO','CDS','MCX','BCD','BFO']:
                Exchange = 'NFO'
                
            if pre_selected_segment != selected_segment:
                oci.range("a2:c500").value = None
                oci.range("i3:j26").value = None
                oco.range("a3:AE500").value = None
                instrument_dict = {} 
                df_exp = pd.DataFrame()
                stop_get_oi_thread = True
                
                df_instrument_temp = df_instrument[ (df_instrument["segment"] == selected_segment)] 
                if len(df_instrument_temp) == 0 :
                    oci.range("f2").value = "Please enter correct symbol"
                else:
                    oci.range("f2").value = None
                df_instrument_temp = df_instrument_temp.drop_duplicates( "name" , keep='first')
                df_instrument_temp =df_instrument_temp[['name']]
                df_instrument_temp.sort_values(by = 'name',inplace = True)
                oci.range("a2").options(index=False, header=False).value = df_instrument_temp
                
            elif pre_selected_symbol != selected_symbol:
                #if symbol only changed
                
                oci.range("b2:c500").value = None
                oci.range("i3:j26").value = None
                oco.range("a3:AE500").value = None
                instrument_dict = {} 
                stop_get_oi_thread = True
                df_exp = pd.DataFrame()
               
            elif pre_selected_expiry != selected_expiry:
                #only expiry changed
                oci.range("i3:j26").value = None
                oco.range("a3:AE500").value = None
                instrument_dict = {} 
                stop_get_oi_thread = True 
                
            pre_selected_segment = selected_segment
            pre_selected_symbol = selected_symbol
            pre_selected_expiry = selected_expiry
            
            if selected_symbol is not None: 
                try: 
                    if len(df_exp)== 0:
                        df_instrument_temp =  df_instrument  
                        df_instrument_temp = df_instrument_temp[ (df_instrument_temp["name"] == selected_symbol) & (df_instrument_temp["segment"] == selected_segment)] 
                        if len(df_instrument_temp) == 0 :
                            oci.range("f3").value = "Please enter correct symbol"
                        else:
                            oci.range("f3").value = None
                        df_exp = df_instrument_temp.drop_duplicates( "expiry" , keep='first')
                        df_exp =df_exp[['expiry','lot_size']]
                        df_exp.sort_values(by = 'expiry',inplace = True)
                        oci.range("b2").options(index=False, header=False).value = df_exp
    
                    if not instrument_dict and selected_expiry is not None:
                        df_instrument_temp = df_instrument
                        df_instrument_temp = df_instrument_temp[(df_instrument_temp["name"] == selected_symbol) & (df_instrument_temp["segment"] == selected_segment) ] 
                        if len(df_instrument_temp) == 0 :
                            oci.range("f3").value = "Please enter correct symbol"
                        else:
                            oci.range("f3").value = None
                            df_instrument_temp = df_instrument_temp[df_instrument_temp["expiry"] == selected_expiry.date()]
                            if len(df_instrument_temp) == 0 :
                                oci.range("f4").value = "Please enter correct expiry in dd-mm-YYYY (date format)"
                            else:
                                oci.range("f4").value = None
                        lot_size = list(df_instrument_temp [ "lot_size"])[0]
                        for i in df_instrument_temp.index: 
                            instrument_dict[f'{Exchange}{":"}{df_instrument_temp["tradingsymbol"][i]}'] = {"strike": float(df_instrument_temp["strike"][i]),
                                                                "instrumentType": df_instrument_temp["instrument_type"][i],
                                                                "token": df_instrument_temp [ "instrument_token"][i]}         
                        stop_get_oi_thread = False 
                        thread = threading.Thread(target=get_oi, args=(instrument_dict,))
                        thread.start() 
                    option_data = {} 
                    
                    SpotPrice = None
                    spot_open = None
                    spot_high = None
                    spot_low = None
                    spot_close = None
                    
                    if Exchange == 'NFO':
                        try:
                            spot_instrument = Symbol_spot[selected_symbol]["Exchange"] + ":"+ Symbol_spot[selected_symbol]["Name"]
                        except Exception as e:
                            #print(f"Exception occur : {e}")
                            spot_instrument = "NSE:" + str(selected_symbol)
                            pass

                    elif Exchange == 'BFO':
                        try:
                            spot_instrument = Symbol_spot[selected_symbol]["Exchange"] + ":"+ Symbol_spot[selected_symbol]["Name"]
                        except Exception as e:
                            #print(f"Exception occur : {e}")
                            spot_instrument = "BSE:" + str(selected_symbol)
                            pass
                    else:
                        spot_instrument = None
                    try:
                        spot_data = kite.quote(spot_instrument)[spot_instrument]
                        SpotPrice = spot_data["last_price"]
                        spot_open = spot_data['ohlc']["open"]
                        spot_high = spot_data['ohlc']["high"]
                        spot_low = spot_data['ohlc']["low"]
                        spot_close = spot_data['ohlc']["close"]
                    except :
                        pass
                
                    
                    df_instrument_temp = df_instrument
                    #get the future instrument
                    df_instrument_temp = df_instrument_temp[ (df_instrument_temp["name"] == selected_symbol) & (df_instrument_temp["segment"] == Exchange + str("-FUT"))]
                    #sort and find latest expiry
                    df_instrument_temp.sort_values(by = 'expiry',inplace = True)
                    #get the intrument
                    instrument_for_future = Exchange + ":" + df_instrument_temp.iloc[0]['tradingsymbol']

                    underlying_future_quote = kite.quote(instrument_for_future)[instrument_for_future]
                    #print(f"underlying_future_quote : {instrument_for_future} : {underlying_future_quote}")
                    underlying_future_price = underlying_future_quote["last_price"]
            
                    #print(underlying_future_price)
                    for symbol, values in kite.quote(list(instrument_dict.keys())).items():
                        try: 
                            #print(f"pro 800:{dt.now()}")
                            try:
                                option_data[symbol] 
                            except Exception as e:
                                option_data[symbol] = {} 

                            option_data[symbol]["strike"] = instrument_dict[symbol]["strike"]
                            option_data[symbol]["instrumentType"] = instrument_dict[symbol]["instrumentType"] 
                            option_data[symbol]["lastPrice"] = values["last_price"]
                            option_data[symbol]["totalTradedvolume"] = int(int(values["volume"])/lot_size)
                            option_data[symbol]["openInterest"] = int(values ["oi"]/lot_size)
                            
                            option_data[symbol]["BIDQTY"] = values['depth']['buy'][0]['quantity']
                            option_data[symbol]["BIDPRICE"] = values['depth']['buy'][0]['price']
                            option_data[symbol]["ASKPRICE"] = values['depth']['sell'][0]['price']
                            option_data[symbol]["ASKQTY"] = values['depth']['sell'][0]['quantity']    
                            
                            option_data[symbol]["change"] = values["last_price"] - values["ohlc"]["close"] if         values["last_price"] != 0 else 0
                            try:
                                option_data[symbol]["changeinopeninterest"] = int((values["oi"] - prev_day_oi[symbol])/lot_size)
                            except Exception as e:
                                #print("Exception occur in changeinopeninterest: " + str(e))
                                option_data[symbol]["changeinopeninterest"] = None
                                
                        except Exception as e:
                            #print("Exception occur in for loop: " + str(e))
                            pass
                    
                    df_oc = pd.DataFrame(option_data).transpose() 
                    #print(df_oc)
                    df_oc_ce = df_oc[df_oc["instrumentType"] == "CE"] 
                    #print(df_oc_ce)
                    df_oc_ce = df_oc_ce [["totalTradedvolume", "change", "lastPrice", "changeinopeninterest", "openInterest", "strike","BIDQTY","BIDPRICE","ASKPRICE","ASKQTY"]]
                    #print(df_oc_ce)            
                    df_oc_ce = df_oc_ce.rename(columns={"openInterest": "CE_OI", "changeinopeninterest": "CE Change in OI","lastPrice": "CE LTP", "change": "CE LTP Change", "totalTradedvolume": "CE Volume" ,"BIDQTY" : "CE_BIDQTY","BIDPRICE":"CE_BIDPRICE","ASKPRICE":"CE_ASKPRICE","ASKQTY":"CE_ASKQTY"}) 
                    #print(df_oc_ce)
                    df_oc_ce.index = df_oc_ce["strike"]
                    #print(df_oc_ce)            
                    df_oc_ce = df_oc_ce.drop(["strike"], axis=1)
                    #df_oc_ce["strike"] = df_oc_ce.index 
                    #print(df_oc_ce)
                    df_oc_pe = df_oc[df_oc["instrumentType"] == "PE"] 
                    df_oc_pe = df_oc_pe[["strike", "openInterest", "changeinopeninterest", "lastPrice", "change", "totalTradedvolume","BIDQTY","BIDPRICE","ASKPRICE","ASKQTY"]] 
                    df_oc_pe = df_oc_pe.rename(columns={"openInterest": "PE_OI", "changeinopeninterest": "PE Change in OI","lastPrice": "PE LTP", "change": "PE LTP Change", "totalTradedvolume" : "PE Volume" ,"BIDQTY" : "PE_BIDQTY","BIDPRICE":"PE_BIDPRICE","ASKPRICE":"PE_ASKPRICE","ASKQTY":"PE_ASKQTY"})
                    
                    df_oc_pe.index = df_oc_pe["strike"] 
                    df_oc_pe = df_oc_pe.drop("strike", axis=1) 
                    #print(df_oc_pe)
                    df_oc_pro = pd.concat([df_oc_ce, df_oc_pe], axis=1).sort_index() 
                    df_oc_pro = df_oc_pro.replace(np.nan, 0) 
                    df_oc_pro["Strike"] = df_oc_pro.index 
                    
                    df_oc_pro['strike_gap'] = abs(df_oc_pro['Strike'] - underlying_future_price)
                    
                    Min_gap = df_oc_pro['strike_gap'].min()
                    ATM_strike = df_oc_pro[df_oc_pro.strike_gap == Min_gap].iloc[0]['Strike']
                    
                    ATM_pos = df_oc_pro.index.get_loc(ATM_strike)
                    
                    df_oc_pro['OI_SUM'] = df_oc_pro["CE_OI"] + df_oc_pro["PE_OI"]
                    
                    Future_LTP = underlying_future_price
                    Max_Pain_at_Strike = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['Strike']
                    Ltp_at_Max_Pain_CE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['CE LTP']
                    Ltp_at_Max_Pain_PE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['PE LTP']
                    #ATM_Strike = 
                    LTP_at_ATM_CE = df_oc_pro[df_oc_pro.Strike == ATM_strike].iloc[0]['CE LTP']
                    LTP_at_ATM_PE = df_oc_pro[df_oc_pro.Strike == ATM_strike].iloc[0]['PE LTP']
                    Total_OI_CE = df_oc_pro["CE_OI"].sum()
                    Total_OI_PE = df_oc_pro["PE_OI"].sum()
                    
                    Max_OI_CE = df_oc_pro["CE_OI"].max()
                    Max_OI_PE = df_oc_pro["PE_OI"].max()
                    Max_OI_at_Strike_CE = df_oc_pro[df_oc_pro.CE_OI == df_oc_pro["CE_OI"].max()].iloc[0]['Strike']
                    Max_OI_at_Strike_PE = df_oc_pro[df_oc_pro.PE_OI == df_oc_pro["PE_OI"].max()].iloc[0]['Strike']
                    LTP_of_Max_OI_Strike_CE = df_oc_pro[df_oc_pro.CE_OI == df_oc_pro["CE_OI"].max()].iloc[0]['CE LTP']
                    LTP_of_Max_OI_Strike_PE = df_oc_pro[df_oc_pro.PE_OI == df_oc_pro["PE_OI"].max()].iloc[0]['PE LTP']

                    Total_Volume_CE = df_oc_pro["CE Volume"].sum()
                    Total_Volume_PE = df_oc_pro["PE Volume"].sum()
                    Max_Vol_at_Strike_CE = df_oc_pro[df_oc_pro['CE Volume'] == df_oc_pro["CE Volume"].max()].iloc[0]['Strike']
                    Max_Vol_at_Strike_PE = df_oc_pro[df_oc_pro['PE Volume'] == df_oc_pro["PE Volume"].max()].iloc[0]['Strike']
                    LTP_of_Max_Vol_Strike_CE = df_oc_pro[df_oc_pro['CE Volume'] == df_oc_pro["CE Volume"].max()].iloc[0]['CE LTP']
                    LTP_of_Max_Vol_Strike_PE = df_oc_pro[df_oc_pro['PE Volume'] == df_oc_pro["PE Volume"].max()].iloc[0]['PE LTP']
                    
                    if stop_get_oi_thread == True:
                        Total_OI_Change_CE = df_oc_pro["CE Change in OI"].sum()
                        Total_OI_Change_PE = df_oc_pro["PE Change in OI"].sum()
                        
                        Max_Change_in_OI_addition_CE = df_oc_pro["CE Change in OI"].max()
                        Max_Change_in_OI_addition_PE = df_oc_pro["PE Change in OI"].max()
                        Max_OI_addition_at_Srike_CE = df_oc_pro[df_oc_pro["CE Change in OI"] == df_oc_pro["CE Change in OI"].max()].iloc[0]['Strike']
                        Max_OI_addition_at_Srike_PE = df_oc_pro[df_oc_pro["PE Change in OI"] == df_oc_pro["PE Change in OI"].max()].iloc[0]['Strike']
                        LTP_of_Max_OI_addition_Strike_CE = df_oc_pro[df_oc_pro["CE Change in OI"] == df_oc_pro["CE Change in OI"].max()].iloc[0]['CE LTP']
                        LTP_of_Max_OI_addition_Strike_PE = df_oc_pro[df_oc_pro["PE Change in OI"] == df_oc_pro["PE Change in OI"].max()].iloc[0]['PE LTP']
                        Max_Change_in_OI_unwinding_CE = -1 * int(df_oc_pro["CE Change in OI"].min())
                        Max_Change_in_OI_unwinding_PE = -1 * int(df_oc_pro["PE Change in OI"].min())
                        Max_OI_unwinding_at_Srike_CE = df_oc_pro[df_oc_pro["CE Change in OI"] == df_oc_pro["CE Change in OI"].min()].iloc[0]['Strike']
                        Max_OI_unwinding_at_Srike_PE = df_oc_pro[df_oc_pro["PE Change in OI"] == df_oc_pro["PE Change in OI"].min()].iloc[0]['Strike']
                        LTP_of_Max_OI_unwinding_Strike_CE = df_oc_pro[df_oc_pro["CE Change in OI"] == df_oc_pro["CE Change in OI"].min()].iloc[0]['CE LTP']
                        LTP_of_Max_OI_unwinding_Strike_PE = df_oc_pro[df_oc_pro["PE Change in OI"] == df_oc_pro["PE Change in OI"].min()].iloc[0]['PE LTP']
                    else:
                        Total_OI_Change_CE = None
                        Total_OI_Change_PE = None
                        
                        Max_Change_in_OI_addition_CE = None
                        Max_Change_in_OI_addition_PE = None
                        Max_OI_addition_at_Srike_CE = None
                        Max_OI_addition_at_Srike_PE = None
                        LTP_of_Max_OI_addition_Strike_CE = None
                        LTP_of_Max_OI_addition_Strike_PE = None

                        Max_Change_in_OI_unwinding_CE = None
                        Max_Change_in_OI_unwinding_PE = None
                        Max_OI_unwinding_at_Srike_CE = None
                        Max_OI_unwinding_at_Srike_PE = None
                        LTP_of_Max_OI_unwinding_Strike_CE = None
                        LTP_of_Max_OI_unwinding_Strike_PE = None
                        
                    
                    df_additional_detail = pd.DataFrame(columns = ['CE','PE'])
                    
                    dic_data = {'CE': SpotPrice, 'PE':Future_LTP}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE': spot_open,'PE':underlying_future_quote["ohlc"]["open"]}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE': spot_high,'PE':underlying_future_quote["ohlc"]["high"]}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE': spot_low,'PE':underlying_future_quote["ohlc"]["low"]}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE': spot_close,'PE':underlying_future_quote["ohlc"]["close"]}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE':underlying_future_quote["oi"]/lot_size}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE':Max_Pain_at_Strike}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE':Ltp_at_Max_Pain_CE,'PE': Ltp_at_Max_Pain_PE}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE':ATM_strike}
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
                    
                    dic_data = {'CE':Total_Volume_CE,'PE':Total_Volume_PE}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    
                    dic_data = {'CE':Max_Vol_at_Strike_CE,'PE':Max_Vol_at_Strike_PE}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE':LTP_of_Max_Vol_Strike_CE,'PE':LTP_of_Max_Vol_Strike_PE}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    oci.range("i3").options(index=False, header=False).value = df_additional_detail
                    
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
                    
                    df_oc_pro = df_oc_pro.loc[:,['CE_Delta','CE_Gamma','CE_Theta','CE_Vega','CE_Rho','CE_OI','CE Change in OI','CE Volume','CE_IV','CE LTP','CE LTP Change','CE_BIDQTY','CE_BIDPRICE','CE_ASKPRICE','CE_ASKQTY','Strike','PE_BIDQTY','PE_BIDPRICE','PE_ASKPRICE','PE_ASKQTY','PE LTP Change','PE LTP','PE_IV','PE Volume','PE Change in OI','PE_OI','PE_Rho','PE_Vega','PE_Theta','PE_Gamma','PE_Delta']]
                    
                    if pre_selected_NoOfStrike != NoOfStrike:
                        oco.range("a3:AE500").value = None
                        
                        oco.range(f"a3:AE500").color = (255,255,255)
                        oco.range(f"p3:p{2 * int(NoOfStrike) + 3}").color = (46,132,198)                
                            
                        pre_selected_NoOfStrike = NoOfStrike
                    
                    if NoOfStrike * 2 < len(df_oc_pro):
                        df_oc_pro = df_oc_pro.iloc[ATM_pos - int(NoOfStrike) : ATM_pos + int(NoOfStrike) + 1]
                        ATM_Row = int(NoOfStrike) + 3
                        oco.range(f"a{ATM_Row}:AE{ATM_Row}").color = (46,132,198) 
                    
                    
                    if Exchange == 'NFO':
                        try:
                            instrument_for_ltp = Symbol_spot[selected_symbol]["Exchange"] + ":"+ Symbol_spot[selected_symbol]["Name"]
                        except Exception as e:
                            #print(f"Exception occur : {e}")
                            instrument_for_ltp = "NSE:" + str(selected_symbol)
                            pass
                        #print(f"@@@@ {instrument_for_ltp}")
                        SpotPrice = kite.quote(instrument_for_ltp)[instrument_for_ltp]["last_price"]
                    elif Exchange == 'BFO':
                        try:
                            instrument_for_ltp = Symbol_spot[selected_symbol]["Exchange"] + ":"+ Symbol_spot[selected_symbol]["Name"]
                            SpotPrice = kite.quote(instrument_for_ltp)[instrument_for_ltp]["last_price"]
                        except Exception as e:
                            SpotPrice = float(underlying_future_price)
                            pass
                        
                    else:
                        SpotPrice = float(underlying_future_price)
                        
                    FuturePrice = float(underlying_future_price)
                    AtmStrike = float(ATM_strike)
                    AtmStrikeCallPrice = float(LTP_at_ATM_CE)
                    AtmStrikePutPrice = float(LTP_at_ATM_PE)
                    ExpiryDateTime = dt(selected_expiry.date().year, selected_expiry.date().month, selected_expiry.date().day, 0, 0, 0)
                    
                    
                    if ExpiryType == 'WEEKLY':
                        ExpiryDateType = ExpType.WEEKLY
                    else:
                        ExpiryDateType = ExpType.MONTHLY
                    
                    
                    if dt.now().time() < time(15, 30, 0):
                        FromDateTime = dt.now() 
                    else:
                        FromDateTime = dt(dt.now().year, dt.now().month,dt.now().day, 15, 30, 0)
                
                    if GreekMatch == "SENSIBULL":
                        tryMatchWith=TryMatchWith.SENSIBULL
                    else:
                        tryMatchWith=TryMatchWith.NSE
                        
                    dayCountType = DayCountType.CALENDARDAYS

                    IvGreeks = CalcIvGreeks( SpotPrice = SpotPrice,  FuturePrice = FuturePrice, AtmStrike = AtmStrike, AtmStrikeCallPrice = AtmStrikeCallPrice, AtmStrikePutPrice = AtmStrikePutPrice, ExpiryDateTime = ExpiryDateTime, ExpiryDateType = ExpiryDateType, FromDateTime = FromDateTime, tryMatchWith = tryMatchWith, dayCountType = dayCountType)
    
                    #print(f"SpotPrice={SpotPrice}, FuturePrice={FuturePrice},  AtmStrike={AtmStrike}, AtmStrikeCallPrice={AtmStrikeCallPrice}, AtmStrikePutPrice={AtmStrikePutPrice}, ExpiryDateTime={ExpiryDateTime},  ExpiryDateType={ExpiryDateType}, FromDateTime={FromDateTime}, tryMatchWith={tryMatchWith}")
                    
                    df_oc_pro.round({"CE LTP":2, 'PE LTP':2})
                    
                    for ind in df_oc_pro.index:
                        
                        StrikePrice= float(df_oc_pro['Strike'][ind])
                        StrikeCallPrice= float(df_oc_pro['CE LTP'][ind])
                        StrikePutPrice= float(df_oc_pro['PE LTP'][ind])
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
                    oco.range("a3").options(index=False, header=False).value = df_oc_pro
                    sleep(selected_refreshrate) 
                    
                except Exception as e:
                    #print("Exception occur in parent loop: " + str(e))
                    pass
        
        sleep(5)
        
def get_oi_pro(data):
    global prev_day_oi_pro, kite, stop_get_oi_pro_thread 
    
    for symbol, v in data.items(): 
        if stop_get_oi_pro_thread:
            break 
        while True: 
            try:
                prev_day_oi_pro[symbol]
                break 
            except Exception as e: 
                try: 
                    pre_day_data = kite.historical_data(v["token"], (dt.now() - timedelta(days=5)).date(),(dt.now() - timedelta(days=1)).date(), "day", oi=True) 
                    try:
                        prev_day_oi_pro[symbol] = pre_day_data[-1]["oi"] 
                    except Exception as e:
                        #print("previous day oi data not found, exception :" + str(e))
                        prev_day_oi_pro[symbol] = 0 
                    sleep(0.5)
                    break 
                except Exception as e:
                    #print("Exception occur while downloading previous day oi :" + str(e))
                    sleep(0.5)
    stop_get_oi_pro_thread = True
    
def start_optionchain_Pro():
    global excel_OC_pro, TradeTerminalFileName
    excel_OC_pro = xw.Book(TradeTerminalFileName)
    global df_instrument
    global prev_day_oi_pro
    pre_selected_segment = pre_selected_symbol = pre_selected_expiry = "" 
    pre_selected_NoOfStrike = 1000
    instrument_dict = {}
    global stop_get_oi_pro_thread
    
    oci_pro = excel_OC_pro.sheets("Option_Chain_Pro_Input")
    oco_pro = excel_OC_pro.sheets("Option_Chain_Pro_Output") 
    
    
    oci_pro.range("d2").value = "Segment==>>"
    oci_pro.range("d3").value, oci_pro.range("d4").value = "Symbol==>>", "Expiry==>>",
    oci_pro.range("d5").value, oci_pro.range("d6").value = "RefreshRate==>>", "NoOfStrike==>>",
    oci_pro.range("d7").value, oci_pro.range("d8").value = "ExpiryType==>>" , "GreekMatch==>>"
    
    print("Excel option chain Pro: Started") 
    #print(f"pro 003:{dt.now()}")
    while True:
        while True: #excel_OC_pro.sheets.active.name in ["Option_Chain_Pro_Input","Option_Chain_Pro_Output"]:
            
            UserInput = oci_pro.range(f"e{2}:e{8}").value
            selected_segment , selected_symbol, selected_expiry, selected_refreshrate, NoOfStrike, ExpiryType, GreekMatch = UserInput[0],UserInput[1], UserInput[2] , UserInput[3] , UserInput[4], UserInput[5] , UserInput[6] 
            
            if selected_refreshrate is None:
                selected_refreshrate = 5
            
            if NoOfStrike is None:
                NoOfStrike = 100
            Exchange = selected_segment[:3]
            if Exchange not in ['NFO','CDS','MCX','BCD','BFO']:
                Exchange = 'NFO'
                
            if pre_selected_segment != selected_segment:
                oci_pro.range("a2:c500").value = None
                oci_pro.range("i3:j26").value = None
                oco_pro.range("a3:AE500").value = None
                instrument_dict = {} 
                df_exp = pd.DataFrame()
                stop_get_oi_pro_thread = True
                
                df_instrument_temp = df_instrument[ (df_instrument["segment"] == selected_segment)] 
                if len(df_instrument_temp) == 0 :
                    oci_pro.range("f2").value = "Please enter correct symbol"
                else:
                    oci_pro.range("f2").value = None
                df_instrument_temp = df_instrument_temp.drop_duplicates( "name" , keep='first')
                df_instrument_temp =df_instrument_temp[['name']]
                df_instrument_temp.sort_values(by = 'name',inplace = True)
                oci_pro.range("a2").options(index=False, header=False).value = df_instrument_temp
                
            elif pre_selected_symbol != selected_symbol:
                #if symbol only changed
                
                oci_pro.range("b2:c500").value = None
                oci_pro.range("i3:j26").value = None
                oco_pro.range("a3:AE500").value = None
                instrument_dict = {} 
                stop_get_oi_pro_thread = True
                df_exp = pd.DataFrame()
               
            elif pre_selected_expiry != selected_expiry:
                #only expiry changed
                oci_pro.range("i3:j26").value = None
                oco_pro.range("a3:AE500").value = None
                instrument_dict = {} 
                stop_get_oi_pro_thread = True 
                
            pre_selected_segment = selected_segment
            pre_selected_symbol = selected_symbol
            pre_selected_expiry = selected_expiry
            
            if selected_symbol is not None: 
                try: 
                    if len(df_exp)== 0:
                        df_instrument_temp =  df_instrument 
                        df_instrument_temp = df_instrument_temp[ (df_instrument_temp["name"] == selected_symbol) & (df_instrument_temp["segment"] == selected_segment)] 
                        if len(df_instrument_temp) == 0 :
                            oci_pro.range("f3").value = "Please enter correct symbol"
                        else:
                            oci_pro.range("f3").value = None
                        df_exp = df_instrument_temp.drop_duplicates( "expiry" , keep='first')
                        df_exp =df_exp[['expiry','lot_size']]
                        df_exp.sort_values(by = 'expiry',inplace = True)
                        oci_pro.range("b2").options(index=False, header=False).value = df_exp
    
                    if not instrument_dict and selected_expiry is not None:
                        df_instrument_temp = df_instrument
                        df_instrument_temp = df_instrument_temp[(df_instrument_temp["name"] == selected_symbol) & (df_instrument_temp["segment"] == selected_segment) ] 
                        if len(df_instrument_temp) == 0 :
                            oci_pro.range("f3").value = "Please enter correct symbol"
                        else:
                            oci_pro.range("f3").value = None
                            df_instrument_temp = df_instrument_temp[df_instrument_temp["expiry"] == selected_expiry.date()]
                            if len(df_instrument_temp) == 0 :
                                oci_pro.range("f4").value = "Please enter correct expiry in dd-mm-YYYY (date format)"
                            else:
                                oci_pro.range("f4").value = None
                        lot_size = list(df_instrument_temp [ "lot_size"])[0]
                        for i in df_instrument_temp.index: 
                            instrument_dict[f'{Exchange}{":"}{df_instrument_temp["tradingsymbol"][i]}'] = {"strike": float(df_instrument_temp["strike"][i]),
                                                                "instrumentType": df_instrument_temp["instrument_type"][i],
                                                                "token": df_instrument_temp [ "instrument_token"][i]}         
                        stop_get_oi_pro_thread = False 
                        thread = threading.Thread(target=get_oi_pro, args=(instrument_dict,))
                        thread.start() 
                    option_data = {} 
                        
                    SpotPrice = None
                    spot_open = None
                    spot_high = None
                    spot_low = None
                    spot_close = None
                    
                    if Exchange == 'NFO':
                        try:
                            spot_instrument = Symbol_spot[selected_symbol]["Exchange"] + ":"+ Symbol_spot[selected_symbol]["Name"]
                        except Exception as e:
                            #print(f"Exception occur : {e}")
                            spot_instrument = "NSE:" + str(selected_symbol)
                            pass

                    elif Exchange == 'BFO':
                        try:
                            spot_instrument = Symbol_spot[selected_symbol]["Exchange"] + ":"+ Symbol_spot[selected_symbol]["Name"]
                        except Exception as e:
                            #print(f"Exception occur : {e}")
                            spot_instrument = "BSE:" + str(selected_symbol)
                            pass
                    else:
                        spot_instrument = None
                    try:
                        spot_data = kite.quote(spot_instrument)[spot_instrument]
                        SpotPrice = spot_data["last_price"]
                        spot_open = spot_data['ohlc']["open"]
                        spot_high = spot_data['ohlc']["high"]
                        spot_low = spot_data['ohlc']["low"]
                        spot_close = spot_data['ohlc']["close"]
                    except :
                        pass
                
                    
                    df_instrument_temp = df_instrument
                    #get the future instrument
                    df_instrument_temp = df_instrument_temp[ (df_instrument_temp["name"] == selected_symbol) & (df_instrument_temp["segment"] == Exchange + str("-FUT"))]
                    #sort and find latest expiry
                    df_instrument_temp.sort_values(by = 'expiry',inplace = True)
                    #get the intrument
                    instrument_for_future = Exchange + ":" + df_instrument_temp.iloc[0]['tradingsymbol']

                    underlying_future_quote = kite.quote(instrument_for_future)[instrument_for_future]
                    #print(f"underlying_future_quote : {instrument_for_future} : {underlying_future_quote}")
                    underlying_future_price = underlying_future_quote["last_price"]
            
                    #print(underlying_future_price)
                    for symbol, values in kite.quote(list(instrument_dict.keys())).items():
                        try: 
                            #print(f"pro 800:{dt.now()}")
                            try:
                                option_data[symbol] 
                            except Exception as e:
                                option_data[symbol] = {} 

                            option_data[symbol]["strike"] = instrument_dict[symbol]["strike"]
                            option_data[symbol]["instrumentType"] = instrument_dict[symbol]["instrumentType"] 
                            option_data[symbol]["lastPrice"] = values["last_price"]
                            option_data[symbol]["totalTradedvolume"] = int(int(values["volume"])/lot_size)
                            option_data[symbol]["openInterest"] = int(values ["oi"]/lot_size)
                            
                            option_data[symbol]["BIDQTY"] = values['depth']['buy'][0]['quantity']
                            option_data[symbol]["BIDPRICE"] = values['depth']['buy'][0]['price']
                            option_data[symbol]["ASKPRICE"] = values['depth']['sell'][0]['price']
                            option_data[symbol]["ASKQTY"] = values['depth']['sell'][0]['quantity']    
                            
                            option_data[symbol]["change"] = values["last_price"] - values["ohlc"]["close"] if         values["last_price"] != 0 else 0
                            try:
                                option_data[symbol]["changeinopeninterest"] = int((values["oi"] - prev_day_oi_pro[symbol])/lot_size)
                            except Exception as e:
                                #print("Exception occur in changeinopeninterest: " + str(e))
                                option_data[symbol]["changeinopeninterest"] = None
                                
                        except Exception as e:
                            #print("Exception occur in for loop: " + str(e))
                            pass
                    
                    df_oc = pd.DataFrame(option_data).transpose() 
                    #print(df_oc)
                    df_oc_ce = df_oc[df_oc["instrumentType"] == "CE"] 
                    #print(df_oc_ce)
                    df_oc_ce = df_oc_ce [["totalTradedvolume", "change", "lastPrice", "changeinopeninterest", "openInterest", "strike","BIDQTY","BIDPRICE","ASKPRICE","ASKQTY"]]
                    #print(df_oc_ce)            
                    df_oc_ce = df_oc_ce.rename(columns={"openInterest": "CE_OI", "changeinopeninterest": "CE Change in OI","lastPrice": "CE LTP", "change": "CE LTP Change", "totalTradedvolume": "CE Volume" ,"BIDQTY" : "CE_BIDQTY","BIDPRICE":"CE_BIDPRICE","ASKPRICE":"CE_ASKPRICE","ASKQTY":"CE_ASKQTY"}) 
                    #print(df_oc_ce)
                    df_oc_ce.index = df_oc_ce["strike"]
                    #print(df_oc_ce)            
                    df_oc_ce = df_oc_ce.drop(["strike"], axis=1)
                    #df_oc_ce["strike"] = df_oc_ce.index 
                    #print(df_oc_ce)
                    df_oc_pe = df_oc[df_oc["instrumentType"] == "PE"] 
                    df_oc_pe = df_oc_pe[["strike", "openInterest", "changeinopeninterest", "lastPrice", "change", "totalTradedvolume","BIDQTY","BIDPRICE","ASKPRICE","ASKQTY"]] 
                    df_oc_pe = df_oc_pe.rename(columns={"openInterest": "PE_OI", "changeinopeninterest": "PE Change in OI","lastPrice": "PE LTP", "change": "PE LTP Change", "totalTradedvolume" : "PE Volume" ,"BIDQTY" : "PE_BIDQTY","BIDPRICE":"PE_BIDPRICE","ASKPRICE":"PE_ASKPRICE","ASKQTY":"PE_ASKQTY"})
                    
                    df_oc_pe.index = df_oc_pe["strike"] 
                    df_oc_pe = df_oc_pe.drop("strike", axis=1) 
                    #print(df_oc_pe)
                    df_oc_pro = pd.concat([df_oc_ce, df_oc_pe], axis=1).sort_index() 
                    df_oc_pro = df_oc_pro.replace(np.nan, 0) 
                    df_oc_pro["Strike"] = df_oc_pro.index 
                    
                    df_oc_pro['strike_gap'] = abs(df_oc_pro['Strike'] - underlying_future_price)
                    
                    Min_gap = df_oc_pro['strike_gap'].min()
                    ATM_strike = df_oc_pro[df_oc_pro.strike_gap == Min_gap].iloc[0]['Strike']
                    
                    ATM_pos = df_oc_pro.index.get_loc(ATM_strike)
                    
                    df_oc_pro['OI_SUM'] = df_oc_pro["CE_OI"] + df_oc_pro["PE_OI"]
                    
                    Future_LTP = underlying_future_price
                    Max_Pain_at_Strike = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['Strike']
                    Ltp_at_Max_Pain_CE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['CE LTP']
                    Ltp_at_Max_Pain_PE = df_oc_pro[df_oc_pro.OI_SUM == df_oc_pro["OI_SUM"].max()].iloc[0]['PE LTP']
                    #ATM_Strike = 
                    LTP_at_ATM_CE = df_oc_pro[df_oc_pro.Strike == ATM_strike].iloc[0]['CE LTP']
                    LTP_at_ATM_PE = df_oc_pro[df_oc_pro.Strike == ATM_strike].iloc[0]['PE LTP']
                    Total_OI_CE = df_oc_pro["CE_OI"].sum()
                    Total_OI_PE = df_oc_pro["PE_OI"].sum()
                    
                    Max_OI_CE = df_oc_pro["CE_OI"].max()
                    Max_OI_PE = df_oc_pro["PE_OI"].max()
                    Max_OI_at_Strike_CE = df_oc_pro[df_oc_pro.CE_OI == df_oc_pro["CE_OI"].max()].iloc[0]['Strike']
                    Max_OI_at_Strike_PE = df_oc_pro[df_oc_pro.PE_OI == df_oc_pro["PE_OI"].max()].iloc[0]['Strike']
                    LTP_of_Max_OI_Strike_CE = df_oc_pro[df_oc_pro.CE_OI == df_oc_pro["CE_OI"].max()].iloc[0]['CE LTP']
                    LTP_of_Max_OI_Strike_PE = df_oc_pro[df_oc_pro.PE_OI == df_oc_pro["PE_OI"].max()].iloc[0]['PE LTP']
                    
                    Total_Volume_CE = df_oc_pro["CE Volume"].sum()
                    Total_Volume_PE = df_oc_pro["PE Volume"].sum()
                    Max_Vol_at_Strike_CE = df_oc_pro[df_oc_pro['CE Volume'] == df_oc_pro["CE Volume"].max()].iloc[0]['Strike']
                    Max_Vol_at_Strike_PE = df_oc_pro[df_oc_pro['PE Volume'] == df_oc_pro["PE Volume"].max()].iloc[0]['Strike']
                    LTP_of_Max_Vol_Strike_CE = df_oc_pro[df_oc_pro['CE Volume'] == df_oc_pro["CE Volume"].max()].iloc[0]['CE LTP']
                    LTP_of_Max_Vol_Strike_PE = df_oc_pro[df_oc_pro['PE Volume'] == df_oc_pro["PE Volume"].max()].iloc[0]['PE LTP']
                    
                    if stop_get_oi_pro_thread == True:
                        Total_OI_Change_CE = df_oc_pro["CE Change in OI"].sum()
                        Total_OI_Change_PE = df_oc_pro["PE Change in OI"].sum()
                        
                        Max_Change_in_OI_addition_CE = df_oc_pro["CE Change in OI"].max()
                        Max_Change_in_OI_addition_PE = df_oc_pro["PE Change in OI"].max()
                        Max_OI_addition_at_Srike_CE = df_oc_pro[df_oc_pro["CE Change in OI"] == df_oc_pro["CE Change in OI"].max()].iloc[0]['Strike']
                        Max_OI_addition_at_Srike_PE = df_oc_pro[df_oc_pro["PE Change in OI"] == df_oc_pro["PE Change in OI"].max()].iloc[0]['Strike']
                        LTP_of_Max_OI_addition_Strike_CE = df_oc_pro[df_oc_pro["CE Change in OI"] == df_oc_pro["CE Change in OI"].max()].iloc[0]['CE LTP']
                        LTP_of_Max_OI_addition_Strike_PE = df_oc_pro[df_oc_pro["PE Change in OI"] == df_oc_pro["PE Change in OI"].max()].iloc[0]['PE LTP']
                        Max_Change_in_OI_unwinding_CE = -1 * int(df_oc_pro["CE Change in OI"].min())
                        Max_Change_in_OI_unwinding_PE = -1 * int(df_oc_pro["PE Change in OI"].min())
                        Max_OI_unwinding_at_Srike_CE = df_oc_pro[df_oc_pro["CE Change in OI"] == df_oc_pro["CE Change in OI"].min()].iloc[0]['Strike']
                        Max_OI_unwinding_at_Srike_PE = df_oc_pro[df_oc_pro["PE Change in OI"] == df_oc_pro["PE Change in OI"].min()].iloc[0]['Strike']
                        LTP_of_Max_OI_unwinding_Strike_CE = df_oc_pro[df_oc_pro["CE Change in OI"] == df_oc_pro["CE Change in OI"].min()].iloc[0]['CE LTP']
                        LTP_of_Max_OI_unwinding_Strike_PE = df_oc_pro[df_oc_pro["PE Change in OI"] == df_oc_pro["PE Change in OI"].min()].iloc[0]['PE LTP']
                    else:
                        Total_OI_Change_CE = None
                        Total_OI_Change_PE = None
                        
                        Max_Change_in_OI_addition_CE = None
                        Max_Change_in_OI_addition_PE = None
                        Max_OI_addition_at_Srike_CE = None
                        Max_OI_addition_at_Srike_PE = None
                        LTP_of_Max_OI_addition_Strike_CE = None
                        LTP_of_Max_OI_addition_Strike_PE = None

                        Max_Change_in_OI_unwinding_CE = None
                        Max_Change_in_OI_unwinding_PE = None
                        Max_OI_unwinding_at_Srike_CE = None
                        Max_OI_unwinding_at_Srike_PE = None
                        LTP_of_Max_OI_unwinding_Strike_CE = None
                        LTP_of_Max_OI_unwinding_Strike_PE = None
                        
                    
                    df_additional_detail = pd.DataFrame(columns = ['CE','PE'])
                    
                    dic_data = {'CE': SpotPrice, 'PE':Future_LTP}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE': spot_open,'PE':underlying_future_quote["ohlc"]["open"]}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE': spot_high,'PE':underlying_future_quote["ohlc"]["high"]}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE': spot_low,'PE':underlying_future_quote["ohlc"]["low"]}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE': spot_close,'PE':underlying_future_quote["ohlc"]["close"]}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE':underlying_future_quote["oi"]/lot_size}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE':Max_Pain_at_Strike}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE':Ltp_at_Max_Pain_CE,'PE': Ltp_at_Max_Pain_PE}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE':ATM_strike}
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
                    
                    dic_data = {'CE':Total_Volume_CE,'PE':Total_Volume_PE}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE':Max_Vol_at_Strike_CE,'PE':Max_Vol_at_Strike_PE}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    dic_data = {'CE':LTP_of_Max_Vol_Strike_CE,'PE':LTP_of_Max_Vol_Strike_PE}
                    df_additional_detail = pd.concat([df_additional_detail, pd.DataFrame.from_dict(dic_data,orient='index').T],ignore_index = True)
                    
                    oci_pro.range("i3").options(index=False, header=False).value = df_additional_detail
                    
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
                    
                    df_oc_pro = df_oc_pro.loc[:,['CE_Delta','CE_Gamma','CE_Theta','CE_Vega','CE_Rho','CE_OI','CE Change in OI','CE Volume','CE_IV','CE LTP','CE LTP Change','CE_BIDQTY','CE_BIDPRICE','CE_ASKPRICE','CE_ASKQTY','Strike','PE_BIDQTY','PE_BIDPRICE','PE_ASKPRICE','PE_ASKQTY','PE LTP Change','PE LTP','PE_IV','PE Volume','PE Change in OI','PE_OI','PE_Rho','PE_Vega','PE_Theta','PE_Gamma','PE_Delta']]
                    
                    if pre_selected_NoOfStrike != NoOfStrike:
                        oco_pro.range("a3:AE500").value = None
                        
                        oco_pro.range(f"a3:AE500").color = (255,255,255)
                        oco_pro.range(f"p3:p{2 * int(NoOfStrike) + 3}").color = (46,132,198)                
                            
                        pre_selected_NoOfStrike = NoOfStrike
                    
                    if NoOfStrike * 2 < len(df_oc_pro):
                        df_oc_pro = df_oc_pro.iloc[ATM_pos - int(NoOfStrike) : ATM_pos + int(NoOfStrike) + 1]
                        ATM_Row = int(NoOfStrike) + 3
                        oco_pro.range(f"a{ATM_Row}:AE{ATM_Row}").color = (46,132,198) 
                    
                    
                    if Exchange == 'NFO':
                        try:
                            instrument_for_ltp = Symbol_spot[selected_symbol]["Exchange"] + ":"+ Symbol_spot[selected_symbol]["Name"]
                        except Exception as e:
                            #print(f"Exception occur : {e}")
                            instrument_for_ltp = "NSE:" + str(selected_symbol)
                            pass
                        #print(f"@@@@ {instrument_for_ltp}")
                        SpotPrice = kite.quote(instrument_for_ltp)[instrument_for_ltp]["last_price"]
                    elif Exchange == 'BFO':
                        try:
                            instrument_for_ltp = Symbol_spot[selected_symbol]["Exchange"] + ":"+ Symbol_spot[selected_symbol]["Name"]
                            SpotPrice = kite.quote(instrument_for_ltp)[instrument_for_ltp]["last_price"]
                        except Exception as e:
                            SpotPrice = float(underlying_future_price)
                            pass
                        
                    else:
                        SpotPrice = float(underlying_future_price)
                        
                    FuturePrice = float(underlying_future_price)
                    AtmStrike = float(ATM_strike)
                    AtmStrikeCallPrice = float(LTP_at_ATM_CE)
                    AtmStrikePutPrice = float(LTP_at_ATM_PE)
                    ExpiryDateTime = dt(selected_expiry.date().year, selected_expiry.date().month, selected_expiry.date().day, 0, 0, 0)
                    
                    
                    if ExpiryType == 'WEEKLY':
                        ExpiryDateType = ExpType.WEEKLY
                    else:
                        ExpiryDateType = ExpType.MONTHLY
                    
                
                    if dt.now().time() < time(15, 30, 0):
                        FromDateTime = dt.now() 
                    else:
                        FromDateTime = dt(dt.now().year, dt.now().month,dt.now().day, 15, 30, 0)
                    
                    if GreekMatch == "SENSIBULL":
                        tryMatchWith=TryMatchWith.SENSIBULL
                    else:
                        tryMatchWith=TryMatchWith.NSE
                        
                    dayCountType = DayCountType.CALENDARDAYS

                    IvGreeks = CalcIvGreeks( SpotPrice = SpotPrice,  FuturePrice = FuturePrice, AtmStrike = AtmStrike, AtmStrikeCallPrice = AtmStrikeCallPrice, AtmStrikePutPrice = AtmStrikePutPrice, ExpiryDateTime = ExpiryDateTime, ExpiryDateType = ExpiryDateType, FromDateTime = FromDateTime, tryMatchWith = tryMatchWith, dayCountType = dayCountType)
    
                    #print(f"SpotPrice={SpotPrice}, FuturePrice={FuturePrice},  AtmStrike={AtmStrike}, AtmStrikeCallPrice={AtmStrikeCallPrice}, AtmStrikePutPrice={AtmStrikePutPrice}, ExpiryDateTime={ExpiryDateTime},  ExpiryDateType={ExpiryDateType}, FromDateTime={FromDateTime}, tryMatchWith={tryMatchWith}")
                    
                    df_oc_pro.round({"CE LTP":2, 'PE LTP':2})
                    
                    for ind in df_oc_pro.index:
                        
                        StrikePrice= float(df_oc_pro['Strike'][ind])
                        StrikeCallPrice= float(df_oc_pro['CE LTP'][ind])
                        StrikePutPrice= float(df_oc_pro['PE LTP'][ind])
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
                    oco_pro.range("a3").options(index=False, header=False).value = df_oc_pro
                    sleep(selected_refreshrate) 
                    
                except Exception as e:
                    #print("Exception occur in parent loop: " + str(e))
                    pass
        
        sleep(5)

def Zerodha_Token():
    global kite, df_instrument
    global Initial_Subscribr_TokenList 
    Initial_Subscribr_TokenList = [256265] #subscribe Nifty 50
    print("Zerodha intrument token download started, may take upto 2-3 minutes ..")
    try:
        instruments = kite.instruments()
        df_instrument = pd.DataFrame(instruments, index=None)
        
        #subscribe active mcx symbol
        df_subscribe_token = df_instrument[(df_instrument.segment == 'MCX-FUT') & (df_instrument.name.isin(['CRUDEOIL','GOLDPETAL','NATURALGAS','SILVERM']))]

        Initial_Subscribr_TokenList.extend(df_subscribe_token['instrument_token'].tolist())
        
        #print(Initial_Subscribr_TokenList)
        try:
            df_instrument.to_csv("Instrument.csv",index=False)       
        except:
            pass
        print("Zerodha intrument token download completed")
    except Exception as e:
        #print(f"Exception in Zerodha_Token : {e}")
        pass
    
def StartThread():
    try:
        global excel_master
    
        Config_sheet = excel_master.sheets['Config']
        
        
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
global excel_master
excel_master = xw.Book(TradeTerminalFileName)
if(Zerodha_login() == 1 ):   
    try:
        Zerodha_Token()
        
        kws = kite.kws()

        # Assign the callbacks.
        kws.on_ticks = on_ticks
        kws.on_order_update = on_order_update
        kws.on_connect = on_connect
        #kws.on_close = on_close

        kws.on_error = on_error
        kws.on_reconnect = on_reconnect
        kws.on_noreconnect = on_noreconnect

        kws.connect(threaded=True)
    
        StartThread()
        print("Done")
    except Exception as e:
        print("Exception : " + str(e))
else:
    print("Credential is not correct")
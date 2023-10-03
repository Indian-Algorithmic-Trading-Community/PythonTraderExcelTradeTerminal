import sys, os

try:
    Version = sys.version[:4]
except Exception as e:
    print(f"{e}")
    pass
    
if Version != "3.11":
    Message = "This trade tool will work with python latest version 3.11 only, so please upgrade python from installed version " + str(Version) + "to python 3.11"
    print(Message)
    sys.exit()
    
try:
    from kiteconnect import KiteConnect, KiteTicker
except (ModuleNotFoundError, ImportError):
    print("KiteConnect module not found")
    os.system(f"{sys.executable} -m pip install -U kiteconnect")
finally:
    from kiteconnect import KiteConnect, KiteTicker 

try:
    from tzlocal import get_localzone
except (ModuleNotFoundError, ImportError):
    print("tzlocal module not found")
    os.system(f"{sys.executable} -m pip install -U tzlocal")
finally:
    from tzlocal import get_localzone
	
try:
    import psutil
except (ModuleNotFoundError, ImportError):
    print("psutil module not found")
    os.system(f"{sys.executable} -m pip install -U psutil")
finally:
    import psutil
    
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
    import pandas as pd
except (ModuleNotFoundError, ImportError):
    print("pandas module not found")
    os.system(f"{sys.executable} -m pip install -U pandas")
finally:
    import pandas as pd

try:
    import scipy
except (ModuleNotFoundError, ImportError):
    print("scipy module not found")
    os.system(f"{sys.executable} -m pip install -U scipy")
finally:
    import scipy
    
try:
    import sourcedefender
except (ModuleNotFoundError, ImportError):
    print("sourcedefender module not found")
    os.system(f"{sys.executable} -m pip install -U sourcedefender")
finally:
    import sourcedefender


try:    
    import Zerodha_Core_V3_003
except Exception as e:
    print(f"Zerodha_Core_V3_003.pye file not found/corrupted, please download the latest file from tinyurl.com/pythontrader : {e}")

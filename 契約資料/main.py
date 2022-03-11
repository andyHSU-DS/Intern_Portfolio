#載入套件
import plotly
import plotly.express as px
import pandas as pd
import plolty.graph_objects as go
import plotly.io as pio
pio.renders.default = 'notebook'
from plotly.subplots import make_subplots
import threading
import sys
import os


#抓資料
def collect_data():
    abspath = r'input/'
    files   = os.listdir(abspath)
    for file in files:
        if 'EC' in file:
            path       = abspath + file
        elif '姓名' in file:
            sales_list = abspath + file
    return path, sales_list
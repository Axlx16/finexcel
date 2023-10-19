# Used to filter tickers by their respective exchanges

from urllib.request import urlopen
import json
import time


def getExchange(ticker):
    try:
        url = "https://financialmodelingprep.com/api/v3/quote/" + ticker + "?apikey=ebfd012c9d42f3c46eee6210a20c90ce"
        response = urlopen(url)
        quote_data = json.loads(response.read())
        return quote_data[0]['exchange']
    
    except(Exception):
        return ""


def getTickerList():
    url = "https://financialmodelingprep.com/api/v3/financial-statement-symbol-lists?apikey=ebfd012c9d42f3c46eee6210a20c90ce"
    response = urlopen(url)
    data_json = json.loads(response.read())

    print(len(data_json))

    x=0
    for i in range(len(data_json)):
        
        time.sleep(0.5)
        exchange_name = getExchange(data_json[x])


        if exchange_name == 'NYSE':
            with open("TextFiles/NYSE_ticker_list.txt", "a") as text_file:
                text_file.write(str(data_json[x]) + "\n")
                print(data_json[x] + " sorted!")
                x+=1
                

        elif exchange_name == 'NASDAQ':
            with open("TextFiles/NASDAQ_ticker_list.txt", "a") as text_file:
                text_file.write(str(data_json[x]) + "\n")
                print(data_json[x] + " sorted!")
                x+=1
                

        else:
           with open("TextFiles/Additonal_ticker_list.txt", "a") as text_file:
                text_file.write(str(data_json[x]) + "\n") 
                print(data_json[x] + " failed!")
                x+=1
        


getTickerList()



print('Done!')
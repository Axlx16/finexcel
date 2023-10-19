import urllib.request
import json

class FundmentalData: 
    
    def __init__(self, ticker, api_key, period):
        self.ticker = ticker
        self.api_key = api_key
        self.period = period


    def getQuote(self):
        try:
            quote_url = "https://financialmodelingprep.com/api/v3/quote/" + self.ticker + "?apikey=" + self.api_key
            response = urllib.request.urlopen(quote_url)
            return json.loads(response.read())
                
        except(Exception):
            print(Exception)


    def getSharesFloat(self):
        try:
            shares_float_url = 'https://financialmodelingprep.com/api/v4/shares_float/?symbol=' + self.ticker + '&apikey=' + self.api_key
            response = urllib.request.urlopen(shares_float_url)
            return json.loads(response.read())
        
        except(Exception):
            print(Exception)


    def getCompanyProfile(self):
        company_profile_url = "https://financialmodelingprep.com/api/v3/profile/" + self.ticker +'?apikey=' + self.api_key
        response = urllib.request.urlopen(company_profile_url)
        return json.loads(response.read())
            

    def getTTMKeyMetrics(self):
        try:
            TTM_key_metrics = "https://financialmodelingprep.com/api/v3/key-metrics-ttm/" + self.ticker + "?apikey=" + self.api_key
            response = urllib.request.urlopen(TTM_key_metrics)
            return json.loads(response.read())
            
        except(Exception):
            print(Exception)


    def getTTMRatio(self):
        try:
            TTM_ratio_url = "https://financialmodelingprep.com/api/v3/ratios-ttm/" + self.ticker + "?apikey=" + self.api_key
            response = urllib.request.urlopen(TTM_ratio_url)
            return json.loads(response.read())
            
        except(Exception):
            print(Exception)
    
    def getEnterpriseValue(self):
        try:
            enterprise_value_url = "https://financialmodelingprep.com/api/v3/enterprise-values/" + self.ticker + "?limit=40&apikey=" + self.api_key 
            response = urllib.request.urlopen(enterprise_value_url)
            return json.loads(response.read())
        
        except(Exception):
            print(Exception)
        

        
    def getIncomeStatement(self):
        try:
            income_statement_url = 'https://financialmodelingprep.com/api/v3/' + "income-statement" + '/' + self.ticker + '?period=' + self.period + '&limit=400&apikey=' + self.api_key
            response = urllib.request.urlopen(income_statement_url)
            return json.loads(response.read())
            
        except(Exception):
            print(Exception)


    def getBalanceSheet(self):
        try:
            income_statement_url = 'https://financialmodelingprep.com/api/v3/' + "balance-sheet-statement" + '/' + self.ticker + '?period=' + self.period + '&limit=400&apikey=' + self.api_key
            response = urllib.request.urlopen(income_statement_url)
            return json.loads(response.read())
            
        except(Exception):
            print(Exception)


    def getCashFlow(self):
        try:
            income_statement_url = 'https://financialmodelingprep.com/api/v3/' + "cash-flow-statement" + '/' + self.ticker + '?period=' + self.period + '&limit=400&apikey=' + self.api_key
            response = urllib.request.urlopen(income_statement_url)
            return json.loads(response.read())
            
        except(Exception):
            print(Exception)


    def getCompanyRatio(self):
        try:
            company_ratio_url = 'https://financialmodelingprep.com/api/v3/ratios/' + self.ticker + '?period=' + self.period + '&limit=140&apikey=' + self.api_key
            response = urllib.request.urlopen(company_ratio_url)
            return json.loads(response.read())

        except(Exception):
            print(Exception)

    def getKeyMetrics(self):
        try:
            key_metrics_url = 'https://financialmodelingprep.com/api/v3/key-metrics/' + self.ticker + '?period=' + self.period + '&limit=40&apikey=' + self.api_key
            response = urllib.request.urlopen(key_metrics_url)
            return json.loads(response.read())

        except(Exception):
            print(Exception)
    
    def getDividend(self):
        try:
            dividend_url = 'https://financialmodelingprep.com/api/v3/historical-price-full/stock_dividend/' + self.ticker + '?apikey=' + self.api_key
            response = urllib.request.urlopen(dividend_url)
            return json.loads(response.read())

        except(Exception):
            print(Exception)

    



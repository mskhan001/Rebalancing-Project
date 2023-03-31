import pandas as pd

def getETFWebsiteData(tickerDatabase_df, etf_name):
    
    # Reading ETF Website Data for Current Shares
    etf_currentHoldings = pd.read_excel('Data/{}_Website_Data.xlsx'.format(etf_name))
    etf_currentHoldings = etf_currentHoldings[['ISIN', 'Shares']]
    etf_currentHoldings.rename(columns = {'Shares': 'Current {} Shares'.format(etf_name)}, inplace=True)
    
    return etf_currentHoldings
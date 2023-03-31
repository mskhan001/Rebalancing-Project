import pandas as pd


def getETFWebsiteData(tickerDatabase_df, etf_name):
    # Skipping the first 2 rows and last row
    etf_currentHoldings = pd.read_excel('Data/{}_Website_Data.xlsx'.format(etf_name), skiprows=4)

    # Keep rows only until 'CASH_USD' is found
    idx = etf_currentHoldings[etf_currentHoldings['Identifier']=='CASH_USD'].index[0]
    etf_currentHoldings =  etf_currentHoldings[:idx]
    etf_currentHoldings.drop(columns=['Name', 'Identifier', 'Weight', 'Sector', 'Local Currency'], inplace=True)
    
    etf_currentHoldings = etf_currentHoldings.merge(tickerDatabase_df[['ISIN', 'SEDOL']], left_on='SEDOL', right_on='SEDOL', how='left')
    etf_currentHoldings.rename(columns={'Shares Held':'Current {} Shares'.format(etf_name)}, inplace=True)
    etf_currentHoldings = etf_currentHoldings[['ISIN', 'Current {} Shares'.format(etf_name)]]

    return etf_currentHoldings
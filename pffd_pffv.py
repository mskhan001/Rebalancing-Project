import pandas as pd


def getETFWebsiteData(tickerDatabase_df, etf_name):
    # PFFD Website Data
    
    # Skipping the first 2 rows and last row
    etf_currentHoldings = pd.read_csv('Data/{}_Website_Data.csv'.format(etf_name), skiprows=2)
    etf_currentHoldings = etf_currentHoldings[:-1]
    
    # Keeping only the required Columns
    etf_currentHoldings = etf_currentHoldings[['SEDOL', 'Shares Held']]
    
    # To get the ISIN by merging on 'SEDOL'
    etf_currentHoldings = etf_currentHoldings.merge(tickerDatabase_df[~tickerDatabase_df['SEDOL'].isna()], left_on='SEDOL', right_on='SEDOL', how='left')
    etf_currentHoldings.rename(columns={'Shares Held':'Current {} Shares'.format(etf_name)}, inplace=True)
    etf_currentHoldings['Current {} Shares'.format(etf_name)] = etf_currentHoldings['Current {} Shares'.format(etf_name)].str.replace(',','').astype(float)
    etf_currentHoldings = etf_currentHoldings[['ISIN', 'Current {} Shares'.format(etf_name)]]
    return etf_currentHoldings
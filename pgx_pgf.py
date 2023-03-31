import pandas as pd


def getETFWebsiteData(tickerDatabase_df, etf_name):
    
    # Reading ETF Website Data for Current Shares
    etf_currentHoldings = pd.read_csv('Data/{}_Website_Data.csv'.format(etf_name))
    etf_currentHoldings = etf_currentHoldings[[' Holding Ticker', ' Shares/Par Value']]
    etf_currentHoldings[' Shares/Par Value'] = etf_currentHoldings[' Shares/Par Value'].str.replace(',','').astype(int) 
    
    website_data_corrections = {
    'ATH-PA': 'ATH.PRA',
    'SCE-PK': 'SCE.PRK',
    'PSA-PM': 'PSA.PRM',
    'CTA-PB': 'CTA.PRB',
    'TRTN-PB': 'TRTN.PRB',
    'SJI': 'SJIJ'
    }
    

    # Changing the Ticker in ETF Website Data to IB Format
    def getIBFormat(x):
        return ".PR".join(x.split(' '))

    etf_currentHoldings['Ticker']=etf_currentHoldings[' Holding Ticker'].apply(getIBFormat)

    ## Applying Corrections in the Ticker format
    for holding_ticker in website_data_corrections:
        etf_currentHoldings.loc[etf_currentHoldings[' Holding Ticker']==holding_ticker, 'Ticker'] = website_data_corrections[holding_ticker]
    
    etf_currentHoldings = etf_currentHoldings.merge(tickerDatabase_df, left_on='Ticker', right_on='Ticker', how='left')
    etf_currentHoldings = etf_currentHoldings[['ISIN','Ticker',' Shares/Par Value']]
    etf_currentHoldings.rename(columns={' Shares/Par Value':'Current {} Shares'.format(etf_name)}, inplace=True)
    etf_currentHoldings = etf_currentHoldings[['ISIN', 'Current {} Shares'.format(etf_name)]]

    return etf_currentHoldings
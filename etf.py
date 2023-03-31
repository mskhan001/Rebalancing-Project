import numpy as np
import pandas as pd

from etfdata import etf_cash, etf_index_mapping, etf_marketCap
from pff import getETFWebsiteData as get_PFF_WebsiteData
from pffd_pffv import getETFWebsiteData as get_PFFD_PFFV_WebsiteData
from pgx_pgf import getETFWebsiteData as get_PGX_PGF_WebsiteData
from psk import getETFWebsiteData as get_PSK_WebsiteData

etf_function_mapping = {
    "PFF" : get_PFF_WebsiteData,
    "PFFD": get_PFFD_PFFV_WebsiteData,
    "PFFV": get_PFFD_PFFV_WebsiteData,
    "PGX" : get_PGX_PGF_WebsiteData,
    "PGF" : get_PGX_PGF_WebsiteData,
    "PSK" : get_PSK_WebsiteData
}

def getTickerDatabase():
    tickerDatabase_df = pd.read_excel('Static Data/TickerDatabase.xlsx')
    tickerDatabase_df = tickerDatabase_df[['ISIN', 'SEDOL', 'Ticker', 'Last Price']]
    
    return tickerDatabase_df

def getProjectedUniverseDF(etf_name):
    
    # Reading ICE Data For Projected Universe
    projected_universe_df = pd.read_excel('Data/{}-Projected.xlsx'.format(etf_index_mapping[etf_name]), skiprows=[0])
    
    # When reading data, below strings are read, they must be marked as np.nan
    msg1 = 'Any unauthorized use or disclosure is prohibited. Nothing herein should in any way be deemed to alter the legal rights and obligations contained in agreements between any ICE Data Services entity ("ICE") and their clients relating to any of the Indices or products or services described herein. The information provided by ICE and contained herein is subject to change without notice and does not constitute any form of representation, or undertaking.  ICE and its affiliates make no warranties whatsoever, either express or implied, as to merchantability, fitness for a particular purpose, or any other matter in connection with the information provided. Without limiting the foregoing, ICE and its affiliates makes no representation or warranty that any information provided hereunder are complete or free from errors, omissions, or defects. All information provided by ICE is owned by or licensed to ICE. ICE retains exclusive ownership of the ICE Indices, including the ICE BofAML Indexes, and the analytics used to create this analysis ICE may in its absolute discretion and without prior notice revise or terminate the ICE information and analytics at any time. The information in this analysis is for internal use only and redistribution of this information to third parties is expressly prohibited.'
    msg2 = 'Neither the analysis nor the information contained therein constitutes investment advice or an offer  or an invitation to make an offer  to buy or sell any securities or any options  futures or other derivatives related to such securities. The information and calculations contained in this analysis have been obtained from a variety of sources  including those other than ICE and ICE does not guarantee their accuracy.  Prior to relying on any ICE information and/or the execution of a security trade based upon such ICE information, you are advised to consult with your broker or other financial representative to verify pricing information. There is no assurance that hypothetical results will be equal to actual performance under any market conditions. THE ICE INFORMATION IS PROVIDED TO THE USERS "AS IS." NEITHER ICE, NOR ITS AFFILIATES, NOR ANY THIRD PARTY DATA PROVIDER WILL BE LIABLE TO ANY USER OR ANYONE ELSE FOR ANY INTERRUPTION, INACCURACY, ERROR OR OMISSION, REGARDLESS OF CAUSE, IN THE ICE INFORMATION OR FOR ANY DAMAGES RESULTING THEREFROM. In no event shall ICE or any of its affiliates, employees  officers  directors or agents of any such persons have any liability to any person or entity relating to or arising out of this information, analysis  or the indices  contained herein.'
    
    # Projected Universe Data
    projected_universe_df = projected_universe_df.replace('NaN', np.nan)
    projected_universe_df = projected_universe_df.replace(msg1, np.nan)
    projected_universe_df= projected_universe_df.replace(msg2, np.nan)

    #drop row if ISIN number is Nan
    projected_universe_df.dropna(subset=['ISIN number'], inplace=True)

    # Storing only the required columns
    projected_universe_df = projected_universe_df[['ISIN number', '% Mkt Value']]
    
    projected_universe_df.rename(columns={'% Mkt Value':'Projected % Mkt Cap'}, inplace=True)
    return projected_universe_df

def getOutputDF(etf_currentHoldings, projected_universe_df, tickerDatabase_df, etf_name):
    
    # DataFrame for our final Output
    etf_df = pd.DataFrame()
    
    # To get all the unique ISIN
    unique_ISIN = pd.concat([etf_currentHoldings['ISIN'], projected_universe_df['ISIN number']]).drop_duplicates().reset_index(drop=True)
    
    etf_df['ISIN'] = unique_ISIN
    etf_df.dropna(subset=['ISIN'], inplace=True)
    etf_df = etf_df.merge(etf_currentHoldings[['ISIN', 'Current {} Shares'.format(etf_name)]], left_on='ISIN', right_on='ISIN', how='left')
    etf_df = etf_df.merge(projected_universe_df, left_on='ISIN', right_on='ISIN number', how='left').drop(columns=['ISIN number'], axis=1)
    etf_df = etf_df.merge(tickerDatabase_df[['ISIN', 'Ticker', 'Last Price']], left_on='ISIN', right_on='ISIN', how='left')


    etf_df.fillna(0, inplace=True)
    
    etf_df['Projected {} Shares'.format(etf_name)] = etf_df.apply(lambda x: getProjectedShares(x, etf_name), axis=1)
    etf_df['Difference'] = etf_df.apply(lambda x: getDifference(x, etf_name), axis=1)
    etf_df.sort_values(by='Difference', key=abs, ascending=False, inplace=True)
    
    return etf_df

def getProjectedShares(x, etf_name):
    if(not isinstance(x['Last Price'], str) and x['Last Price']!=0):
        return np.rint(((etf_marketCap[etf_name]-etf_cash[etf_name])*x['Projected % Mkt Cap'])/(x['Last Price']*100))
    return 0

def getDifference(x, etf_name):
    if((not isinstance(x['Projected {} Shares'.format(etf_name)], str)) and (not isinstance(x['Current {} Shares'.format(etf_name)], str))):
        return x['Projected {} Shares'.format(etf_name)]-x['Current {} Shares'.format(etf_name)]
    return np.nan

def writeToFile(etf_df, etf_name):
    # Output File Name
    outputFile = 'Output/ETF/{} ETF vs ICE.xlsx'.format(etf_name)
    
    with pd.ExcelWriter(outputFile, mode="w", engine='xlsxwriter') as writer:
        etf_df.to_excel(writer, sheet_name=etf_name, columns=['ISIN', 'Ticker', 'Last Price', 'Projected % Mkt Cap', 'Current {} Shares'.format(etf_name), 'Projected {} Shares'.format(etf_name), 'Difference'])
        
        # Formatting
        workbook = writer.book
        worksheet = writer.sheets[etf_name]

        cellFormat = workbook.add_format({'num_format': '#,##0'})
        worksheet.set_column('F:H', 10, cellFormat)

def main():
    for etf_name in etf_function_mapping.keys():
        tickerDatabase_df = getTickerDatabase()
        projected_universe_df = getProjectedUniverseDF(etf_name)
        etf_currentHoldings = etf_function_mapping[etf_name](tickerDatabase_df, etf_name)
        etf_df = getOutputDF(etf_currentHoldings, projected_universe_df, tickerDatabase_df, etf_name)
        writeToFile(etf_df, etf_name)

main()
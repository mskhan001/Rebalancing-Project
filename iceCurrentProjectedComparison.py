#!/usr/bin/env python
# coding: utf-8

# In[1]:

import os

import numpy as np
import pandas as pd
from babel.numbers import format_currency

# In[2]:


# # Total market Cap for PFF/PGX, needs to be manually updated Can be automated later
current_market_cap = {
    "PFF": 13679614009,
    "PGX": 4902900000,
    "PSK": 1067240000,
    "PFFD": 2360000000,
    "PGF": 1179500000,
    "PFXF": 1000000000,
    "PFFV": 282740000,
    "FPE": 6076702207
}

current_etf_cash = {
    "PFF": 90000000,
    "PGX": 40000000,
    "PSK": 1000000,
    "PFFD": 0,
    "PGF": 0,
    "PFXF": 0,
    "PFFV": 0,
    "FPE": 0
}

index_etf_mapping = {
    "PFF": "PHGY",
    "PGX": "P0P4",
    "PSK": "PFAR",
    "PFFD": "PLCR",
    "PGF": "PFAF",
    "PFXF": "PFAN",
    "PFFV": "PFTF"
}

outputFile = 'Rebalancing Numbers - ICE Index.xlsx'


# In[3]:


def getICEData(index_name, etf_name):
    '''
    Reads Data from the ICE Current and ICE Projected Universe Excel Files and returns the DataFrames after cleaning Data
    '''

    # Reading the ICE Data for Current and Projected Universe
    ice_currentUniverse_df = pd.read_excel(
        'Data/{}-Current.xlsx'.format(index_name), skiprows=[0])
    ice_projectedUniverse_df = pd.read_excel(
        'Data/{}-Projected.xlsx'.format(index_name), skiprows=[0])

    # DataFrame for our final Output
    output_df = pd.DataFrame()

    # When reading data, below strings are read, they must be marked as np.nan
    msg1 = 'Any unauthorized use or disclosure is prohibited. Nothing herein should in any way be deemed to alter the legal rights and obligations contained in agreements between any ICE Data Services entity ("ICE") and their clients relating to any of the Indices or products or services described herein. The information provided by ICE and contained herein is subject to change without notice and does not constitute any form of representation, or undertaking.  ICE and its affiliates make no warranties whatsoever, either express or implied, as to merchantability, fitness for a particular purpose, or any other matter in connection with the information provided. Without limiting the foregoing, ICE and its affiliates makes no representation or warranty that any information provided hereunder are complete or free from errors, omissions, or defects. All information provided by ICE is owned by or licensed to ICE. ICE retains exclusive ownership of the ICE Indices, including the ICE BofAML Indexes, and the analytics used to create this analysis ICE may in its absolute discretion and without prior notice revise or terminate the ICE information and analytics at any time. The information in this analysis is for internal use only and redistribution of this information to third parties is expressly prohibited.'
    msg2 = 'Neither the analysis nor the information contained therein constitutes investment advice or an offer  or an invitation to make an offer  to buy or sell any securities or any options  futures or other derivatives related to such securities. The information and calculations contained in this analysis have been obtained from a variety of sources  including those other than ICE and ICE does not guarantee their accuracy.  Prior to relying on any ICE information and/or the execution of a security trade based upon such ICE information, you are advised to consult with your broker or other financial representative to verify pricing information. There is no assurance that hypothetical results will be equal to actual performance under any market conditions. THE ICE INFORMATION IS PROVIDED TO THE USERS "AS IS." NEITHER ICE, NOR ITS AFFILIATES, NOR ANY THIRD PARTY DATA PROVIDER WILL BE LIABLE TO ANY USER OR ANYONE ELSE FOR ANY INTERRUPTION, INACCURACY, ERROR OR OMISSION, REGARDLESS OF CAUSE, IN THE ICE INFORMATION OR FOR ANY DAMAGES RESULTING THEREFROM. In no event shall ICE or any of its affiliates, employees  officers  directors or agents of any such persons have any liability to any person or entity relating to or arising out of this information, analysis  or the indices  contained herein.'

    # Current Universe Data
    ice_currentUniverse_df = ice_currentUniverse_df.replace('NaN', np.nan)
    ice_currentUniverse_df = ice_currentUniverse_df.replace(msg1, np.nan)
    ice_currentUniverse_df = ice_currentUniverse_df.replace(msg2, np.nan)

    # drop row if ISIN number is Nan
    ice_currentUniverse_df.dropna(subset=['ISIN number'], inplace=True)

    # Projected Universe Data
    ice_projectedUniverse_df = ice_projectedUniverse_df.replace('NaN', np.nan)
    ice_projectedUniverse_df = ice_projectedUniverse_df.replace(msg1, np.nan)
    ice_projectedUniverse_df = ice_projectedUniverse_df.replace(msg2, np.nan)

    # drop row if ISIN number is Nan
    ice_projectedUniverse_df.dropna(subset=['ISIN number'], inplace=True)

    return ice_currentUniverse_df, ice_projectedUniverse_df, output_df


# In[4]:


def initiateOutputDF(output_df, ice_currentUniverse_df, ice_projectedUniverse_df):
    # Storing the unique ISIN ID's from both the current and projected universe data
    unique_ISIN = pd.concat([ice_currentUniverse_df['ISIN number'],
                            ice_projectedUniverse_df['ISIN number']]).drop_duplicates().reset_index(drop=True)
    output_df['ISIN number'] = unique_ISIN

    # Storing the Prices against the
    isin_to_ticker_df = pd.read_excel('Static Data/ISINtoTicker.xlsx')
    output_df = output_df.merge(
        isin_to_ticker_df, left_on='ISIN number', right_on='ISIN', how='left')
    output_df = output_df.drop(
        ['ISIN', 'Activ Last Price Formula', 'Activ Ticker', 'Bloomberg File Ticker'], axis=1)
    return output_df


# In[5]:


def getCurrentMarketCap(x, ice_currentUniverse_df):
    isid = x['ISIN number']
    ticker = ice_currentUniverse_df[ice_currentUniverse_df['ISIN number'] == isid]
    if (len(ticker) > 0):
        return ticker.iloc[0]['% Mkt Value']
    return 0


def getProjectedMarketCap(x, ice_projectedUniverse_df):
    isid = x['ISIN number']
    ticker = ice_projectedUniverse_df[ice_projectedUniverse_df['ISIN number'] == isid]
    if (len(ticker) > 0):
        return ticker.iloc[0]['% Mkt Value']
    return 0


def getCurrentETFShares(x, index_name, etf_name):
    if (not isinstance(x['Last Price'], str)):
        return np.rint(((current_market_cap[etf_name]-current_etf_cash[etf_name])*x['Current % Mkt Cap'])/(x['Last Price']*100))
    return np.nan


def getProjectedETFShares(x, index_name, etf_name):
    if (not isinstance(x['Last Price'], str)):
        return np.rint(((current_market_cap[etf_name]-current_etf_cash[etf_name])*x['Projected % Mkt Cap'])/(x['Last Price']*100))
    return np.nan


# In[6]:


def modifyOutputDF(output_df, ice_currentUniverse_df, ice_projectedUniverse_df, index_name, etf_name):
    output_df['Current % Mkt Cap'] = output_df.apply(
        lambda x: getCurrentMarketCap(x, ice_currentUniverse_df), axis=1)
    output_df['Projected % Mkt Cap'] = output_df.apply(
        lambda x: getProjectedMarketCap(x, ice_projectedUniverse_df), axis=1)

    output_df['Current {} Shares'.format(etf_name)] = output_df.apply(
        lambda x: getCurrentETFShares(x, index_name, etf_name), axis=1)
    output_df['Projected {} Shares'.format(etf_name)] = output_df.apply(
        lambda x: getProjectedETFShares(x, index_name, etf_name), axis=1)
    output_df['Difference'] = output_df.apply(lambda x: x['Projected {} Shares'.format(
        etf_name)]-x['Current {} Shares'.format(etf_name)], axis=1)

    # Sorting by Ticker Name
    output_df = output_df.sort_values(
        'Ticker').reset_index().drop('index', axis=1)
    return output_df


# In[7]:


def calculateAggregateNumbers(output_df, index_name, etf_name):
    # Number of Shares
    total_buys_ice = output_df['Difference'].where(
        output_df['Difference'] > 0).sum()
    total_sells_ice = output_df['Difference'].where(
        output_df['Difference'] < 0).sum()

    # Total Dollar Buying
    total_dollar_buying = output_df[output_df['Difference'] > 0].apply(
        lambda x: x['Last Price']*x['Difference'], axis=1).sum()
    total_dollar_selling = output_df[output_df['Difference'] < 0].apply(
        lambda x: x['Last Price']*x['Difference'], axis=1).sum()

    total_transactions_df = pd.DataFrame({" ": [total_buys_ice,
                                                total_sells_ice,
                                                total_dollar_buying,
                                                abs(total_dollar_selling),
                                                current_market_cap[etf_name],
                                                current_etf_cash[etf_name]]},
                                         index=['Total {} Buys (Number of Shares)'.format(etf_name),
                                                'Total {} Sells (Number of Shares)'.format(
                                                    etf_name),
                                                'Total {} Buying (in $)'.format(
                                                    etf_name),
                                                'Total {} Selling (in $)'.format(
                                                    etf_name),
                                                "{} NAV".format(etf_name),
                                                "{} Cash".format(etf_name)])

    return output_df, total_transactions_df


# In[8]:


def compareWithOld(output_df, index_name, etf_name):
    old_df = pd.read_excel('Old/{}'.format(outputFile),
                           sheet_name='{}'.format(etf_name))

    output_df = output_df.merge(old_df[['ISIN number', 'Difference Today']],
                                left_on='ISIN number', right_on='ISIN number', how='left')
    output_df.rename(columns={'Difference Today': 'Difference Yesterday',
                     'Difference': 'Difference Today'}, inplace=True)
    output_df['Difference from Yesterday'] = output_df['Difference Today'] - \
        output_df['Difference Yesterday']

    output_df['Difference $'] = (output_df['Difference Today'] * output_df['Last Price']).apply(
        lambda x: format_currency(x, currency="USD", locale="en_US"))
    output_df['Difference of Difference $'] = (
        output_df['Difference from Yesterday'] * output_df['Last Price'])
    output_df.sort_values(by='Difference of Difference $',
                          key=abs, ascending=False, inplace=True)

    output_df['Difference of Difference $'] = output_df['Difference of Difference $'].apply(
        lambda x: format_currency(x, currency="USD", locale="en_US"))

    return output_df


# In[9]:


def getFileWriter():
    '''Return Append Mode if file is present else Write Mode '''
    if (os.path.isfile(outputFile)):
        return pd.ExcelWriter(outputFile, mode="a", if_sheet_exists="replace", engine='openpyxl')
    return pd.ExcelWriter(outputFile, mode="w", engine='openpyxl')


def writeToFile(output_df, total_transactions_df, index_name, etf_name):
    writer = getFileWriter()
    output_df.to_excel(writer, sheet_name='{}'.format(etf_name))
    total_transactions_df.to_excel(
        writer, sheet_name='Total {} Transactions ICE'.format(etf_name))
    writer.close()


# In[10]:


def main():

    for key in index_etf_mapping:
        etf_name = key
        index_name = index_etf_mapping[key]

        ice_currentUniverse_df, ice_projectedUniverse_df, output_df = getICEData(
            index_name, etf_name)
        output_df = initiateOutputDF(
            output_df, ice_currentUniverse_df, ice_projectedUniverse_df)
        output_df = modifyOutputDF(
            output_df, ice_currentUniverse_df, ice_projectedUniverse_df, index_name, etf_name)
        output_df, total_transactions_df = calculateAggregateNumbers(
            output_df, index_name, etf_name)
        output_df = compareWithOld(output_df, index_name, etf_name)
        writeToFile(output_df, total_transactions_df, index_name, etf_name)


# In[11]:


main()

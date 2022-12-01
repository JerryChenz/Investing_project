import re
import json
from bs4 import BeautifulSoup
import requests
import pandas as pd


def scrap_data(ticker, url):
    """Scrap data from Yahoo Finance for an input ticker"""

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome'
                      '/71.0.3578.98 Safari/537.36'}
    response = requests.get(url.format(ticker, ticker), headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')

    pattern = re.compile(r'\s--\sData\s--\s')
    script_data = soup.find('script', text=pattern).contents[0]
    start = script_data.find("context") - 2
    json_data = json.loads(script_data[start:-12])

    return json_data


def _parse_table(json_info):
    """Parse the raw list. Return a clean_dict with t-i year and value pair"""
    clean_dict = {}

    year = int(list(reversed(json_info))[0]['asOfDate'][:4])
    for yearly in reversed(json_info):
        if yearly:
            clean_dict[year] = yearly['reportedValue']['raw']
        else:
            clean_dict[year] = 0
        year -= 1

    return clean_dict


def get_income_statement(ticker):
    """Scrape income statement from Yahoo Finance for a given ticker"""

    url_financials = "https://finance.yahoo.com/quote/{}/financials?p={}"
    is_dict = {}

    json_data = scrap_data(ticker, url_financials)
    json_is = json_data['context']['dispatcher']['stores']['QuoteTimeSeriesStore']['timeSeries']

    # sales
    sales_dict = _parse_table(json_is['annualTotalRevenue'])
    is_dict['sales'] = sales_dict
    # cogs
    cogs_dict = _parse_table(json_is['annualCostOfRevenue'])
    is_dict['cogs'] = cogs_dict
    # operating expenses
    op_cost_dict = _parse_table(json_is['annualOperatingExpense'])
    is_dict['op_cost'] = op_cost_dict
    # interest expense
    interest_dict = _parse_table(json_is['annualInterestExpense'])
    is_dict['interest'] = interest_dict
    # annualNetIncome
    ni_dict = _parse_table(json_is['annualNetIncome'])
    is_dict['net_income'] = ni_dict

    return pd.DataFrame(is_dict).transpose()


def get_balance_sheet(ticker):
    """Scrape balance sheet from Yahoo Finance for a given ticker"""

    bs_dict = {}

    # Scrap Data
    url_bs = "https://finance.yahoo.com/quote/{}/balance-sheet?p={}"
    json_data = scrap_data(ticker, url_bs)
    json_bs = json_data['context']['dispatcher']['stores']['QuoteTimeSeriesStore']['timeSeries']

    # current assets
    current_assets_dict = _parse_table(json_bs['annualCurrentAssets'])
    # current liabilities
    current_liabilities_dict = _parse_table(json_bs['annualCurrentLiabilities'])
    # ST Interest-bearing Debt
    short_debt_dict = _parse_table(json_bs['annualCurrentDebtAndCapitalLeaseObligation'])
    # LT Interest-bearing Debt
    long_debt_dict = _parse_table(json_bs['annualLongTermDebtAndCapitalLeaseObligation'])
    # Equity
    equity_dict = _parse_table(json_bs['annualTotalEquityGrossMinorityInterest'])
    # Minority interests
    minority_interest_dict = _parse_table(json_bs['annualMinorityInterest'])
    # Cash & Cash Equivalents
    cash_dict = _parse_table(json_bs['annualCashAndCashEquivalents'])
    # PP&E
    ppe_dict = _parse_table(json_bs['annualNetPPE'])
    # Non-operating Asset Estimate
    non_op_asset_dict = {}
    short_investment_dict = _parse_table(json_bs['annualOtherShortTermInvestments'])
    non_current_assets_dict = _parse_table(json_bs['annualTotalNonCurrentAssets'])
    goodwill_dict = _parse_table(json_bs['annualGoodwillAndOtherIntangibleAssets'])
    prepaid_dict = _parse_table(json_bs['annualNonCurrentPrepaidAssets'])

    for i in non_current_assets_dict.keys():
        non_op_asset_dict[i] = short_investment_dict[i] + non_current_assets_dict[i] \
                               - ppe_dict[i] - goodwill_dict[i] - prepaid_dict[i]

    # Initialize bs_dict
    bs_dict['current_assets'] = current_assets_dict
    bs_dict['current_liabilities'] = current_liabilities_dict
    bs_dict['short_debt'] = short_debt_dict
    bs_dict['long_debt'] = long_debt_dict
    bs_dict['equity'] = equity_dict
    bs_dict['minority_interest'] = minority_interest_dict
    bs_dict['cash'] = cash_dict
    bs_dict['non_op_asset'] = non_op_asset_dict
    bs_dict['ppe'] = ppe_dict

    return pd.DataFrame(bs_dict).transpose()

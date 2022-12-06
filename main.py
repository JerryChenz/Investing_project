import security_mod
import yfinance as yf
from datetime import datetime
import pipline_mod


def gen_val_xlsx(ticker):
    """generate or update a valuation file with argument, ticker"""

    s = security_mod.Stock(ticker)
    # load from yahoo finance
    s.load_from_yf()
    # generates or update the valuation file
    s.create_val_xlsx()


if __name__ == '__main__':
    # stare_list = ['0806.HK', '1475.HK', '1766.HK', '6186.HK']
    # for s in stare_list:
    #     gen_val_xlsx(s)
    o = pipline_mod.Pipeline()
    o.load_opportunities()
    print(o.assets)

    # stock_info = yahoo_fin.get_quote_table(stock)
    # company_info = yahoo_fin.get_quote_data(stock)
    # print(round(stock_info['Quote Price'],2))
    # print(company_info['currency'])
    # print(company_info['exchange'])
    # print(stock_info['Earnings Date'])
    # print(stock_info['Forward Dividend & Yield'])
    # print(stock_info['Ex-Dividend Date'])

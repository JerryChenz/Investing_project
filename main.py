import security_mod

if __name__ == '__main__':
    ticker = '1475.HK'
    s = security_mod.Stock(ticker)
    s.create_val_xlsx()

    # stock_info = yahoo_fin.get_quote_table(stock)
    # company_info = yahoo_fin.get_quote_data(stock)
    # print(round(stock_info['Quote Price'],2))
    # print(company_info['currency'])
    # print(company_info['exchange'])
    # print(stock_info['Earnings Date'])
    # print(stock_info['Forward Dividend & Yield'])
    # print(stock_info['Ex-Dividend Date'])

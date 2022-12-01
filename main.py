import scrap_mod
import yahoo_fin
import excel_mod


class Securities:
    """Parent class"""

    def __init__(self, security_code):
        self.security_code = security_code


class Stock(Securities):
    """child class"""

    def __init__(self, security_code):
        """ """
        super().__init__(security_code)
        self.is_df = scrap_mod.get_income_statement(security_code)
        self.bs_df = scrap_mod.get_balance_sheet(security_code)


if __name__ == '__main__':
    ticker = '1475.HK'
    raw_data_filename = 'Raw_fin_data.xlsx'

    #s = Stock(ticker)
    df = scrap_mod.get_income_statement(ticker)
    print(df)
    #print(s.is_df)

    # is_df.to_csv(f'{stock}_income_statement.csv', sep=',', encoding='utf-8')
    # bs_df.to_csv(f'{stock}_balance_sheet.csv', sep=',', encoding='utf-8')

    # stock_info = yahoo_fin.get_quote_table(stock)
    # company_info = yahoo_fin.get_quote_data(stock)
    # print(round(stock_info['Quote Price'],2))
    # print(company_info['currency'])
    # print(company_info['exchange'])
    # print(stock_info['Earnings Date'])
    # print(stock_info['Forward Dividend & Yield'])
    # print(stock_info['Ex-Dividend Date'])

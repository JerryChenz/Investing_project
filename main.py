import security_mod
import pipline_mod


def gen_val_xlsx(ticker):
    """generate or update a valuation file with argument, ticker"""

    stock = security_mod.Stock(ticker)
    try:
        # load from yahoo finance
        stock.load_from_yf()
    except KeyError:
        print("Check your stock ticker")
    else:
        # generates or update the valuation file
        stock.create_val_xlsx()


def update_pipeline_monitor():
    """Update the pipeline monitor"""

    o = pipline_mod.Pipeline()
    o.load_opportunities()


if __name__ == '__main__':
    # stare_list = ['0806.HK', '1475.HK', '1766.HK', '6186.HK']
    # for s in stare_list:
    #     gen_val_xlsx(s)
    update_pipeline_monitor()

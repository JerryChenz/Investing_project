from datetime import datetime
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


def days_between(d1, d2):
    d1 = datetime.strptime(d1, "%Y-%m-%d")
    d2 = datetime.strptime(d2, "%Y-%m-%d")
    return abs((d2 - d1).days)


if __name__ == '__main__':
    #stare_list = ['600751.SS', '603338.SS']
    #for s in stare_list:
    #    gen_val_xlsx(s)
    update_pipeline_monitor()

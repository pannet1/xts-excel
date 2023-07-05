import json
import pandas as pd
from pandas import json_normalize
from omspy_brokers.XTConnect.xts import Xts
from omspy_brokers.XTConnect.Connect import XTSConnect
from toolkit.fileutils import Fileutils
from toolkit.utilities import Utilities
from toolkit.logger import Logger
from instrument import sec_dir, msg_code, exch
from instrument import dump_instruments, get_str_inst_fm_id, get_lst_dct_inst, get_exch_lst_inst, get_id_fm_str_inst
from msexcel import Msexcel
from pprint import pprint


logging = Logger(30)
glb_tline = {}
glb_dct_inst = {}


def get_interactive(i):
    """
    # login methods
    """
    xt = Xts(i['api'], i["secret"], i["userid"])
    if not xt.authenticate():
        logging.error("interactive api could not login")
        SystemExit()
    else:
        return xt


def get_marketdata(m):
    md = XTSConnect(m['api'], m['secret'], source="WEBAPI")
    resp = md.marketdata_login()
    print(resp)
    return md


def get_exch_colon_inst(exch_id: int, inst_id: int) -> str:
    global glb_dct_inst
    logging.info(glb_dct_inst)
    exch_inst = str(exch_id) + "_" + str(inst_id)
    exch_colon_inst = glb_dct_inst.get(exch_inst, False)
    if exch_colon_inst:
        return exch_colon_inst
    for key, value in exch.items():
        if value["id"] == exch_id:
            exch_key = key
            break
    inst = get_str_inst_fm_id(exch_key, inst_id)
    exch_colon_inst = exch_key + ":" + inst
    glb_dct_inst[exch_inst] = exch_colon_inst
    return exch_colon_inst


def resp_to_quote(data: dict):
    global glb_tline
    dct_1 = json.loads(data)
    keys_to_extract = [
        'ExchangeSegment',
        'ExchangeInstrumentID',
        'Open',
        'High',
        'Low',
        'Close',
        'LastTradedPrice',
        'AverageTradedPrice',
        'AskInfo',
        'BidInfo'
    ]
    dct = {k: v for k, v in dct_1.items() if k in keys_to_extract}
    dct['Ask'] = dct['AskInfo'].get('Price')
    dct['Bid'] = dct['BidInfo'].get('Price')
    exch_inst = get_exch_colon_inst(
        dct["ExchangeSegment"], dct["ExchangeInstrumentID"])
    dct.pop('ExchangeSegment')
    dct.pop('ExchangeInstrumentID')
    dct.pop('AskInfo')
    dct.pop('BidInfo')
    glb_tline[exch_inst] = dct


futil = Fileutils()
i = futil.get_lst_fm_yml(sec_dir + "arham_interactive.yaml")
iApi = get_interactive(i)

"""
TODO: refactor to dump all files
"""
exchangesegments = [iApi.broker.EXCHANGE_NSEFO]
dump_file = sec_dir + "NFO" + ".txt"
if futil.is_file_not_2day(dump_file):
    resp = iApi.broker.get_master(exchangesegments)
    if (
        resp is not None
        and isinstance(resp, dict)
        and isinstance(resp['result'], dict)
    ):
        data = resp.get('result')
        dump_instruments(dump_file, data)
    else:
        print("unable to dump instruments")
else:
    print(f"{dump_file} file not modified")

str_active_sht = ""
m = futil.get_lst_fm_yml(sec_dir + "arham_marketdata.yaml")
mApi = get_marketdata(m)
msxl = Msexcel("../../../excel.xlsm")
sht_live = msxl.sheet("LIVE")
sht_marg = msxl.sheet("MARGIN")
sht_hold = msxl.sheet("HOLDINGS")
sht_postn = msxl.sheet("POSITION")
sht_order = msxl.sheet("ORDERBOOK")

while not msxl.books.active:
    sleep(1)
else:
    sht_hold.range('A1:AK100').value = ""
    sht_postn.range('A1:AK100').value = ""
    sht_order.range('A1:AK100').value = ""
while True:
    obj_xls_active = msxl.books.active
    if (obj_xls_active):
        str_active_sht = obj_xls_active.sheets.active.name
        if str_active_sht == "ORDERBOOK":
            resp = iApi.orders
            if resp:
                df = pd.DataFrame(resp)
                sht_order['A1'].options(index=False, header=True).value = df
        elif str_active_sht == "MARGIN":
            resp = iApi.margins
            if resp:
                # Flatten the nested dictionaries
                df = json_normalize(resp, sep='_')
                # Convert DataFrame to float
                df = df.astype(float)
                # Fill NaN values with 0
                df.fillna(0, inplace=True)
                # Keep only columns with values greater than zero
                df = df.loc[:, (df > 0).any()]
                # Drop rows that contain only 0 values
                df = df.loc[(df > 0).any(axis=1)]
                # Reset the index to numeric values
                df.reset_index(drop=True, inplace=True)
                sht_marg['A1'].options(index=False, header=True).value = df
        elif str_active_sht == "HOLDINGS":
            resp = iApi.holdings.get('Holdings', False)
            if resp:
                df = pd.DataFrame.from_dict(resp).T
                sht_hold['A1'].options(index=False, header=True).value = df
        else:
            resp = iApi.positions
            if resp:
                resp = [{k: v for k, v in item.items() if k != 'childPositions'} for item in resp]
                df = pd.DataFrame(resp).sort_values('Quantity')
                sht_postn['A1'].options(index=False, header=True).value = df
        Utilities().slp_til_nxt_sec()

    # get market watch from excel sheet
    col, dat = msxl.get_col_dat(sht_live, "B", "C", 1, 101)
    df_mw = pd.DataFrame(dat, columns=col).dropna(axis=0)
    # create exchange key and corresponding list of instruments
    dct_exch_lst_inst = get_exch_lst_inst(df_mw)
    # count number of exchanges
    num_of_exch = len(dct_exch_lst_inst.keys())
    # iterate for each exchange and its instruments
    for k, v in dct_exch_lst_inst.items():
        # convert exch and inst to numerical codes
        lst_dct_inst = get_lst_dct_inst(k, v)
        # get quote from API
        resp = mApi.get_quote(lst_dct_inst, msg_code["touchline"], "JSON")
        if (
            resp is not None
            and isinstance(resp, dict)
            and resp.get('result', False)
            and resp['result'].get('listQuotes', False)
        ):
            for quote in resp['result']['listQuotes']:
                # update glb_tline sack
                resp_to_quote(quote)
        # delta sleep after each iteration
        if num_of_exch>1:
            Utilities().slp_til_nxt_sec()

    # read the sheet again to get the updated market watch
    col, dat = msxl.get_col_dat(sht_live, "B", "C", 1, 101)
    df_mw = pd.DataFrame(dat, columns=col)
    df_mw['row_idx'] = df_mw.index
    df_mw.dropna(axis=0, inplace=True)
    # prepare data by splitting "Exch" and "Sym" from keys of global quote
    if any(glb_tline):
        lst_dct_tline = [{"Exch": key.split(":")[0], "Sym": key.split(
            ":")[1], **value} for key, value in glb_tline.items()]
        # create the dataframe
        df_quotes = pd.DataFrame(lst_dct_tline)
        # merge quotes and mw so that mw can contain duplicates now
        df_new = df_mw.merge(df_quotes, on=["Exch", "Sym"], how="left").dropna()
        df_new.set_index('row_idx', inplace=True)
        # Assign the additional column values to the Excel sheet
        for i, (index, row) in enumerate(df_new.iterrows()):
            addr = 'B' + str(index + 2)
            if not sht_live:
                break
            sht_live.range(addr).value = row.values.tolist()
            # checking for empty qty, dir and product
            lst_cell_values = sht_live.range("L"+str(index+2)+":"+"N"+str(index+2)).value
            lst_order_values = sht_live.range("P" + str(index + 2) + ":" + "Q" + str(index + 2)).value
            # atleast one values of side, qty and otype is empty
            if any(item is None for item in lst_cell_values):
                # delete order number and status
                sht_live.range("P" + str(index + 2) + ":" + "Q" + str(index + 2)).value = ""
            # at least one of the order value is empty
            elif any(item is None for item in lst_order_values):
                """
                'exchangeSegment', 'exchangeInstrumentID', 'productType', 'orderType', 
                'orderSide', 'timeInForce', 'disclosedQuantity', 'orderQuantity', 'limitPrice', 
                'stopPrice', and 'orderUniqueIdentifier'
                """
                is_full, lst = msxl.get_lst_fm_rng(sht_live, "B", "N", index + 2)
                if(is_full):
                    exchangeSegment = exch.get(lst[0])['code']
                    exchangeInstrumentID = get_id_fm_str_inst(lst[0], lst[1])
                    dct_order = {
                        'exchangeSegment': exchangeSegment,
                        'exchangeInstrumentID': exchangeInstrumentID,
                        'product': lst[12], 'order_type': 'MARKET', 'side': lst[11],
                        'validity': 'DAY', 'quantity': lst[10], 'trigger_price': 0,
                        'price': 0,
                    }
                    order_id = iApi.order_place(**dct_order)
                    if order_id:
                        sht_live.range("P" + str(index + 2)).value = order_id
                        sht_live.range("Q" + str(index + 2)).value = "Success"
                    else:
                        sht_live.range("P" + str(index + 2)).value = 0
                        sht_live.range("Q" + str(index + 2)).value = "FAILED"


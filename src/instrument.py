
from typing import Union

# user settings
sec_dir = "../../../"

# system constants
msg_code = {
    "instrument_change": 1105,
    "touchline": 1501,
    "market_data": 1502,
    "candle_data": 1505,
    "market_status": 1507,
    "oi": 1510,
    "ltp": 1512,
}
exch = {
    "NSE": {"id": 1, "code": "NSECM"},
    "NFO": {"id": 2, "code": "NSEFO"},
}


def dump_instruments(dump_file: str, data: Union[str, str]) -> None:
    header = "ExchangeSegment | ExchangeInstrumentID | InstrumentType | Name | Description | Series | NameWithSeries | InstrumentID | PriceBand.High | PriceBand.Low | FreezeQty | TickSize | LotSize | Multiplier | UnderlyingInstrumentId | UnderlyingIndexName | ContractExpiration | StrikePrice | OptionType | displayName | PriceNumerator | PriceDenominator"
    header += "\n"
    with open(dump_file, "w") as file:
        file.write(header)
        file.write(data)


def get_str_inst_fm_id(exch_code: str, inst_id: int) -> str:
    def is_int(string):
        try:
            int(string)
            return True
        except ValueError:
            return False
    data_file = sec_dir + exch_code + ".txt"
    with open(data_file, "r") as file:
        contents = file.read()
        records = contents.split("\n")
        print(f"searching in {len(records)} records for {inst_id}")
        for record in records:
            fields = record.split("|")
            if is_int(fields[1]) and int(fields[1]) == inst_id:
                inst = fields[4]
                return inst
    return "INSTRUMENT_NOT_FOUND"

def get_id_fm_str_inst(exch_code: str, inst: str) -> int:
    data_file = sec_dir + exch_code + ".txt"
    with open(data_file, "r") as file:
        contents = file.read()
        records = contents.split("\n")
        for record in records:
            fields = record.split("|")
            if fields[4] == inst:
                int_inst = fields[1]
                return int_inst
    return 0

def get_lst_dct_inst(exch_key: str, lst_inst: list) -> list:
    """
        consumed by get quotes method
    """
    lst_dct_inst = []
    for inst in lst_inst:
        dct_inst = {}
        exch_id = exch[exch_key].get('id')
        dct_inst['exchangeSegment'] = exch_id
        dct_inst['exchangeInstrumentID'] = get_id_fm_str_inst(
            exch_key, inst)
        lst_dct_inst.append(dct_inst)
    return lst_dct_inst


def get_exch_lst_inst(df):
    unique_exch = df['Exch'].unique().tolist()
    dct_unique_sym = {}
    for exchange in unique_exch:
        filtered_df = df[df['Exch'] == exchange]
        dct_unique_sym[exchange] = filtered_df['Sym'].unique().tolist()
    return dct_unique_sym

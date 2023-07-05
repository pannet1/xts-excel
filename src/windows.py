from toolkit.fileutils import Fileutils
from toolkit.utilities import Utilities
from msexcel import Msexcel
import pandas as pd
from time import sleep

msxl = Msexcel("../../../excel.xlsm")
sht_live = msxl.sheet("Live")

while True:
    col, dat = msxl.get_col_dat(sht_live, "B", "B",  1 , 10)
    df_mwatch = pd.DataFrame(dat, columns=col).dropna(axis=0)
    print(df_mwatch)
    Utilities().slp_til_nxt_sec()
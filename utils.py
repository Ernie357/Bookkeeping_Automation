import pathlib
from typing import NamedTuple
from openpyxl.utils import column_index_from_string
from datetime import datetime

def get_full_script_dir():
    return pathlib.Path(__file__).parent.resolve()

merge_name_map = {
    "Inv_Nbr": column_index_from_string("G") - 1,
    "Estate_No": column_index_from_string("I") - 1,
    "Bill_To": column_index_from_string("J") - 1,
    "Address_1": column_index_from_string("K") - 1,
    "Address_2": column_index_from_string("L") - 1,
    "Estate_of": column_index_from_string("M") - 1,
    "M_1st_Run": column_index_from_string("O") - 1,
    "M_2nd_Run": column_index_from_string("P") - 1,
    "M_3rd_Run": column_index_from_string("Q") - 1, 
    "price": column_index_from_string("R") - 1,
    "QR_Image": column_index_from_string("X") - 1,
    "QR_Link": column_index_from_string("Y") - 1
}

merge_id_key = "Inv_Nbr"

def get_formatted_value(key: str, value: str):
    if key == "price" and value is not None:
        return float(value)
    if "Run" in key:
        if value == "None":
            return ""
        dt = datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%#m/%#d/%Y")
    return value
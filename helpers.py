from collections import namedtuple
import json

with open("config.json") as config_file:
    config = json.load(config_file)

invert = {
    "Cr": "Dr",
    "Dr": "Cr",
}

to_be_removed = "To be removed"

column_names = namedtuple(
    "column_names",
    [
        "vd", # Voucher Date
        "vtn", # Voucher Type Name
        "ln", # Ledger Name
        "la", # Ledger Amount
        "dr_cr", # Ledger Amount Dr/Cr
        "vn", # Voucher Narration
    ],
)

column_names = column_names(*config["column_names2"])

def clean(bank):
    return bank if "A/c" in bank or "Account" in bank or "A/c." in bank else f"{bank} A/c"

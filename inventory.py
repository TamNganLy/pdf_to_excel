# inventory.py

import os
from datetime import date, timedelta

BASE_PATH = r"\\SERVER2\Tech\Tesla Turnkey Inventory"
MAX_DAYS = 30

def find_previous_inventory():
    day = date.today() - timedelta(days=1)

    for _ in range(MAX_DAYS):
        filename = day.strftime("%m-%d-%Y") + " Tesla Turnkey Inventory.xlsx"
        full_path = os.path.join(BASE_PATH, filename)
        if os.path.exists(full_path):
            formula = (f",'{BASE_PATH}\\[{filename}]Sheet1'!$A$6:$C$1000,3,FALSE)")
            return formula
        day -= timedelta(days=1)
    return None
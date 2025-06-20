# utils.py
import os
import re
import pandas as pd
from glob import glob
from collections import namedtuple
import sys

Line = namedtuple('Line', 'PO_Number Part_Number Qty Date Inventory_Qty')

OPEN_ORDER_FOLDER = 'open_order_pdf/*'
SPECIAL_PART_FOLDER = 'open_order_pdf/*'
# Pattern for PO# -> PO #: 4900412040
PO_NUMBER_RE = re.compile(r'(PO #:) (.*)')

# Pattern for page#
PAGE_PATTERN = re.compile(r'Page \d+ of \d+')

# Pattern for part# style 1 (part# and price): 1636784-00-A 25.00 EA 2023-09-08 ...
PART_NUMBER_RE_1 = re.compile(r'(\d{7}-\w{2}-\w{1}) (\d+).\d{2}')
# Pattern for part# style 2(have -US): 1636784-00-A-US
PART_NUMBER_RE_2 = re.compile(r'(\d{7}-\d{2}-\w{1})')
# Pattern for part# style 3 (no rev):1636784-00
PART_NUMBER_RE_3 = re.compile(r'(\d{7}-\d{2})')
# Pattern for part# style 4 (part# and des change position): 1636784-00 Glass 25.00 EA 2023-09-08 ...
PART_NUMBER_RE_4 = re.compile(r'(\d{7}-\d{2}-\w{1}) .* (\d+).\d{2} \w+ \d{4}-\d{2}-\d{2}')
# Pattern for part# style 5 (part# and des are the same): 1127845-00-A 1127845-00-A 1.00EA 2025-05-27 ...
PART_NUMBER_RE_5 = re.compile(r'(\d{7}-\w{2}-\w{1}) (\d{7}-\d{2}-\w{1}) (\d+).\d{2} \w+ \d{4}-\d{2}-\d{2}')
# Pattern for part# style 6 (part# doesn't include '-'): 1658568 00 A 1.00EA 2025-05-27 ...
PART_NUMBER_RE_6 = re.compile(r'(\d{7} \d{2} \w{1}) (\d+).\d{2} \w+ \d{4}-\d{2}-\d{2}')


# Pattern for quality style 1 (no part# but descri):Glass 25.00 EA 2023-09-08 ...
QUANTITY_RE_1 = re.compile(r'(.*) (\d+).\d{2} \w+ \d{4}-\d{2}-\d{2}')
# Pattern for quality style 2 (not part#): 25.00 EA 2023-09-08 ...
QUANTITY_RE_3 = re.compile(r'(\d+).\d{2} \w+ \d{4}-\d{2}-\d{2}')

def collect_pdf_files(folder):
    path_file = pd.DataFrame(glob(OPEN_ORDER_FOLDER))
    path_file.columns = ['location']
    file_folder = {os.path.getmtime(path): path for path in path_file['location']}
    return dict(sorted(file_folder.items()))

def print_excel(lines, p_order, part, quality, d1, search_function, search_file, col):
    if search_file: 
        search_function = "=VLOOKUP(B" + str(col) + search_file
    else:
        search_function = ''
    lines.append(Line(p_order, part, quality, d1, search_function))

def get_base_folder():
    """Returns the directory where the .exe or script is located."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)  # .exe
    else:
        return os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))  # .py


def load_keyword_patterns(filename):
    filepath = os.path.join(get_base_folder(), filename)
    if not os.path.exists(filepath):
        print(f"Excel file not found: {filepath}")
        return []

    df = pd.read_excel(filepath)
    special_patterns = [
        [re.compile(r'(\d{3}.\d{3}.\d{2}) -\w{2} (\d+).\d{2}'), ''],
        [re.compile(r'(\d{3}.\d{3}.\d{2})-\w{2} (\d+).\d{2}'), '']
    ]

    for _, row in df.iterrows():
        keyword = row['keyword']
        part_name = row['part_name'] if pd.notna(row['part_name']) else ''

        # Build a safe regex: match keyword + quantity format
        pattern_str = rf'({re.escape(keyword)}) (\d+).\d{{2}}'
        pattern = re.compile(pattern_str)

        special_patterns.append([pattern, part_name])
    
    return special_patterns
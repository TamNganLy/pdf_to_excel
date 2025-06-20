# order_processor.py

import os
import pdfplumber
import pandas as pd
from datetime import date
from inventory import find_previous_inventory
from pdf_parser import extract_part_number, extract_description
from utils import Line, collect_pdf_files, load_keyword_patterns

class OrderProcessor:
    def __init__(self, pdf_folder):
        self.pdf_folder = pdf_folder
        self.col = 2
        self.first = True

    def run(self):      
        file_folder = collect_pdf_files(self.pdf_folder)

        for _, file in file_folder.items():
            lines = []
            self.col = self.process_file(file, lines, self.col)
            df = pd.DataFrame(lines)
            if self.first:
                df.to_csv('daily_open_order.csv', index=False)
                self.first = False
            else:
                df.to_csv('daily_open_order.csv', mode='a', index=False, header=False)

    def process_file(self, file_name, lines, col):
        SPECIAL_PATTERN = load_keyword_patterns('special_parts.xlsx')
        today = date.today().strftime("%m/%d/%Y")
        search_file = find_previous_inventory()
        has_part = False

        with pdfplumber.open(file_name) as pdf:
            has_part, col = extract_part_number(pdf, lines, today, has_part, search_file, col, SPECIAL_PATTERN)
            extract_description(pdf, lines, today, has_part)
        return col
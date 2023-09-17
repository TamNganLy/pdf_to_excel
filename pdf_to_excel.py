import re

import pdfplumber
import pandas as pd
from collections import namedtuple
from datetime import date
from glob import glob

Line = namedtuple('Line', 'PO_Number Part_Number Qty Date')

# Pattern for PO# -> PO #: 4900412040
PO_number_re = re.compile(r'(PO #:) (.*)')

# Pattern for part# style 1 (part# and price): 1636784-00-A 25.00 EA 2023-09-08 ...
part_number_re_1 = re.compile(r'(\d{7}-\d{2}-\w{1}) (\d+).\d{2}')
# Pattern for part# style 2(have -US): 1636784-00-A-US
part_number_re_2 = re.compile(r'(\d{7}-\d{2}-\w{1})-')
# Pattern for part# style 3 (no rev):1636784-00
part_number_re_3 = re.compile(r'(\d{7}-\d{2})')
# Pattern for part# style 4 (part# and des change position):1636784-00 Glass 25.00 EA 2023-09-08 ...
part_number_re_4 = re.compile(r'(\d{7}-\d{2}-\w{1}) .* (\d+).\d{2} \w+ \d{4}-\d{2}-\d{2}')

# Pattern for quality style 1 (no part# but descri):Glass 25.00 EA 2023-09-08 ...
quality_re_1 = re.compile(r'(.*) (\d+).\d{2} \w+ \d{4}-\d{2}-\d{2}')
# Pattern for quality style 2 (not part#): 25.00 EA 2023-09-08 ...
quality_re_2 = re.compile(r'(\d+).\d{2} \w+ \d{4}-\d{2}-\d{2}')


# Special part pattern
spe_part_re_P10 = re.compile(r'(P10) (\d+).\d{2}')
spe_part_re_IKEA = re.compile(r'(\d{3}.\d{3}.\d{2}) -\w{2} (\d+).\d{2}')
spe_part_re_IKEA_US = re.compile(r'(\d{3}.\d{3}.\d{2})-\w{2} (\d+).\d{2}')

path_file = pd.DataFrame(glob('open_order_pdf/*'))
path_file.columns = ['location']

import os.path
file_folder = {}
for path in path_file['location']:
    file_folder[os.path.getmtime(path)] = path

file_folder = dict(sorted(file_folder.items()))

def extractPartNumber(pdf, file_name, d1, has_part, date_pattern):
    part = ''
    
    for page in pdf.pages:
        text = page.extract_text() 
        text_line = text.split('\n')
        i = 0
        while i < len(text_line):
            line = text_line[i]
            PO_number = PO_number_re.search(line)
            # Fine PO number
            if PO_number:
                p_order = PO_number.group(2)  
                print(p_order)
            # Find page break
            elif date_pattern.search(line):
                i += 1
                continue
            #Special case: P10 1744035-00-A
            elif spe_part_re_P10.search(line):
                items = spe_part_re_P10.search(line)
                part = '1744035-00-A'
                quality = int(float(items.group(2)))
                lines.append(Line(p_order, part, quality, d1))
                print('Part: ', part, '. Qty: ', quality)
                has_part = True
            #Special case: IKEA 190.063.23
            elif spe_part_re_IKEA.search(line):
                items = spe_part_re_IKEA.search(line)
                part = items.group(1)
                quality = int(float(items.group(2)))
                lines.append(Line(p_order, part, quality, d1))
                print('Part: ', part, '. Qty: ', quality)
                has_part = True
            #Special case: IKEA 190.063.23_US
            elif spe_part_re_IKEA_US.search(line):
                items = spe_part_re_IKEA_US.search(line)
                part = items.group(1)
                quality = int(float(items.group(2)))
                lines.append(Line(p_order, part, quality, d1))
                print('Part: ', part, '. Qty: ', quality)
                has_part = True
            # First case: part# and qty on the same line:1636784-00-A 25.00 or part# and des change position         
            elif part_number_re_1.search(line) or part_number_re_4.search(line):
                if part_number_re_1.search(line):
                    items = part_number_re_1.search(line)
                elif part_number_re_4.search(line):
                    items = part_number_re_4.search(line)
                part = items.group(1)
                #Special case: 1895242-00-A_300P
                if (part == "1895242-00-A") and (text_line[i+1].split('.')[0] == '300P'):
                    part = '1895242-00-A_300P'
                quality = int(float(items.group(2)))
                lines.append(Line(p_order, part, quality, d1))
                print('Part: ', part, '. Qty: ', quality)
                has_part = True
            #Second case: part# (-US) and qty are not on the same line
            elif part_number_re_2.search(line):
                part = part_number_re_2.search(line).group(1)
                i += 1
                line = text_line[i]
                items = quality_re_2.search(line)
                quality = int(float(items.group(1)))
                lines.append(Line(p_order, part, quality, d1))
                print('Part: ', part, '. Qty: ', quality)
                has_part = True
            #Third case: (no rev):1636784-00
            elif part_number_re_3.search(line):
                part = part_number_re_3.search(line).group(1)
                check_qty = False

                if quality_re_2.search(line):
                    items = quality_re_2.search(line)
                    check_qty = True
                else:    
                    i += 1
                    line = text_line[i]
                    if quality_re_2.search(line):
                        items = quality_re_2.search(line)
                        check_qty = True

                if check_qty:
                    quality = int(float(items.group(1)))
                    i += 1
                    line = text_line[i]
                    temp = line.split()
                    for c in temp:
                        if '-A' in c or '-B' in c or '-C' in c or '-D' in c or '-E' in c:
                            part = part + '-' + c.split('-')[1]
                            break
                    lines.append(Line(p_order, part, quality, d1))
                    print('Part: ', part, '. Qty: ', quality)
                    has_part = True
            i += 1
    return has_part

def extractDescription(pdf, file_name, d1, has_part, date_pattern):
    start_search = False
    n = 10
    keyword = 'PN: '

    if has_part:
          return  
  
    for page in pdf.pages:
        text = page.extract_text()
        for line in text.split('\n'):
            PO_number = PO_number_re.search(line)
            # Fine PO number
            if PO_number:
                p_order = PO_number.group(2)
            elif start_search:
                if line == 'Number Assembly':
                    continue
                # First case: des and qty on the same line: Glass 25.00 
                if quality_re_1.search(line):
                    quality = quality_re_1.search(line).group(2)
                    # Check if it is the first line: 10 Glass 25.00 
                    if line.startswith(str(n)) or line.startswith(str(n+10)) or line.startswith(str(n+20)):
                        if n > 10:
                            if re.compile(r'(PN:) (\w+)').search(part):
                                before_keyword, keyword, part = part.partition(keyword)
                            lines.append(Line(p_order, part, quality, d1))
                            print('Part: ', part, '. Qty: ', quality)
                        m = len(str(n)) + 1
                        part = quality_re_1.search(line).group(1)[m:]
                        if line.startswith(str(n)):
                            n += 10
                        else:
                            n = int(line.split()[0])
                    else:
                        part = part + ' ' + quality_re_1.search(line).group(1)
                # Second case: only qty: 25.00 EA 2023-09-08 ...
                elif quality_re_2.search(line):
                    quality = quality_re_2.search(line).group(1)
 
                #Special case: Service Kit 3.0
                elif line.startswith('1554480-10-B'):
                    part = line
                    break

                #Third case: first line of parts and no qty: 10 Glass
                elif line.startswith(str(n)) or line.startswith(str(n+10)) or line.startswith(str(n+20)):
                        if n > 10:
                            if re.compile(r'(PN:) (\w+)').search(part):
                                before_keyword, keyword, part = part.partition(keyword)
                            lines.append(Line(p_order, part, quality, d1))
                            print('Part: ', part, '. Qty: ', quality)
                        m = len(str(n)) + 1
                        part = line[m:]
                        if line.startswith(str(n)):
                            n += 10
                        else:
                            n = int(line.split()[0])
                # Find page break
                elif date_pattern.search(line):
                    total = re.compile(r'\d+.\d{2}')
                    if total.search(part.split()[-1]):
                        part = ' '.join(part.split()[:-1])
                else:
                    part = part + ' ' + line
            if line.endswith('(USD)'):
                start_search = True
            if line.startswith('Notes'):
                break
    if re.compile(r'(PN:) (\w+)').search(part):
        before_keyword, keyword, part = part.partition(keyword)
    lines.append(Line(p_order, part, quality, d1))
    print('Part: ', part, '. Qty: ', quality)

def inputPartNumber(file_name):
    today = date.today()
    d1 = today.strftime("%m/%d/%Y")
    has_part = False
    date_pattern = re.compile(r'Page \d+ of \d+')

    with pdfplumber.open(file_name) as pdf:
        pages = pdf.pages
        
        has_part = extractPartNumber(pdf, file_name, d1, has_part, date_pattern)     
        extractDescription(pdf, file_name, d1, has_part, date_pattern)

first = True
for key, value in file_folder.items():
    lines = []
    file = value
    inputPartNumber(file)
    if first:
        df = pd.DataFrame(lines)
        df.to_csv('daily_open_order.csv', index=False)
        first = False
    else:
        df = pd.DataFrame(lines, columns=df.columns)
        df.to_csv('daily_open_order.csv', mode='a', index=False, header=False)

df.info()

while True:
    close = input('Close the program Y/N: ')
    if (close == 'Y' or close =='y'):
        break

import re
from utils import *


def extract_part_number(pdf, lines, d1, has_part, search_file, col, SPECIAL_PATTERN):
    p_order = ''
    part = ''
    search_function = ''
    try:
        for page in pdf.pages:
            text = page.extract_text()
            text_line = text.split('\n')
            i = 0
            while i < len(text_line):
                line = text_line[i]
                PO_number = PO_NUMBER_RE.search(line)
                if PO_number:
                    p_order = PO_number.group(2)
                    print(p_order)
                if PAGE_PATTERN.search(line):
                    i += 1
                    continue
                special_exist, col = export_special_parts(line, lines, p_order, d1, search_function, search_file, col, SPECIAL_PATTERN)
                if special_exist:
                    i += 1
                    has_part = True
                    continue
                if PART_NUMBER_RE_1.search(line) or PART_NUMBER_RE_4.search(line):
                    items = PART_NUMBER_RE_1.search(line) or PART_NUMBER_RE_4.search(line)
                    part = items.group(1)
                    quality = QUANTITY_RE_1.search(line).group(2)
                    print_excel(lines, p_order, part, quality, d1, search_function, search_file, col)
                    col += 1
                    print('Part: ', part, '. Qty: ', quality)
                    has_part = True
                elif PART_NUMBER_RE_2.search(line):
                    part = PART_NUMBER_RE_2.search(line).group(1)
                    i += 1
                    line = text_line[i]
                    items = QUANTITY_RE_3.search(line)
                    try:
                        quality = int(float(items.group(1)))
                    except:
                        continue
                    print_excel(lines, p_order, part, quality, d1, search_function, search_file, col)
                    col += 1
                    print('Part: ', part, '. Qty: ', quality)
                    has_part = True
                elif PART_NUMBER_RE_3.search(line):
                    part = PART_NUMBER_RE_3.search(line).group(1)
                    check_qty = False
                    if QUANTITY_RE_3.search(line):
                        items = QUANTITY_RE_3.search(line)
                        check_qty = True
                    else:
                        i += 1
                        line = text_line[i]
                        if QUANTITY_RE_3.search(line):
                            items = QUANTITY_RE_3.search(line)
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
                        print_excel(lines, p_order, part, quality, d1, search_function, search_file, col)
                        col += 1
                        print('Part: ', part, '. Qty: ', quality)
                        has_part = True
                elif PART_NUMBER_RE_6.search(line):
                    items = PART_NUMBER_RE_6.search(line)
                    part = '-'.join(items.group(1).split())
                    quality = int(float(items.group(2)))
                    print_excel(lines, p_order, part, quality, d1, search_function, search_file, col)
                    col += 1
                    has_part = True
                i += 1
    except:
        print('except error')
    return has_part, col

def extract_description(pdf, lines, d1, has_part):
    start_search = False
    n = 10
    keyword = 'PN: '

    if has_part:
        return  
    try:
        for page in pdf.pages:
            text = page.extract_text()
            for line in text.split('\n'):
                PO_number = PO_NUMBER_RE.search(line)
                # Fine PO number
                if PO_number:
                    p_order = PO_number.group(2)
                elif start_search:
                    if line == 'Number Assembly':
                        continue
                    # First case: des and qty on the same line: Glass 25.00 
                    if QUANTITY_RE_1.search(line):
                        quality = QUANTITY_RE_1.search(line).group(2)
                        # Check if it is the first line: 10 Glass 25.00 
                        if line.startswith(str(n)) or line.startswith(str(n+10)) or line.startswith(str(n+20)):
                            if n > 10:
                                if re.compile(r'(PN:) (\w+)').search(part):
                                    before_keyword, keyword, part = part.partition(keyword)
                                part = part.replace(" ", "")
                                lines.append(Line(p_order, part, quality, d1, ''))
                                print('Part: ', part, '. Qty: ', quality)
                            m = len(str(n)) + 1
                            part = QUANTITY_RE_1.search(line).group(1)[m:]
                            if line.startswith(str(n)):
                                n += 10
                            else:
                                n = int(line.split()[0])
                        else:
                            part = part + ' ' + QUANTITY_RE_1.search(line).group(1)
                    # Second case: only qty: 25.00 EA 2023-09-08 ...
                    elif QUANTITY_RE_3.search(line):
                        quality = QUANTITY_RE_3.search(line).group(1)

                    #Special case: Service Kit 3.0
                    elif line.startswith('1554480-10-B'):
                        part = line
                        break

                    #Third case: first line of parts and no qty: 10 Glass
                    elif line.startswith(str(n)) or line.startswith(str(n+10)) or line.startswith(str(n+20)):
                            if n > 10:
                                if re.compile(r'(PN:) (\w+)').search(part):
                                    before_keyword, keyword, part = part.partition(keyword)
                                part = part.replace(" ", "")
                                lines.append(Line(p_order, part, quality, d1, ''))
                                print('Part: ', part, '. Qty: ', quality)
                            m = len(str(n)) + 1
                            part = line[m:]
                            if line.startswith(str(n)):
                                n += 10
                            else:
                                n = int(line.split()[0])
                    # Find page break
                    elif PAGE_PATTERN.search(line):
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
        part = part.replace(" ", "")
        lines.append(Line(p_order, part, quality, d1, ''))
        print('Part: ', part, '. Qty: ', quality)
    except:
        print('except error')

def export_special_part(items, part, lines, p_order, d1, search_function, search_file, col):
    quality = int(float(items.group(2)))
    print_excel(lines, p_order, part, quality, d1, search_function, search_file, col)
    print('Part: ', part, '. Qty: ', quality)

def export_special_parts(line, lines, p_order, d1, search_function, search_file, col, SPECIAL_PATTERN):
    
    for pattern, partName in SPECIAL_PATTERN:
        if pattern.search(line):
            items = pattern.search(line)
            if partName == '':
                partName = items.group(1)
            export_special_part(items, partName, lines, p_order, d1, search_function, search_file, col)
            col += 1
            return True, col
    return False, col
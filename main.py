import logging
from order_processor import OrderProcessor

def main():
    logging.getLogger("pdfminer").setLevel(logging.ERROR)
    processor = OrderProcessor(pdf_folder='open_order_pdf')
    processor.run()

    while True:
        close = input('Close the program Y/N: ')
        if (close == 'Y' or close =='y'):
            break

if __name__ == '__main__':
    main()
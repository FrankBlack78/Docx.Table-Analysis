from os import system, name
import sys
from docx import Document
import pandas as pd


# Functions
def clear_screen():
    if name == 'nt':
        _ = system('cls')
    else:
        _ = system('clear')

def logo():
    print('+-+ +-+ +-+ +-+ +-+   +-+ +-+ +-+ +-+ +-+ +-+')
    print('|B| |L| |A| |C| |K|   |C| |o| |d| |i| |n| |g|')
    print('+-+ +-+ +-+ +-+ +-+   +-+ +-+ +-+ +-+ +-+ +-+')
    print('\n' + '-' * 45)
    print('Program: Docx.Table Analysis')


# Main
while True:
    clear_screen()
    logo()
    print('\n' + '-' * 45)
    print('Commands:')
    print('[1] Read Docx.Table')
    # print('[2] test')
    # print('[3] test')
    # print('[4] test')
    # print('[5] test')
    # print('\n')
    print('[9] Quit program')
    print('-' * 45)
    input_ = input('>>> ')

    # Quit program
    if input_ == '9':
        clear_screen()
        sys.exit(1)

    # Read Docx.Table
    if input_ == '1':

        while True:
            clear_screen()
            logo()
            print('Menue: Read Docx.Table')
            print('-' * 45)
            print('-No menue-')
            print('-' * 45)
            file_ = input('Enter file-name: >>> ')

            # Test-Document
            # doc = Document('Test.docx')

            try:
                doc = Document(file_)
            except:
                print('File not found. Script terminated.')
                sys.exit(1)

            while True:
                try:
                    tables = doc.tables
                except:
                    print('No tables in this Docx.File. Script terminated.')
                    sys.exit(1)

                clear_screen()
                logo()
                print('Menue: Read Docx.Table')
                print(len(tables), ' tables(s) detected.')
                print('-' * 45)
                print('-No menue-')
                print('-' * 45)
                table_ = int(input('Enter table-number: >>> '))-1

                result = []

                try:
                    for row in tables[table_].rows:
                        interim = []
                        result.append(interim)
                        for cell in row.cells:
                            interim.append(cell.text)
                except:
                    print('Table not found. Script terminated.')
                    sys.exit(1)

                clear_screen()
                logo()
                print('Menue: Read Docx.Table')
                print(len(tables), ' tables(s) detected.')
                print('Table ', table_ + 1, ' selected.')
                print('-' * 45)
                print('-No menue-')
                print('-' * 45)

                print('\n>>> Printing results <<<')
                # print(result)
                labels = result[0]
                df = pd.DataFrame.from_records(result[1:], columns=labels)

                print(df)
                sys.exit(1)

from datetime import datetime
import os
import re
from typing import Union

import pandas as pd

from biedronka import ExcelDataDumper


class FileReader:
    def __init__(self, path_to_folder: str):
        self.folder_path = path_to_folder
        self.file_path = None
        self.data = None

    def get_newest_file(self, name_start: str, file_format: str) -> None:
        files = [file for file in os.listdir(self.folder_path)
                 if os.path.isfile(os.path.join(self.folder_path, file))
                 and file.startswith(name_start)
                 and file.endswith(file_format)]
        self.file_path = os.path.join(self.folder_path,
                                      max(files, key=lambda file: os.path.getctime(
                                          os.path.join(self.folder_path, file))))

    def read_csv(self) -> pd.DataFrame:
        if not self.file_path:
            raise ValueError("File path is not set. Call get_newest_file() first.")
        self.data = pd.read_csv(self.file_path, delimiter=';', skiprows=20, skipfooter=1,
                                usecols=[0, 2, 3, 8, 9, 15], engine='python',
                                on_bad_lines='skip', encoding="windows-1250")
        return self.data

    def process_transactions(self) -> pd.DataFrame:
        self.get_newest_file('Lista_transakcji_nr', 'csv')
        return self.read_csv()


class DataCleaner:
    def __init__(self):
        self.cleaned_data = None

    def clean_data(self, data: pd.DataFrame) -> None:
        kwota = 'Kwota transakcji (waluta rachunku)'
        data_new = data[data['Waluta'] == 'PLN'] \
            .dropna(axis='rows', how='any') \
            .map(lambda x: x.strip() if isinstance(x, str) else x)

        data_new['Wydatek'] = data_new[kwota] \
            .apply(lambda x: x.startswith('-'))

        data_new[kwota] = data_new[kwota] \
            .astype(str) \
            .apply(lambda x: x.lstrip('-')) \
            .str.replace(',', '.') \
            .astype(float) \
            .round(2)

        data_new['Data transakcji'] = pd.to_datetime(data_new['Data transakcji'],
                                                     format="%Y-%m-%d")
        self.cleaned_data = data_new

    def sub_blik_payment_titles(self, row: pd.Series) -> pd.Series:
        pattern = r'\W?Przelew na telefon \+\d{2}x{6}\d{3} (.+) Dla (?:[\w| ]+) Od (?:[\w| ]+)\W?'
        row['Tytuł'] = re.sub(pattern, r'\1' + ' ' + row['Dane kontrahenta'], row['Tytuł'])
        row['Tytuł'] = row['Tytuł'] + ' ' + row['Dane kontrahenta'] \
            if row['Tytuł'] == 'Przelew na telefon BLIK' else row['Tytuł']

        pattern_platnosc_blik = (
            r'Płatność BLIK \d{2}\.\d{2}\.\d{4} Nr transakcji \d{11} '
            r'(?:https?:\/\/)?(?:www\.)?([\w]+)\.\w{2,3}\/?'
        )
        row['Tytuł'] = re.sub(pattern_platnosc_blik, r'\1', row['Tytuł'])
        return row

    def sub_card_payment_titles(self, row: pd.Series) -> pd.Series:
        if self.cleaned_data is None:
            raise ValueError("Data has not been cleaned. Call clean_data() first.")
        mapping = {
            'WWW.BILET.INTERCITY.PL  WARSZAWA  P': 'Bilet IC',
            'www.bilet.intercity.pl    Warszawa': 'Bilet IC',
            'ZABKA': 'Żabka',
            'JMP S.A. BIEDRONKA': 'Biedronka',
            'OLX_': 'OLX',
        }
        matching_key = next((key for key in mapping.keys()
                             if row['Dane kontrahenta'].startswith(key)), None)
        card_payment = row['Tytuł'].startswith('Płatność kartą')
        # card_return = row['Tytuł'].startswith('Zwrot płatności')
        if matching_key:
            row['Tytuł'] = mapping[matching_key]
        elif card_payment:
            row['Tytuł'] = row['Dane kontrahenta']
        return row

    def process_data(self, data: pd.DataFrame) -> pd.DataFrame:
        self.clean_data(data)
        self.cleaned_data = self.cleaned_data \
                                .apply(self.sub_blik_payment_titles, axis=1) \
                                .apply(self.sub_card_payment_titles, axis=1) \
                                .drop(columns=['Dane kontrahenta', 'Waluta']) \
                                .rename(columns={'Kwota transakcji (waluta rachunku)': 'Kwota'}) \
                                .iloc[::-1] \
            .reset_index(drop=True)
        self.cleaned_data['Kategoria'] = self.payment_category()
        return self.cleaned_data[['Data transakcji', 'Tytuł', 'Kwota',
                                  'Kategoria', 'Saldo po transakcji', 'Wydatek']]

    def payment_category(self) -> pd.Series:
        category_mapping = {
            'Bilet IC': 'transport',
            'Pieczywo': 'spożywcze',
        }
        return self.cleaned_data['Tytuł'] \
            .map(category_mapping) \
            .fillna('')


class ExcelDataInserter(ExcelDataDumper):
    def __init__(self, excel_path: str):
        super().__init__(excel_path)
        self.latest_date = None
        self.ws = None
        self.categories = ['transport', 'mieszkanie', 'spożywcze', 'media', 'rozrywka',
                           'prezenty', 'ubrania', 'wyjazdy', 'domowe', 'inne']

    def get_latest_transaction_date(self) -> datetime:
        return super().latest_grocery_date()

    def rename_create_spreadsheets(self, names: list[str]) -> None:
        return super().handle_sheets(names)

    def month_mapping(self, month: int) -> str:
        months = {
            1: 'styczeń',
            2: 'luty',
            3: 'marzec',
            4: 'kwiecień',
            5: 'maj',
            6: 'czerwiec',
            7: 'lipiec',
            8: 'sierpień',
            9: 'wrzesień',
            10: 'październik',
            11: 'listopad',
            12: 'grudzień'
        }
        return months[month]

    def filter_data_on_date(self, data: pd.DataFrame) -> pd.DataFrame:
        return data[data['Data transakcji'] > self.latest_date]

    def last_data_row(self, column: str = 'A') -> int:
        return max((cell.row for cell in self.ws[column]
                    if cell.value is not None),
                   default=1)

    def set_currency_format(self, rows: Union[list, int], columns: Union[list, int]) -> None:
        if isinstance(rows, int) and isinstance(columns, int):
            self.ws.cell(row=rows, column=columns).number_format = \
                ExcelDataDumper.CURRENCY_FORMAT
        elif isinstance(rows, int) and isinstance(columns, list):
            for col in columns:
                self.ws.cell(row=rows,
                             column=col).number_format = ExcelDataDumper.CURRENCY_FORMAT
        elif isinstance(rows, list) and isinstance(columns, int):
            for row in rows:
                self.ws.cell(row=row,
                             column=columns).number_format = ExcelDataDumper.CURRENCY_FORMAT
        elif isinstance(rows, int) and isinstance(columns, int):
            for row in rows:
                for col in columns:
                    self.ws.cell(row=row,
                                 column=col).number_format = ExcelDataDumper.CURRENCY_FORMAT

    def fill_spreadsheet(self, data: pd.DataFrame):
        months_to_fill = sorted(
            list(data['Data transakcji'].dt.month.unique()),
            key=lambda x: (x <= 9, x))
        months_str = [self.month_mapping(month) for month in months_to_fill]
        self.rename_create_spreadsheets(months_str)

        for month, month_num in zip(months_str, months_to_fill):
            self.ws = self.wb[month]
            last_row = self.last_data_row()
            m_data = data[data['Data transakcji'].dt.month == month_num].copy()
            m_data['Data transakcji'] = m_data['Data transakcji'] \
                .dt.strftime('%d.%m.%Y')

            headers = ['Data', 'Opis', 'Kwota', 'Kategoria']
            self.insert_header(last_row, headers)
            if last_row == 1:
                for col, header in enumerate(headers, start=1):
                    self.ws.cell(row=1, column=col, value=header)
                last_row += 1

            m_data_expenses = m_data[m_data['Wydatek'] == True].iloc[:, :-2]
            m_data_incomes = m_data[m_data['Wydatek'] == False].iloc[:, 1:-3]
            m_data_incomes = m_data_incomes[['Kwota', 'Tytuł']]

            self.insert_expenses(m_data_expenses, last_row)

            # entry incomes
            column_n = 7  # G
            column_letter = chr(65 + column_n - 1)
            last_row_incomes = self.insert_incomes(m_data_incomes, column_n)

            sum_incomes_row = last_row_incomes + 1
            self.insert_sum_incomes(sum_incomes_row, column_n, column_letter, last_row_incomes)

            # entry categories and its formulas
            self.empty_cells([sum_incomes_row + 1, sum_incomes_row + 2, sum_incomes_row + 3], [column_n, column_n + 1])

            start_row_expenses = sum_incomes_row + 3
            last_row_expenses = self.insert_categorized_expenses(start_row_expenses, column_n)

            sum_expenses_row = last_row_expenses + 1
            self.insert_sum_expenses(sum_expenses_row, column_n, column_letter,
                                     start_row_expenses, last_row_expenses)

            self.insert_balance(last_row_expenses, column_n, sum_incomes_row, sum_expenses_row, column_letter)

        self.save_excel_workbook()

    def insert_expenses(self, data_expenses: pd.DataFrame, last_row: int):
        for r_idx, row in enumerate(data_expenses.itertuples(index=False),
                                    start=last_row):
            for c_idx, value in enumerate(row, start=1):
                self.ws.cell(row=r_idx, column=c_idx, value=value)
                if c_idx == 3:
                    self.set_currency_format(r_idx, 3)

    def insert_incomes(self, data_incomes: pd.DataFrame, column_number: int):
        self.ws.cell(row=1, column=column_number, value='wpływy')
        row_ix = 2
        for row_ix, row in enumerate(data_incomes.itertuples(index=False), start=2):
            for col_ix, value in enumerate(row, start=column_number):
                self.ws.cell(row=row_ix, column=col_ix, value=value)
                if col_ix == column_number:
                    self.set_currency_format(row_ix, column_number)
        return row_ix

    def insert_header(self, last_row: int, headers: list):
        if last_row == 1:
            for col, header in enumerate(headers, start=1):
                self.ws.cell(row=1, column=col, value=header)
            last_row += 1

    def empty_cells(self, rows: list, columns: list):
        for row in rows:
            for col in columns:
                self.ws.cell(row=row, column=col, value='')

    def insert_categorized_expenses(self, start_row: int, column: int):
        column_letter2 = chr(65 + column)
        self.ws.cell(row=start_row, column=column, value='wydatki wg kategorii')

        row_n = start_row + 1
        for row_n, category in enumerate(self.categories, start=start_row + 1):
            self.ws.cell(row=row_n, column=column + 1, value=category)
            self.ws.cell(row=row_n, column=column,
                         value=f'=SUMIF($D$2:$D$250,{column_letter2}{row_n},$C$2:$C$250)')
            self.set_currency_format(row_n, column)
        return row_n

    def insert_sum_expenses(self, sum_expenses_row: int, column_n: int, column_letter: str,
                            start_row_expenses: int, last_row_expenses: int):
        self.ws.cell(row=sum_expenses_row, column=column_n + 1, value='suma')
        self.ws.cell(row=sum_expenses_row, column=column_n,
                     value=f'=SUM({column_letter}{start_row_expenses + 1}:{column_letter}{last_row_expenses})')
        self.set_currency_format(sum_expenses_row, column_n)

    def insert_balance(self, last_row_expenses: int, column: int,
                       sum_incomes_row: int, sum_expenses_row: int,
                       column_letter: str):
        self.ws.cell(row=last_row_expenses + 3, column=column, value='bilans')
        self.ws.cell(row=last_row_expenses + 4, column=column,
                     value=f'={column_letter}{sum_incomes_row}-{column_letter}{sum_expenses_row}')
        self.set_currency_format(last_row_expenses + 4, column)

    def insert_sum_incomes(self, sum_incomes_row: int, column_n: int, column_letter: str, last_row_incomes: int):
        self.ws.cell(row=sum_incomes_row, column=column_n + 1, value='suma')
        self.ws.cell(row=sum_incomes_row, column=column_n,
                     value=f'=SUM({column_letter}2:{column_letter}{last_row_incomes})')
        self.set_currency_format(sum_incomes_row, column_n)
    def insert_data_to_excel(self, data: pd.DataFrame):
        self.latest_date = self.get_latest_transaction_date()
        filtered_data = self.filter_data_on_date(data)
        self.fill_spreadsheet(filtered_data)


folder_path = r'C:\Users\xx\xx'
path_excel = r'C:\Users\xx\xx.xlsx'

file_reader = FileReader(folder_path)
raw_data = file_reader.process_transactions()

data_cleaner = DataCleaner()
cleaned_data = data_cleaner.process_data(raw_data)

excel_inserter = ExcelDataInserter(path_excel)
excel_inserter.insert_data_to_excel(cleaned_data)

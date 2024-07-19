import openpyxl
import pandas as pd

from common import debug_print


class DataHandler:
    def __init__(self):
        self.data = None
        self.column_info = {}

    def load_data(self, file_path):

        wb = openpyxl.load_workbook(file_path)
        sheet_names = wb.get_sheet_names()
        name = sheet_names[0]
        sheet_ranges = wb[name]
        self.data = pd.DataFrame(sheet_ranges.values)
        self.data.fillna(0, inplace=True)

        new_header = self.data.iloc[0]  # 첫 번째 row를 헤더로 사용
        self.data = self.data[1:]  # 첫 번째 row를 데이터로 사용하므로 제거
        self.data.columns = new_header  # 헤더 설정
        self.create_column_info() # 컬럼 정보를 저장

    def create_column_info(self):
        # 첫번째 row 값을 조사하여 컬럼 정보를 저장한다

        # 알파벳 순서대로 맵핑
        for i, value in enumerate(self.data.columns):
            # 알파벳 맵핑 (A부터 시작, ASCII 코드 기준)
            alphabet = chr(65 + i)

            is_formula = False
            cell_data = self.data.at[1, value]
            if str(cell_data).startswith('='):
                is_formula = True

            self.add_data(i, alphabet, value, is_formula, cell_data)

        # 결과 출력
        debug_print(self.column_info)

    def add_data(self, index, alphabet, name, is_formula, formula):
        self.column_info[index] = {
            'alphabet': alphabet,
            'name': name,
            'is_formula': is_formula,
            'formula': formula
        }

    def get_data(self):
        return self.data

    def get_column_info(self):
        return self.column_info

    def save_data(self, file_path):
        if self.data is not None:
            debug_print("=========================== save data ==========================")
            debug_print(self.data)
            debug_print("=========================== save data ==========================")

            writer = pd.ExcelWriter(file_path, engine='openpyxl')
            self.data.to_excel(writer, index=False, sheet_name='Sheet1')
            writer.close()


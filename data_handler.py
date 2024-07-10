import pandas as pd
import openpyxl

class DataHandler:
    def __init__(self):
        self.data = None
        self.formula_columns = []
        self.column_info = {}

    def load_data(self, file_path):
        self.data = pd.read_excel(file_path, header=0, na_values=['NA', '?'], keep_default_na=True)
        print(self.data)
        self.formula_columns = self._detect_formula_columns(file_path)
        self.data.fillna(0, inplace=True)
        self._calculate_formula_columns()
        print('load_data finished!')

    def load_data2(self, file_path):
        # self.data = pd.read_excel(file_path, header=0, keep_default_na=False, engine='openpyxl', dtype=str)
        #print(self.data)
        # print(self.data.dtypes)
        wb = openpyxl.load_workbook(file_path)
        sheet_names = wb.get_sheet_names()
        name = sheet_names[0]
        sheet_ranges = wb[name]
        self.data = pd.DataFrame(sheet_ranges.values)

        new_header = self.data.iloc[0]  # 첫 번째 row를 헤더로 사용
        self.data = self.data[1:]  # 첫 번째 row를 데이터로 사용하므로 제거
        self.data.columns = new_header  # 헤더 설정
        #print(self.data.columns)
        print(self.data)

        self.create_column_info()
        # print(self.data)
        # print(self.data.dtypes)
        print('load_data 2 finished!')




        # self.data['Total'] = self.data.apply(self.calculate_total, axis=1)

        print(self.data)
        print('calculate_data finished!')



    def create_column_info(self):
        # 첫번째 row 값을 조사하여 컬럼 정보를 저장한다

        # 알파벳 순서대로 맵핑
        for i, value in enumerate(self.data.columns):
            print(f"i: {i}, value: {value}")
            # 알파벳 맵핑 (A부터 시작, ASCII 코드 기준)
            alphabet = chr(65 + i)  # 65는 'A'의 ASCII 코드 값

            is_formula = False
            cell_data = self.data.at[1, value]
            print(f"cell_data: {cell_data}")
            if str(cell_data).startswith('='):
                is_formula = True

            self.add_data(i, alphabet, value, is_formula, cell_data)

        # 결과 출력
        print(self.column_info)

    def add_data(self, index, alphabet, name, is_formula, formula):
        self.column_info[index] = {
            'alphabet': alphabet,
            'name': name,
            'is_formula': is_formula,
            'formula': formula
        }

    def calculate_total(self, row):
        print(f"row : {row}")
        formula = self.data.at[row, 'Total'].replace('=', '')  # Remove '=' sign
        print(f"formula : {formula}")
        variables = formula.split('+')  # Split by '+'
        print(f"variables : {variables}")

        total = 0
        for var in variables:
            var = var.strip()  # Remove leading/trailing whitespace
            if '*' in var:
                c1, c2 = var.split('*')
                print(c1)
                print(c2)
                col1 = self.column_info[c1(0)]
                col2 = self.column_info[c2(0)]
                v1 = self.data(c1[1]-1, col1)
                v2 = self.data(c2[1]-1, col2)
                print(v1)
                print(v2)
                total += row.iloc[col1] * row.iloc[col2]
            else:
                col_index = ord(var[0]) - ord('D')  # Convert 'D' to column index
                total += row.iloc[col_index]

        return total


    def _detect_formula_columns(self, file_path):
        formula_columns = []
        wb = openpyxl.load_workbook(file_path, data_only=False, read_only=True)
        sheet = wb.active

        results = {}

        # 첫 번째 row에서 column 명을 가져오기
        columns = [cell.value for cell in sheet[1]]

        # 각 컬럼의 두 번째 row 값을 읽어 저장
        for i, column in enumerate(columns, start=1):
            cell = sheet.cell(row=2, column=i)
            if cell.value is not None:
                if cell.data_type == 'f':
                    formula_columns.append(column)
                    results[column] = 'Formula'
                else:
                    results[column] = 'Value'
            else:
                results[column] = 'Empty'

        # 결과 출력
        for column, cell_type in results.items():
            print(f"Column: {column}, Type: {cell_type}")

        print(formula_columns)
        return formula_columns

    def _calculate_formula_columns(self):
        for col in self.formula_columns:
            try:
                self.data[col] = self.data.apply(lambda row: self._evaluate_formula(row, col), axis=1)
            except Exception as e:
                print(f"Error calculating formula for column {col}: {e}")

    def _evaluate_formula(self, row, col):
        formula = row[col]
        if not isinstance(formula, str) or not formula.startswith('='):
            return formula  # 수식이 아니거나 올바르지 않은 경우 값을 그대로 반환

        formula = formula[1:]  # '=' 제거

        try:
            # 수식 내의 셀 참조를 실제 값으로 대체
            formula_with_values = self._replace_with_values(row, formula)
            return eval(formula_with_values)
        except Exception as e:
            print(f"Error evaluating formula '{formula}' for column {col}: {e}")
            return row[col]  # 수식 평가 중 오류 발생 시 원래 값을 반환

    def _replace_with_values(self, row, formula):
        import re
        cell_references = re.findall(r'[A-Z]+[0-9]+', formula)  # 셀 참조 찾기 (예: 'C2', 'D2', 등)
        print(cell_references)
        for cell in cell_references:
            col_letter = ''.join(filter(str.isalpha, cell))  # 컬럼 문자 추출 (예: 'C')
            col_index = openpyxl.utils.cell.column_index_from_string(col_letter) - 1  # 컬럼 인덱스 추출
            cell_value = row[col_index]  # 현재 행에서 해당 셀 값 가져오기

            formula = formula.replace(cell, str(cell_value))  # 셀 참조를 값으로 대체

        return formula

    def get_data(self):
        return self.data

    def get_column_info(self):
        return self.column_info

    def save_data(self, file_path):
        if self.data is not None:
            print("save_data in!")
            print(self.data)

            # pandas 사용해서 일단 저장
            writer = pd.ExcelWriter(file_path, engine='openpyxl')
            self.data.to_excel(writer, index=False, sheet_name='Sheet1')
            writer.close()

            # openpyxl을 사용하여 수식을 업데이트
            wb = openpyxl.load_workbook(file_path)
            sheet = wb['Sheet1']

            for col in self.formula_columns:
                print(f"col: {col}")
                for row_idx, value in enumerate(self.data[col], start=2):
                    cell_reference = f"{col}{row_idx}"
                    print(f"cell_reference: {cell_reference}")
                    original_formula = value
                    print(f"original_formula: {original_formula}")
                    if isinstance(original_formula, str) and original_formula.startswith('='):
                        sheet[cell_reference].value = original_formula

            wb.save(file_path)

            # # 수식을 원래 상태로 되돌려 저장
            # for col in self.formula_columns:
            #     print(f"col: {col}")
            #     for row in range(1, len(self.data) + 1):
            #         print(f"row: {row}")
            #         cell_reference = f"{col}{row + 1}"
            #         print(f"cell_reference: {cell_reference}")
            #         original_formula = self.data[col][row - 1]
            #         print(f"original_formula: {original_formula}")
            #         if isinstance(original_formula, str) and original_formula.startswith('='):
            #             worksheet[cell_reference].value = original_formula
            #
            # writer.save()

    def get_chart_data(self, x_column, y_column):
        if self.data is not None:
            return self.data[[x_column, y_column, 'Key', 'Summary']]
        return None
import pandas as pd

class DataHandler:
    def __init__(self):
        self.data = None

    def load_data(self, file_path):
        self.data = pd.read_excel(file_path, header=0, dtype={
            'Key': str,
            'Summary': str
        }, na_values=['NA', '?'], keep_default_na=True)
        self.data.fillna(0, inplace=True)
        # print("--- load_data")
        # print(self.data.dtypes)
        # print("load_data ---")
        # print(self.data)

    def get_data(self):
        return self.data

    def save_data(self, file_path):
        if self.data is not None:
            self.data.to_excel(file_path, index=False)

    def get_chart_data(self, x_column, y_column):
        if self.data is not None:
            # print("data_handler.get_chart_data")
            # print("x_column: " + x_column)
            # print("y_column: " + y_column)
            # print(self.data[[x_column, y_column, 'Key', 'Summary']])
            return self.data[[x_column, y_column, 'Key', 'Summary']]
        return None
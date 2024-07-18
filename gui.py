import random

from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QHBoxLayout, QPushButton, \
    QTableWidget, QTableWidgetItem, QFileDialog, QComboBox, QSplitter, QMessageBox, QHeaderView, QAbstractItemView
from PyQt5.QtCore import Qt, pyqtSlot
from PyQt5.QtGui import QFont, QColor
from chart import ChartCanvas
from data_handler import DataHandler
import re
import pandas as pd


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Quadrant Chart Tool')

        self.resize(1920, 1080)  # 윈도우 기본 크기 설정

        self.x_column = ''
        self.y_column = ''
        self.x_column_index = -1
        self.y_column_index = -1
        self.chart_data = None

        self.data_handler = DataHandler()
        self.column_info = self.data_handler.get_column_info()
        self.colors = []
        self.previous_selected_row = -1  # 이전 선택된 행의 인덱스를 추적하는 변수

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QHBoxLayout(self.central_widget)

        self.left_layout = QVBoxLayout()
        self.right_layout = QVBoxLayout()

        # 버튼 레이아웃
        self.button_layout = QHBoxLayout()

        self.load_button = QPushButton('Load Data')
        self.load_button.clicked.connect(self.load_data)
        self.button_layout.addWidget(self.load_button)

        self.x_combo_box = QComboBox()
        self.x_combo_box.setPlaceholderText('Select X axis')
        self.button_layout.addWidget(self.x_combo_box)

        self.y_combo_box = QComboBox()
        self.y_combo_box.setPlaceholderText('Select Y axis')
        self.button_layout.addWidget(self.y_combo_box)

        self.swap_axes_button = QPushButton('Swap Axes')
        self.swap_axes_button.clicked.connect(self.swap_axes)
        self.button_layout.addWidget(self.swap_axes_button)

        self.plot_button = QPushButton('Plot Chart')
        self.plot_button.clicked.connect(self.plot_chart)
        self.button_layout.addWidget(self.plot_button)

        self.save_button = QPushButton('Save Changes')
        self.save_button.clicked.connect(self.save_changes)
        self.button_layout.addWidget(self.save_button)

        # 버튼 레이아웃을 왼쪽 레이아웃의 상단에 추가
        self.left_layout.addLayout(self.button_layout)

        # 데이터 테이블 레이아웃
        self.table_widget = QTableWidget()

        # 테이블 레이아웃을 왼쪽 레이아웃의 하단에 추가
        self.left_layout.addWidget(self.table_widget)

        # 차트 정보 레이아웃
        self.chartview_layout = QHBoxLayout()
        self.selected_point_label = QLabel('Selected Point: None')
        self.chartview_layout.addWidget(self.selected_point_label, 3)
        h_widget = QWidget()
        h_widget.setLayout(self.chartview_layout)
        h_widget.setFixedHeight(100)

        # 차트 정보 레이아웃을 오른쪽 레이아웃의 상단에 추가
        self.right_layout.addWidget(h_widget)

        # 차트 캔버스
        self.chart_canvas = ChartCanvas(self)
        self.chart_canvas.point_clicked.connect(self.highlight_selected_row)  # Signal 연결
        self.chart_canvas.point_selected.connect(self.display_selected_point)  # Signal 연결
        self.chart_canvas.point_dropped.connect(self.handle_point_drop)  # Signal 연결

        # 차트 캔버스를 오른쪽 레이아웃의 하단에 추가
        self.right_layout.addWidget(self.chart_canvas, 8)

        # 차트 정보 레이아웃에 Reverse Chart 버튼 추가
        # self.reverse_chart_button = QPushButton('Reverse Chart')
        # self.reverse_chart_button.clicked.connect(self.chart_canvas.reverse_chart)
        # self.chartview_layout.addWidget(self.reverse_chart_button)
        self.reverse_x_axis_button = QPushButton('Reverse X Axis')
        self.reverse_x_axis_button.clicked.connect(self.chart_canvas.reverse_x_axis)
        self.chartview_layout.addWidget(self.reverse_x_axis_button, 1)

        self.reverse_y_axis_button = QPushButton('Reverse Y Axis')
        self.reverse_y_axis_button.clicked.connect(self.chart_canvas.reverse_y_axis)
        self.chartview_layout.addWidget(self.reverse_y_axis_button, 1)

        # QSplitter를 사용하여 레이아웃 나누기
        self.splitter = QSplitter(Qt.Horizontal)
        left_widget = QWidget()
        left_widget.setLayout(self.left_layout)
        self.splitter.addWidget(left_widget)

        right_widget = QWidget()
        right_widget.setLayout(self.right_layout)
        self.splitter.addWidget(right_widget)

        self.splitter.setSizes([self.size().width() // 2, self.size().width() // 2])
        self.main_layout.addWidget(self.splitter)

        # 테이블 값 변경 시 차트 갱신
        self.table_widget.itemChanged.connect(self.on_item_changed)

    # def resizeEvent(self, event):
    #     super().resizeEvent(event)
    #     # 윈도우가 리사이즈 될 때 차트 캔버스를 정사각형으로 유지
    #     self.chart_canvas.setFixedSize(self.splitter.sizes()[1], self.splitter.sizes()[1])

    def load_data(self):

        file_path, _ = QFileDialog.getOpenFileName(self, 'Open File', '', 'Excel Files (*.xlsx)')
        if file_path:
            self.table_widget.itemChanged.disconnect()  # 신호 연결 해제
            self.data_handler.load_data(file_path)
            self.display_data()
            self.populate_combo_boxes()
            self.generate_colors()
            self.chart_canvas.initialize(is_swap=False)
            self.table_widget.itemChanged.connect(self.on_item_changed)  # 신호 다시 연결

    def plot_chart(self):

        # 현재 테이블에 노출되는 데이터를 업데이트하여 저장한다
        if self.get_table_data() is not None:
            self.table_widget.itemChanged.disconnect()  # 신호 연결 해제
            self.update_data_from_gui()
            self.reset_table_style()
            self.table_widget.itemChanged.connect(self.on_item_changed)  # 신호 다시 연결

        # X, Y축 선택 여부를 체크하고 축 정보를 저장한다
        if self.check_axes_selection() is False:
            return

        # GUI 테이블로부터 데이터 가져와서 차트 그리기
        chart_data = self.get_table_data()[[self.x_column, self.y_column, 'Key', 'Summary']]
        if chart_data is not None:
            print("================== chart data =======================")
            print(chart_data)
            print("================== chart data =======================")
            self.chart_canvas.plot(chart_data, self.x_column, self.y_column, self.colors)

    def save_changes(self):
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save File', '', 'Excel Files (*.xlsx)')
        if file_path:
            self.data_handler.save_data(file_path)

    def populate_combo_boxes(self):
        data = self.data_handler.get_data()
        columns = data.columns

        exclude_texts = ['Key', 'Summary']  # 2개의 컬럼은 고정이라고 가정

        filtered_columns = [col for col in columns if not any(text in col for text in exclude_texts)]

        self.x_combo_box.clear()
        self.y_combo_box.clear()
        self.x_combo_box.addItems(filtered_columns)
        self.y_combo_box.addItems(filtered_columns)

    def swap_axes(self):
        current_x = self.x_combo_box.currentText()
        current_y = self.y_combo_box.currentText()

        self.x_combo_box.setCurrentText(current_y)
        self.y_combo_box.setCurrentText(current_x)

        self.chart_canvas.initialize(is_swap=True)

    def check_axes_selection(self):
        self.x_column = self.x_combo_box.currentText()
        self.y_column = self.y_combo_box.currentText()

        # X, Y축 선택 여부를 체크한다
        if self.x_column == "" or self.y_column == "":
            QMessageBox.critical(self, 'Error', 'X and Y axes must be selected.')
            return False
        # X, Y축 값 다름을 체크한다
        if self.x_column == self.y_column:
            QMessageBox.critical(self, 'Error', 'X and Y axes must be different.')
            self.y_combo_box.setCurrentIndex(-1)
            return False

        # 선택된 X, Y축의 인덱스 값을 저장한다
        header = self.table_widget.horizontalHeader()
        model = header.model()
        column_count = header.count()

        self.x_column_index = -1
        self.y_column_index = -1

        for i in range(column_count):
            text = model.headerData(i, Qt.Horizontal)
            if text == self.x_column:
                self.x_column_index = i

            if text == self.y_column:
                self.y_column_index = i

            if self.x_column_index != -1 and self.y_column_index != -1:
                break

        print(f"check_axes_selection -> x_column_index: {self.x_column_index}, y_column_index: {self.y_column_index}")
        return True

    def display_data(self):
        data = self.data_handler.get_data()
        self.table_widget.setRowCount(data.shape[0])
        self.table_widget.setColumnCount(data.shape[1])
        self.table_widget.setHorizontalHeaderLabels(data.columns)

        # 추가 계산이 필요한 컬럼을 먼저 조사
        column_info = self.data_handler.get_column_info()
        true_formula_indices = [index for index, details in column_info.items() if details['is_formula']]

        # 결과 출력
        print(true_formula_indices)

        # 이 부분에 수식을 계산해서 숫자로 변환하는 코드 추가
        for i in range(data.shape[0]):
            for j in range(data.shape[1]):
                print(f"i: {i}, j: {j}")
                if j in true_formula_indices:
                    print("to calculate for j")
                    calculated_value = self.calculate_formula(i, j)
                    self.table_widget.setItem(i, j, QTableWidgetItem(str(calculated_value)))
                else:
                    self.table_widget.setItem(i, j, QTableWidgetItem(str(data.iat[i, j])))

    def calculate_formula(self, row, col):
        data = self.data_handler.get_data()
        formula = self.column_info[col]['formula']
        formula = formula.lstrip('=')
        formula = re.sub(r'\d+', '', formula)

        print(f"formula: {formula}")

        # 변수에 대응하는 값을 저장할 딕셔너리
        variables = {}

        # 각 변수에 대응하는 숫자를 사용자로부터 입력받기
        for char in formula:
            if char.isalpha():  # 알파벳인지 확인
                if char not in variables:  # 이미 입력받은 변수는 건너뜀
                    index_by_alphabet = None
                    for index, details in self.column_info.items():
                        if details.get('alphabet') == char:
                            index_by_alphabet = index
                            break

                    value = data.iat[row, index_by_alphabet]
                    variables[char] = value

        # 변수 대입 및 계산
        for var, value in variables.items():
            formula = formula.replace(var, str(value))

        try:
            # 계산 수행
            result = eval(formula)
            rounded_result = round(result, 1)
        except ZeroDivisionError:
            print("Error: Division by zero")
            rounded_result = 0
        except Exception as e:
            print(f"Error: {e}")
            rounded_result = 0

        # 결과 출력
        print(f"계산 결과: {rounded_result}")
        return rounded_result

    @pyqtSlot(dict)
    def display_selected_point(self, point):
        info = (
            f"<b>Selected Point:</b><br>"
            f"<b>Key:</b> {point['Key']}<br>"
            f"<b>Summary:</b> {point['Summary']}<br>"
            f"<b>X ({self.x_column}):</b> {point['x']}<br>"
            f"<b>Y ({self.y_column}):</b> {point['y']}<br>"
        )
        self.selected_point_label.setText(info)



    def update_row_by_key(self, key_to_update, new_x_value, new_y_value):
        print("update_row_by_key")

        for row in range(self.table_widget.rowCount()):
            item = self.table_widget.item(row, 0)
            if item is not None and item.text() == key_to_update:
                print("Key found!")
                print(f"row: {row}, x_column_index: {self.x_column_index}, y_column_index: {self.y_column_index}")
                # 보여지는 테이블 데이터 변경 X
                self.table_widget.item(row, self.x_column_index).setText(str(new_x_value))
                self.table_widget.item(row, self.x_column_index).setBackground(Qt.yellow)
                font = self.table_widget.item(row, self.x_column_index).font()
                font.setItalic(True)
                self.table_widget.item(row, self.x_column_index).setForeground(Qt.darkBlue)
                self.table_widget.item(row, self.x_column_index).setFont(font)
                # 보여지는 테이블 데이터 변경 Y
                self.table_widget.item(row, self.y_column_index).setText(str(new_y_value))
                self.table_widget.item(row, self.y_column_index).setBackground(Qt.yellow)
                font = self.table_widget.item(row, self.y_column_index).font()
                font.setItalic(True)
                self.table_widget.item(row, self.y_column_index).setForeground(Qt.darkBlue)
                self.table_widget.item(row, self.y_column_index).setFont(font)

                break

    def update_data_from_gui(self):
        print("update_data_from_gui")
        # GUI 테이블에서 변경된 값을 DataFrame에 저장
        # 수식이 없는 컬럼은 값을 그대로 저장하고
        # 수식이 있는 컬럼은 저장 없이 계산만 하여 보여준다

        # 저장할 DataFrame 가져오기
        data = self.data_handler.get_data()

        # 추가 계산이 필요한 컬럼을 먼저 조사
        column_info = self.data_handler.get_column_info()
        true_formula_indices = [index for index, details in column_info.items() if details['is_formula']]

        # 결과 출력
        print(true_formula_indices)

        print("========================== update_data =============================")
        print(data)
        print("========================== update_data =============================")

        # 이 부분에 수식을 계산해서 숫자로 변환하는 코드 추가
        for i in range(data.shape[0]):
            for j in range(data.shape[1]):
                print(f"i: {i}, j: {j}")
                if j in true_formula_indices:
                    print("to calculate for j")
                    calculated_value = self.calculate_formula(i, j)
                    self.table_widget.setItem(i, j, QTableWidgetItem(str(calculated_value)))
                else:
                    print("only save data for j")
                    cell_value = self.table_widget.item(i, j).text()
                    if cell_value is None:
                        cell_value = ''  # None을 빈 문자열로 대체
                    data.iat[i, j] = cell_value

    def get_table_data(self):
        print("get_table_data!")
        # 테이블의 값을 기준으로 차트를 그릴 DataFrame 생성
        rows = self.table_widget.rowCount()
        cols = self.table_widget.columnCount()
        data = []

        headers = [self.table_widget.horizontalHeaderItem(j).text() for j in range(cols)]

        for i in range(rows):
            row_data = []
            for j in range(cols):
                item = self.table_widget.item(i, j)
                row_data.append(item.text() if item else '')
            data.append(row_data)

        print("================== table data =======================")
        print(data)
        print("================== table data =======================")
        df = pd.DataFrame(data, columns=headers)
        df.index = range(1, len(df) + 1)

        return df

    @pyqtSlot(str, float, float)
    def handle_point_drop(self, key, x, y):
        print("handle_point_drop")
        self.table_widget.itemChanged.disconnect()  # 신호 연결 해제

        print(f"Point dropped: Key={key}, X={x}, Y={y}")
        self.update_row_by_key(key, x, y)

        self.table_widget.itemChanged.connect(self.on_item_changed)  # 신호 다시 연결

    @pyqtSlot(str)
    def highlight_selected_row(self, key):
        print("highlight_selected_row")
        self.table_widget.itemChanged.disconnect()  # 신호 연결 해제

        # key = point['Key']
        data = self.data_handler.get_data()
        row = data.index[data['Key'] == key].tolist()[0] - 1

        print(f"row: {row}")

        # 이전 선택된 행이 있고, 그 행이 현재 행과 다르면 원래 상태로 되돌림
        if self.previous_selected_row != -1 and self.previous_selected_row != row:
            print("1")
            for col in range(self.table_widget.columnCount()):
                item = self.table_widget.item(self.previous_selected_row, col)
                if item.background().color() != QColor(Qt.yellow):
                    item.setBackground(QColor(Qt.white))

        # 현재 선택된 행을 강조
        for col in range(self.table_widget.columnCount()):
            print(f"for range is {self.table_widget.columnCount()}")
            item = self.table_widget.item(row, col)
            if item.background().color() != QColor(Qt.yellow):
                item.setBackground(Qt.green)
            # self.table_widget.item(row, col).setBackground(Qt.green)

        # 선택된 행이 보이도록 스크롤
        #self.table_widget.scrollToItem(self.table_widget.item(row, 0), QAbstractItemView.PositionAtCenter)
        self.table_widget.verticalScrollBar().setValue(row)

        self.previous_selected_row = row

        self.table_widget.itemChanged.connect(self.on_item_changed)  # 신호 다시 연결



    def reset_table_style(self):
        # 기본 폰트 및 색상 설정
        default_font = QFont()
        default_background = QColor(Qt.white)
        default_foreground = QColor(Qt.black)

        for row in range(self.table_widget.rowCount()):
            for column in range(self.table_widget.columnCount()):
                item = self.table_widget.item(row, column)
                if item:
                    item.setBackground(default_background)
                    item.setForeground(default_foreground)
                    item.setFont(default_font)



    def on_item_changed(self, item):
        # X, Y축이 선택되어 있다면 차트를 업데이트
        if self.x_column and self.y_column:
            self.plot_chart()

    def generate_colors(self):
        print("generate_colors")
        if self.data_handler is not None:
            print("generate random color!")
            self.colors = [self.random_color() for _ in range(self.data_handler.get_data().shape[0])]

    def random_color(self):
        r = lambda: random.randint(0, 255)
        return f'#{r():02x}{r():02x}{r():02x}'

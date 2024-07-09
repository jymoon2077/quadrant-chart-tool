from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QHBoxLayout, QPushButton, \
    QTableWidget, QTableWidgetItem, QFileDialog, QComboBox, QSplitter, QMessageBox
from PyQt5.QtCore import Qt, pyqtSlot
from chart import ChartCanvas
from data_handler import DataHandler
import sys


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Quadrant Chart Tool')

        self.resize(1920, 1080)  # 윈도우 기본 크기 설정

        self.x_column = ''
        self.y_column = ''
        self.x_column_index = -1
        self.y_column_index = -1

        self.data_handler = DataHandler()

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QHBoxLayout(self.central_widget)

        self.left_layout = QVBoxLayout()
        self.right_layout = QVBoxLayout()

        # 버튼들을 위한 레이아웃
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

        # X와 Y 축 서로 변경 버튼 추가
        self.switch_axes_button = QPushButton('Switch Axes')
        self.switch_axes_button.clicked.connect(self.switch_axes)
        self.button_layout.addWidget(self.switch_axes_button)

        self.plot_button = QPushButton('Plot Chart')
        self.plot_button.clicked.connect(self.plot_chart)
        self.button_layout.addWidget(self.plot_button)

        self.save_button = QPushButton('Save Changes')
        self.save_button.clicked.connect(self.save_changes)
        self.button_layout.addWidget(self.save_button)

        # 버튼 레이아웃을 왼쪽 레이아웃의 상단에 추가
        self.left_layout.addLayout(self.button_layout)

        # 데이터 테이블 생성
        self.table_widget = QTableWidget()
        self.left_layout.addWidget(self.table_widget)

        # 선택된 데이터 포인트 정보를 표시할 레이블
        self.selected_point_label = QLabel('Selected Point: None')
        self.right_layout.addWidget(self.selected_point_label, 1)

        # 차트 캔버스 생성
        self.chart_canvas = ChartCanvas(self)
        self.chart_canvas.point_selected.connect(self.display_selected_point)  # Signal 연결
        self.chart_canvas.point_dropped.connect(self.handle_point_drop)  # Signal 연결
        self.right_layout.addWidget(self.chart_canvas, 9)

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

    # def resizeEvent(self, event):
    #     super().resizeEvent(event)
    #     # 윈도우가 리사이즈 될 때 차트 캔버스를 정사각형으로 유지
    #     self.chart_canvas.setFixedSize(self.splitter.sizes()[1], self.splitter.sizes()[1])

    # 버튼

    def load_data(self):
        file_path, _ = QFileDialog.getOpenFileName(self, 'Open File', '', 'Excel Files (*.xlsx)')
        if file_path:
            self.data_handler.load_data(file_path)
            self.display_data()
            self.populate_combo_boxes()

    def plot_chart(self):
        self.x_column = self.x_combo_box.currentText()
        self.y_column = self.y_combo_box.currentText()

        if self.x_column == "" or self.y_column == "":
            QMessageBox.critical(self, 'Error', 'X and Y axes must be selected.')
            return

        if self.x_column == self.y_column:
            QMessageBox.critical(self, 'Error', 'X and Y axes must be different.')
            self.y_combo_box.setCurrentIndex(-1)
            return

        header = self.table_widget.horizontalHeader()
        model = header.model()
        column_count = header.count()
        print(column_count)

        for i in range(column_count):
            text = model.headerData(i, Qt.Horizontal)
            print(i)
            print(text)
            # item = header
            # print(item.text())
            # print(self.x_column)
            if text == self.x_column:
                self.x_column_index = i
            elif text == self.y_column:
                self.y_column_index = i
            elif self.x_column_index != -1 & self.y_column_index != -1:
                break


        chart_data = self.data_handler.get_chart_data(self.x_column, self.y_column)
        if chart_data is not None:
            self.chart_canvas.plot(chart_data, self.x_column, self.y_column)

    def save_changes(self):
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save File', '', 'Excel Files (*.xlsx)')
        if file_path:
            self.data_handler.save_data(file_path)

    # 콤보 박스

    def populate_combo_boxes(self):
        data = self.data_handler.get_data()
        columns = data.columns

        exclude_texts = ['Key', 'Summary']  # 여러 개의 텍스트를 리스트로 정의

        filtered_columns = [col for col in columns if not any(text in col for text in exclude_texts)]

        self.x_combo_box.clear()
        self.y_combo_box.clear()
        self.x_combo_box.addItems(filtered_columns)
        self.y_combo_box.addItems(filtered_columns)

    def switch_axes(self):
        current_x = self.x_combo_box.currentText()
        current_y = self.y_combo_box.currentText()

        # Swap current X and Y selections
        self.x_combo_box.setCurrentText(current_y)
        self.y_combo_box.setCurrentText(current_x)

        # X와 Y 축이 동일한지 검사하는 슬롯 추가
    def check_axes_selection(self):
        selected_x = self.x_combo_box.currentText()
        selected_y = self.y_combo_box.currentText()

        if selected_x == selected_y:
            QMessageBox.critical(self, 'Error', 'X and Y axes must be different.')
            # X와 Y 축이 동일한 경우, X 축 콤보 박스를 초기 상태로 되돌림
            self.y_combo_box.setCurrentIndex(-1)

    def display_data(self):
        data = self.data_handler.get_data()
        self.table_widget.setRowCount(data.shape[0])
        self.table_widget.setColumnCount(data.shape[1])
        self.table_widget.setHorizontalHeaderLabels(data.columns)

        for i in range(data.shape[0]):
            for j in range(data.shape[1]):
                self.table_widget.setItem(i, j, QTableWidgetItem(str(data.iat[i, j])))

    @pyqtSlot(dict)
    def display_selected_point(self, point):
        info = f"Selected Point: Key={point['Key']}, X={point['x']}, Y={point['y']}, Summary={point['Summary']}"
        self.selected_point_label.setText(info)


    def update_table_cell_color(self, key):
        # 테이블의 셀 색상 변경 함수
        for row in range(self.table_widget.rowCount()):
            item = self.table_widget.item(row, 0)  # Key 열에 해당하는 item 가져오기
            if item.text() == key:
                item.setBackground(Qt.yellow)  # 변경된 데이터의 셀을 노란색으로 표시

    @pyqtSlot(str, float, float)
    def handle_point_drop(self, key, x, y):
        # 이 메서드에서 왼쪽 테이블의 값을 변경하고 표시를 업데이트할 수 있음
        print(f"Point dropped: Key={key}, X={x}, Y={y}")

        self.update_row_by_key(key, x, y)

    def update_row_by_key(self, key_to_update, new_x_value, new_y_value):
        print("update_row_by_key")
        # Find the row index where 'Key' column matches key_to_update
        for row in range(self.table_widget.rowCount()):
            item = self.table_widget.item(row, 0)  # Assuming 'Key' column is the second column (index 1)
            if item is not None and item.text() == key_to_update:
                print("Key found!")
                # Update X and Y values in the same row
                # 보여지는 테이블 데이터 변경
                self.table_widget.item(row, self.x_column_index).setText(str(new_x_value))  # Update X value
                self.table_widget.item(row, self.x_column_index).setBackground(Qt.yellow)
                self.table_widget.item(row, self.y_column_index).setText(str(new_y_value))  # Update Y value
                self.table_widget.item(row, self.y_column_index).setBackground(Qt.yellow)

                # 실제 DataFrame 변경
                self.data_handler.data.at[row, self.x_column] = new_x_value
                self.data_handler.data.at[row, self.y_column] = new_y_value

                break



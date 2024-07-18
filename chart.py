import math
import random
from matplotlib import lines
from matplotlib.backend_bases import MouseButton
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt5.QtCore import pyqtSignal, pyqtSlot, Qt
from PyQt5.QtWidgets import QTableWidgetItem, QMessageBox
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
import numpy as np
import pandas as pd
import adjustText as aT


class ChartCanvas(FigureCanvas):
    point_selected = pyqtSignal(dict)
    point_clicked = pyqtSignal(str)
    point_dropped = pyqtSignal(str, float, float)

    def __init__(self, parent=None):

        fig = Figure()
        self.axes = fig.add_subplot(111)
        super().__init__(fig)
        self.setParent(parent)
        self.selected_point = None
        self.data = None
        self.y_label = None
        self.x_label = None
        # self.is_reversed = False
        self.is_x_reversed = False
        self.is_y_reversed = False
        self.annotates = []
        self.chart_size_x = 0
        self.chart_size_y = 0
        self.hline = None
        self.vline = None
        self.dragging_line = None
        self.press = None
        self.x_mid = None  # 초기화
        self.y_mid = None  # 초기화
        self.prev_mouse_x = None
        self.prev_mouse_y = None

        # 테스트용 기본 툴바
        self.toolbar = NavigationToolbar(self, parent)
        parent.addToolBar(self.toolbar)

        self.cid = self.mpl_connect('button_press_event', self.on_click)
        self.cidmotion = self.mpl_connect('motion_notify_event', self.on_motion)
        self.cidrelease = self.mpl_connect('button_release_event', self.on_release)

    def initialize(self, is_swap):
        if is_swap is True:
            self.x_mid, self.y_mid = self.y_mid, self.x_mid
        else:
            self.data = None
            self.x_mid = None
            self.y_mid = None

    @pyqtSlot(pd.DataFrame, str, str, list)
    def plot(self, data, x_label, y_label, colors):
        print("plot")
        self.data = data
        self.x_label = x_label
        self.y_label = y_label

        if 'Color' not in self.data.columns:
            # self.data['Color'] = [self.random_color() for _ in range(len(self.data))]
            self.data['Color'] = colors

        self.update_plot(True)

    def update_plot(self, force_redraw):
        print("update_plot")
        self.axes.clear()
        self.annotates = []

        print(f"before -> x_mid: {self.x_mid}, y_mid: {self.y_mid}")

        if force_redraw is True:
            self.get_chart_max_size()

        if self.data is not None:
            x_data = self.data.iloc[:, 0].astype(np.float64)
            y_data = self.data.iloc[:, 1].astype(np.float64)

            # 사분면 크기 설정
            x_max, y_max = self.chart_size_x, self.chart_size_y
            # x_mid, y_mid = x_max / 2, y_max / 2
            if self.x_mid is None or self.y_mid is None:
                print("line /2")
                self.x_mid = x_max / 2
                self.y_mid = y_max / 2

            if self.x_mid >= x_max:
                t = x_max * 0.99
                self.x_mid = math.floor(t * 10) / 10
            if self.y_mid >= y_max:
                t = y_max * 0.99
                self.y_mid = math.floor(t * 10) / 10

            print(f"after -> x_mid: {self.x_mid}, y_mid: {self.y_mid}")

            # 사분면 선 그리기
            # self.axes.axhline(y=y_mid, color='black', linestyle='--')
            # self.axes.axvline(x=x_mid, color='black', linestyle='--')

            self.hline = lines.Line2D([0, x_max], [self.y_mid, self.y_mid], color='black', linestyle='--')
            self.vline = lines.Line2D([self.x_mid, self.x_mid], [0, y_max], color='black', linestyle='--')

            self.axes.add_line(self.hline)
            self.axes.add_line(self.vline)

            # # 사분면 배경 그리기
            # self.axes.fill_between([0, x_mid], y_mid, y_max, color='red', alpha=0.1)
            # self.axes.fill_between([x_mid, x_max], y_mid, y_max, color='blue', alpha=0.1)
            # self.axes.fill_between([0, x_mid], 0, y_mid, color='green', alpha=0.1)
            # self.axes.fill_between([x_mid, x_max], 0, y_mid, color='yellow', alpha=0.1)

            colors = self.data['Color'].tolist()
            self.scatter = self.axes.scatter(x_data, y_data, c=colors, picker=True)

            # 차트 데이터 생성
            for i, key in enumerate(self.data['Key']):
                color = self.data['Color'].iloc[i]
                annotate = self.axes.annotate(key, (x_data.iloc[i], y_data.iloc[i]),
                                              bbox=dict(facecolor=color, alpha=0.5))
                self.annotates.append(annotate)

            # 사분면 X, Y 범위 설정
            if self.is_x_reversed:
                self.axes.set_xlim(x_max, 0)
            else:
                self.axes.set_xlim(0, x_max)

            if self.is_y_reversed:
                self.axes.set_ylim(y_max, 0)
            else:
                self.axes.set_ylim(0, y_max)

            self.axes.set_title('Quadrant Chart', color='green')
            self.axes.set_xlabel(self.x_label, color='red')
            self.axes.set_ylabel(self.y_label, color='blue')

        self.draw()


    def get_chart_max_size(self):
        # 차트 사이즈를 현재 데이터 기준으로 갱신해야 할 때 호출
        if self.data is not None:
            x_data = self.data.iloc[:, 0].astype(np.float64)
            y_data = self.data.iloc[:, 1].astype(np.float64)

            self.chart_size_x = x_data.max()
            self.chart_size_y = y_data.max()

    def on_click(self, event):
        if event.inaxes != self.axes:
            return

        print("on_click!")

        if event.button == MouseButton.LEFT:
            contains, _ = self.hline.contains(event)
            if contains:
                self.dragging_line = self.hline
                self.press = self.hline.get_ydata()
            contains, _ = self.vline.contains(event)
            if contains:
                self.dragging_line = self.vline
                self.press = self.vline.get_xdata()

            self.prev_mouse_x, self.prev_mouse_y = event.xdata, event.ydata

        for annotate in self.annotates:
            contains, attr = annotate.contains(event)
            if contains:
                key = annotate.get_text()
                if key in self.data['Key'].values:
                    ind = self.data.index[self.data['Key'] == key].tolist()[0]
                    self.selected_point = {
                        "Key": key,
                        "x": self.data.loc[ind, self.x_label],
                        "y": self.data.loc[ind, self.y_label],
                        "Summary": self.data.loc[ind, "Summary"]
                    }
                    self.point_selected.emit(self.selected_point)
                    self.point_clicked.emit(key)
                break

        #
        self.background = self.figure.canvas.copy_from_bbox(self.axes.bbox)
        self.axes.draw_artist(self.scatter)
        self.figure.canvas.blit(self.axes.bbox)

    def on_motion(self, event):
        if event.inaxes != self.axes:
            return

        if self.selected_point is not None:
            print("on_motion")
            # Update the data frame with new coordinates
            if event.xdata is not None and event.ydata is not None:
                key = self.selected_point['Key']

                if self.is_x_reversed:
                    new_x = max(min(self.axes.get_xlim()[0], event.xdata), self.axes.get_xlim()[1])
                else:
                    new_x = min(max(event.xdata, self.axes.get_xlim()[0]), self.axes.get_xlim()[1])

                if self.is_y_reversed:
                    new_y = max(min(self.axes.get_ylim()[0], event.ydata), self.axes.get_ylim()[1])
                else:
                    new_y = min(max(event.ydata, self.axes.get_ylim()[0]), self.axes.get_ylim()[1])

                self.data.loc[self.data['Key'] == key, self.x_label] = new_x
                self.data.loc[self.data['Key'] == key, self.y_label] = new_y
                self.update_plot(False)

        if self.dragging_line is None:
            return

        if self.dragging_line == self.hline:
            y0 = event.ydata
            self.hline.set_ydata([y0, y0])
        elif self.dragging_line == self.vline:
            x0 = event.xdata
            self.vline.set_xdata([x0, x0])

        self.figure.canvas.draw()

    def on_release(self, event):
        if self.selected_point is not None:
            print("on_release")

            # 마우스 커서가 전혀 움직이지 않았다면
            if self.prev_mouse_x == event.xdata and self.prev_mouse_y == event.ydata:
                self.selected_point = None
                return

            key = self.selected_point['Key']
            new_x = round(event.xdata, 1)
            new_y = round(event.ydata, 1)

            if self.is_x_reversed:
                new_x = max(min(self.axes.get_xlim()[0], new_x), self.axes.get_xlim()[1])

            if self.is_y_reversed:
                new_y = max(min(self.axes.get_ylim()[0], new_y), self.axes.get_ylim()[1])

            self.data.loc[self.data['Key'] == key, self.x_label] = new_x
            self.data.loc[self.data['Key'] == key, self.y_label] = new_y
            self.update_plot(False)
            self.point_dropped.emit(key, new_x, new_y)

            # 업데이트된 data point 의 정보를 상단에 노출
            self.selected_point['x'] = new_x
            self.selected_point['y'] = new_y
            self.point_selected.emit(self.selected_point)

            self.selected_point = None

        if event.button == MouseButton.LEFT:
            # self.x_mid = event.xdata
            # self.y_mid = event.ydata

            if self.dragging_line == self.vline:
                self.x_mid = event.xdata
                print(f"Vertical line dropped at x = {self.x_mid}")

            elif self.dragging_line == self.hline:
                self.y_mid = event.ydata
                print(f"Horizontal line dropped at y = {self.y_mid}")

            self.dragging_line = None
            self.press = None

    @pyqtSlot()
    def reverse_chart(self):
        self.is_reversed = not self.is_reversed
        self.update_plot(False)

    def reverse_x_axis(self):
        self.is_x_reversed = not self.is_x_reversed
        self.update_plot(False)

    def reverse_y_axis(self):
        self.is_y_reversed = not self.is_y_reversed
        self.update_plot(False)

    @pyqtSlot(int)
    def highlight_point(self, row):
        print("highlight_point!")
        key = self.data.iloc[row]['Key']
        print(f"highlight_point > key: {key}")
        for i, annotate in enumerate(self.annotates):
            if annotate.get_text() == key:
                print(f"highlight_point > for loop: {key}th annotate found!")
                # QMessageBox.critical(self, 'Error', 'highlight_point <- signal accepted!')
                self.selected_point = {
                    "Key": key,
                    "x": self.data.iloc[i][self.x_label],
                    "y": self.data.iloc[i][self.y_label],
                    "Summary": self.data.iloc[i]["Summary"]
                }
                self.point_selected.emit(self.selected_point)
                self.point_clicked.emit(key)
                break

        self.update_plot(False)  # 주석 강조 업데이트
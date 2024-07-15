from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt5.QtCore import pyqtSignal, pyqtSlot, Qt
from PyQt5.QtWidgets import QTableWidgetItem
import numpy as np
import pandas as pd
import random

class ChartCanvas(FigureCanvas):
    point_selected = pyqtSignal(dict)
    point_dropped = pyqtSignal(str, float, float)

    def __init__(self, parent=None):
        fig = Figure()
        self.axes = fig.add_subplot(111)
        super().__init__(fig)
        self.setParent(parent)
        self.selected_point = None
        self.data = None
        # self.is_reversed = False
        self.is_x_reversed = False
        self.is_y_reversed = False
        self.annotates = []
        self.chart_size_x = 0
        self.chart_size_y = 0

        self.cid = self.mpl_connect('button_press_event', self.on_click)
        self.cidmotion = self.mpl_connect('motion_notify_event', self.on_motion)
        self.cidrelease = self.mpl_connect('button_release_event', self.on_release)

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

        if force_redraw is True:
            self.get_chart_max_size()

        if self.data is not None:
            x_data = self.data.iloc[:, 0].astype(np.float64)
            y_data = self.data.iloc[:, 1].astype(np.float64)

            # 사분면 크기 설정
            x_max, y_max = self.chart_size_x, self.chart_size_y
            x_mid, y_mid = x_max / 2, y_max / 2

            # 사분면 선 그리기
            self.axes.axhline(y=y_mid, color='black', linestyle='--')
            self.axes.axvline(x=x_mid, color='black', linestyle='--')

            # 사분면 배경 그리기
            self.axes.fill_between([0, x_mid], y_mid, y_max, color='red', alpha=0.1)
            self.axes.fill_between([x_mid, x_max], y_mid, y_max, color='blue', alpha=0.1)
            self.axes.fill_between([0, x_mid], 0, y_mid, color='green', alpha=0.1)
            self.axes.fill_between([x_mid, x_max], 0, y_mid, color='yellow', alpha=0.1)

            print("===================== update plot ========================")
            print(self.data)
            print("===================== update plot ========================")

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

    def random_color(self):
        r = lambda: random.randint(0, 255)
        return f'#{r():02x}{r():02x}{r():02x}'

    def on_click(self, event):
        if event.inaxes != self.axes:
            return

        print("on_click!")

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
                break

        #
        self.background = self.figure.canvas.copy_from_bbox(self.axes.bbox)
        self.axes.draw_artist(self.scatter)
        self.figure.canvas.blit(self.axes.bbox)

    def on_motion(self, event):
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

    def on_release(self, event):
        if self.selected_point is not None:
            print("on_release")
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
from PyQt5.QtWidgets import QWidget
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

        self.cid = self.mpl_connect('button_press_event', self.on_click)
        self.cidmotion = self.mpl_connect('motion_notify_event', self.on_motion)
        self.cidrelease = self.mpl_connect('button_release_event', self.on_release)

    @pyqtSlot(pd.DataFrame, str, str)
    def plot(self, data, x_label, y_label):
        self.data = data
        self.x_label = x_label
        self.y_label = y_label

        if 'Color' not in self.data.columns:
            self.data['Color'] = [self.random_color() for _ in range(len(self.data))]

        self.update_plot()

    def update_plot(self):
        self.axes.clear()
        if self.data is not None:
            x_data = self.data.iloc[:, 0].astype(np.float64)
            y_data = self.data.iloc[:, 1].astype(np.float64)

            # 사분면 색상 설정
            x_max, y_max = x_data.max(), y_data.max()
            x_mid, y_mid = x_max / 2, y_max / 2

            self.axes.axhline(y=y_mid, color='black', linestyle='--')
            self.axes.axvline(x=x_mid, color='black', linestyle='--')

            self.axes.fill_between([0, x_mid], y_mid, y_max, color='red', alpha=0.1)
            self.axes.fill_between([x_mid, x_max], y_mid, y_max, color='blue', alpha=0.1)
            self.axes.fill_between([0, x_mid], 0, y_mid, color='green', alpha=0.1)
            self.axes.fill_between([x_mid, x_max], 0, y_mid, color='yellow', alpha=0.1)

            colors = self.data['Color'].tolist()
            self.scatter = self.axes.scatter(x_data, y_data, c=colors, picker=True)

            for i, key in enumerate(self.data['Key']):
                self.axes.annotate(key, (x_data.iloc[i], y_data.iloc[i]), bbox=dict(facecolor=colors[i], alpha=0.5))

            # X축과 Y축의 범위를 데이터의 최대 값에 맞추어 설정
            self.axes.set_xlim(0, x_max)
            self.axes.set_ylim(0, y_max)

            self.axes.set_title('Quadrant Chart', color='green')
            self.axes.set_xlabel(self.x_label, color='red')
            self.axes.set_ylabel(self.y_label, color='blue')

        self.draw()

    def random_color(self):
        r = lambda: random.randint(0, 255)
        return f'#{r():02x}{r():02x}{r():02x}'

    def on_click(self, event):
        if event.inaxes != self.axes:
            return

        contains, attr = self.scatter.contains(event)
        if not contains:
            return

        ind = attr['ind'][0]
        self.press = ind, self.data.iloc[ind][self.x_label], self.data.iloc[ind][self.y_label]

        self.selected_point = {
            "Key": self.data.iloc[ind]["Key"],
            "x": self.data.iloc[ind][self.x_label],
            "y": self.data.iloc[ind][self.y_label],
            "Summary": self.data.iloc[ind]["Summary"]
        }

        self.point_selected.emit(self.selected_point)

        self.background = self.figure.canvas.copy_from_bbox(self.axes.bbox)
        self.axes.draw_artist(self.scatter)
        self.figure.canvas.blit(self.axes.bbox)

    def on_motion(self, event):
        if self.selected_point is not None:
            print("on_motion")
            # Update the data frame with new coordinates
            key = self.selected_point['Key']
            new_x = min(max(event.xdata, self.axes.get_xlim()[0]), self.axes.get_xlim()[1])
            new_y = min(max(event.ydata, self.axes.get_ylim()[0]), self.axes.get_ylim()[1])
            self.data.loc[self.data['Key'] == key, self.x_label] = new_x
            self.data.loc[self.data['Key'] == key, self.y_label] = new_y
            self.update_plot()

    def on_release(self, event):
        if self.selected_point is not None:
            print("on_release!")
            new_x = min(max(event.xdata, self.axes.get_xlim()[0]), self.axes.get_xlim()[1])
            new_y = min(max(event.ydata, self.axes.get_ylim()[0]), self.axes.get_ylim()[1])
            self.point_dropped.emit(self.selected_point['Key'], new_x, new_y)
            self.selected_point = None

    def update_table(self, key, new_x, new_y):
        # key를 이용하여 테이블에서 해당 행을 찾고, X, Y 값을 업데이트
        for row in range(self.table_widget.rowCount()):
            if self.table_widget.item(row, 0).text() == key:
                self.table_widget.setItem(row, 1, QTableWidgetItem(str(new_x)))
                self.table_widget.setItem(row, 2, QTableWidgetItem(str(new_y)))
                # 셀 색상 변경 등 추가적인 표시를 원하면 여기에서 처리 가능
                break

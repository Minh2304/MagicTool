from PySide6.QtCore import QPoint, QRect
from PySide6.QtGui import QMouseEvent, Qt, QColor, QPainter, QPen, QCursor, QFont
from PySide6.QtWidgets import QWidget


class GridCanvas(QWidget):
    def __init__(self, fields, field_selected_callback):
        super().__init__()
        self.setMinimumSize(600, 600)

        self.dpi = self.logicalDpiX()
        self.grid_size = int(self.dpi / 2.54)

        self.occupied_cells = set()
        self.rects = []  # (QRect, QColor, field_name)

        self.start_point = None
        self.end_point = None

        self.active_field = None
        self.fields = fields
        self.used_fields = set()
        self.field_selected_callback = field_selected_callback

    def set_active_field(self, field_name):
        self.active_field = field_name

    def clear_rect_by_field(self, field_name):
        new_rects = []
        new_occupied = set()
        for rect, color, name in self.rects:
            if name != field_name:
                new_rects.append((rect, color, name))
                new_occupied |= self.get_cells_in_rect(rect)
        self.rects = new_rects
        self.occupied_cells = new_occupied
        if field_name in self.used_fields:
            self.used_fields.remove(field_name)
            self.field_selected_callback(field_name, remove=True)
        self.update()

    def snap_to_grid(self, point):
        x = point.x() // self.grid_size * self.grid_size
        y = point.y() // self.grid_size * self.grid_size
        return QPoint(x, y)

    def get_cells_in_rect(self, rect: QRect):
        cells = set()
        left = rect.left() // self.grid_size
        top = rect.top() // self.grid_size
        right = (rect.right() - 1) // self.grid_size
        bottom = (rect.bottom() - 1) // self.grid_size
        for x in range(left, right + 1):
            for y in range(top, bottom + 1):
                cells.add((x, y))
        return cells

    def mousePressEvent(self, event: QMouseEvent):
        if self.active_field and event.button() == Qt.LeftButton:
            self.start_point = self.snap_to_grid(event.position().toPoint())
            self.end_point = self.start_point
            self.update()

    def mouseMoveEvent(self, event: QMouseEvent):
        if self.start_point:
            self.end_point = self.snap_to_grid(event.position().toPoint())
        self.update()

    def mouseReleaseEvent(self, event: QMouseEvent):
        if not self.active_field:
            return

        if self.start_point and self.end_point:
            rect = QRect(self.start_point, self.end_point).normalized()

            # Bỏ nếu vẽ trong vùng bị chặn
            if self.is_in_frozen_area(rect):
                self.start_point = None
                self.end_point = None
                self.update()
                return

            # Bỏ qua hình nếu chiều rộng hoặc chiều cao = 0
            if rect.width() == 0 or rect.height() == 0:
                self.start_point = None
                self.end_point = None
                self.update()
                return

            new_cells = self.get_cells_in_rect(rect)

            if not (new_cells & self.occupied_cells):
                color = QColor(0, 100, 255)
                self.rects.append((rect, color, self.active_field))
                self.occupied_cells.update(new_cells)
                self.used_fields.add(self.active_field)
                self.field_selected_callback(self.active_field)
                self.active_field = None

            self.start_point = None
            self.end_point = None
            self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        # Tô nền 3 hàng đầu
        frozen_rect = QRect(0, 0, self.width(), self.grid_size * 3)
        painter.fillRect(frozen_rect, QColor(200, 200, 200))  # Xám đậm hơn lưới
        self.draw_grid(painter)
        self.draw_rects(painter)

        if self.active_field and self.start_point:
            temp_rect = QRect(self.start_point, self.end_point).normalized()
            painter.setPen(QPen(Qt.red, 1, Qt.DashLine))
            painter.drawRect(temp_rect)

        if self.active_field and not self.start_point:
            cursor_pos = self.mapFromGlobal(QCursor.pos())
            painter.setPen(Qt.black)
            painter.setFont(QFont("Arial", 12))
            display_field = self.active_field.replace("3==D", " ")  # Replace 3==D thành dấu cách
            painter.drawText(cursor_pos + QPoint(10, -10), display_field)

    def draw_grid(self, painter):
        pen = QPen(QColor(220, 220, 220), 1)
        painter.setPen(pen)
        width = self.width()
        height = self.height()
        step = self.grid_size

        for x in range(0, width, step):
            painter.drawLine(x, 0, x, height)
        for y in range(0, height, step):
            painter.drawLine(0, y, width, y)

    def draw_rects(self, painter):
        for rect, color, field in self.rects:
            painter.setPen(QPen(color, 2))
            painter.drawRect(rect)
            painter.setFont(QFont("Arial", 10))
            display_field = field.replace("3==D", " ")  # Replace 3==D thành dấu cách
            painter.drawText(rect.topLeft() + QPoint(4, -4), display_field)

    def is_in_frozen_area(self, rect: QRect) -> bool:
        top_row = rect.top() // self.grid_size
        bottom_row = rect.bottom() // self.grid_size
        return top_row < 3  # 3 hàng đầu (0, 1, 2) bị cấm
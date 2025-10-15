from PySide6.QtCore import QPoint, QRect, Qt
from PySide6.QtGui import QMouseEvent, QColor, QPainter, QPen, QCursor, QFont
from PySide6.QtWidgets import QWidget, QMenu


class GridCanvas(QWidget):
    def __init__(self, fields, field_selected_callback):
        super().__init__()
        self.setMinimumSize(600, 600)

        self.dpi = self.logicalDpiX()
        # Đảm bảo grid_size luôn hợp lệ (>0)
        try:
            calculated = int(self.dpi / 2.54)
        except Exception:
            calculated = 0
        self.grid_size = max(calculated, 10)

        self.occupied_cells = set()
        self.rects = []  # (QRect, QColor, field_name)

        self.start_point = None
        self.end_point = None

        self.active_field = None
        self.fields = fields
        self.used_fields = set()
        self.field_selected_callback = field_selected_callback

        # Biến cho chức năng di chuyển và resize
        self.selected_rect_index = None  # Index của rect đang được chọn
        self.is_moving = False
        self.is_resizing = False
        self.resize_handle = None  # 'tl', 'tr', 'bl', 'br', 't', 'b', 'l', 'r'
        self.drag_start_pos = None
        self.original_rect = None
        
        self.setMouseTracking(True)  # Để theo dõi con trỏ chuột
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)

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

    def get_rect_at_pos(self, pos):
        """Tìm rect tại vị trí pos, trả về index"""
        for i in range(len(self.rects) - 1, -1, -1):  # Kiểm tra từ trên xuống
            rect, _, _ = self.rects[i]
            if rect.contains(pos):
                return i
        return None

    def get_resize_handle(self, pos, rect):
        """Xác định handle resize nào đang được hover (góc hoặc cạnh)"""
        handle_size = 8  # Kích thước vùng handle
        
        # Kiểm tra các góc
        tl = QRect(rect.left() - handle_size//2, rect.top() - handle_size//2, handle_size, handle_size)
        tr = QRect(rect.right() - handle_size//2, rect.top() - handle_size//2, handle_size, handle_size)
        bl = QRect(rect.left() - handle_size//2, rect.bottom() - handle_size//2, handle_size, handle_size)
        br = QRect(rect.right() - handle_size//2, rect.bottom() - handle_size//2, handle_size, handle_size)
        
        if tl.contains(pos):
            return 'tl'  # top-left
        if tr.contains(pos):
            return 'tr'  # top-right
        if bl.contains(pos):
            return 'bl'  # bottom-left
        if br.contains(pos):
            return 'br'  # bottom-right
        
        # Kiểm tra các cạnh
        edge_margin = 5
        if abs(pos.x() - rect.left()) < edge_margin and rect.top() <= pos.y() <= rect.bottom():
            return 'l'  # left
        if abs(pos.x() - rect.right()) < edge_margin and rect.top() <= pos.y() <= rect.bottom():
            return 'r'  # right
        if abs(pos.y() - rect.top()) < edge_margin and rect.left() <= pos.x() <= rect.right():
            return 't'  # top
        if abs(pos.y() - rect.bottom()) < edge_margin and rect.left() <= pos.x() <= rect.right():
            return 'b'  # bottom
        
        return None

    def update_cursor(self, pos):
        """Cập nhật cursor dựa trên vị trí chuột"""
        if self.is_moving or self.is_resizing:
            return
        
        rect_index = self.get_rect_at_pos(pos)
        if rect_index is not None:
            rect, _, _ = self.rects[rect_index]
            handle = self.get_resize_handle(pos, rect)
            
            if handle in ['tl', 'br']:
                self.setCursor(Qt.SizeFDiagCursor)
            elif handle in ['tr', 'bl']:
                self.setCursor(Qt.SizeBDiagCursor)
            elif handle in ['t', 'b']:
                self.setCursor(Qt.SizeVerCursor)
            elif handle in ['l', 'r']:
                self.setCursor(Qt.SizeHorCursor)
            else:
                self.setCursor(Qt.SizeAllCursor)  # Move cursor
        else:
            self.setCursor(Qt.ArrowCursor)

    def mousePressEvent(self, event: QMouseEvent):
        pos = event.position().toPoint()
        
        # Kiểm tra nếu click vào rect hiện có
        rect_index = self.get_rect_at_pos(pos)
        
        if rect_index is not None and event.button() == Qt.LeftButton:
            rect, color, field = self.rects[rect_index]
            handle = self.get_resize_handle(pos, rect)
            
            self.selected_rect_index = rect_index
            self.drag_start_pos = pos
            self.original_rect = QRect(rect)
            
            # Xóa cells cũ của rect này khỏi occupied_cells
            old_cells = self.get_cells_in_rect(rect)
            self.occupied_cells -= old_cells
            
            if handle:
                self.is_resizing = True
                self.resize_handle = handle
            else:
                self.is_moving = True
        
        elif self.active_field and event.button() == Qt.LeftButton:
            # Vẽ rect mới
            self.start_point = self.snap_to_grid(pos)
            self.end_point = self.start_point
            self.update()

    def mouseMoveEvent(self, event: QMouseEvent):
        pos = event.position().toPoint()
        
        if self.is_moving and self.selected_rect_index is not None:
            # Di chuyển rect
            delta = pos - self.drag_start_pos
            rect, color, field = self.rects[self.selected_rect_index]
            new_rect = self.original_rect.translated(delta)
            new_rect = self.snap_rect_to_grid(new_rect)
            
            # Kiểm tra không vào frozen area và không overlap
            if not self.is_in_frozen_area(new_rect):
                new_cells = self.get_cells_in_rect(new_rect)
                # Chỉ check overlap với các rect khác (không tính rect hiện tại)
                if not (new_cells & self.occupied_cells):
                    self.rects[self.selected_rect_index] = (new_rect, color, field)
            
        elif self.is_resizing and self.selected_rect_index is not None:
            # Resize rect
            rect, color, field = self.rects[self.selected_rect_index]
            snapped_pos = self.snap_to_grid(pos)
            new_rect = self.resize_rect(self.original_rect, snapped_pos, self.resize_handle)
            
            # Kiểm tra kích thước tối thiểu
            if new_rect.width() >= self.grid_size and new_rect.height() >= self.grid_size:
                if not self.is_in_frozen_area(new_rect):
                    new_cells = self.get_cells_in_rect(new_rect)
                    if not (new_cells & self.occupied_cells):
                        self.rects[self.selected_rect_index] = (new_rect, color, field)
        
        elif self.start_point:
            # Vẽ rect mới
            self.end_point = self.snap_to_grid(pos)
        
        # Cập nhật cursor
        self.update_cursor(pos)
        self.update()

    def mouseReleaseEvent(self, event: QMouseEvent):
        if self.is_moving or self.is_resizing:
            # Hoàn tất di chuyển/resize
            if self.selected_rect_index is not None:
                rect, color, field = self.rects[self.selected_rect_index]
                new_cells = self.get_cells_in_rect(rect)
                self.occupied_cells.update(new_cells)
            
            self.is_moving = False
            self.is_resizing = False
            self.selected_rect_index = None
            self.resize_handle = None
            self.drag_start_pos = None
            self.original_rect = None
            self.update()
            return
        
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

    def snap_rect_to_grid(self, rect):
        """Snap toàn bộ rect về grid"""
        left = rect.left() // self.grid_size * self.grid_size
        top = rect.top() // self.grid_size * self.grid_size
        width = rect.width()
        height = rect.height()
        return QRect(left, top, width, height)

    def resize_rect(self, original_rect, new_pos, handle):
        """Tính toán rect mới khi resize"""
        new_rect = QRect(original_rect)
        
        if 'l' in handle:  # Left
            new_rect.setLeft(new_pos.x())
        if 'r' in handle:  # Right
            new_rect.setRight(new_pos.x())
        if 't' in handle:  # Top
            new_rect.setTop(new_pos.y())
        if 'b' in handle:  # Bottom
            new_rect.setBottom(new_pos.y())
        
        return new_rect.normalized()

    def show_context_menu(self, pos):
        """Hiển thị menu chuột phải để xóa rect"""
        rect_index = self.get_rect_at_pos(pos)
        if rect_index is not None:
            rect, color, field = self.rects[rect_index]
            
            menu = QMenu(self)
            delete_action = menu.addAction("❌ Xóa vùng")
            action = menu.exec(self.mapToGlobal(pos))
            
            if action == delete_action:
                # Xóa rect
                old_cells = self.get_cells_in_rect(rect)
                self.occupied_cells -= old_cells
                self.rects.pop(rect_index)
                
                # Kiểm tra nếu không còn rect nào của field này
                field_still_exists = any(f == field for _, _, f in self.rects)
                if not field_still_exists:
                    self.used_fields.discard(field)
                    self.field_selected_callback(field, remove=True)
                
                self.update()

    def paintEvent(self, event):
        # Tránh vẽ khi widget chưa sẵn sàng
        if self.width() <= 0 or self.height() <= 0 or self.grid_size <= 0:
            return

        painter = QPainter(self)
        if not painter.isActive():
            return

        # Tô nền 3 hàng đầu
        frozen_rect = QRect(0, 0, self.width(), self.grid_size * 3)
        painter.fillRect(frozen_rect, QColor(200, 200, 200))  # Xám đậm hơn lưới
        self.draw_grid(painter)
        self.draw_rects(painter)

        if self.active_field and self.start_point and self.end_point:
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
        for i, (rect, color, field) in enumerate(self.rects):
            if rect.width() <= 0 or rect.height() <= 0:
                continue
            
            # Vẽ rect
            pen_width = 3 if i == self.selected_rect_index else 2
            painter.setPen(QPen(color, pen_width))
            painter.drawRect(rect)
            
            # Vẽ text
            painter.setFont(QFont("Arial", 10))
            display_field = field.replace("3==D", " ")  # Replace 3==D thành dấu cách
            top_left = rect.topLeft() + QPoint(4, 14)
            painter.drawText(top_left, display_field)
            
            # Vẽ resize handles nếu đang được chọn
            if i == self.selected_rect_index:
                self.draw_resize_handles(painter, rect)

    def draw_resize_handles(self, painter, rect):
        """Vẽ các handle để resize"""
        handle_size = 8
        painter.setBrush(QColor(0, 100, 255))
        painter.setPen(QPen(Qt.white, 1))
        
        # Vẽ 4 góc
        handles = [
            rect.topLeft(),
            rect.topRight(),
            rect.bottomLeft(),
            rect.bottomRight()
        ]
        
        for handle_pos in handles:
            handle_rect = QRect(
                handle_pos.x() - handle_size // 2,
                handle_pos.y() - handle_size // 2,
                handle_size,
                handle_size
            )
            painter.drawRect(handle_rect)

    def is_in_frozen_area(self, rect: QRect) -> bool:
        top_row = rect.top() // self.grid_size
        bottom_row = rect.bottom() // self.grid_size
        return top_row < 3  # 3 hàng đầu (0, 1, 2) bị cấm
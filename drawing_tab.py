import os
import sys
import threading

from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QPushButton, QVBoxLayout, QListWidget,
    QListWidgetItem, QMenu, QFrame, QScrollArea, QFileDialog, QMessageBox, QLineEdit, QLabel, QSizePolicy
)
import requests
from PySide6.QtGui import QColor
from PySide6.QtCore import Qt, QPoint, QThread, Signal, QRect, QTimer
from back_end import eu
from grid_canvas import GridCanvas  # nếu bạn để GridCanvas ở file khác


class NotificationWidget(QLabel):
    """Widget thông báo tự động ẩn sau 4 giây"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet(
            "QLabel { "
            "background-color: rgba(45, 140, 255, 230); "
            "color: white; "
            "padding: 15px 25px; "
            "border-radius: 10px; "
            "font-size: 12pt; "
            "font-weight: bold; "
            "}"
        )
        self.hide()
        
        # Timer để tự động ẩn sau 4 giây
        self.timer = QTimer()
        self.timer.timeout.connect(self.fade_out)
        
    def show_message(self, message):
        """Hiển thị thông báo"""
        self.setText(message)
        self.adjustSize()
        
        # Căn giữa theo chiều ngang, đặt ở phía dưới
        parent_width = self.parent().width()
        parent_height = self.parent().height()
        x = (parent_width - self.width()) // 2
        y = parent_height - self.height() - 50  # Cách đáy 50px
        self.move(x, y)
        
        self.show()
        self.raise_()  # Đưa lên trên cùng
        
        # Bắt đầu timer 4 giây
        self.timer.start(4000)
    
    def fade_out(self):
        """Ẩn thông báo"""
        self.timer.stop()
        self.hide()

class ImportWorker(QThread):
    """Worker thread để import Excel không block UI"""
    finished = Signal(dict)  # Signal khi thành công, trả về result
    error = Signal(str)      # Signal khi có lỗi, trả về error message
    
    def __init__(self, file_path, ip):
        super().__init__()
        self.file_path = file_path
        self.ip = ip
    
    def run(self):
        """Chạy trong thread riêng"""
        try:
            # Gửi file Excel lên server
            url = f"http://{self.ip}:5000/upload_excel"
            with open(self.file_path, 'rb') as f:
                response = requests.post(url, files={'file': f})
            
            if response.status_code != 200:
                raise Exception(response.json().get('error', 'Unknown error'))
            
            result = response.json()
            self.finished.emit(result)  # Emit signal với kết quả
            
        except Exception as e:
            self.error.emit(str(e))  # Emit signal lỗi


class DrawingTab(QWidget):
    def __init__(self, fields, main_window):
        super().__init__()
        self.main_window = main_window
        self.fields = fields
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.txt_path = ""
        self.import_worker = None  # Worker thread cho import

        # Canvas vẽ
        self.canvas = GridCanvas(fields, self.mark_field_used)
        self.main_layout.addWidget(self.canvas)

        # Cửa sổ overlay (header nằm phía trên canvas)
        # KHÔNG set geometry cố định, để nó tự điều chỉnh theo resizeEvent
        self.overlay = QWidget(self)
        self.overlay.setStyleSheet("background-color: rgba(240, 240, 240, 0.95); border: 1px solid gray; border-radius: 5px;")
        self.overlay.setAttribute(Qt.WA_TransparentForMouseEvents, False)

        overlay_layout = QHBoxLayout(self.overlay)
        overlay_layout.setContentsMargins(10, 8, 10, 8)
        overlay_layout.setSpacing(10)

        # ===== PHẦN 1: IP + Import Excel + Import TXT + Preview =====
        left_frame = QFrame()
        left_frame.setStyleSheet(
            "QFrame { "
            "background-color: rgba(255, 255, 255, 0.9); "
            "border: 2px solid #2d8cff; "
            "border-radius: 8px; "
            "padding: 5px; "
            "}"
        )
        left_layout = QHBoxLayout(left_frame)
        left_layout.setContentsMargins(10, 0, 10, 0)  # Padding trên/dưới = 0 để căn giữa theo chiều cao
        left_layout.setSpacing(8)
        left_layout.setAlignment(Qt.AlignVCenter)  # Căn giữa theo chiều dọc
        
        # IP server + 2 nút Import + Preview
        self.ip_label = QLabel("IP:")
        self.ip_label.setStyleSheet("QLabel { color: #1f3b75; font-weight: bold; font-size: 10pt; }")
        
        self.ip_input = QLineEdit()
        self.ip_input.setPlaceholderText("IP API")
        self.ip_input.setText("107.98.33.94")
        self.ip_input.setFixedSize(100, 28)
        self.ip_input.setStyleSheet("QLineEdit { padding: 4px; border: 1px solid #aaa; border-radius: 4px; font-size: 9pt; }")
        
        self.import_excel_button = QPushButton("Import Excel")
        self.import_excel_button.setFixedSize(95, 28)
        self.import_excel_button.clicked.connect(self.import_excel_file)
        
        self.import_txt_button = QPushButton("Import TXT")
        self.import_txt_button.setFixedSize(95, 28)
        self.import_txt_button.clicked.connect(self.import_txt_file)
        
        # Nút Preview - Ban đầu ẩn, khi hiện sẽ kéo dài
        self.preview_button = QPushButton("Preview ...")
        self.preview_button.setMinimumWidth(120)  # Chiều rộng tối thiểu
        self.preview_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)  # Kéo dài theo chiều ngang
        self.preview_button.setFixedHeight(28)
        self.preview_button.clicked.connect(self.preview_txt_file)
        self.preview_button.hide()
        
        left_layout.addWidget(self.ip_label)
        left_layout.addWidget(self.ip_input)
        left_layout.addWidget(self.import_excel_button)
        left_layout.addWidget(self.import_txt_button)
        left_layout.addWidget(self.preview_button, 1)  # stretch factor = 1 để kéo dài
        left_layout.addStretch()

        
        # ===== PHẦN 2: List field draw =====
        center_frame = QFrame()
        center_frame.setStyleSheet(
            "QFrame { "
            "background-color: rgba(255, 255, 255, 0.9); "
            "border: 2px solid #2d8cff; "
            "border-radius: 8px; "
            "padding: 5px; "
            "}"
        )
        center_layout = QHBoxLayout(center_frame)
        center_layout.setContentsMargins(10, 0, 10, 0)  # Padding trên/dưới = 0
        center_layout.setSpacing(10)
        center_layout.setAlignment(Qt.AlignVCenter)  # Căn giữa theo chiều dọc
        center_layout.addStretch()
        
        self.toggle_button = QPushButton("List field draw")
        self.toggle_button.setFixedSize(120, 28)
        self.toggle_button.clicked.connect(self.toggle_field_popup)
        center_layout.addWidget(self.toggle_button)
        
        # Nút "Tạo tất cả các trường"
        self.auto_draw_button = QPushButton("Tạo tất cả các trường")
        self.auto_draw_button.setFixedSize(150, 28)
        self.auto_draw_button.clicked.connect(self.auto_draw_all_fields)
        center_layout.addWidget(self.auto_draw_button)
        
        center_layout.addStretch()
        
        # ===== PHẦN 3: Hoàn tất + Reset =====
        right_frame = QFrame()
        right_frame.setStyleSheet(
            "QFrame { "
            "background-color: rgba(255, 255, 255, 0.9); "
            "border: 2px solid #2d8cff; "
            "border-radius: 8px; "
            "padding: 5px; "
            "}"
        )
        right_layout = QHBoxLayout(right_frame)
        right_layout.setContentsMargins(10, 0, 10, 0)  # Padding trên/dưới = 0
        right_layout.setSpacing(8)
        right_layout.setAlignment(Qt.AlignVCenter)  # Căn giữa theo chiều dọc
        right_layout.addStretch()
        
        self.done_button = QPushButton("Hoàn tất")
        self.done_button.setFixedSize(80, 28)
        self.done_button.clicked.connect(self.handle_done_clicked)
        right_layout.addWidget(self.done_button)
        
        self.reset_button = QPushButton("Reset")
        self.reset_button.setFixedSize(70, 28)
        self.reset_button.clicked.connect(self.reset_all)
        right_layout.addWidget(self.reset_button)
        
        right_layout.addStretch()
        
        # Thêm 3 frame vào overlay với tỷ lệ 1:1:1
        overlay_layout.addWidget(left_frame, 1)
        overlay_layout.addWidget(center_frame, 1)
        overlay_layout.addWidget(right_frame, 1)

        # Style cho các nút - Nhỏ gọn và đẹp hơn
        rounded_button_style = (
            "QPushButton {"
            " background-color: #2d8cff; color: white; border: none;"
            " border-radius: 6px; padding: 4px 10px; font-weight: bold; font-size: 9pt;"
            "}"
            "QPushButton:hover { background-color: #1f7ae0; }"
            "QPushButton:pressed { background-color: #1667bf; }"
        )
        for btn in [self.toggle_button, self.import_excel_button, self.import_txt_button, self.preview_button, self.done_button]:
            btn.setStyleSheet(rounded_button_style)
        
        # Style riêng cho nút Reset (màu đỏ)
        reset_button_style = (
            "QPushButton {"
            " background-color: #ff4444; color: white; border: none;"
            " border-radius: 6px; padding: 4px 10px; font-weight: bold; font-size: 9pt;"
            "}"
            "QPushButton:hover { background-color: #dd2222; }"
            "QPushButton:pressed { background-color: #bb1111; }"
        )
        self.reset_button.setStyleSheet(reset_button_style)

        reset_button_style = (
            "QPushButton {"
            " background-color: #ff4444; color: white; border: none;"
            " border-radius: 12px; padding: 6px 12px; font-weight: bold;"
            "}"
            "QPushButton:hover { background-color: #dd2222; }"
            "QPushButton:pressed { background-color: #bb1111; }"
        )
        self.reset_button.setStyleSheet(reset_button_style)

        # Popup danh sách trường (ẩn/hiện bên dưới nút)
        self.popup = QFrame(self)
        self.popup.setWindowFlags(Qt.Window | Qt.FramelessWindowHint | Qt.NoDropShadowWindowHint)
        self.popup.setStyleSheet("background-color: white; border: 1px solid gray;")
        self.popup.setMinimumWidth(200)

        popup_layout = QVBoxLayout(self.popup)
        popup_layout.setContentsMargins(0, 0, 0, 0)

        self.list_widget = QListWidget()
        # Thêm items với 3==D được replace thành dấu cách
        for field in fields:
            display_name = field.replace("3==D", " ").strip()
            item = QListWidgetItem(display_name)
            item.setData(Qt.UserRole, field)  # Lưu tên gốc để sử dụng
            self.list_widget.addItem(item)
        self.list_widget.itemDoubleClicked.connect(self.on_item_double_clicked)
        self.list_widget.setContextMenuPolicy(Qt.CustomContextMenu)
        self.list_widget.customContextMenuRequested.connect(self.show_context_menu)
        self.list_widget.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        popup_layout.addWidget(self.list_widget)
        self.popup.setLayout(popup_layout)
        
        # Tạo notification widget
        self.notification = NotificationWidget(self)

    def resizeEvent(self, event):
        """Điều chỉnh overlay khi resize cửa sổ"""
        super().resizeEvent(event)
        # Đặt overlay kéo dài gần hết chiều rộng, trừ margin 2 bên
        margin = 20
        overlay_width = self.width() - 2 * margin
        overlay_height = 80  # Tăng chiều cao để vừa với các nút đã giảm kích thước và căn giữa theo chiều dọc
        self.overlay.setGeometry(margin, 10, overlay_width, overlay_height)

    def toggle_field_popup(self):
        if self.popup.isVisible():
            self.popup.hide()
        else:
            # Update list trước khi hiển thị: tô xanh các field đã vẽ
            self.update_list_colors()
            
            # Tính vị trí popup nằm dưới button
            button_pos = self.toggle_button.mapToGlobal(QPoint(0, self.toggle_button.height()))
            self.popup.move(button_pos)
            self.popup.show()

    def update_list_colors(self):
        """Cập nhật màu nền cho các item trong list dựa trên used_fields"""
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            field = item.data(Qt.UserRole)
            if field in self.canvas.used_fields:
                item.setBackground(QColor(144, 238, 144))  # Xanh lá nhạt
            else:
                item.setBackground(QColor("white"))

    def on_item_double_clicked(self, item):
        field = item.data(Qt.UserRole)  # Lấy tên gốc thay vì tên hiển thị
        if field not in self.canvas.used_fields:
            self.canvas.set_active_field(field)
            self.popup.hide()  # Ẩn popup sau khi chọn

    def mark_field_used(self, field, remove=False):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.data(Qt.UserRole) == field:  # So sánh với tên gốc
                if remove:
                    item.setBackground(QColor("white"))
                else:
                    item.setBackground(QColor(144, 238, 144))

    def show_context_menu(self, pos):
        item = self.list_widget.itemAt(pos)
        if item:
            menu = QMenu(self)
            delete_action = menu.addAction("❌ Xoá vùng vẽ")
            action = menu.exec(self.list_widget.mapToGlobal(pos))
            if action == delete_action:
                field = item.data(Qt.UserRole)  # Lấy tên gốc
                self.canvas.clear_rect_by_field(field)
                self.mark_field_used(field, remove=True)
                self.popup.hide()  # 🔺 Thêm dòng này để ẩn popup sau khi xóa

    def handle_done_clicked(self):
        rects = self.canvas.rects  # (QRect, QColor, field_name)
        eu.save_config(rects, fields=self.fields)  # ✅ Luôn lưu cả fields
        
        # Đảm bảo Trang chính biết đường dẫn TXT hiện tại và đã load dữ liệu
        if getattr(self, 'txt_path', None):
            self.main_window.current_file_path = self.txt_path
            try:
                if not getattr(self.main_window.sm, 'sentences', []):
                    from sentence_manager import SentenceManager
                    # Giữ instance hiện tại nếu có để tránh mất state nút
                    if not isinstance(self.main_window.sm, SentenceManager):
                        self.main_window.sm = SentenceManager()
                    self.main_window.sm.load_from_txt(self.txt_path)
            except Exception as e:
                print(f"DEBUG: handle_done_clicked load_from_txt error: {e}")

        self.popup.hide()
        self.main_window.on_done(rects)  # Gọi lên MainWindow xử lý

    def load_saved_rects(self, rect_data):
        for rect, field in rect_data:
            color = QColor(0, 100, 255)
            self.canvas.rects.append((rect, color, field))
            self.canvas.occupied_cells.update(self.canvas.get_cells_in_rect(rect))
            self.canvas.used_fields.add(field)
            self.mark_field_used(field)
        self.canvas.update()

    def auto_draw_all_fields(self):
        """Tự động vẽ tất cả các trường chưa được vẽ"""
        if not self.fields:
            QMessageBox.warning(self, "Cảnh báo", "Chưa có danh sách fields!\nVui lòng Import Excel hoặc TXT trước.")
            return
        
        # Tìm các field chưa được vẽ
        undrawn_fields = [field for field in self.fields if field not in self.canvas.used_fields]
        
        if not undrawn_fields:
            self.notification.show_message("Tất cả các trường đã được vẽ!")
            return
        
        # Hỏi xác nhận
        reply = QMessageBox.question(
            self, "Xác nhận",
            f"Bạn có muốn tự động tạo {len(undrawn_fields)} trường chưa được vẽ không?\n\n"
            f"Các trường sẽ được vẽ tự động vào vị trí trống phù hợp.",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        # Tham số để vẽ tự động
        grid_size = int(self.canvas.logicalDpiX() // 2.54)  # Kích thước 1 ô lưới (1cm)
        default_width = int(grid_size * 8)   # Chiều rộng mặc định: 8cm (đã là bội số)
        default_height = int(grid_size * 2)  # Chiều cao mặc định: 2cm (đã là bội số)
        margin = int(grid_size * 1)          # Margin: 1cm (đã là bội số)
        
        canvas_width = int(self.canvas.width())
        canvas_height = int(self.canvas.height())
        
        # Vùng header không được vẽ vào (overlay height = 80px + top margin = 10px)
        header_height = 100  # Chiều cao vùng header cần tránh
        
        # Hàm snap tọa độ về bội số của grid_size
        def snap_to_grid(value):
            """Làm tròn giá trị về bội số gần nhất của grid_size"""
            return int(round(value / grid_size) * grid_size)
        
        # Snap header_height và margin về lưới
        header_height_snapped = snap_to_grid(header_height)
        
        # Hàm tìm vị trí trống phù hợp
        def find_empty_position(width, height):
            """Tìm vị trí trống đầu tiên có thể chứa hình chữ nhật với kích thước cho trước"""
            # Bắt đầu từ dưới header xuống, từ trái sang phải
            # Đảm bảo vị trí bắt đầu nằm trên lưới
            start_x = margin  # margin đã là bội số của grid_size
            start_y = max(margin, header_height_snapped + margin)  # Snap về lưới
            
            # Duyệt theo grid_size để đảm bảo tất cả vị trí đều nằm trên lưới
            y = start_y
            while y + height <= canvas_height:
                x = start_x
                while x + width <= canvas_width:
                    # Tạo rect thử nghiệm - tất cả tọa độ đều là bội số của grid_size
                    test_rect = QRect(x, y, width, height)
                    
                    # Kiểm tra không được vẽ vào vùng header
                    if test_rect.top() < header_height_snapped:
                        x += grid_size
                        continue
                    
                    # Kiểm tra xem có trùng với rect nào đã vẽ không
                    is_collision = False
                    for existing_rect, _, _ in self.canvas.rects:
                        # Thêm margin để tránh dính sát nhau
                        expanded_existing = existing_rect.adjusted(-margin, -margin, margin, margin)
                        if test_rect.intersects(expanded_existing):
                            is_collision = True
                            break
                    
                    if not is_collision:
                        # Tìm thấy vị trí trống
                        return test_rect
                    
                    x += grid_size  # Di chuyển sang phải theo lưới
                y += grid_size  # Di chuyển xuống dưới theo lưới
            
            # Nếu không tìm thấy vị trí trống, đặt ở cuối canvas
            return None
        
        # Danh sách các kích thước để thử (từ nhỏ đến lớn)
        size_options = [
            (int(grid_size * 8), int(grid_size * 2)),   # 8cm x 2cm (mặc định)
            (int(grid_size * 6), int(grid_size * 2)),   # 6cm x 2cm (nhỏ hơn)
            (int(grid_size * 5), int(grid_size * 2)),   # 5cm x 2cm
            (int(grid_size * 4), int(grid_size * 2)),   # 4cm x 2cm (rất nhỏ)
            (int(grid_size * 3), int(grid_size * 2)),   # 3cm x 2cm (tối thiểu)
        ]
        
        created_count = 0
        failed_fields = []
        
        # Vẽ từng field chưa được vẽ
        for field in undrawn_fields:
            rect = None
            
            # Thử các kích thước khác nhau để tìm vị trí phù hợp
            for width, height in size_options:
                rect = find_empty_position(width, height)
                if rect:
                    break
            
            if rect is None:
                # Không tìm được vị trí trống, đặt ở cuối canvas (dưới header)
                # Tìm vị trí y thấp nhất và snap về lưới
                last_y = max(margin, header_height_snapped + margin)
                for existing_rect, _, _ in self.canvas.rects:
                    if existing_rect.bottom() > last_y:
                        last_y = existing_rect.bottom()
                
                # Snap last_y về lưới và thêm margin
                last_y_snapped = snap_to_grid(last_y + margin)
                
                rect = QRect(margin, last_y_snapped, default_width, default_height)
                
                # Kiểm tra nếu vượt quá canvas height
                if rect.bottom() > canvas_height:
                    failed_fields.append(field)
                    continue
            
            # Thêm vào canvas
            color = QColor(0, 100, 255)
            self.canvas.rects.append((rect, color, field))
            
            # Cập nhật occupied_cells
            cells = self.canvas.get_cells_in_rect(rect)
            self.canvas.occupied_cells.update(cells)
            self.canvas.used_fields.add(field)
            
            # Đánh dấu field đã được vẽ trong list
            self.mark_field_used(field)
            created_count += 1
        
        # Cập nhật canvas
        self.canvas.update()
        
        # Thông báo kết quả
        if failed_fields:
            QMessageBox.warning(
                self, "Hoàn tất (có lỗi)",
                f"Đã tự động tạo {created_count}/{len(undrawn_fields)} trường!\n\n"
                f"Không thể tạo {len(failed_fields)} trường do hết không gian:\n" +
                "\n".join(failed_fields[:5]) + 
                (f"\n... và {len(failed_fields) - 5} trường khác" if len(failed_fields) > 5 else "")
            )
        else:
            self.notification.show_message(f"Đã tự động tạo {created_count} trường vào vị trí trống!")

    def import_excel_file(self):
        self.popup.hide()

        file_path, _ = QFileDialog.getOpenFileName(self, "Chọn file Excel", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return
        
        # Lưu file_path để dùng sau
        self.current_import_file = file_path
        self.user_chose_overwrite = False  # Flag để tránh hỏi 2 lần
        
        # ✅ Kiểm tra xem file txt đã tồn tại chưa
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))
        
        file_name = os.path.basename(file_path).replace(".xlsx", "").replace(".xls", "")
        txt_path = os.path.join(app_dir, f"{file_name}.txt")
        txt_file_name = f"{file_name}.txt"
        
        # Nếu file txt đã tồn tại → hỏi người dùng
        if os.path.exists(txt_path):
            reply = QMessageBox.question(
                self, "File đã tồn tại",
                f"File {txt_file_name} đã tồn tại trong thư mục chương trình.\n\nBạn có muốn sử dụng file cũ không?\n\n• Yes: Sử dụng file cũ (giữ nguyên dữ liệu)\n• No: Gọi server để tạo file mới từ Excel",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                # Sử dụng file txt cũ
                self.txt_path = txt_path
                from sentence_manager import SentenceManager
                self.main_window.sm = SentenceManager()
                self.main_window.sm.load_from_txt(self.txt_path)
                self.main_window.current_file_path = self.txt_path
                
                # Load fields từ file txt
                with open(self.txt_path, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    if len(lines) > 1:
                        self.fields = lines[1].strip().split('\t')
                        self.load_fields_to_list()
                
                # Hiển thị nút Preview
                self.preview_button.setText(f"Preview {txt_file_name}")
                self.preview_button.show()
                return
            else:
                # Người dùng chọn No → đánh dấu đã chọn ghi đè
                self.user_chose_overwrite = True
        
        # Nếu chưa có file hoặc người dùng chọn No → gọi server
        # Disable nút Import và đổi text
        self.import_excel_button.setEnabled(False)
        self.import_excel_button.setText("Đang Import...")
        self.import_excel_button.setStyleSheet(
            "QPushButton {"
            " background-color: #aaaaaa; color: white; border: none;"
            " border-radius: 6px; padding: 4px 10px; font-size: 9pt;"
            "}"
        )
        
        # Tạo và chạy worker thread
        ip = self.ip_input.text().strip() or "107.98.33.94"
        self.import_worker = ImportWorker(file_path, ip)
        self.import_worker.finished.connect(self.on_import_success)
        self.import_worker.error.connect(self.on_import_error)
        self.import_worker.start()

    def on_import_success(self, result):
        """Callback khi import thành công"""
        try:
            file_path = self.current_import_file
            fields_raw = result['fields_raw']          # Thứ tự và tên cột gốc từ Excel (không strip)
            header_fields = result['fields']           # Header đã sanitize (3==D thay cho xuống dòng/tab)
            data = result['data']
            
            # Reset canvas và các trường cũ trước khi import file mới
            self.canvas.rects.clear()
            self.canvas.occupied_cells.clear()
            self.canvas.used_fields.clear()
            self.canvas.update()
            
            # In ra terminal để kiểm tra
            print("=== EXCEL IMPORT DEBUG ===")
            print(f"Fields raw ({len(fields_raw)}): {fields_raw}")
            print(f"Header fields ({len(header_fields)}): {header_fields}")
            print(f"Data rows: {len(data)}")
            if data:
                print("First 3 rows of data:")
                for i, row in enumerate(data[:3]):
                    print(f"Row {i+1}: {row}")
            print("==========================")

            # ✅ Chuẩn bị file txt - Lấy thư mục chứa file .exe hoặc script đang chạy
            if getattr(sys, 'frozen', False):
                # Nếu chạy từ file .exe (PyInstaller)
                app_dir = os.path.dirname(sys.executable)
            else:
                # Nếu chạy từ script Python
                app_dir = os.path.dirname(os.path.abspath(__file__))
            
            file_name = os.path.basename(file_path).replace(".xlsx", "").replace(".xls", "")
            self.txt_path = os.path.join(app_dir, f"{file_name}.txt")
            txt_file_name = f"{file_name}.txt"

            use_existing = False
            # Chỉ hỏi nếu file tồn tại VÀ người dùng chưa chọn ghi đè trước đó
            if os.path.exists(self.txt_path) and not getattr(self, 'user_chose_overwrite', False):
                reply = QMessageBox.question(
                    self, "File đã tồn tại",
                    f"File {txt_file_name} đã tồn tại trong thư mục chương trình.\n\nBạn có muốn sử dụng file cũ không?\n\n• Yes: Sử dụng file cũ (giữ nguyên dữ liệu)\n• No: Ghi đè bằng dữ liệu mới từ Excel",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.Yes:
                    use_existing = True

            if use_existing:
                from sentence_manager import SentenceManager
                self.main_window.sm = SentenceManager()
                self.main_window.sm.load_from_txt(self.txt_path)
                self.main_window.current_file_path = self.txt_path  # Lưu đường dẫn
            else:
                # ✅ Ghi file .txt mới theo đúng thứ tự cột gốc
                with open(self.txt_path, "w", encoding="utf-8") as f:
                    f.write("0\n")  # Dòng đầu tiên là index mặc định
                    f.write("\t".join(header_fields) + "\n")
                    for row in data:
                        values = []
                        for raw in fields_raw:
                            val = row.get(raw, "")
                            if isinstance(val, str):
                                val = val.strip().replace('\r\n', '3==D').replace('\n', '3==D').replace('\r', '3==D')
                            values.append(val)
                        # Thêm status "Not Done" vào cột cuối
                        values.append("Not Done")
                        f.write("\t".join(values) + "\n")
                
                # Load dữ liệu mới vào main_window.sm
                from sentence_manager import SentenceManager
                self.main_window.sm = SentenceManager()
                self.main_window.sm.load_from_txt(self.txt_path)
                self.main_window.current_file_path = self.txt_path  # Lưu đường dẫn

            # ✅ Cập nhật canvas & popup field theo header_fields để hiển thị đúng
            self.fields = header_fields
            self.canvas.fields = self.fields
            self.list_widget.clear()
            # Thêm items với 3==D được replace thành dấu cách
            for field in self.fields:
                display_name = field.replace("3==D", " ").strip()
                item = QListWidgetItem(display_name)
                # Lưu tên gốc mapping theo header (vì SentenceManager dùng header làm key)
                item.setData(Qt.UserRole, field)
                self.list_widget.addItem(item)

            eu.save_config(self.canvas.rects, fields=self.fields)

            for field in self.canvas.used_fields:
                self.mark_field_used(field)

            self.canvas.update()

            # ✅ Hiện nút Preview với tên file
            self.preview_button.setText(f"Preview {txt_file_name}")
            self.preview_button.show()

            # ✅ Gọi cập nhật Trang chính
            self.main_window.on_done(self.canvas.rects, stay_on_current_tab=True)
            
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Lỗi khi xử lý dữ liệu:\n{e}")
        
        finally:
            # Restore nút Import
            self.restore_import_button()

    def on_import_error(self, error_message):
        """Callback khi import gặp lỗi"""
        QMessageBox.critical(self, "Lỗi", f"Lỗi khi tải file Excel:\n{error_message}")
        self.restore_import_button()

    def restore_import_button(self):
        """Khôi phục trạng thái nút Import Excel"""
        self.import_excel_button.setEnabled(True)
        self.import_excel_button.setText("Import Excel")
        self.import_excel_button.setStyleSheet(
            "QPushButton {"
            " background-color: #2d8cff; color: white; border: none;"
            " border-radius: 6px; padding: 4px 10px; font-size: 9pt; font-weight: bold;"
            "}"
            "QPushButton:hover { background-color: #1f7ae0; }"
            "QPushButton:pressed { background-color: #1667bf; }"
        )

    def import_txt_file(self):
        """Import file TXT đã được convert (không cần gọi server)"""
        self.popup.hide()
        
        file_path, _ = QFileDialog.getOpenFileName(self, "Chọn file TXT", "", "Text Files (*.txt)")
        if not file_path:
            return
        
        try:
            # Lấy thư mục chứa chương trình
            if getattr(sys, 'frozen', False):
                app_dir = os.path.dirname(sys.executable)
            else:
                app_dir = os.path.dirname(os.path.abspath(__file__))
            
            # Đọc file txt để lấy fields
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                if len(lines) < 2:
                    raise Exception("File TXT không đúng định dạng!")
                
                # Dòng 2 là header (fields)
                header_line = lines[1].strip()
                header_fields = header_line.split('\t')
            
            # Copy file vào thư mục chương trình với tên gốc
            file_name = os.path.basename(file_path)
            self.txt_path = os.path.join(app_dir, file_name)
            
            # Nếu file đã tồn tại, hỏi có muốn ghi đè không
            if os.path.exists(self.txt_path) and os.path.abspath(file_path) != os.path.abspath(self.txt_path):
                reply = QMessageBox.question(
                    self, "File đã tồn tại",
                    f"File {file_name} đã tồn tại trong thư mục chương trình.\n\nBạn có muốn ghi đè không?",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.No:
                    return
            
            # Copy file nếu khác thư mục
            if os.path.abspath(file_path) != os.path.abspath(self.txt_path):
                import shutil
                shutil.copy2(file_path, self.txt_path)
            
            # Cập nhật fields và UI
            self.fields = header_fields
            self.canvas.fields = self.fields
            self.list_widget.clear()
            
            for field in self.fields:
                display_name = field.replace("3==D", " ").strip()
                item = QListWidgetItem(display_name)
                item.setData(Qt.UserRole, field)
                self.list_widget.addItem(item)
            
            # Lưu config
            eu.save_config(self.canvas.rects, fields=self.fields)
            
            # Update màu cho các field đã vẽ
            for field in self.canvas.used_fields:
                self.mark_field_used(field)
            
            self.canvas.update()
            
            # Hiện nút Preview
            self.preview_button.setText(f"Preview {file_name}")
            self.preview_button.show()
            
            # Load dữ liệu vào main window
            from sentence_manager import SentenceManager
            self.main_window.sm = SentenceManager()
            self.main_window.sm.load_from_txt(self.txt_path)
            self.main_window.current_file_path = self.txt_path
            
            # Cập nhật Trang chính
            self.main_window.on_done(self.canvas.rects, stay_on_current_tab=True)
            
            self.notification.show_message(f"Đã import file TXT thành công! File: {file_name}")
            
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Lỗi khi import file TXT:\n{e}")

    def preview_txt_file(self):
        """Mở file txt cho người dùng xem"""
        if not self.txt_path or not os.path.exists(self.txt_path):
            QMessageBox.warning(self, "Không tìm thấy file", "Chưa có file txt để xem!")
            return
        
        try:
            # Mở file txt bằng chương trình mặc định của hệ thống
            if sys.platform == 'win32':
                os.startfile(self.txt_path)
            elif sys.platform == 'darwin':  # macOS
                os.system(f'open "{self.txt_path}"')
            else:  # Linux
                os.system(f'xdg-open "{self.txt_path}"')
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể mở file:\n{e}")

    def reset_all(self):
        """Reset toàn bộ: xóa list, các vùng vẽ"""
        reply = QMessageBox.question(
            self, "Xác nhận Reset",
            "Bạn có chắc chắn muốn xóa toàn bộ:\n\n• Danh sách các trường\n• Tất cả các vùng đã vẽ\n• Cấu hình đã lưu\n\nHành động này không thể hoàn tác!",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # Xóa tất cả các vùng vẽ
            self.canvas.rects.clear()
            self.canvas.occupied_cells.clear()
            self.canvas.used_fields.clear()
            self.canvas.update()
            
            # Xóa list widget
            self.list_widget.clear()
            
            # Reset fields
            self.fields = []
            self.canvas.fields = []
            
            # Reset đường dẫn file txt
            self.txt_path = ""
            
            # Ẩn nút Preview
            self.preview_button.hide()
            
            # Xóa config đã lưu
            eu.save_config([], fields=[])
            
            # Ẩn popup nếu đang hiển thị
            self.popup.hide()
            
            self.notification.show_message("Đã reset toàn bộ thành công!")


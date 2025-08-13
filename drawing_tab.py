import os

from PySide6.QtWidgets import (
    QWidget, QHBoxLayout, QPushButton, QVBoxLayout, QListWidget,
    QListWidgetItem, QMenu, QFrame, QScrollArea, QFileDialog, QMessageBox, QLineEdit, QLabel
)
import requests
from PySide6.QtGui import QColor
from PySide6.QtCore import Qt, QPoint
from back_end import eu
from grid_canvas import GridCanvas  # nếu bạn để GridCanvas ở file khác

class DrawingTab(QWidget):
    def __init__(self, fields, main_window):
        super().__init__()
        self.main_window = main_window
        self.fields = fields
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.txt_path = ""

        # Canvas vẽ
        self.canvas = GridCanvas(fields, self.mark_field_used)
        self.main_layout.addWidget(self.canvas)

        # Cửa sổ overlay (nút nằm phía trên canvas)
        self.overlay = QWidget(self)
        self.overlay.setGeometry(10, 10, 520, 72)
        self.overlay.setStyleSheet("background-color: rgba(255,255,255,0.8); border: 1px solid gray;")
        self.overlay.setAttribute(Qt.WA_TransparentForMouseEvents, False)

        overlay_layout = QVBoxLayout(self.overlay)
        overlay_layout.setContentsMargins(8, 6, 8, 6)
        overlay_layout.setSpacing(6)

        # Dòng trên: Label + Ô nhập IP server API
        ip_row = QHBoxLayout()
        ip_row.setSpacing(6)
        self.ip_label = QLabel("IP server", self)
        self.ip_label.setStyleSheet("QLabel { background-color: #f2f6ff; border-radius: 8px; padding: 4px 8px; color: #1f3b75; }")
        self.ip_input = QLineEdit(self)
        self.ip_input.setPlaceholderText("IP API")
        self.ip_input.setText("107.98.33.94")
        self.ip_input.setFixedWidth(180)
        ip_row.addWidget(self.ip_label)
        ip_row.addWidget(self.ip_input)
        ip_row.addStretch(1)
        overlay_layout.addLayout(ip_row)

        # Dòng dưới: các nút chức năng
        button_row = QHBoxLayout()
        button_row.setSpacing(6)
        self.toggle_button = QPushButton("List")
        self.toggle_button.clicked.connect(self.toggle_field_popup)
        button_row.addWidget(self.toggle_button)

        self.import_button = QPushButton("Import")
        self.import_button.clicked.connect(self.import_excel_file)
        button_row.addWidget(self.import_button)

        self.done_button = QPushButton("Hoàn tất")
        self.done_button.clicked.connect(self.handle_done_clicked)
        button_row.addWidget(self.done_button)
        button_row.addStretch(1)
        overlay_layout.addLayout(button_row)

        # Bo tròn đẹp cho các nút
        rounded_button_style = (
            "QPushButton {"
            " background-color: #2d8cff; color: white; border: none;"
            " border-radius: 12px; padding: 6px 12px;"
            "}"
            "QPushButton:hover { background-color: #1f7ae0; }"
            "QPushButton:pressed { background-color: #1667bf; }"
        )
        for btn in [self.toggle_button, self.import_button, self.done_button]:
            btn.setStyleSheet(rounded_button_style)

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

    def toggle_field_popup(self):
        if self.popup.isVisible():
            self.popup.hide()
        else:
            # Tính vị trí popup nằm dưới button
            button_pos = self.toggle_button.mapToGlobal(QPoint(0, self.toggle_button.height()))
            self.popup.move(button_pos)
            self.popup.show()

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

    def import_excel_file(self):
        self.popup.hide()

        file_path, _ = QFileDialog.getOpenFileName(self, "Chọn file Excel", "", "Excel Files (*.xlsx *.xls)")
        if not file_path:
            return

        try:
            # Gửi file Excel lên server
            ip = self.ip_input.text().strip() or "107.98.33.94"
            url = f"http://{ip}:5000/upload_excel"
            with open(file_path, 'rb') as f:
                response = requests.post(url, files={'file': f})
            if response.status_code != 200:
                raise Exception(response.json().get('error', 'Unknown error'))

            result = response.json()
            fields_raw = result['fields_raw']          # Thứ tự và tên cột gốc từ Excel (không strip)
            header_fields = result['fields']           # Header đã sanitize (3==D thay cho xuống dòng/tab)
            data = result['data']
            
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

            # ✅ Chuẩn bị file txt
            file_name = os.path.basename(file_path).replace(".xlsx", "").replace(".xls", "")
            self.txt_path = os.path.join(os.path.dirname(__file__), f"{file_name}.txt")

            use_existing = False
            if os.path.exists(self.txt_path):
                reply = QMessageBox.question(
                    self, "File đã tồn tại",
                    "File data.txt đã tồn tại. Bạn có muốn tiếp tục làm không?\nNhấn No để ghi đè lên file cũ",
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
                        f.write("\t".join(values) + "\n")

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

            # ✅ Gọi cập nhật Trang chính
            self.main_window.on_done(self.canvas.rects, stay_on_current_tab=True)

        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Lỗi khi tải file Excel:\n{e}")

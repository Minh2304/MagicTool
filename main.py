from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout,
    QLabel, QLineEdit, QTextEdit, QFrame, QPushButton, QFileDialog,
    QHBoxLayout, QMessageBox, QCheckBox, QComboBox, QGraphicsOpacityEffect
)
from PySide6.QtCore import Qt, QTimer, QEvent, QPropertyAnimation, QEasingCurve
from PySide6.QtGui import QColor, QKeySequence, QTextCursor
from drawing_tab import DrawingTab
from back_end import eu
from sentence_manager import SentenceManager


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


class CustomTextEdit(QTextEdit):
    """Custom QTextEdit để xử lý bỏ selection khi focus vào ô khác"""
    def __init__(self, parent=None, main_window=None):
        super().__init__(parent)
        self.main_window = main_window
    
    def focusInEvent(self, event):
        """Khi focus vào ô này, bỏ selection của tất cả các ô khác"""
        super().focusInEvent(event)
        if self.main_window and hasattr(self.main_window, 'field_widgets'):
            for widget in self.main_window.field_widgets.values():
                if widget != self:
                    # Bỏ selection bằng cách đặt cursor về vị trí hiện tại
                    cursor = widget.textCursor()
                    cursor.clearSelection()
                    widget.setTextCursor(cursor)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Review Text Tool")
        self.sm = SentenceManager()  # Quản lý câu
        self.current_file_path = None  # Lưu đường dẫn file hiện tại

        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # Tab 1: Trang chính
        self.tab1 = QWidget()
        self.tab1.setObjectName("MainCanvas")

        # Header (vùng đóng băng 3 dòng đầu)
        self.header = QFrame(self.tab1)
        self.header.setStyleSheet("background-color: rgb(180, 180, 180); border: 1px solid gray;")
        self.header.setGeometry(0, 0, 1000, 0)  # sẽ chỉnh chiều cao sau

        self.header_layout = QHBoxLayout(self.header)
        self.header_layout.setContentsMargins(10, 5, 10, 5)
        self.header_layout.setSpacing(10)

        # Chia thành 3 phần với viền đẹp
        # Phần 1: STT ở hàng 1, 3 nút ở hàng 2
        left_frame = QFrame()
        left_frame.setStyleSheet(
            "QFrame { "
            "background-color: rgba(255, 255, 255, 0.7); "
            "border: 2px solid #2d8cff; "
            "border-radius: 8px; "
            "padding: 5px; "
            "}"
        )
        left_section = QVBoxLayout(left_frame)  # Đổi sang VBoxLayout
        left_section.setContentsMargins(8, 5, 8, 5)
        left_section.setSpacing(8)
        
        # Hàng 1: STT và ô input (căn lề trái)
        stt_row = QHBoxLayout()
        stt_row.setSpacing(8)
        
        stt_label = QLabel("STT:")
        stt_label.setStyleSheet(
            "QLabel { "
            "color: #1f3b75; "
            "font-weight: bold; "
            "font-size: 11pt; "
            "border: none; "
            "background: transparent; "
            "}"
        )
        stt_row.addWidget(stt_label)
        
        self.stt_input = QLineEdit()
        self.stt_input.setFixedSize(60, 30)
        self.stt_input.setAlignment(Qt.AlignCenter)
        self.stt_input.setStyleSheet(
            "QLineEdit { "
            "padding: 4px; "
            "border: 1px solid #2d8cff; "
            "border-radius: 4px; "
            "font-size: 11pt; "
            "font-weight: bold; "
            "background-color: white; "
            "}"
            "QLineEdit:focus { "
            "border: 2px solid #2d8cff; "
            "}"
        )
        self.stt_input.returnPressed.connect(self.jump_to_sentence)
        stt_row.addWidget(self.stt_input)
        
        # Label hiển thị tổng số câu
        self.stt_total_label = QLabel("/ 0")
        self.stt_total_label.setStyleSheet(
            "QLabel { "
            "font-size: 11pt; "
            "font-weight: bold; "
            "color: #666666; "
            "background: transparent; "
            "border: none; "
            "}"
        )
        stt_row.addWidget(self.stt_total_label)
        
        stt_row.addStretch()
        left_section.addLayout(stt_row)
        
        # Hàng 2: 3 nút Trước, Sau, Xuất Excel
        buttons_row = QHBoxLayout()
        buttons_row.setSpacing(10)
        
        self.prev_btn = QPushButton("← Trước")
        self.prev_btn.setFixedSize(100, 32)
        self.prev_btn.clicked.connect(self.prev_sentence)
        buttons_row.addWidget(self.prev_btn)
        
        self.next_btn = QPushButton("Sau →")
        self.next_btn.setFixedSize(100, 32)
        self.next_btn.clicked.connect(self.next_sentence)
        buttons_row.addWidget(self.next_btn)
        
        self.save_btn = QPushButton("Xuất Excel")
        self.save_btn.setFixedSize(100, 32)
        self.save_btn.clicked.connect(self.export_excel)
        buttons_row.addWidget(self.save_btn)
        
        buttons_row.addStretch()
        left_section.addLayout(buttons_row)
        
        # Phần 2: Checkbox và text - có nền như phần 1
        center_frame = QFrame()
        center_frame.setStyleSheet(
            "QFrame { "
            "background-color: rgba(255, 255, 255, 0.7); "
            "border: 2px solid #2d8cff; "
            "border-radius: 8px; "
            "padding: 5px; "
            "}"
        )
        center_section = QVBoxLayout(center_frame)
        center_section.setContentsMargins(10, 8, 10, 8)
        center_section.setSpacing(5)
        center_section.addStretch()
        
        # Enter để Next - ô tick cùng hàng với text
        checkbox_container = QHBoxLayout()
        checkbox_container.setSpacing(10)
        
        # Thêm spacing bên trái
        checkbox_container.addSpacing(30)
        
        # Container chứa checkbox và label "Enter để Next" trên cùng một hàng
        first_line = QHBoxLayout()
        first_line.setSpacing(8)
        
        # Tạo checkbox với dấu tick đẹp bằng Unicode
        self.enter_to_next_checkbox = QCheckBox("✓")
        self.enter_to_next_checkbox.setFixedSize(24, 24)  # Cố định kích thước để không bị cắt
        self.enter_to_next_checkbox.setStyleSheet(
            "QCheckBox { "
            "spacing: -22px; "  # Đưa text vào trong checkbox
            "color: transparent; "  # Ẩn text khi chưa check
            "padding: 0px; "
            "margin: 0px; "
            "}"
            "QCheckBox:checked { "
            "color: white; "  # Hiện text màu trắng khi checked
            "font-size: 14pt; "
            "font-weight: bold; "
            "}"
            "QCheckBox::indicator { "
            "width: 22px; "
            "height: 22px; "
            "border: 2px solid #2d8cff; "
            "border-radius: 4px; "
            "background-color: white; "
            "}"
            "QCheckBox::indicator:checked { "
            "background-color: #2d8cff; "
            "border: 2px solid #2d8cff; "
            "}"
            "QCheckBox::indicator:hover { "
            "border: 2px solid #1f7ae0; "
            "background-color: rgba(45, 140, 255, 0.1); "
            "}"
        )
        self.enter_to_next_checkbox.setChecked(False)
        first_line.addWidget(self.enter_to_next_checkbox, 0, Qt.AlignVCenter)
        
        # Label "Enter để Next" cùng hàng với checkbox
        enter_label = QLabel("Enter để Next")
        enter_label.setStyleSheet(
            "QLabel { "
            "color: #1f3b75; "
            "font-weight: bold; "
            "font-size: 11pt; "
            "border: none; "
            "background: transparent; "
            "}"
        )
        first_line.addWidget(enter_label, 0, Qt.AlignVCenter)
        
        # Container chính chứa 2 dòng
        text_container = QVBoxLayout()
        text_container.setSpacing(5)
        text_container.addLayout(first_line)
        
        # Label "F4 để back" ở dòng thứ 2
        f4_label = QLabel("F4 để back")
        f4_label.setStyleSheet(
            "QLabel { "
            "color: #1f3b75; "
            "font-weight: bold; "
            "font-size: 11pt; "
            "border: none; "
            "background: transparent; "
            "margin-left: 32px; "  # Thụt lề để thẳng hàng với "Enter để Next"
            "}"
        )
        text_container.addWidget(f4_label)
        
        checkbox_container.addLayout(text_container)
        checkbox_container.addStretch()
        
        center_section.addLayout(checkbox_container)
        center_section.addStretch()
        
        # Phần 3: Filter
        right_frame = QFrame()
        right_frame.setStyleSheet(
            "QFrame { "
            "background-color: rgba(255, 255, 255, 0.7); "
            "border: 2px solid #2d8cff; "
            "border-radius: 8px; "
            "padding: 5px; "
            "}"
        )
        right_section = QVBoxLayout(right_frame)
        right_section.setContentsMargins(10, 5, 10, 5)
        right_section.setSpacing(8)
        right_section.addStretch()
        
        # Label "Filter"
        filter_label = QLabel("Filter:")
        filter_label.setStyleSheet(
            "QLabel { "
            "color: #1f3b75; "
            "font-weight: bold; "
            "font-size: 11pt; "
            "border: none; "
            "background: transparent; "
            "}"
        )
        filter_label.setAlignment(Qt.AlignCenter)
        right_section.addWidget(filter_label)
        
        # Hàng chứa ComboBox và nút Filter (ngang)
        filter_row = QHBoxLayout()
        filter_row.setSpacing(8)
        
        # ComboBox với 3 lựa chọn (bên trái)
        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["All", "Not Done", "Done"])
        self.filter_combo.setFixedHeight(30)
        self.filter_combo.setStyleSheet(
            "QComboBox { "
            "padding: 4px 8px; "
            "border: 1px solid #2d8cff; "
            "border-radius: 4px; "
            "font-size: 10pt; "
            "background-color: white; "
            "}"
            "QComboBox:hover { "
            "border: 2px solid #2d8cff; "
            "}"
            "QComboBox::drop-down { "
            "border: none; "
            "width: 20px; "
            "}"
        )
        filter_row.addWidget(self.filter_combo)
        
        # Nút Filter (bên phải)
        self.filter_btn = QPushButton("Filter")
        self.filter_btn.setFixedSize(70, 30)
        self.filter_btn.clicked.connect(self.apply_filter)
        self.filter_btn.setStyleSheet(
            "QPushButton { "
            "background-color: #2d8cff; "
            "color: white; "
            "border: none; "
            "border-radius: 6px; "
            "padding: 4px 10px; "
            "font-size: 10pt; "
            "font-weight: bold; "
            "}"
            "QPushButton:hover { "
            "background-color: #1f7ae0; "
            "}"
            "QPushButton:pressed { "
            "background-color: #1667bf; "
            "}"
        )
        filter_row.addWidget(self.filter_btn)
        
        right_section.addLayout(filter_row)
        right_section.addStretch()
        
        # Label thông báo lỗi STT - hiện bên dưới header khi cần
        self.stt_error_label = QLabel("")
        self.stt_error_label.setStyleSheet(
            "QLabel { "
            "color: red; "
            "font-size: 8pt; "
            "font-weight: bold; "
            "background-color: rgba(255, 200, 200, 0.8); "
            "padding: 2px 8px; "
            "border-radius: 3px; "
            "}"
        )
        self.stt_error_label.setAlignment(Qt.AlignCenter)
        self.stt_error_label.hide()
        
        # Thêm 3 frame vào header với tỷ lệ 1:1:1 (bằng nhau)
        self.header_layout.addWidget(left_frame, 1)
        self.header_layout.addWidget(center_frame, 1)
        self.header_layout.addWidget(right_frame, 1)

        self.header.setLayout(self.header_layout)
        
        # Đặt label lỗi STT vào tab1 (hiển thị bên dưới header khi cần)
        self.stt_error_label.setParent(self.tab1)
        self.stt_error_label.setGeometry(0, 0, 0, 0)  # Sẽ set lại trong update_stt_error_position

        # Font mặc định cho QTextEdit (đặt sớm để on_done có thể sử dụng)
        self.text_font_point_size = 11
        QApplication.instance().installEventFilter(self)

        # Tab 2: Vẽ vùng
        saved_rects, saved_fields = eu.load_config()
        fields = saved_fields or []

        # ✅ Khởi tạo SentenceManager rỗng để vẽ layout nếu chưa import
        self.sm = SentenceManager()
        self.sm.fields = fields
        self.sm.sentences = []  # Chưa có câu nào

        self.tab2 = DrawingTab(fields, self)

        if saved_rects:
            self.tab2.load_saved_rects(saved_rects)
            self.on_done([(rect, QColor(0, 100, 255), field) for rect, field in saved_rects])

        self.tabs.addTab(self.tab1, "Trang chính")
        self.tabs.addTab(self.tab2, "Vẽ vùng")
        
        # Tạo notification widget (đặt trên tab1)
        self.notification = NotificationWidget(self.tab1)
        
        # Đảm bảo header giãn chiều rộng đúng sau khi hiển thị
        QTimer.singleShot(0, self.update_header_width)

        # Style nút bo tròn tại Trang chính
        rounded_button_style = (
            "QPushButton {"
            " background-color: #2d8cff; color: white; border: none;"
            " border-radius: 12px; padding: 6px 12px;"
            "}"
            "QPushButton:hover { background-color: #1f7ae0; }"
            "QPushButton:pressed { background-color: #1667bf; }"
        )
        for btn in [self.prev_btn, self.next_btn, self.save_btn]:
            btn.setStyleSheet(rounded_button_style)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'header') and self.header:
            grid_size = self.logicalDpiX() // 2.54
            reserved_height = grid_size * 3
            self.header.setGeometry(0, 0, self.tab1.width(), reserved_height)
            
            # Cập nhật vị trí label lỗi STT (nằm ngay dưới header, căn phải)
            if hasattr(self, 'stt_error_label') and self.stt_error_label:
                try:
                    label_width = 250
                    label_height = 25
                    x = self.tab1.width() - label_width - 20  # Căn phải, cách lề 20px
                    y = reserved_height + 5  # Cách header 5px
                    self.stt_error_label.setGeometry(x, y, label_width, label_height)
                except RuntimeError:
                    # Label đã bị xóa, bỏ qua
                    pass

    def on_done(self, rects, stay_on_current_tab=False):
        # Nếu chưa có dữ liệu sentences nhưng đã có đường dẫn file, load lại
        if (not hasattr(self, 'sm')):
            self.sm = SentenceManager()
        
        # Luôn load lại từ file nếu có current_file_path để đảm bảo dữ liệu mới nhất
        if getattr(self, 'current_file_path', None):
            try:
                self.sm.load_from_txt(self.current_file_path)
                print(f"DEBUG: Loaded {len(self.sm.sentences)} sentences from {self.current_file_path}")
            except Exception as e:
                print(f"DEBUG: Failed to load TXT in on_done: {e}")
        
        # Xóa widgets cũ (trừ stt_error_label và notification)
        for child in self.tab1.findChildren(QTextEdit) + self.tab1.findChildren(QLabel):
            if child.parent() == self.tab1 and child != self.stt_error_label and child != getattr(self, 'notification', None):
                child.deleteLater()

        self.field_widgets = {}

        grid_size = self.logicalDpiX() // 2.54
        reserved_height = grid_size * 3
        self.header.setGeometry(0, 0, self.tab1.width(), reserved_height)

        current_sentence = None
        if hasattr(self, 'sm') and self.sm.sentences:
            current_sentence = self.sm.current()

        for rect, color, name in rects:
            value = ""
            if current_sentence:
                value = current_sentence.get(name).replace("3==D", "\n")
            print(f"DEBUG: Field '{name}' = '{value}'")
            print(f"DEBUG: Available fields: {list(current_sentence.fields.keys()) if current_sentence else 'No sentence'}")

            input_box = CustomTextEdit(self.tab1, main_window=self)
            input_box.setGeometry(rect)
            input_box.setPlainText(value)
            input_box.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            input_box.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            input_box.setLineWrapMode(QTextEdit.WidgetWidth)
            # Thêm viền rõ nét cho các khung
            input_box.setStyleSheet(
                "QTextEdit { "
                "padding: 4px; "
                "border: 2px solid #0064ff; "  # Viền xanh dương rõ nét
                "border-radius: 3px; "
                "}"
            )
            # Đặt font Times New Roman và cỡ chữ hiện tại
            font = input_box.font()
            font.setFamily("Times New Roman")
            font.setPointSize(self.text_font_point_size)
            input_box.setFont(font)
            input_box.show()

            label = QLabel(name.replace("3==D", " "), self.tab1)
            label.move(rect.x() + 4, rect.y() - 18)
            label.setStyleSheet("QLabel { font-size: 10pt; color: #444444; }")
            label.adjustSize()
            label.show()

            self.field_widgets[name] = input_box

        if not stay_on_current_tab:
            self.tabs.setCurrentWidget(self.tab1)
        
        # Cập nhật STT sau khi load xong
        self.update_stt_display()

    def import_data(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Chọn file TXT", "", "Text Files (*.txt)")
        if file_path:
            self.current_file_path = file_path  # Lưu đường dẫn file
            self.sm.load_from_txt(file_path)
            self.show_sentence()

    def show_sentence(self):
        if not hasattr(self, 'field_widgets'):
            return
        sentence = self.sm.current()
        if sentence is None:
            return
        for field, widget in self.field_widgets.items():
            value = sentence.get(field).replace("3==D", "\n")  # Convert 3==D thành xuống dòng khi hiển thị
            widget.setPlainText(value)
        
        # Cập nhật STT (index + 1 vì bắt đầu từ 1)
        self.update_stt_display()

    def update_stt_display(self):
        """Cập nhật số thứ tự hiển thị trong ô STT và tổng số câu"""
        if hasattr(self, 'stt_input'):
            if self.sm.sentences:
                current_index = self.sm.current_index + 1  # Chuyển từ index 0 sang số thứ tự 1
                total_sentences = len(self.sm.sentences)
                self.stt_input.setText(str(current_index))
                
                # Cập nhật label tổng số câu
                if hasattr(self, 'stt_total_label'):
                    self.stt_total_label.setText(f"/ {total_sentences}")
            else:
                # Không có dữ liệu, hiển thị 0
                self.stt_input.setText("0")
                if hasattr(self, 'stt_total_label'):
                    self.stt_total_label.setText("/ 0")
            
            # Ẩn thông báo lỗi nếu có
            if hasattr(self, 'stt_error_label'):
                self.stt_error_label.hide()

    def jump_to_sentence(self):
        """Nhảy đến câu có số thứ tự được nhập"""
        if not self.sm.sentences:
            return
        
        try:
            stt = int(self.stt_input.text().strip())
            total_sentences = len(self.sm.sentences)
            
            # Kiểm tra số thứ tự có hợp lệ không (từ 1 đến total)
            if stt < 1 or stt > total_sentences:
                self.stt_error_label.setText(f"Số thứ tự phải từ 1 đến {total_sentences}")
                self.stt_error_label.show()
                return
            
            # Ẩn thông báo lỗi nếu hợp lệ
            self.stt_error_label.hide()
            
            # Lưu câu hiện tại trước khi chuyển
            self.save_current_sentence()
            if self.current_file_path:
                self.sm.save_to_txt(self.current_file_path)
            
            # Chuyển đến câu mới (stt - 1 vì index bắt đầu từ 0)
            self.sm.current_index = stt - 1
            self.update_text_boxes()
            
        except ValueError:
            self.stt_error_label.setText("Vui lòng nhập số hợp lệ")
            self.stt_error_label.show()

    def apply_filter(self):
        """Lọc và hiển thị các câu theo filter được chọn"""
        if not self.sm.sentences and not self.sm.all_sentences:
            QMessageBox.warning(self, "Cảnh báo", "Chưa có dữ liệu để lọc!")
            return
        
        # Lưu câu hiện tại trước khi filter
        self.save_current_sentence()
        if self.current_file_path:
            self.sm.save_to_txt(self.current_file_path)
        
        filter_type = self.filter_combo.currentText()
        
        # Áp dụng filter
        self.sm.apply_filter(filter_type)
        
        # Kiểm tra nếu không có câu nào sau khi filter
        if not self.sm.sentences:
            self.show_notification(f"Không có câu nào có trạng thái '{filter_type}'!")
            # Quay về filter All
            self.filter_combo.setCurrentText("All")
            self.sm.apply_filter("All")
        
        # Cập nhật giao diện
        self.update_text_boxes()
        
        # Hiển thị thông báo
        total = len(self.sm.sentences)
        self.show_notification(f"Đã lọc: {filter_type} - Số câu: {total}")

    def save_current_sentence(self):
        if not hasattr(self, 'field_widgets'):
            return
        sentence = self.sm.current()
        if sentence is None:
            return
        for field, widget in self.field_widgets.items():
            text = widget.toPlainText().strip().replace("\n", "3==D")  # ✅ Chuyển ngược lại
            sentence.set(field, text)

    def next_sentence(self):
        # Kiểm tra nếu đang ở câu cuối cùng
        if self.sm.current_index == len(self.sm.sentences) - 1:
            # Lưu câu hiện tại và đánh dấu Done
            self.save_current_sentence()
            current_sentence = self.sm.current()
            if current_sentence:
                current_sentence.mark_as_done()
            if self.current_file_path:
                self.sm.save_to_txt(self.current_file_path)
            
            # Hiển thị thông báo
            reply = QMessageBox.question(
                self, "Câu cuối cùng",
                "Bạn đang ở câu cuối cùng!\n\nBạn có muốn xuất file Excel không?",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.export_excel()
            return
        
        # Nếu chưa phải câu cuối, tiếp tục bình thường
        self.save_current_sentence()
        # Đánh dấu câu hiện tại là Done trước khi chuyển
        current_sentence = self.sm.current()
        if current_sentence:
            current_sentence.mark_as_done()
        self.sm.next()
        if self.current_file_path:
            self.sm.save_to_txt(self.current_file_path)
        self.update_text_boxes()

    def prev_sentence(self):
        self.save_current_sentence()
        # Đánh dấu câu hiện tại là Done trước khi chuyển (back cũng tính là Done)
        current_sentence = self.sm.current()
        if current_sentence:
            current_sentence.mark_as_done()
        self.sm.previous()
        if self.current_file_path:
            self.sm.save_to_txt(self.current_file_path)
        self.update_text_boxes()

    def save_sentence(self):
        if hasattr(self, 'field_widgets') and hasattr(self, 'sm'):
            sentence = self.sm.current()
            for field, widget in self.field_widgets.items():
                text = widget.toPlainText().strip().replace("\n", "3==D")  # Chuyển \n thành 3==D khi lưu
                sentence.set(field, text)

    def update_text_boxes(self):
        sentence = self.sm.current()
        if sentence is None:
            return
        for field, widget in self.field_widgets.items():
            value = sentence.get(field).replace("3==D", "\n")  # Convert 3==D thành xuống dòng khi hiển thị
            widget.setPlainText(value)
        
        # Cập nhật STT
        self.update_stt_display()

    def update_header_width(self):
        if hasattr(self, 'header'):
            grid_size = self.logicalDpiX() // 2.54
            reserved_height = grid_size * 3
            self.header.setGeometry(0, 0, self.tab1.width(), reserved_height)

    def apply_text_font(self):
        if not hasattr(self, 'field_widgets'):
            return
        for widget in self.field_widgets.values():
            font = widget.font()
            font.setFamily("Times New Roman")
            font.setPointSize(self.text_font_point_size)
            widget.setFont(font)

    def eventFilter(self, obj, event):
        try:
            # Xử lý phím F4 -> Về câu trước
            if event.type() == QEvent.KeyPress and event.key() == Qt.Key_F4:
                self.prev_sentence()
                return True
            
            # Xử lý phím Enter
            if event.type() == QEvent.KeyPress and event.key() in (Qt.Key_Return, Qt.Key_Enter):
                # Nếu nhấn Shift+Enter -> Luôn xuống dòng (không xử lý gì)
                if QApplication.keyboardModifiers() & Qt.ShiftModifier:
                    return False  # Để Qt xử lý bình thường (xuống dòng)
                
                # Kiểm tra focus widget
                focus_widget = QApplication.focusWidget()
                
                # Nếu đang focus ở ô STT -> Để returnPressed xử lý (jump_to_sentence)
                if focus_widget == self.stt_input:
                    return False  # Để Qt xử lý signal returnPressed
                
                # Nếu checkbox được bật và đang ở tab Trang chính -> Next
                if self.enter_to_next_checkbox.isChecked() and self.tabs.currentWidget() == self.tab1:
                    # Kiểm tra focus không phải ở nút bấm
                    if not isinstance(focus_widget, QPushButton):
                        self.next_sentence()
                        return True
            
            # Xử lý Ctrl + Scroll để zoom text
            if event.type() == QEvent.Wheel and QApplication.keyboardModifiers() & Qt.ControlModifier:
                # Xác định widget đang focus là QTextEdit (viewport cũng được tính)
                focus_widget = QApplication.focusWidget()
                is_text_focus = isinstance(focus_widget, QTextEdit) or (
                    hasattr(self, 'field_widgets') and any(w.hasFocus() for w in getattr(self, 'field_widgets', {}).values())
                )
                if is_text_focus:
                    delta = event.angleDelta().y()
                    if delta > 0:
                        self.text_font_point_size = min(self.text_font_point_size + 1, 36)
                    elif delta < 0:
                        self.text_font_point_size = max(self.text_font_point_size - 1, 8)
                    self.apply_text_font()
                    return True
        except Exception as e:
            print(f"DEBUG: eventFilter error: {e}")
        return super().eventFilter(obj, event)

    def show_notification(self, message):
        """Hiển thị thông báo toast ở phía dưới màn hình"""
        if hasattr(self, 'notification'):
            self.notification.show_message(message)

    def closeEvent(self, event):
        try:
            # Đảm bảo đóng mọi popup/top-level widget còn mở
            try:
                if hasattr(self, 'tab2') and hasattr(self.tab2, 'popup') and self.tab2.popup is not None:
                    self.tab2.popup.hide()
                    self.tab2.popup.close()
            except Exception:
                pass
            for w in QApplication.topLevelWidgets():
                if w is not self:
                    try:
                        w.close()
                    except Exception:
                        pass
        finally:
            event.accept()
            QApplication.instance().quit()

    def export_excel(self):
        # Lưu thay đổi hiện tại vào bộ nhớ trước khi xuất
        try:
            self.save_current_sentence()
        except Exception as e:
            print(f"DEBUG: save_current_sentence error before export: {e}")

        if not getattr(self.sm, 'fields', []) or not getattr(self.sm, 'sentences', []):
            QMessageBox.warning(self, "Xuất Excel", "Không có dữ liệu để xuất.")
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "Chọn nơi lưu Excel", "", "Excel Files (*.xlsx)")
        if not file_path:
            return

        # Đảm bảo đuôi .xlsx
        if not file_path.lower().endswith('.xlsx'):
            file_path += '.xlsx'

        # Chuẩn bị dữ liệu theo đúng thứ tự cột như file đầu vào
        columns = [col.replace("3==D", "\n") for col in self.sm.fields]
        rows = []
        for sentence in self.sm.sentences:
            row = []
            for field in self.sm.fields:
                value = sentence.get(field)
                row.append(value.replace("3==D", "\n") if isinstance(value, str) else value)
            rows.append(row)

        # Ghi Excel bằng pandas; nếu thiếu thư viện thì báo lỗi rõ ràng
        try:
            import pandas as pd
            import numpy as np  # An toàn cho dữ liệu trống
            df = pd.DataFrame(rows, columns=columns)
            # Tránh NaN hiển thị lạ trong Excel
            df = df.replace({np.nan: ""})
            df.to_excel(file_path, index=False)
            self.show_notification("Xuất Excel thành công!")
        except ImportError:
            QMessageBox.critical(self, "Thiếu thư viện", "Thiếu pandas để xuất Excel. Vui lòng cài đặt:\n\npy -m pip install pandas openpyxl")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể xuất Excel:\n{e}")

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(1000, 700)
    window.show()
    sys.exit(app.exec())
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout,
    QLabel, QLineEdit, QTextEdit, QFrame, QPushButton, QFileDialog,
    QHBoxLayout
)
from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor
from drawing_tab import DrawingTab
from back_end import eu
from sentence_manager import SentenceManager

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Tool PySide6 với 2 tab")
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
        self.header_layout.setAlignment(Qt.AlignLeft)  # ✅ Căn trái

        # Nút ← Trước
        self.prev_btn = QPushButton("← Trước")
        self.prev_btn.setFixedSize(100, 32)
        self.prev_btn.clicked.connect(self.prev_sentence)
        self.header_layout.addWidget(self.prev_btn)

        # Nút Sau →
        self.next_btn = QPushButton("Sau →")
        self.next_btn.setFixedSize(100, 32)
        self.next_btn.clicked.connect(self.next_sentence)
        self.header_layout.addWidget(self.next_btn)

        # Nút Lưu 💾
        self.save_btn = QPushButton("💾 Lưu file")
        self.save_btn.setFixedSize(100, 32)
        self.save_btn.clicked.connect(self.save_sentence)
        self.header_layout.addWidget(self.save_btn)

        self.header.setLayout(self.header_layout)

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
        # Đảm bảo header giãn chiều rộng đúng sau khi hiển thị
        QTimer.singleShot(0, self.update_header_width)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, 'header') and self.header:
            grid_size = self.logicalDpiX() // 2.54
            reserved_height = grid_size * 3
            self.header.setGeometry(0, 0, self.tab1.width(), reserved_height)

    def on_done(self, rects, stay_on_current_tab=False):
        # Nếu chưa có dữ liệu sentences nhưng đã có đường dẫn file, load lại
        if (not hasattr(self, 'sm')):
            self.sm = SentenceManager()
        if getattr(self.sm, 'sentences', []) == [] and getattr(self, 'current_file_path', None):
            try:
                self.sm.load_from_txt(self.current_file_path)
            except Exception as e:
                print(f"DEBUG: Failed to load TXT in on_done: {e}")
        
        # Xóa widgets cũ
        for child in self.tab1.findChildren(QTextEdit) + self.tab1.findChildren(QLabel):
            if child.parent() == self.tab1:
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

            input_box = QTextEdit(self.tab1)
            input_box.setGeometry(rect)
            input_box.setPlainText(value)
            input_box.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            input_box.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            input_box.setLineWrapMode(QTextEdit.WidgetWidth)
            input_box.setStyleSheet("QTextEdit { font-size: 11pt; padding: 4px; }")
            input_box.show()

            label = QLabel(name.replace("3==D", " "), self.tab1)
            label.move(rect.x() + 4, rect.y() - 18)
            label.setStyleSheet("QLabel { font-size: 10pt; color: #444444; }")
            label.adjustSize()
            label.show()

            self.field_widgets[name] = input_box

        if not stay_on_current_tab:
            self.tabs.setCurrentWidget(self.tab1)

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
        self.save_current_sentence()
        self.sm.next()
        if self.current_file_path:
            self.sm.save_to_txt(self.current_file_path)
        self.update_text_boxes()

    def prev_sentence(self):
        self.save_current_sentence()
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

    def update_header_width(self):
        if hasattr(self, 'header'):
            grid_size = self.logicalDpiX() // 2.54
            reserved_height = grid_size * 3
            self.header.setGeometry(0, 0, self.tab1.width(), reserved_height)

if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = MainWindow()
    window.resize(1000, 700)
    window.show()
    sys.exit(app.exec())
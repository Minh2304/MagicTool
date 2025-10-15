import os
import sys
import json
from PySide6.QtCore import QRect

class EngineerUnderground:
    def __init__(self):
        self.main_ui = None
        
        # Lấy thư mục chứa file .exe hoặc script đang chạy
        if getattr(sys, 'frozen', False):
            # Nếu chạy từ file .exe (PyInstaller)
            app_dir = os.path.dirname(sys.executable)
        else:
            # Nếu chạy từ script Python
            app_dir = os.path.dirname(os.path.abspath(__file__))
        
        self.config_path = os.path.join(app_dir, "config.json")

    def save_config(self, rects, fields=None):
        data = {
            "rects": [],
            "fields": fields if fields is not None else []
        }
        for rect, color, field in rects:
            data["rects"].append({
                "field": field,
                "x": rect.x(),
                "y": rect.y(),
                "width": rect.width(),
                "height": rect.height()
            })
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def load_config(self):
        try:
            if not os.path.exists(self.config_path):
                return [], []

            with open(self.config_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            rects = []
            for item in data.get("rects", []):
                rect = QRect(item["x"], item["y"], item["width"], item["height"])
                field = item["field"]
                rects.append((rect, field))

            fields = data.get("fields", [])
            return rects, fields

        except Exception as e:
            print("Lỗi khi đọc config:", e)
            return [], []  # ✅ Luôn trả về đúng định dạng

    def read_txt(self, file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                first_line = f.readline()
                field_names = [field.strip() for field in first_line.strip().split("\t") if field.strip()]
                print("Danh sách trường từ file:", field_names)
                return field_names
        except Exception as e:
            print("Lỗi khi đọc file TXT:", e)
            return []
# Khởi tạo sẵn để sử dụng trong các file khác
eu = EngineerUnderground()



class Sentence:
    def __init__(self, field_names: list[str], values: list[str]):
        self.fields = dict(zip(field_names, values))

    def get(self, field: str) -> str:
        return self.fields.get(field, "")

    def set(self, field: str, value: str):
        self.fields[field] = value

    def to_list(self) -> list[str]:
        return [self.fields.get(key, "") for key in self.fields]
class SentenceManager:
    def __init__(self):
        self.sentences: list[Sentence] = []
        self.fields: list[str] = []
        self.current_index: int = 0

    def load_from_txt(self, file_path: str):
        self._last_loaded_path = file_path  # Lưu đường dẫn để sử dụng khi save
        print(f"DEBUG: Loading from {file_path}")
        with open(file_path, "r", encoding="utf-8") as f:
            # Không dùng strip() để không mất tab ở cuối (bảo toàn số cột)
            raw_lines = f.readlines()
            lines = [line.rstrip("\n\r") for line in raw_lines]
            print(f"DEBUG: Total lines read: {len(lines)}")
            if len(lines) < 2:
                print("DEBUG: Not enough lines, clearing data")
                self.fields = []
                self.sentences = []
                self.current_index = 0
                return

            self.fields = lines[1].split("\t")  # ✅ Dòng thứ 2 là danh sách field

            # Dòng 3 trở đi là data; đảm bảo đủ số cột bằng cách pad rỗng
            self.sentences = []
            for line in lines[2:]:
                values = line.split("\t")
                if len(values) < len(self.fields):
                    values += [""] * (len(self.fields) - len(values))
                elif len(values) > len(self.fields):
                    values = values[:len(self.fields)]
                self.sentences.append(Sentence(self.fields, values))

            self.current_index = int(lines[0]) if lines[0].isdigit() else 0  # ✅ đọc index từ dòng 1
            print(f"DEBUG: Loaded {len(self.sentences)} sentences, fields={len(self.fields)}, current_index={self.current_index}")
            print(f"DEBUG: Fields: {self.fields}")

    def save_to_txt(self, file_path: str = None):
        if file_path is None:
            # Nếu không có path, sử dụng path từ lần load cuối
            if not hasattr(self, '_last_loaded_path'):
                raise ValueError("Không có đường dẫn file để lưu. Vui lòng cung cấp file_path hoặc load file trước.")
            file_path = self._last_loaded_path
            
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(f"{self.current_index}\n")  # Dòng đầu là index
            f.write("\t".join(self.fields) + "\n")  # Dòng 2 là fields
            for sentence in self.sentences:
                row = []
                for field in self.fields:
                    # Không strip để giữ nguyên khoảng trắng người dùng nhập; chỉ thay \n bằng 3==D
                    value = sentence.get(field).replace("\n", "3==D")
                    row.append(value)
                f.write("\t".join(row) + "\n")

    def current(self) -> Sentence:
        print(f"DEBUG: current_index={self.current_index}, len(sentences)={len(self.sentences)}")
        if not self.sentences:
            print("DEBUG: No sentences available!")
            return None
        if self.current_index >= len(self.sentences):
            print(f"DEBUG: current_index {self.current_index} >= len(sentences) {len(self.sentences)}")
            self.current_index = len(self.sentences) - 1
        if self.current_index < 0:
            print(f"DEBUG: current_index {self.current_index} < 0")
            self.current_index = 0
        return self.sentences[self.current_index]

    def next(self):
        print(f"DEBUG: next() - current_index={self.current_index}, len(sentences)={len(self.sentences)}")
        if self.current_index < len(self.sentences) - 1:
            self.current_index += 1
            print(f"DEBUG: next() - new current_index={self.current_index}")
        else:
            print(f"DEBUG: next() - already at last sentence")

    def previous(self):
        print(f"DEBUG: previous() - current_index={self.current_index}, len(sentences)={len(self.sentences)}")
        if self.current_index > 0:
            self.current_index -= 1
            print(f"DEBUG: previous() - new current_index={self.current_index}")
        else:
            print(f"DEBUG: previous() - already at first sentence")
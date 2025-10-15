"""
Script tự động build file EXE từ main.py
Chạy file này để build chương trình thành file exe
"""
import os
import sys
import subprocess
import shutil

def check_pyinstaller():
    """Kiểm tra xem PyInstaller đã được cài đặt chưa"""
    try:
        import PyInstaller
        print("✓ PyInstaller đã được cài đặt")
        return True
    except ImportError:
        print("✗ PyInstaller chưa được cài đặt")
        return False

def install_pyinstaller():
    """Cài đặt PyInstaller"""
    print("\n📦 Đang cài đặt PyInstaller...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ Cài đặt PyInstaller thành công")
        return True
    except subprocess.CalledProcessError:
        print("✗ Không thể cài đặt PyInstaller")
        return False

def clean_build_folders():
    """Xóa các thư mục build và dist cũ"""
    folders_to_remove = ['build', 'dist']
    spec_files = [f for f in os.listdir('.') if f.endswith('.spec')]
    
    print("\n🧹 Dọn dẹp các file build cũ...")
    
    for folder in folders_to_remove:
        if os.path.exists(folder):
            try:
                shutil.rmtree(folder)
                print(f"  ✓ Đã xóa thư mục: {folder}")
            except Exception as e:
                print(f"  ⚠ Không thể xóa {folder}: {e}")
    
    for spec_file in spec_files:
        try:
            os.remove(spec_file)
            print(f"  ✓ Đã xóa file: {spec_file}")
        except Exception as e:
            print(f"  ⚠ Không thể xóa {spec_file}: {e}")

def build_exe():
    """Build file exe bằng PyInstaller"""
    print("\n🔨 Đang build file EXE...")
    print("=" * 60)
    
    # Các tham số cho PyInstaller
    # Sử dụng sys.executable để đảm bảo dùng đúng Python
    pyinstaller_args = [
        sys.executable,                # Python executable
        '-m', 'PyInstaller',           # Chạy PyInstaller như một module
        '--name=MagicTool',            # Tên file exe
        '--onefile',                   # Build thành 1 file duy nhất
        '--windowed',                  # Không hiện console (GUI app)
        '--clean',                     # Làm sạch cache trước khi build
        '--noconfirm',                 # Không hỏi xác nhận ghi đè
        # Thêm file dữ liệu nếu có
        '--add-data=Book1.txt;.',      # Thêm Book1.txt vào exe
        # Đảm bảo các module được import
        '--hidden-import=PySide6.QtCore',
        '--hidden-import=PySide6.QtWidgets',
        '--hidden-import=PySide6.QtGui',
        '--hidden-import=openpyxl',
        '--hidden-import=pandas',
        '--hidden-import=numpy',
        '--hidden-import=back_end',
        '--hidden-import=drawing_tab',
        '--hidden-import=grid_canvas',
        '--hidden-import=sentence_manager',
        # Loại bỏ các module không cần thiết để giảm dung lượng
        '--exclude-module=matplotlib',
        '--exclude-module=scipy',
        '--exclude-module=PIL',
        '--exclude-module=tkinter',
        'main.py'                      # File chính
    ]
    
    try:
        # Hiển thị lệnh đang chạy
        print(f"Lệnh: python -m PyInstaller --name=MagicTool --onefile --windowed ...\n")
        
        # Chạy PyInstaller
        result = subprocess.run(
            pyinstaller_args,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )
        
        # Hiển thị output trong quá trình build
        if result.stdout:
            # Chỉ hiển thị các dòng quan trọng
            for line in result.stdout.split('\n'):
                if any(keyword in line.lower() for keyword in ['warning', 'error', 'info:', 'building', 'completed']):
                    print(line)
        
        if result.returncode == 0:
            print("=" * 60)
            print("\n✅ BUILD THÀNH CÔNG!")
            
            # Kiểm tra file exe có tồn tại không
            exe_path = os.path.abspath('dist/MagicTool.exe')
            if os.path.exists(exe_path):
                file_size = os.path.getsize(exe_path) / (1024 * 1024)  # MB
                print(f"\n📁 File exe: {exe_path}")
                print(f"📊 Dung lượng: {file_size:.2f} MB")
                print("\n🎉 Bạn có thể chạy file exe từ thư mục 'dist'")
            else:
                print(f"\n⚠ Không tìm thấy file exe tại: {exe_path}")
            
            return True
        else:
            print("=" * 60)
            print("\n❌ BUILD THẤT BẠI!")
            print("\nLỗi:")
            if result.stderr:
                print(result.stderr)
            return False
            
    except Exception as e:
        print(f"\n❌ Lỗi khi build: {e}")
        print(f"Loại lỗi: {type(e).__name__}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Hàm chính"""
    print("=" * 60)
    print("🚀 MAGIC TOOL - AUTO BUILD EXE")
    print("=" * 60)
    
    # Kiểm tra đang ở đúng thư mục
    if not os.path.exists('main.py'):
        print("\n❌ Không tìm thấy file main.py!")
        print("Vui lòng chạy script này trong thư mục chứa main.py")
        input("\nNhấn Enter để thoát...")
        return
    
    # Kiểm tra và cài đặt PyInstaller nếu cần
    if not check_pyinstaller():
        response = input("\nBạn có muốn cài đặt PyInstaller không? (y/n): ")
        if response.lower() in ['y', 'yes', 'có', '']:
            if not install_pyinstaller():
                print("\n❌ Không thể tiếp tục build mà không có PyInstaller")
                input("\nNhấn Enter để thoát...")
                return
        else:
            print("\n❌ Cần PyInstaller để build exe")
            input("\nNhấn Enter để thoát...")
            return
    
    # Hỏi có muốn dọn dẹp không
    response = input("\nBạn có muốn xóa các file build cũ không? (y/n): ")
    if response.lower() in ['y', 'yes', 'có', '']:
        clean_build_folders()
    
    # Build exe
    print("\n⏳ Quá trình build có thể mất vài phút, vui lòng đợi...")
    success = build_exe()
    
    if success:
        print("\n" + "=" * 60)
        print("✨ HOÀN THÀNH!")
        print("=" * 60)
    else:
        print("\n" + "=" * 60)
        print("⚠ Build không thành công. Vui lòng kiểm tra lỗi ở trên.")
        print("=" * 60)
    
    input("\nNhấn Enter để thoát...")

if __name__ == "__main__":
    main()

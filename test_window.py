import sys
import traceback

try:
    print("Importing modules...")
    from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QPushButton
    from PyQt5.QtCore import Qt
    print("PyQt5 imported successfully")
    
    app = QApplication(sys.argv)
    print("QApplication created")
    
    window = QMainWindow()
    window.setWindowTitle('测试窗口')
    window.setGeometry(100, 100, 800, 600)
    print("MainWindow created")
    
    central = QWidget()
    layout = QVBoxLayout(central)
    
    label = QLabel('发票识别工具 - 测试窗口')
    label.setAlignment(Qt.AlignCenter)
    label.setStyleSheet('font-size: 24px;')
    layout.addWidget(label)
    
    btn = QPushButton('点击测试')
    btn.clicked.connect(lambda: print('Button clicked!'))
    layout.addWidget(btn)
    
    window.setCentralWidget(central)
    window.show()
    print("Window shown, starting event loop...")
    
    sys.exit(app.exec_())
    
except Exception as e:
    print(f"Error: {e}")
    traceback.print_exc()
    input("Press Enter to exit...")

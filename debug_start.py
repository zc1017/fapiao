import sys
import traceback

try:
    print("开始导入模块...")
    from PyQt5.QtWidgets import QApplication
    print("PyQt5导入成功")
    
    from invoice_app import InvoiceMainWindow
    print("InvoiceMainWindow导入成功")
    
    print("创建QApplication...")
    app = QApplication(sys.argv)
    print("QApplication创建成功")
    
    app.setStyle('Fusion')
    print("设置样式成功")
    
    from PyQt5.QtGui import QFont
    font = QFont('Microsoft YaHei', 9)
    app.setFont(font)
    print("设置字体成功")
    
    print("创建主窗口...")
    window = InvoiceMainWindow()
    print("主窗口创建成功")
    
    print("显示窗口...")
    window.show()
    print("窗口显示成功")
    
    print("启动事件循环...")
    sys.exit(app.exec_())
except Exception as e:
    print(f"错误: {e}")
    traceback.print_exc()
    input("按回车键退出...")

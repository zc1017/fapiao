import sys
import traceback

try:
    print("Step 1: Importing modules...")
    import os
    import re
    import io
    import zipfile
    import base64
    import tempfile
    from datetime import datetime
    print("Step 2: Standard modules imported")
    
    from PyQt5.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
        QHeaderView, QMessageBox, QLabel, QListWidget, QListWidgetItem,
        QSplitter, QProgressBar, QStatusBar
    )
    from PyQt5.QtCore import Qt, QThread, pyqtSignal
    from PyQt5.QtGui import QFont
    print("Step 3: PyQt5 imported")
    
    import cv2
    import numpy as np
    from pyzbar.pyzbar import decode, ZBarSymbol
    from defusedxml import ElementTree as ET
    import openpyxl
    from openpyxl.styles import Font as ExcelFont, Alignment, Border, Side
    from pdf2image import convert_from_path
    from PIL import Image
    print("Step 4: All modules imported")
    
    print("Step 5: Creating QApplication...")
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    print("Step 6: QApplication created")
    
    print("Step 7: Importing InvoiceMainWindow...")
    from invoice_app import InvoiceMainWindow
    print("Step 8: InvoiceMainWindow imported")
    
    print("Step 9: Creating main window...")
    window = InvoiceMainWindow()
    print("Step 10: Main window created")
    
    window.show()
    print("Step 11: Window shown, starting event loop...")
    
    ret = app.exec_()
    print(f"Step 12: Event loop ended with code {ret}")
    sys.exit(ret)
    
except Exception as e:
    print(f"ERROR: {e}")
    traceback.print_exc()
    input("Press Enter to exit...")

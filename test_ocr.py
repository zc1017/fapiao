import sys
sys.path.insert(0, r'c:\Users\zc\Desktop\fapiao')

from invoice_app import InvoiceParser, get_ocr
import os
import cv2
import numpy as np
import fitz

print("初始化OCR引擎...")
ocr = get_ocr()
if ocr is None:
    print("OCR初始化失败！")
    sys.exit(1)

print("OCR初始化成功！\n")

test_dir = r'c:\Users\zc\Desktop\fapiao'
files = [f for f in os.listdir(test_dir) if f.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg', '.bmp'))]

if not files:
    print("没有找到测试文件")
else:
    file = files[0]
    file_path = os.path.join(test_dir, file)
    print(f"测试文件: {file}")
    print('='*50)
    
    # 读取PDF并显示OCR识别的所有文字
    doc = fitz.open(file_path)
    page = doc.load_page(0)
    mat = fitz.Matrix(2.0, 2.0)
    pix = page.get_pixmap(matrix=mat)
    img_data = pix.tobytes("png")
    nparr = np.frombuffer(img_data, np.uint8)
    image = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
    doc.close()
    
    print("\nOCR识别所有文字:")
    result = ocr.readtext(image)
    
    for i, item in enumerate(result):
        if len(item) >= 2:
            text = item[1]
            confidence = item[2] if len(item) >= 3 else 0
            print(f"  [{i}] {text} (置信度: {confidence:.2f})")
    
    print("\n" + "="*50)
    print("测试销售方信息提取:")
    seller_info, error = InvoiceParser.ocr_extract_seller_info(file_path)
    if seller_info:
        print(f"  销售方名称: {seller_info.get('销售方名称', 'N/A')}")
        print(f"  销售方税号: {seller_info.get('销售方纳税人识别号', 'N/A')}")
    else:
        print(f"  提取失败: {error}")

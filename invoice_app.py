import sys
import os
import re
import io
import zipfile
import base64
import tempfile
import urllib.request
import urllib.error
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QLabel, QListWidget, QListWidgetItem,
    QSplitter, QProgressBar, QStatusBar, QDialog, QCheckBox, QScrollArea
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor
import cv2
import numpy as np
from pyzbar.pyzbar import decode, ZBarSymbol
from defusedxml import ElementTree as ET
import openpyxl
from openpyxl.styles import Font as ExcelFont, Alignment, Border, Side
import fitz
from PIL import Image

_ocr_instance = None

def get_ocr():
    global _ocr_instance
    if _ocr_instance is None:
        try:
            from paddleocr import PaddleOCR
            import paddle
            
            try:
                if paddle.is_compiled_with_cuda():
                    paddle.device.set_device('gpu')
                    print("检测到NVIDIA GPU，使用GPU加速")
                elif hasattr(paddle.device, 'is_compiled_with_rocm') and paddle.device.is_compiled_with_rocm():
                    paddle.device.set_device('gpu:gpu')
                    print("检测到AMD GPU，使用GPU加速")
                else:
                    print("使用CPU模式")
            except:
                print("使用CPU模式")
            
            _ocr_instance = PaddleOCR(lang='ch')
            print("PaddleOCR初始化成功")
        except Exception as e:
            print(f"PaddleOCR初始化失败: {e}")
            return None
    return _ocr_instance


class InvoiceParser:
    """发票解析器"""
    
    @staticmethod
    def fetch_invoice_xml_from_url(url):
        """从发票二维码URL获取XML数据"""
        try:
            if not url.startswith('http'):
                return None, '无效的URL'
            
            req = urllib.request.Request(
                url,
                headers={
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
                    'Accept': '*/*',
                    'Accept-Language': 'zh-CN,zh;q=0.9',
                }
            )
            
            with urllib.request.urlopen(req, timeout=10) as response:
                content = response.read()
                
                if content[:4] == b'PK\x03\x04':
                    with zipfile.ZipFile(io.BytesIO(content), 'r') as zf:
                        xml_files = [f for f in zf.namelist() if f.endswith('.xml')]
                        if xml_files:
                            xml_content = zf.read(xml_files[0])
                            try:
                                return xml_content.decode('utf-8'), None
                            except:
                                return xml_content.decode('gbk'), None
                
                try:
                    xml_content = content.decode('utf-8')
                except UnicodeDecodeError:
                    try:
                        xml_content = content.decode('gbk')
                    except:
                        xml_content = content.decode('utf-8', errors='ignore')
                
                if '<?xml' in xml_content.lower() or '<EInvoice' in xml_content:
                    return xml_content, None
                
                if re.match(r'^[A-Za-z0-9+/=]+$', xml_content.strip()) and len(xml_content.strip()) > 50:
                    decoded = InvoiceParser.decode_base64_to_xml(xml_content.strip())
                    if decoded:
                        return decoded, None
                
                return None, f'URL返回的数据格式无法识别: {xml_content[:100]}...'
                
        except urllib.error.HTTPError as e:
            return None, f'HTTP错误: {e.code}'
        except urllib.error.URLError as e:
            return None, f'网络错误: {str(e.reason)}'
        except Exception as e:
            return None, f'获取数据失败: {str(e)}'
    
    @staticmethod
    def ocr_extract_seller_info(image_path):
        """使用OCR提取销售方和购买方信息"""
        try:
            ocr = get_ocr()
            if ocr is None:
                return None, 'OCR未初始化'
            
            if image_path.lower().endswith('.pdf'):
                doc = fitz.open(image_path)
                page = doc.load_page(0)
                mat = fitz.Matrix(3.0, 3.0)
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("png")
                nparr = np.frombuffer(img_data, np.uint8)
                image = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
                doc.close()
            else:
                image = cv2.imread(image_path)
            
            if image is None:
                return None, '无法读取图片'
            
            def preprocess_image(img):
                """图像预处理提高OCR准确度"""
                if len(img.shape) == 3:
                    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                else:
                    gray = img.copy()
                
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
                enhanced = clahe.apply(gray)
                
                denoised = cv2.fastNlMeansDenoising(enhanced, None, h=10, searchWindowSize=21, templateWindowSize=7)
                
                _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
                
                kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
                sharpened = cv2.filter2D(binary, -1, kernel)
                
                return sharpened
            
            processed_image = preprocess_image(image)
            
            try:
                result = ocr.predict(processed_image)
            except:
                try:
                    result = ocr.ocr(processed_image)
                except:
                    result = None
            
            if not result:
                try:
                    result = ocr.predict(image)
                except:
                    try:
                        result = ocr.ocr(image)
                    except:
                        result = None
            
            if not result:
                return None, 'OCR未识别到文字'
            
            all_text = []
            if isinstance(result, dict) and 'rec_text' in result:
                texts = result.get('rec_text', [])
                for text in texts:
                    all_text.append((text, 1.0))
            elif isinstance(result, list) and len(result) > 0:
                for line in result[0] if isinstance(result[0], list) else result:
                    if line and len(line) >= 2:
                        if isinstance(line[1], tuple) and len(line[1]) >= 2:
                            text = line[1][0]
                            confidence = line[1][1]
                        else:
                            text = str(line[1])
                            confidence = 1.0
                        all_text.append((text, confidence))
            
            full_text = '\n'.join([t[0] for t in all_text])
            
            info = {}
            
            tax_id_pattern = r'[0-9A-Z]{15,20}|[0-9]{15,20}'
            tax_ids = re.findall(tax_id_pattern, full_text)
            
            lines = full_text.split('\n')
            
            seller_idx = -1
            buyer_idx = -1
            
            for i, line in enumerate(lines):
                if '销售方' in line or '收款方' in line or '开票方' in line:
                    seller_idx = i
                if '购买方' in line or '付款方' in line or '收票方' in line:
                    buyer_idx = i
            
            company_keywords = ['有限公司', '有限责任公司', '股份公司', '集团', '公司', '旅行社', '酒店', '餐厅', '超市', '商场', '银行', '保险', '医院', '学校', '大学', '研究所', '中心', '店', '部', '厂', '社', '院', '所', '站', '馆', '场', '分公司']
            
            def is_company_name(text):
                if len(text) < 3:
                    return False
                if re.match(r'^[0-9\-]+$', text):
                    return False
                if re.match(r'^[A-Z0-9]+$', text) and len(text) < 5:
                    return False
                for kw in company_keywords:
                    if kw in text:
                        return True
                if re.search(r'[省市县区].*[店社厂院]', text):
                    return True
                return False
            
            def clean_company_name(text):
                text = re.sub(r'^[名称：:\s]+', '', text)
                text = re.sub(r'[名称：:\s]+$', '', text)
                text = text.strip()
                
                ocr_corrections = {
                    '汊': '汉',
                    '卄': '廿',
                    '〇': '零',
                    '①': '一',
                    '②': '二',
                    '③': '三',
                    '④': '四',
                    '⑤': '五',
                    '⑥': '六',
                    '⑦': '七',
                    '⑧': '八',
                    '⑨': '九',
                    '㈠': '一',
                    '㈡': '二',
                    '㈢': '三',
                    '㈣': '四',
                    '㈤': '五',
                    '㈥': '六',
                    '㈦': '七',
                    '㈧': '八',
                    '㈨': '九',
                    '．': '.',
                    '－': '-',
                    '（': '(',
                    '）': ')',
                    '［': '[',
                    '］': ']',
                    '｛': '{',
                    '｝': '}',
                    '，': ',',
                    '：': ':',
                    '；': ';',
                }
                
                for wrong, correct in ocr_corrections.items():
                    text = text.replace(wrong, correct)
                
                return text
            
            names = []
            for i, line in enumerate(lines):
                line = line.strip()
                if is_company_name(line):
                    names.append((i, line))
                elif '名称' in line:
                    name_match = re.search(r'名称[：:\s]*(.+?)(?:\s|$|纳税人|地址|电话|统一)', line)
                    if name_match:
                        name = clean_company_name(name_match.group(1))
                        if len(name) > 2 and is_company_name(name):
                            names.append((i, name))
            
            if not names:
                for i, line in enumerate(lines):
                    line = line.strip()
                    if is_company_name(line):
                        names.append((i, line))
            
            company_names = [(idx, name) for idx, name in names if is_company_name(name)]
            
            if len(company_names) >= 2:
                info['销售方名称'] = company_names[-1][1]
                info['购买方名称'] = company_names[0][1]
            elif seller_idx >= 0 and buyer_idx >= 0:
                seller_names = [(idx, name) for idx, name in company_names if seller_idx < idx < buyer_idx]
                buyer_names = [(idx, name) for idx, name in company_names if idx > buyer_idx]
                
                if seller_names:
                    info['销售方名称'] = seller_names[0][1]
                if buyer_names:
                    info['购买方名称'] = buyer_names[0][1]
            elif seller_idx >= 0:
                for idx, name in company_names:
                    if idx > seller_idx and idx < seller_idx + 15:
                        info['销售方名称'] = name
                        break
                
                if '销售方名称' not in info:
                    for idx, name in names:
                        if idx > seller_idx and idx < seller_idx + 15:
                            info['销售方名称'] = name
                            break
            
            if '销售方名称' not in info and company_names:
                info['销售方名称'] = company_names[-1][1]
            
            if tax_ids:
                if seller_idx >= 0 and buyer_idx >= 0:
                    for tax_id in tax_ids:
                        for i, line in enumerate(lines):
                            if tax_id in line:
                                if abs(i - seller_idx) < abs(i - buyer_idx):
                                    info['销售方纳税人识别号'] = tax_id
                                else:
                                    info['购买方纳税人识别号'] = tax_id
                                break
                elif len(tax_ids) >= 2:
                    info['销售方纳税人识别号'] = tax_ids[-1]
                    info['购买方纳税人识别号'] = tax_ids[0]
                else:
                    info['销售方纳税人识别号'] = tax_ids[0]
            
            return info if info else None, None
            
        except Exception as e:
            return None, f'OCR识别错误: {str(e)}'
    
    @staticmethod
    def decode_qrcode(image_path):
        """识别图片或PDF中的二维码"""
        try:
            if image_path.lower().endswith('.pdf'):
                return InvoiceParser.decode_qrcode_from_pdf(image_path)
            
            image = cv2.imread(image_path)
            if image is None:
                return None, '无法读取图片文件'
            
            decoded_objects = decode(image, symbols=[ZBarSymbol.QRCODE])
            
            if not decoded_objects:
                gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
                decoded_objects = decode(gray, symbols=[ZBarSymbol.QRCODE])
            
            if not decoded_objects:
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
                enhanced = clahe.apply(gray if len(image.shape) == 2 else cv2.cvtColor(image, cv2.COLOR_BGR2GRAY))
                decoded_objects = decode(enhanced, symbols=[ZBarSymbol.QRCODE])
            
            if decoded_objects:
                return decoded_objects[0].data.decode('utf-8'), None
            return None, '未检测到二维码'
        except Exception as e:
            return None, f'二维码识别错误: {str(e)}'
    
    @staticmethod
    def decode_qrcode_from_pdf(pdf_path):
        """从PDF文件中识别二维码"""
        try:
            doc = fitz.open(pdf_path)
            
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                mat = fitz.Matrix(2.0, 2.0)
                pix = page.get_pixmap(matrix=mat)
                
                img_data = pix.tobytes("png")
                nparr = np.frombuffer(img_data, np.uint8)
                open_cv_image = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
                
                decoded_objects = decode(open_cv_image, symbols=[ZBarSymbol.QRCODE])
                
                if not decoded_objects:
                    gray = cv2.cvtColor(open_cv_image, cv2.COLOR_BGR2GRAY)
                    decoded_objects = decode(gray, symbols=[ZBarSymbol.QRCODE])
                
                if not decoded_objects:
                    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
                    enhanced = clahe.apply(gray)
                    decoded_objects = decode(enhanced, symbols=[ZBarSymbol.QRCODE])
                
                if decoded_objects:
                    doc.close()
                    return decoded_objects[0].data.decode('utf-8'), None
            
            doc.close()
            return None, 'PDF中未检测到二维码'
        except Exception as e:
            return None, f'PDF处理错误: {str(e)}'
    
    @staticmethod
    def parse_invoice_qrcode(qrcode_data):
        """解析发票二维码数据，支持多种格式"""
        if not qrcode_data:
            return None, '二维码数据为空'
        
        try:
            if qrcode_data.startswith('http'):
                xml_content, error = InvoiceParser.fetch_invoice_xml_from_url(qrcode_data)
                if xml_content:
                    result = InvoiceParser.parse_xml_to_dict(xml_content)
                    if result:
                        return result, None
                
                return InvoiceParser.parse_url_format(qrcode_data)
            
            if '<' in qrcode_data and '>' in qrcode_data and '<?xml' in qrcode_data.lower():
                return InvoiceParser.parse_xml_to_dict(qrcode_data), None
            
            if re.match(r'^[A-Za-z0-9+/=]+$', qrcode_data) and len(qrcode_data) > 50:
                result = InvoiceParser.decode_base64_to_xml(qrcode_data)
                if result:
                    if '<' in result and '>' in result:
                        return InvoiceParser.parse_xml_to_dict(result), None
            
            parts = qrcode_data.split(',')
            if len(parts) >= 4:
                return InvoiceParser.parse_csv_format(parts), None
            
            return None, f'无法识别的二维码格式: {qrcode_data[:100]}...'
        except Exception as e:
            return None, f'解析错误: {str(e)}'
    
    @staticmethod
    def parse_url_format(url):
        """解析URL格式的二维码"""
        invoice_data = {}
        
        patterns = {
            '发票代码': [r'fpdm=([^&]+)', r'd=([^&]+)', r'code=([^&]+)'],
            '发票号码': [r'fphm=([^&]+)', r'e=([^&]+)', r'no=([^&]+)', r'number=([^&]+)'],
            '校验码': [r'jym=([^&]+)', r'k=([^&]+)', r'checkCode=([^&]+)', r'check=([^&]+)'],
            '开票日期': [r'kprq=([^&]+)', r't=([^&]+)', r'date=([^&]+)', r'time=([^&]+)'],
            '合计金额': [r'kjje=([^&]+)', r'm=([^&]+)', r'amount=([^&]+)', r'money=([^&]+)'],
            '价税合计': [r'kpje=([^&]+)', r'total=([^&]+)'],
            '购买方名称': [r'gmfmc=([^&]+)', r'buyerName=([^&]+)'],
            '销售方名称': [r'xsfmc=([^&]+)', r'sellerName=([^&]+)'],
        }
        
        for field, pats in patterns.items():
            for pat in pats:
                match = re.search(pat, url, re.IGNORECASE)
                if match:
                    invoice_data[field] = match.group(1)
                    break
        
        if '开票日期' in invoice_data:
            date_str = invoice_data['开票日期']
            if len(date_str) == 8 and date_str.isdigit():
                invoice_data['开票日期'] = f'{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}'
            elif len(date_str) == 10 and '-' in date_str:
                pass
        
        if invoice_data:
            invoice_data['发票类型'] = '增值税发票'
            return invoice_data, None
        
        return None, 'URL中未找到发票信息'
    
    @staticmethod
    def parse_csv_format(parts):
        """解析CSV格式的二维码数据
        
        全电发票格式：版本号,发票种类,发票代码,发票号码,金额,开票日期,校验码,加密字段
        例如：01,32,,26427000000316294884,100.76,20260330,,5F98
        
        传统格式：发票类型代码,发票代码,发票号码,金额,日期,校验码(后6位)
        或：发票类型,发票代码,发票号码,金额,税额,日期,校验码
        """
        invoice_data = {}
        
        try:
            parts = [p.strip() for p in parts]
            
            invoice_type_map = {
                '01': '增值税专用发票',
                '02': '货运专票',
                '03': '机动车发票',
                '04': '增值税普通发票',
                '10': '电子普通发票',
                '11': '卷式发票',
                '14': '电子专票',
                '32': '全电发票（电子发票）',
                '33': '全电发票（增值税专用发票）',
            }
            
            if len(parts) >= 6:
                if parts[0] == '01' and parts[1].isdigit():
                    invoice_data['版本号'] = parts[0]
                    type_code = parts[1]
                    invoice_data['发票类型代码'] = type_code
                    invoice_data['发票类型'] = invoice_type_map.get(type_code, f'发票类型{type_code}')
                    invoice_data['发票代码'] = parts[2] if len(parts) > 2 else ''
                    invoice_data['发票号码'] = parts[3] if len(parts) > 3 else ''
                    invoice_data['价税合计'] = parts[4] if len(parts) > 4 else ''
                    invoice_data['开票日期'] = parts[5] if len(parts) > 5 else ''
                    if len(parts) > 6:
                        invoice_data['校验码'] = parts[6]
                    if len(parts) > 7:
                        invoice_data['加密字段'] = parts[7]
                elif parts[0].isdigit() and len(parts[0]) == 2:
                    type_code = parts[0]
                    invoice_data['发票类型代码'] = type_code
                    invoice_data['发票类型'] = invoice_type_map.get(type_code, f'发票类型{type_code}')
                    invoice_data['发票代码'] = parts[1]
                    invoice_data['发票号码'] = parts[2]
                    invoice_data['价税合计'] = parts[3]
                    invoice_data['开票日期'] = parts[4]
                    invoice_data['校验码'] = parts[5] if len(parts) > 5 else ''
                else:
                    invoice_data['发票代码'] = parts[0]
                    invoice_data['发票号码'] = parts[1]
                    invoice_data['价税合计'] = parts[2]
                    invoice_data['开票日期'] = parts[3]
                    if len(parts) > 4:
                        invoice_data['校验码'] = parts[4]
                    if len(parts) > 5:
                        invoice_data['合计税额'] = parts[5]
                    invoice_data['发票类型'] = '增值税发票'
            elif len(parts) >= 4:
                invoice_data['发票代码'] = parts[0]
                invoice_data['发票号码'] = parts[1]
                invoice_data['价税合计'] = parts[2]
                invoice_data['开票日期'] = parts[3]
                invoice_data['发票类型'] = '增值税发票'
            
            if '开票日期' in invoice_data:
                date_str = invoice_data['开票日期']
                if len(date_str) == 8 and date_str.isdigit():
                    invoice_data['开票日期'] = f'{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}'
            
            if '发票类型' not in invoice_data:
                invoice_data['发票类型'] = '增值税发票'
            
            return invoice_data
        except Exception as e:
            return {'发票类型': '增值税发票', '错误': str(e)}
    
    @staticmethod
    def decode_base64_to_xml(base64_data):
        """将Base64数据解码为XML"""
        try:
            decoded_bytes = base64.b64decode(base64_data)
            
            if decoded_bytes[:4] == b'PK\x03\x04':
                with zipfile.ZipFile(io.BytesIO(decoded_bytes), 'r') as zf:
                    xml_files = [f for f in zf.namelist() if f.endswith('.xml')]
                    if xml_files:
                        content = zf.read(xml_files[0])
                        try:
                            return content.decode('utf-8')
                        except:
                            return content.decode('gbk')
            
            try:
                return decoded_bytes.decode('utf-8')
            except UnicodeDecodeError:
                try:
                    return decoded_bytes.decode('gbk')
                except:
                    return decoded_bytes.decode('utf-8', errors='ignore')
        except Exception as e:
            print(f"Base64解码错误: {e}")
            return None
    
    @staticmethod
    def parse_xml_to_dict(xml_content):
        """解析XML内容为字典"""
        if not xml_content:
            return None
        
        try:
            root = ET.fromstring(xml_content)
            
            namespace = ''
            if root.tag.startswith('{'):
                namespace = root.tag.split('}')[0] + '}'
            
            def find_element(parent, *tags):
                for tag in tags:
                    elem = parent.find(namespace + tag) if namespace else parent.find(tag)
                    if elem is not None and elem.text:
                        return elem.text.strip()
                return ''
            
            def find_all_elements(parent, *tags):
                for tag in tags:
                    elems = parent.findall('.//' + namespace + tag) if namespace else parent.findall('.//' + tag)
                    if elems:
                        return elems
                    elems = parent.findall(namespace + tag) if namespace else parent.findall(tag)
                    if elems:
                        return elems
                return []
            
            def get_attr(parent, attr_name):
                return parent.get(attr_name, '') if parent is not None else ''
            
            invoice_data = {}
            
            if root.tag == 'EInvoice' or root.find('EInvoiceData') is not None or root.find('.//SellerInformation') is not None:
                invoice_data['发票类型'] = '电子发票'
                
                seller_info = root.find('.//SellerInformation')
                if seller_info is not None:
                    invoice_data['销售方名称'] = find_element(seller_info, 'SellerName')
                    invoice_data['销售方纳税人识别号'] = find_element(seller_info, 'SellerIdNum')
                    seller_addr = find_element(seller_info, 'SellerAddr')
                    seller_tel = find_element(seller_info, 'SellerTelNum')
                    invoice_data['销售方地址电话'] = f"{seller_addr} {seller_tel}".strip()
                
                buyer_info = root.find('.//BuyerInformation')
                if buyer_info is not None:
                    invoice_data['购买方名称'] = find_element(buyer_info, 'BuyerName')
                    invoice_data['购买方纳税人识别号'] = find_element(buyer_info, 'BuyerIdNum')
                    invoice_data['购买方地址电话'] = find_element(buyer_info, 'BuyerAddr')
                    invoice_data['购买方开户行及账号'] = find_element(buyer_info, 'BuyerBankAccount')
                
                basic_info = root.find('.//BasicInformation')
                if basic_info is not None:
                    invoice_data['合计金额'] = find_element(basic_info, 'TotalAmWithoutTax')
                    invoice_data['合计税额'] = find_element(basic_info, 'TotalTaxAm')
                    invoice_data['价税合计'] = find_element(basic_info, 'TotalTax-includedAmount')
                    invoice_data['价税合计大写'] = find_element(basic_info, 'TotalTax-includedAmountInChinese')
                    invoice_data['开票人'] = find_element(basic_info, 'Drawer')
                    invoice_data['开票日期'] = find_element(basic_info, 'RequestTime')
                
                tax_info = root.find('.//TaxSupervisionInfo')
                if tax_info is not None:
                    invoice_data['发票号码'] = find_element(tax_info, 'InvoiceNumber')
                    invoice_data['开票日期'] = find_element(tax_info, 'IssueTime') or invoice_data.get('开票日期', '')
                    invoice_data['税务机关'] = find_element(tax_info, 'TaxBureauName')
                
                header_info = root.find('.//Header')
                if header_info is not None:
                    invoice_data['电子发票ID'] = find_element(header_info, 'EIid')
                
                goods_list = []
                items = root.findall('.//IssuItemInformation')
                for item in items:
                    goods = {
                        '货物名称': find_element(item, 'ItemName'),
                        '规格型号': find_element(item, 'SpecMod'),
                        '单位': find_element(item, 'MeaUnits'),
                        '数量': find_element(item, 'Quantity'),
                        '单价': find_element(item, 'UnPrice'),
                        '金额': find_element(item, 'Amount'),
                        '税率': find_element(item, 'TaxRate'),
                        '税额': find_element(item, 'ComTaxAm'),
                    }
                    if any(goods.values()):
                        goods_list.append(goods)
                
                invoice_data['商品明细'] = goods_list
                return invoice_data
            
            invoice_data = {
                '发票类型': get_attr(root, '发票类型') or get_attr(root, 'invoiceType') or find_element(root, '发票类型名称', 'invoiceType', 'InvoiceType') or '增值税发票',
                '发票代码': find_element(root, '发票代码', 'invoiceCode', 'fpdm', 'InvoiceCode'),
                '发票号码': find_element(root, '发票号码', 'invoiceNumber', 'fphm', 'InvoiceNumber', 'InvoiceNo'),
                '开票日期': find_element(root, '开票日期', '开票时间', 'issueDate', 'kprq', 'IssueDate', 'IssueTime'),
                '校验码': find_element(root, '校验码', 'checkCode', 'jym', 'CheckCode'),
                '机器编号': find_element(root, '机器编号', 'machineNumber', 'MachineNo'),
                '购买方名称': find_element(root, '购买方名称', 'buyerName', 'gmfmc', 'BuyerName', 'Buyer'),
                '购买方纳税人识别号': find_element(root, '购买方纳税人识别号', 'buyerTaxId', 'gmfsbh', 'BuyerTaxID', 'BuyerCode'),
                '购买方地址电话': find_element(root, '购买方地址电话', 'buyerAddressPhone', 'gmfdzdh', 'BuyerAddressPhone'),
                '购买方开户行及账号': find_element(root, '购买方开户行及账号', 'buyerBankAccount', 'gmfyhzh', 'BuyerBankAccount'),
                '销售方名称': find_element(root, '销售方名称', 'sellerName', 'xsfmc', 'SellerName', 'Seller'),
                '销售方纳税人识别号': find_element(root, '销售方纳税人识别号', 'sellerTaxId', 'xsfsbh', 'SellerTaxID', 'SellerCode'),
                '销售方地址电话': find_element(root, '销售方地址电话', 'sellerAddressPhone', 'xsfdzdh', 'SellerAddressPhone'),
                '销售方开户行及账号': find_element(root, '销售方开户行及账号', 'sellerBankAccount', 'xsfyhzh', 'SellerBankAccount'),
                '合计金额': find_element(root, '合计金额', 'totalAmount', 'hjje', 'TotalAmount', 'Amount'),
                '合计税额': find_element(root, '合计税额', 'totalTax', 'hjse', 'TotalTax', 'Tax'),
                '价税合计': find_element(root, '价税合计', 'totalPriceTax', 'jshj', 'TotalPriceTax', 'Total'),
                '价税合计大写': find_element(root, '价税合计大写', 'totalPriceTaxCN', 'jshjdx', 'TotalPriceTaxCN'),
                '收款人': find_element(root, '收款人', 'payee', 'skr', 'Payee'),
                '复核人': find_element(root, '复核人', 'reviewer', 'fhr', 'Reviewer'),
                '开票人': find_element(root, '开票人', 'issuer', 'kpr', 'Issuer', 'Drawer'),
                '备注': find_element(root, '备注', 'remarks', 'bz', 'Remarks', 'Memo'),
            }
            
            goods_list = []
            
            item_tags = ['货物或应税劳务名称', '商品明细', 'invoiceItem', 'hwxx', 'xmx', 'Item', 'Goods', 'InvoiceItem', 'TaxRateItem']
            for item_tag in item_tags:
                items = find_all_elements(root, item_tag)
                if items:
                    for item in items:
                        goods = {
                            '货物名称': find_element(item, '货物或应税劳务名称', '商品名称', 'goodsName', 'hwmc', 'xmmc', 'GoodsName', 'ItemName', 'Name'),
                            '规格型号': find_element(item, '规格型号', 'specification', 'ggxh', 'Specification', 'Model'),
                            '单位': find_element(item, '单位', 'unit', 'dw', 'Unit', 'MeasureUnit'),
                            '数量': find_element(item, '数量', 'quantity', 'sl', 'Quantity', 'Num'),
                            '单价': find_element(item, '单价', 'unitPrice', 'dj', 'UnitPrice', 'Price'),
                            '金额': find_element(item, '金额', 'amount', 'je', 'Amount', 'Money'),
                            '税率': find_element(item, '税率', 'taxRate', 'slv', 'TaxRate'),
                            '税额': find_element(item, '税额', 'tax', 'se', 'Tax', 'TaxAmount'),
                        }
                        if any(goods.values()):
                            goods_list.append(goods)
                    if goods_list:
                        break
            
            invoice_data['商品明细'] = goods_list
            
            return invoice_data
        except Exception as e:
            print(f"XML解析错误: {e}")
            return None


class ProcessThread(QThread):
    """处理发票的线程"""
    progress = pyqtSignal(int, int)
    result_ready = pyqtSignal(dict)
    finished = pyqtSignal(list)
    error = pyqtSignal(str)
    
    def __init__(self, files):
        super().__init__()
        self.files = files
    
    def run(self):
        results = []
        total = len(self.files)
        
        for i, file_path in enumerate(self.files):
            self.progress.emit(i + 1, total)
            
            qrcode_data, qrcode_error = InvoiceParser.decode_qrcode(file_path)
            if not qrcode_data:
                result = {
                    '文件名': os.path.basename(file_path),
                    '文件路径': file_path,
                    '状态': '失败',
                    '错误信息': qrcode_error or '无法识别二维码'
                }
                results.append(result)
                self.result_ready.emit(result)
                continue
            
            invoice_data, parse_error = InvoiceParser.parse_invoice_qrcode(qrcode_data)
            
            if isinstance(invoice_data, str):
                xml_content = invoice_data
                invoice_data = InvoiceParser.parse_xml_to_dict(xml_content)
                if not invoice_data:
                    result = {
                        '文件名': os.path.basename(file_path),
                        '文件路径': file_path,
                        '状态': '失败',
                        '错误信息': '无法解析XML数据'
                    }
                    results.append(result)
                    self.result_ready.emit(result)
                    continue
            
            if not invoice_data:
                result = {
                    '文件名': os.path.basename(file_path),
                    '文件路径': file_path,
                    '状态': '失败',
                    '错误信息': parse_error or '无法解析发票数据'
                }
                results.append(result)
                self.result_ready.emit(result)
                continue
            
            if not invoice_data.get('销售方名称') or not invoice_data.get('销售方纳税人识别号') or not invoice_data.get('购买方名称') or not invoice_data.get('购买方纳税人识别号'):
                ocr_info, ocr_error = InvoiceParser.ocr_extract_seller_info(file_path)
                if ocr_info:
                    if not invoice_data.get('销售方名称') and ocr_info.get('销售方名称'):
                        invoice_data['销售方名称'] = ocr_info['销售方名称']
                    if not invoice_data.get('销售方纳税人识别号') and ocr_info.get('销售方纳税人识别号'):
                        invoice_data['销售方纳税人识别号'] = ocr_info['销售方纳税人识别号']
                    if not invoice_data.get('购买方名称') and ocr_info.get('购买方名称'):
                        invoice_data['购买方名称'] = ocr_info['购买方名称']
                    if not invoice_data.get('购买方纳税人识别号') and ocr_info.get('购买方纳税人识别号'):
                        invoice_data['购买方纳税人识别号'] = ocr_info['购买方纳税人识别号']
            
            invoice_data['文件名'] = os.path.basename(file_path)
            invoice_data['文件路径'] = file_path
            invoice_data['状态'] = '成功'
            invoice_data['错误信息'] = ''
            results.append(invoice_data)
            self.result_ready.emit(invoice_data)
        
        self.finished.emit(results)


class InvoiceMainWindow(QMainWindow):
    """发票识别主窗口"""
    
    def __init__(self):
        super().__init__()
        self.file_list = []
        self.invoice_results = []
        self.init_ui()
    
    def init_ui(self):
        self.setWindowTitle('发票识别工具')
        self.setGeometry(100, 100, 1400, 900)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(10)
        
        toolbar_widget = QWidget()
        toolbar_layout = QHBoxLayout(toolbar_widget)
        toolbar_layout.setContentsMargins(0, 0, 0, 0)
        
        self.btn_add = QPushButton('添加文件')
        self.btn_add.clicked.connect(self.add_files)
        self.btn_add.setMinimumHeight(40)
        self.btn_add.setMinimumWidth(100)
        
        self.btn_add_folder = QPushButton('添加文件夹')
        self.btn_add_folder.clicked.connect(self.add_folder)
        self.btn_add_folder.setMinimumHeight(40)
        self.btn_add_folder.setMinimumWidth(100)
        
        self.btn_remove = QPushButton('删除选中')
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_remove.setMinimumHeight(40)
        self.btn_remove.setMinimumWidth(100)
        
        self.btn_clear = QPushButton('清空列表')
        self.btn_clear.clicked.connect(self.clear_list)
        self.btn_clear.setMinimumHeight(40)
        self.btn_clear.setMinimumWidth(100)
        
        self.btn_process = QPushButton('开始识别')
        self.btn_process.clicked.connect(self.process_files)
        self.btn_process.setMinimumHeight(40)
        self.btn_process.setMinimumWidth(120)
        self.btn_process.setStyleSheet('background-color: #2196F3; color: white; font-size: 14px; font-weight: bold;')
        
        self.btn_export = QPushButton('导出Excel')
        self.btn_export.clicked.connect(self.export_to_excel)
        self.btn_export.setMinimumHeight(40)
        self.btn_export.setMinimumWidth(100)
        self.btn_export.setEnabled(False)
        
        toolbar_layout.addWidget(self.btn_add)
        toolbar_layout.addWidget(self.btn_add_folder)
        toolbar_layout.addWidget(self.btn_remove)
        toolbar_layout.addWidget(self.btn_clear)
        toolbar_layout.addSpacing(20)
        toolbar_layout.addWidget(self.btn_process)
        toolbar_layout.addWidget(self.btn_export)
        toolbar_layout.addStretch()
        
        main_layout.addWidget(toolbar_widget)
        
        file_section = QWidget()
        file_layout = QVBoxLayout(file_section)
        file_layout.setContentsMargins(0, 0, 0, 0)
        file_layout.setSpacing(5)
        
        file_header = QHBoxLayout()
        file_header.addWidget(QLabel('待识别文件列表:'))
        self.file_count_label = QLabel('共 0 个文件')
        file_header.addWidget(self.file_count_label)
        file_header.addStretch()
        file_layout.addLayout(file_header)
        
        self.file_listbox = QListWidget()
        self.file_listbox.setSelectionMode(QListWidget.ExtendedSelection)
        self.file_listbox.setAlternatingRowColors(True)
        self.file_listbox.setMinimumHeight(150)
        file_layout.addWidget(self.file_listbox)
        
        main_layout.addWidget(file_section)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        result_section = QWidget()
        result_layout = QVBoxLayout(result_section)
        result_layout.setContentsMargins(0, 0, 0, 0)
        result_layout.setSpacing(5)
        
        result_header = QHBoxLayout()
        result_header.addWidget(QLabel('识别结果:'))
        self.result_count_label = QLabel('共 0 条记录')
        result_header.addWidget(self.result_count_label)
        result_header.addStretch()
        result_layout.addLayout(result_header)
        
        self.result_table = QTableWidget()
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.result_table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.result_table.cellDoubleClicked.connect(self.on_cell_double_clicked)
        self.result_table.cellChanged.connect(self.on_cell_changed)
        result_layout.addWidget(self.result_table)
        
        main_layout.addWidget(result_section, 1)
        
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage('就绪')
    
    def on_cell_double_clicked(self, row, col):
        """双击单元格"""
        if col == 0:
            if row < len(self.invoice_results):
                result = self.invoice_results[row]
                file_path = result.get('文件路径', '')
                if file_path and os.path.exists(file_path):
                    import subprocess
                    import platform
                    try:
                        if platform.system() == 'Windows':
                            os.startfile(file_path)
                        elif platform.system() == 'Darwin':
                            subprocess.run(['open', file_path])
                        else:
                            subprocess.run(['xdg-open', file_path])
                    except Exception as e:
                        QMessageBox.warning(self, '提示', f'无法打开文件: {str(e)}')
                else:
                    QMessageBox.warning(self, '提示', '文件路径不存在')
    
    def on_cell_changed(self, row, col):
        """单元格内容改变时保存"""
        if row < len(self.invoice_results):
            headers = [
                '文件名', '状态', '发票类型', '发票代码', '发票号码', '开票日期',
                '购买方名称', '购买方纳税人识别号', '销售方名称', '销售方纳税人识别号',
                '合计金额', '合计税额', '价税合计', '错误信息'
            ]
            if col < len(headers):
                item = self.result_table.item(row, col)
                if item:
                    self.invoice_results[row][headers[col]] = item.text()
    
    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            '选择发票文件',
            '',
            '图片文件 (*.png *.jpg *.jpeg *.bmp *.pdf);;所有文件 (*.*)'
        )
        if files:
            self.add_files_to_list(files)
    
    def add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, '选择文件夹')
        if folder:
            files = []
            for ext in ['*.png', '*.jpg', '*.jpeg', '*.bmp', '*.pdf']:
                files.extend(self.get_files_by_pattern(folder, ext))
            if files:
                self.add_files_to_list(files)
    
    def get_files_by_pattern(self, folder, pattern):
        import glob
        return glob.glob(os.path.join(folder, pattern))
    
    def add_files_to_list(self, files):
        for f in files:
            if f not in self.file_list:
                self.file_list.append(f)
                self.file_listbox.addItem(f)
        self.update_file_count()
    
    def remove_selected(self):
        selected_items = self.file_listbox.selectedItems()
        for item in selected_items:
            row = self.file_listbox.row(item)
            self.file_listbox.takeItem(row)
            if item.text() in self.file_list:
                self.file_list.remove(item.text())
        self.update_file_count()
    
    def clear_list(self):
        self.file_list.clear()
        self.file_listbox.clear()
        self.update_file_count()
    
    def update_file_count(self):
        self.file_count_label.setText(f'共 {len(self.file_list)} 个文件')
    
    def process_files(self):
        if not self.file_list:
            QMessageBox.warning(self, '提示', '请先添加发票文件！')
            return
        
        self.btn_process.setEnabled(False)
        self.btn_add.setEnabled(False)
        self.btn_add_folder.setEnabled(False)
        self.btn_remove.setEnabled(False)
        self.btn_clear.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_bar.showMessage('正在识别...')
        
        self.process_thread = ProcessThread(self.file_list.copy())
        self.process_thread.progress.connect(self.update_progress)
        self.process_thread.result_ready.connect(self.add_result)
        self.process_thread.finished.connect(self.on_process_finished)
        self.process_thread.start()
    
    def update_progress(self, current, total):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.status_bar.showMessage(f'正在识别: {current}/{total}')
    
    def add_result(self, result):
        """实时添加一条识别结果"""
        self.invoice_results.append(result)
        
        print(f"\n识别结果: {result.get('文件名', '')}")
        for key, value in result.items():
            if value and key not in ['文件路径', '二维码内容']:
                print(f"  {key}: {value}")
        
        headers = [
            '文件名', '状态', '发票类型', '发票号码', '开票日期',
            '购买方名称', '购买方纳税人识别号', '销售方名称', '销售方纳税人识别号',
            '合计金额', '合计税额', '价税合计', '错误信息'
        ]
        
        if self.result_table.columnCount() == 0:
            self.result_table.setColumnCount(len(headers))
            self.result_table.setHorizontalHeaderLabels(headers)
        
        self.result_table.blockSignals(True)
        
        row = self.result_table.rowCount()
        self.result_table.insertRow(row)
        
        for col, header in enumerate(headers):
            value = result.get(header, '')
            item = QTableWidgetItem(str(value) if value else '')
            if header == '状态':
                if value == '成功':
                    item.setBackground(Qt.green)
                else:
                    item.setBackground(Qt.red)
            elif header == '文件名':
                item.setForeground(QColor(0, 0, 255))
                font = item.font()
                font.setUnderline(True)
                item.setFont(font)
                file_path = result.get('文件路径', '')
                if file_path:
                    item.setToolTip(f'双击打开: {file_path}')
            self.result_table.setItem(row, col, item)
        
        self.result_table.blockSignals(False)
        
        header = self.result_table.horizontalHeader()
        for i in range(len(headers)):
            header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
        self.result_table.resizeColumnsToContents()
        
        self.result_count_label.setText(f'共 {len(self.invoice_results)} 条记录')
        
        success_count = sum(1 for r in self.invoice_results if r.get('状态') == '成功')
        self.status_bar.showMessage(f'识别中: 成功 {success_count} 个, 失败 {len(self.invoice_results) - success_count} 个')
    
    def on_process_finished(self, results):
        self.invoice_results = results
        self.display_results(results)
        
        self.btn_process.setEnabled(True)
        self.btn_add.setEnabled(True)
        self.btn_add_folder.setEnabled(True)
        self.btn_remove.setEnabled(True)
        self.btn_clear.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.btn_export.setEnabled(True)
        
        success_count = sum(1 for r in results if r.get('状态') == '成功')
        self.status_bar.showMessage(f'识别完成: 成功 {success_count} 个, 失败 {len(results) - success_count} 个')
    
    def display_results(self, results):
        if not results:
            return
        
        headers = [
            '文件名', '状态', '发票类型', '发票号码', '开票日期',
            '购买方名称', '购买方纳税人识别号', '销售方名称', '销售方纳税人识别号',
            '合计金额', '合计税额', '价税合计', '错误信息'
        ]
        
        self.result_table.setColumnCount(len(headers))
        self.result_table.setHorizontalHeaderLabels(headers)
        self.result_table.setRowCount(len(results))
        
        for row, result in enumerate(results):
            for col, header in enumerate(headers):
                value = result.get(header, '')
                item = QTableWidgetItem(str(value) if value else '')
                if header == '状态':
                    if value == '成功':
                        item.setBackground(Qt.green)
                    else:
                        item.setBackground(Qt.red)
                elif header == '文件名':
                    item.setForeground(QColor(0, 0, 255))
                    font = item.font()
                    font.setUnderline(True)
                    item.setFont(font)
                    file_path = result.get('文件路径', '')
                    if file_path:
                        item.setToolTip(f'双击打开: {file_path}')
                self.result_table.setItem(row, col, item)
        
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.result_table.horizontalHeader().setStretchLastSection(True)
        
        self.result_count_label.setText(f'共 {len(results)} 条记录')
    
    def export_to_excel(self):
        if not self.invoice_results:
            QMessageBox.warning(self, '提示', '没有可导出的数据！')
            return
        
        all_headers = [
            '文件名', '状态', '发票类型', '发票代码', '发票号码', '开票日期',
            '购买方名称', '购买方纳税人识别号', '销售方名称', '销售方纳税人识别号',
            '合计金额', '合计税额', '价税合计', '错误信息'
        ]
        
        default_selected = all_headers.copy()
        default_selected.remove('状态')
        default_selected.remove('发票类型')
        default_selected.remove('错误信息')
        
        dialog = FieldSelectDialog(all_headers, default_selected, self)
        if dialog.exec_() != QDialog.Accepted:
            return
        
        selected_headers = dialog.get_selected_fields()
        if not selected_headers:
            QMessageBox.warning(self, '提示', '请至少选择一个字段！')
            return
        
        default_name = f'发票识别结果_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            '保存Excel文件',
            default_name,
            'Excel文件 (*.xlsx)'
        )
        
        if not file_path:
            return
        
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = '发票信息'
            
            header_font = ExcelFont(bold=True, size=11)
            header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell_alignment = Alignment(vertical='center', wrap_text=True)
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            for col, header in enumerate(selected_headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = header_font
                cell.alignment = header_alignment
                cell.border = thin_border
            
            for row, result in enumerate(self.invoice_results, 2):
                for col, header in enumerate(selected_headers, 1):
                    value = result.get(header, '')
                    cell = ws.cell(row=row, column=col, value=str(value) if value else '')
                    cell.alignment = cell_alignment
                    cell.border = thin_border
            
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column].width = adjusted_width
            
            ws.freeze_panes = 'A2'
            
            if any(r.get('商品明细') for r in self.invoice_results if r.get('状态') == '成功'):
                ws_detail = wb.create_sheet(title='商品明细')
                
                detail_headers = [
                    '发票号码', '货物名称', '规格型号', '单位', '数量', '单价', '金额', '税率', '税额'
                ]
                
                for col, header in enumerate(detail_headers, 1):
                    cell = ws_detail.cell(row=1, column=col, value=header)
                    cell.font = header_font
                    cell.alignment = header_alignment
                    cell.border = thin_border
                
                detail_row = 2
                for result in self.invoice_results:
                    if result.get('状态') == '成功' and result.get('商品明细'):
                        invoice_no = result.get('发票号码', '')
                        for goods in result['商品明细']:
                            ws_detail.cell(row=detail_row, column=1, value=invoice_no).border = thin_border
                            for col, key in enumerate(['货物名称', '规格型号', '单位', '数量', '单价', '金额', '税率', '税额'], 2):
                                cell = ws_detail.cell(row=detail_row, column=col, value=str(goods.get(key, '')))
                                cell.border = thin_border
                            detail_row += 1
                
                for col in ws_detail.columns:
                    max_length = 0
                    column = col[0].column_letter
                    for cell in col:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    ws_detail.column_dimensions[column].width = adjusted_width
            
            wb.save(file_path)
            QMessageBox.information(self, '成功', f'导出成功！\n文件保存在: {file_path}')
            
        except Exception as e:
            QMessageBox.critical(self, '错误', f'导出失败: {str(e)}')


class FieldSelectDialog(QDialog):
    """字段选择对话框"""
    
    def __init__(self, all_fields, default_selected, parent=None):
        super().__init__(parent)
        self.setWindowTitle('选择导出字段')
        self.setMinimumWidth(400)
        self.setMinimumHeight(500)
        
        layout = QVBoxLayout(self)
        
        label = QLabel('请选择要导出的字段：')
        layout.addWidget(label)
        
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        self.checkboxes = {}
        for field in all_fields:
            cb = QCheckBox(field)
            cb.setChecked(field in default_selected)
            self.checkboxes[field] = cb
            scroll_layout.addWidget(cb)
        
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        layout.addWidget(scroll)
        
        btn_layout = QHBoxLayout()
        
        btn_select_all = QPushButton('全选')
        btn_select_all.clicked.connect(self.select_all)
        btn_layout.addWidget(btn_select_all)
        
        btn_deselect_all = QPushButton('取消全选')
        btn_deselect_all.clicked.connect(self.deselect_all)
        btn_layout.addWidget(btn_deselect_all)
        
        layout.addLayout(btn_layout)
        
        btn_box = QHBoxLayout()
        btn_ok = QPushButton('确定')
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton('取消')
        btn_cancel.clicked.connect(self.reject)
        btn_box.addStretch()
        btn_box.addWidget(btn_ok)
        btn_box.addWidget(btn_cancel)
        layout.addLayout(btn_box)
    
    def select_all(self):
        for cb in self.checkboxes.values():
            cb.setChecked(True)
    
    def deselect_all(self):
        for cb in self.checkboxes.values():
            cb.setChecked(False)
    
    def get_selected_fields(self):
        return [field for field, cb in self.checkboxes.items() if cb.isChecked()]


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    font = QFont('Microsoft YaHei', 9)
    app.setFont(font)
    
    window = InvoiceMainWindow()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()

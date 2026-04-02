import requests
import json
import time
import os
from datetime import datetime
from defusedxml import ElementTree as ET


class TaxBureauClient:
    """电子税务局发票查询客户端"""
    
    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'zh-CN,zh;q=0.9',
            'Content-Type': 'application/json;charset=UTF-8',
        })
        self.cookie_file = 'tax_cookies.json'
        self.base_url = 'https://etax.chinatax.gov.cn'
        
    def load_cookies(self):
        """加载保存的cookies"""
        if os.path.exists(self.cookie_file):
            try:
                with open(self.cookie_file, 'r') as f:
                    cookies = json.load(f)
                    for cookie in cookies:
                        self.session.cookies.set(cookie['name'], cookie['value'])
                return True
            except:
                pass
        return False
    
    def save_cookies(self):
        """保存cookies到文件"""
        cookies = []
        for cookie in self.session.cookies:
            cookies.append({
                'name': cookie.name,
                'value': cookie.value,
                'domain': cookie.domain,
                'path': cookie.path,
            })
        with open(self.cookie_file, 'w') as f:
            json.dump(cookies, f)
    
    def check_login_status(self):
        """检查登录状态"""
        try:
            url = f'{self.base_url}/api/user/info'
            resp = self.session.get(url, timeout=10)
            if resp.status_code == 200:
                data = resp.json()
                return data.get('success', False)
        except:
            pass
        return False
    
    def query_invoice_by_qrcode(self, invoice_code, invoice_number, total_amount, issue_date, check_code=''):
        """通过发票信息查询发票详情
        
        Args:
            invoice_code: 发票代码
            invoice_number: 发票号码
            total_amount: 价税合计金额
            issue_date: 开票日期 (YYYYMMDD 或 YYYY-MM-DD)
            check_code: 校验码后6位（可选）
        """
        if '-' in issue_date:
            issue_date = issue_date.replace('-', '')
        
        params = {
            'fpdm': invoice_code,
            'fphm': invoice_number,
            'je': total_amount,
            'kprq': issue_date,
        }
        
        if check_code:
            params['jym'] = check_code
        
        try:
            url = f'{self.base_url}/api/invoice/query'
            resp = self.session.post(url, json=params, timeout=15)
            
            if resp.status_code == 200:
                data = resp.json()
                if data.get('success'):
                    return data.get('data'), None
                else:
                    return None, data.get('message', '查询失败')
            else:
                return None, f'HTTP错误: {resp.status_code}'
        except requests.exceptions.Timeout:
            return None, '请求超时'
        except requests.exceptions.ConnectionError:
            return None, '网络连接失败'
        except Exception as e:
            return None, f'查询错误: {str(e)}'
    
    def query_invoice_xml(self, invoice_number, invoice_code=''):
        """获取发票XML数据
        
        Args:
            invoice_number: 发票号码/发票唯一标识
            invoice_code: 发票代码（全电发票可为空）
        """
        params = {
            'fphm': invoice_number,
            'fpdm': invoice_code,
        }
        
        try:
            url = f'{self.base_url}/api/invoice/xml'
            resp = self.session.post(url, json=params, timeout=15)
            
            if resp.status_code == 200:
                data = resp.json()
                if data.get('success'):
                    xml_content = data.get('data', {}).get('xmlContent', '')
                    if xml_content:
                        return xml_content, None
                    return None, '未获取到XML数据'
                else:
                    return None, data.get('message', '获取XML失败')
            else:
                return None, f'HTTP错误: {resp.status_code}'
        except requests.exceptions.Timeout:
            return None, '请求超时'
        except requests.exceptions.ConnectionError:
            return None, '网络连接失败'
        except Exception as e:
            return None, f'获取错误: {str(e)}'
    
    def query_full_electronic_invoice(self, invoice_number):
        """查询全电发票详情
        
        Args:
            invoice_number: 全电发票号码/发票唯一标识
        """
        try:
            url = f'{self.base_url}/api/invoice/full/query'
            params = {'fphm': invoice_number}
            resp = self.session.post(url, json=params, timeout=15)
            
            if resp.status_code == 200:
                data = resp.json()
                if data.get('success'):
                    return data.get('data'), None
                else:
                    return None, data.get('message', '查询失败')
            else:
                return None, f'HTTP错误: {resp.status_code}'
        except requests.exceptions.Timeout:
            return None, '请求超时'
        except requests.exceptions.ConnectionError:
            return None, '网络连接失败'
        except Exception as e:
            return None, f'查询错误: {str(e)}'


class InvoiceQueryService:
    """发票查询服务"""
    
    def __init__(self):
        self.client = TaxBureauClient()
        self.client.load_cookies()
    
    def query_invoice(self, invoice_data):
        """查询发票信息
        
        Args:
            invoice_data: 包含发票信息的字典，支持以下字段：
                - 发票号码
                - 发票代码
                - 价税合计
                - 合计金额
                - 开票日期
                - 校验码
        """
        invoice_number = invoice_data.get('发票号码', '')
        invoice_code = invoice_data.get('发票代码', '')
        total_amount = invoice_data.get('价税合计') or invoice_data.get('合计金额', '')
        issue_date = invoice_data.get('开票日期', '')
        check_code = invoice_data.get('校验码', '')
        
        if not invoice_number:
            return None, '发票号码不能为空'
        
        if check_code and len(check_code) > 6:
            check_code = check_code[-6:]
        
        if issue_date:
            issue_date = issue_date.replace('-', '')[:8]
        
        xml_content, error = self.client.query_invoice_xml(invoice_number, invoice_code)
        if xml_content:
            return xml_content, None
        
        result, error = self.client.query_invoice_by_qrcode(
            invoice_code, invoice_number, total_amount, issue_date, check_code
        )
        if result:
            return result, None
        
        return None, error
    
    def parse_invoice_xml(self, xml_content):
        """解析发票XML"""
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
            
            invoice_data = {}
            
            if root.tag == 'EInvoice' or root.find('.//SellerInformation') is not None:
                invoice_data['发票类型'] = '电子发票'
                
                seller_info = root.find('.//SellerInformation')
                if seller_info is not None:
                    invoice_data['销售方名称'] = find_element(seller_info, 'SellerName')
                    invoice_data['销售方纳税人识别号'] = find_element(seller_info, 'SellerIdNum')
                
                buyer_info = root.find('.//BuyerInformation')
                if buyer_info is not None:
                    invoice_data['购买方名称'] = find_element(buyer_info, 'BuyerName')
                    invoice_data['购买方纳税人识别号'] = find_element(buyer_info, 'BuyerIdNum')
                
                basic_info = root.find('.//BasicInformation')
                if basic_info is not None:
                    invoice_data['合计金额'] = find_element(basic_info, 'TotalAmWithoutTax')
                    invoice_data['合计税额'] = find_element(basic_info, 'TotalTaxAm')
                    invoice_data['价税合计'] = find_element(basic_info, 'TotalTax-includedAmount')
                    invoice_data['开票人'] = find_element(basic_info, 'Drawer')
                    invoice_data['开票日期'] = find_element(basic_info, 'RequestTime')
                
                tax_info = root.find('.//TaxSupervisionInfo')
                if tax_info is not None:
                    invoice_data['发票号码'] = find_element(tax_info, 'InvoiceNumber')
            else:
                invoice_data['发票类型'] = find_element(root, '发票类型名称', 'invoiceType') or '增值税发票'
                invoice_data['发票代码'] = find_element(root, '发票代码', 'invoiceCode', 'fpdm')
                invoice_data['发票号码'] = find_element(root, '发票号码', 'invoiceNumber', 'fphm')
                invoice_data['开票日期'] = find_element(root, '开票日期', '开票时间', 'issueDate', 'kprq')
                invoice_data['购买方名称'] = find_element(root, '购买方名称', 'buyerName', 'gmfmc')
                invoice_data['购买方纳税人识别号'] = find_element(root, '购买方纳税人识别号', 'buyerTaxId', 'gmfsbh')
                invoice_data['销售方名称'] = find_element(root, '销售方名称', 'sellerName', 'xsfmc')
                invoice_data['销售方纳税人识别号'] = find_element(root, '销售方纳税人识别号', 'sellerTaxId', 'xsfsbh')
                invoice_data['合计金额'] = find_element(root, '合计金额', 'totalAmount', 'hjje')
                invoice_data['合计税额'] = find_element(root, '合计税额', 'totalTax', 'hjse')
                invoice_data['价税合计'] = find_element(root, '价税合计', 'totalPriceTax', 'jshj')
            
            return invoice_data
        except Exception as e:
            return None

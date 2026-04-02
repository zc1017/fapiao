import sys
sys.path.insert(0, r'c:\Users\zc\Desktop\fapiao')

from invoice_app import InvoiceParser

# 全电发票二维码数据
qrcode_data = "01,32,,26427000000316294884,100.76,20260330,,5F98"

result = InvoiceParser.parse_csv_format(qrcode_data.split(','))

print("全电发票解析结果:")
for key, value in result.items():
    print(f"  {key}: {value}")

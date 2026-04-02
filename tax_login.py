import asyncio
import json
import os
from datetime import datetime
from playwright.async_api import async_playwright, Page, Browser


class TaxBureauLogin:
    """电子税务局自动登录"""
    
    def __init__(self, headless=False):
        self.headless = headless
        self.browser = None
        self.context = None
        self.page = None
        self.storage_state_file = 'tax_auth_state.json'
        
    async def start(self):
        """启动浏览器"""
        self.playwright = await async_playwright().start()
        self.browser = await self.playwright.chromium.launch(
            headless=self.headless,
            args=['--disable-blink-features=AutomationControlled']
        )
        
        if os.path.exists(self.storage_state_file):
            self.context = await self.browser.new_context(
                storage_state=self.storage_state_file
            )
        else:
            self.context = await self.browser.new_context()
        
        self.page = await self.context.new_page()
        
        await self.page.set_extra_http_headers({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept-Language': 'zh-CN,zh;q=0.9',
        })
        
    async def close(self):
        """关闭浏览器"""
        if self.context:
            await self.context.close()
        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()
    
    async def save_auth_state(self):
        """保存认证状态"""
        await self.context.storage_state(path=self.storage_state_file)
    
    async def load_auth_state(self):
        """加载认证状态"""
        if os.path.exists(self.storage_state_file):
            try:
                with open(self.storage_state_file, 'r') as f:
                    return json.load(f)
            except:
                pass
        return None
    
    async def check_login_status(self) -> bool:
        """检查登录状态"""
        try:
            await self.page.goto('https://etax.chinatax.gov.cn/', wait_until='networkidle')
            
            current_url = self.page.url
            
            if 'login' not in current_url.lower() and 'auth' not in current_url.lower():
                return True
            
            return False
        except Exception as e:
            print(f"检查登录状态错误: {e}")
            return False
    
    async def login_with_qr_code(self, timeout=300):
        """二维码扫码登录
        
        Args:
            timeout: 等待扫码超时时间(秒)
        """
        try:
            await self.page.goto('https://etax.chinatax.gov.cn/', wait_until='networkidle')
            
            await asyncio.sleep(2)
            
            qr_selector = 'img[src*="qrcode"], img[src*="qr"], .qrcode img, #qrcode, .login-qr img'
            qr_element = await self.page.query_selector(qr_selector)
            
            if not qr_element:
                login_tab = await self.page.query_selector('text=扫码登录, text=二维码登录, [class*="qr"]')
                if login_tab:
                    await login_tab.click()
                    await asyncio.sleep(1)
            
            print("请使用个人所得税APP或电子税务局APP扫描二维码登录...")
            
            start_time = datetime.now()
            while (datetime.now() - start_time).seconds < timeout:
                current_url = self.page.url
                if 'login' not in current_url.lower() and 'auth' not in current_url.lower():
                    print("登录成功!")
                    await self.save_auth_state()
                    return True
                
                await asyncio.sleep(2)
            
            print("扫码登录超时")
            return False
            
        except Exception as e:
            print(f"登录错误: {e}")
            return False
    
    async def login_with_password(self, username: str, password: str, captcha: str = None):
        """账号密码登录
        
        Args:
            username: 用户名/手机号
            password: 密码
            captcha: 验证码(如果需要)
        """
        try:
            await self.page.goto('https://etax.chinatax.gov.cn/', wait_until='networkidle')
            
            login_tab = await self.page.query_selector('text=账号登录, text=密码登录')
            if login_tab:
                await login_tab.click()
                await asyncio.sleep(1)
            
            username_input = await self.page.query_selector('input[placeholder*="用户名"], input[placeholder*="手机号"], input[name="username"], #username')
            if username_input:
                await username_input.fill(username)
            
            password_input = await self.page.query_selector('input[placeholder*="密码"], input[type="password"], #password')
            if password_input:
                await password_input.fill(password)
            
            if captcha:
                captcha_input = await self.page.query_selector('input[placeholder*="验证码"], #captcha')
                if captcha_input:
                    await captcha_input.fill(captcha)
            
            login_button = await self.page.query_selector('button:has-text("登录"), input[type="submit"]')
            if login_button:
                await login_button.click()
            
            await asyncio.sleep(3)
            
            current_url = self.page.url
            if 'login' not in current_url.lower():
                print("登录成功!")
                await self.save_auth_state()
                return True
            
            error_msg = await self.page.query_selector('.error-message, .login-error')
            if error_msg:
                error_text = await error_msg.text_content()
                print(f"登录失败: {error_text}")
            
            return False
            
        except Exception as e:
            print(f"登录错误: {e}")
            return False
    
    async def query_invoice(self, invoice_data: dict):
        """查询发票信息
        
        Args:
            invoice_data: 发票信息字典
        """
        try:
            await self.page.goto('https://etax.chinatax.gov.cn/', wait_until='networkidle')
            
            invoice_menu = await self.page.query_selector('text=发票业务, text=发票查询, [href*="invoice"]')
            if invoice_menu:
                await invoice_menu.click()
                await asyncio.sleep(2)
            
            invoice_code = invoice_data.get('发票代码', '')
            invoice_number = invoice_data.get('发票号码', '')
            total_amount = invoice_data.get('价税合计') or invoice_data.get('合计金额', '')
            issue_date = invoice_data.get('开票日期', '').replace('-', '')[:8]
            check_code = invoice_data.get('校验码', '')
            
            code_input = await self.page.query_selector('input[name="fpdm"], input[placeholder*="发票代码"]')
            if code_input and invoice_code:
                await code_input.fill(invoice_code)
            
            number_input = await self.page.query_selector('input[name="fphm"], input[placeholder*="发票号码"]')
            if number_input:
                await number_input.fill(invoice_number)
            
            amount_input = await self.page.query_selector('input[name="je"], input[placeholder*="金额"]')
            if amount_input and total_amount:
                await amount_input.fill(str(total_amount))
            
            date_input = await self.page.query_selector('input[name="kprq"], input[placeholder*="日期"]')
            if date_input and issue_date:
                await date_input.fill(issue_date)
            
            check_input = await self.page.query_selector('input[name="jym"], input[placeholder*="校验码"]')
            if check_input and check_code:
                await check_input.fill(check_code[-6:] if len(check_code) > 6 else check_code)
            
            query_button = await self.page.query_selector('button:has-text("查询"), button:has-text("查验")')
            if query_button:
                await query_button.click()
                await asyncio.sleep(3)
            
            result_table = await self.page.query_selector('table, .result-table')
            if result_table:
                rows = await result_table.query_selector_all('tr')
                results = []
                for row in rows[1:]:
                    cells = await row.query_selector_all('td')
                    if cells:
                        row_data = []
                        for cell in cells:
                            text = await cell.text_content()
                            row_data.append(text.strip())
                        results.append(row_data)
                return results
            
            return None
            
        except Exception as e:
            print(f"查询发票错误: {e}")
            return None
    
    async def get_invoice_xml(self, invoice_number: str) -> str:
        """获取发票XML数据
        
        Args:
            invoice_number: 发票号码
        """
        try:
            detail_button = await self.page.query_selector(f'tr:has-text("{invoice_number}") button:has-text("详情")')
            if detail_button:
                await detail_button.click()
                await asyncio.sleep(2)
            
            xml_button = await self.page.query_selector('button:has-text("XML"), button:has-text("下载XML")')
            if xml_button:
                async with self.page.expect_download() as download_info:
                    await xml_button.click()
                download = await download_info.value
                xml_content = await download.path()
                
                with open(xml_content, 'r', encoding='utf-8') as f:
                    return f.read()
            
            return None
            
        except Exception as e:
            print(f"获取XML错误: {e}")
            return None


async def auto_login_and_query(invoice_data: dict = None, headless: bool = False):
    """自动登录并查询发票
    
    Args:
        invoice_data: 发票信息(可选)
        headless: 是否无头模式
    """
    client = TaxBureauLogin(headless=headless)
    
    try:
        await client.start()
        
        is_logged_in = await client.check_login_status()
        
        if not is_logged_in:
            print("未登录，请扫码登录...")
            success = await client.login_with_qr_code(timeout=300)
            if not success:
                return None, "登录失败"
        
        if invoice_data:
            result = await client.query_invoice(invoice_data)
            return result, None
        
        return True, None
        
    finally:
        await client.close()


def run_async_login():
    """运行异步登录"""
    asyncio.run(auto_login_and_query())


if __name__ == '__main__':
    asyncio.run(auto_login_and_query())

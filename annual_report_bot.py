"""工商年报自动申报主程序

流程：
1. 联络员变更（如需要）
2. 联络员登录
3. 填写年报表单
4. 提交并记录结果
"""
import os
import time
import json
from datetime import datetime
from playwright.sync_api import sync_playwright, Page, expect
from loguru import logger

import config
from captcha_solver import CaptchaSolver
from data_reader import read_enterprise_data, read_annual_report_data
from sms_handler import SmsHandler


class AnnualReportBot:
    def __init__(self):
        self.captcha = CaptchaSolver()
        self.sms = SmsHandler()
        self.results = []  # 记录每家企业的处理结果
        
        # 确保目录存在
        os.makedirs(config.SCREENSHOT_DIR, exist_ok=True)
        os.makedirs("logs", exist_ok=True)
        
        # 配置日志
        logger.add(config.LOG_FILE, rotation="10 MB", encoding="utf-8")
    
    def take_screenshot(self, page: Page, name: str):
        """截图保存"""
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        path = os.path.join(config.SCREENSHOT_DIR, f"{name}_{ts}.png")
        page.screenshot(path=path, full_page=True)
        logger.info(f"截图已保存: {path}")
        return path
    
    def solve_captcha_with_retry(self, page: Page, captcha_img_selector: str, 
                                  captcha_input_selector: str) -> bool:
        """识别并填写图形验证码，支持重试
        
        Args:
            page: 页面对象
            captcha_img_selector: 验证码图片选择器
            captcha_input_selector: 验证码输入框选择器
        Returns:
            是否成功填写
        """
        for attempt in range(config.CAPTCHA_MAX_RETRY):
            try:
                # 等待验证码图片加载
                page.wait_for_selector(captcha_img_selector, timeout=10000)
                time.sleep(0.5)  # 等图片完全渲染
                
                # 识别验证码
                code = self.captcha.solve_from_element(page, captcha_img_selector)
                
                if not code or len(code) < 4:
                    logger.warning(f"验证码识别结果异常: {code}，点击刷新重试")
                    page.click(captcha_img_selector)  # 点击图片刷新验证码
                    time.sleep(1)
                    continue
                
                # 填入验证码
                page.fill(captcha_input_selector, code)
                logger.info(f"验证码已填入: {code} (第{attempt+1}次)")
                return True
                
            except Exception as e:
                logger.error(f"验证码处理失败(第{attempt+1}次): {e}")
                time.sleep(1)
        
        logger.error("验证码多次识别失败")
        return False
    
    # ==================== 联络员变更 ====================
    
    def change_liaison(self, page: Page, enterprise: dict) -> bool:
        """执行联络员变更

        Args:
            page: 页面对象
            enterprise: 企业信息字典，包含：
                - 注册号/统一社会信用代码
                - 法定代表人
                - 身份证
        Returns:
            变更是否成功
        """
        reg_no = enterprise.get("注册号/统一社会信用代码", enterprise.get("注册号", ""))
        logger.info(f"开始联络员变更: {enterprise.get('企业名称', '')} ({reg_no})")

        try:
            # 打开联络员变更页面
            page.goto(config.CHANGE_LIAISON_URL, wait_until="networkidle")
            time.sleep(2)

            # ---- 填写表单 ----
            # 注意：以下选择器需要根据实际页面HTML调整

            # 统一社会信用代码/注册号
            page.fill('input[name="regno"]', reg_no)

            # 法定代表人姓名
            page.fill('input[name="lerep"]', enterprise.get("法定代表人", ""))

            # 法定代表人证件号码
            page.fill('input[name="lerepCerno"]', enterprise.get("身份证", ""))

            # ---- 新联络员信息 ----
            page.fill('input[name="newLiaisonName"]', config.NEW_LIAISON["name"])
            
            # 联络员证件类型（下拉选择）
            page.select_option('select[name="newLiaisonCertype"]', 
                             label=config.NEW_LIAISON["id_type"])
            
            # 新联络员证件号码
            page.fill('input[name="newLiaisonCerno"]', config.NEW_LIAISON["id_number"])
            
            # 新联络员手机号码
            page.fill('input[name="newLiaisonPhone"]', config.NEW_LIAISON["phone"])
            
            # ---- 图形验证码 ----
            if not self.solve_captcha_with_retry(
                page, 
                'img#captchaImg',      # 验证码图片选择器 - 需调整
                'input[name="captcha"]' # 验证码输入框选择器 - 需调整
            ):
                return False
            
            # ---- 短信验证码 ----
            # 点击获取验证码
            page.click('button:has-text("获取验证码")')
            time.sleep(1)
            
            # 等待人工输入短信验证码
            sms_code = self.sms.wait_for_sms_code(
                config.NEW_LIAISON["phone"], 
                purpose=f"联络员变更-{enterprise.get('企业名称', '')}"
            )
            if not sms_code:
                logger.error("未获取到短信验证码")
                return False
            
            page.fill('input[name="smsCode"]', sms_code)
            
            # ---- 提交 ----
            page.click('button:has-text("保 存"), input[value="保 存"]')
            time.sleep(3)
            
            # 判断是否成功
            self.take_screenshot(page, f"change_liaison_{reg_no}")
            
            # 检查页面是否有成功提示
            if page.locator('text=成功').count() > 0 or page.locator('text=变更成功').count() > 0:
                logger.info(f"联络员变更成功: {reg_no}")
                return True
            else:
                page_text = page.inner_text("body")
                logger.warning(f"联络员变更结果不确定: {page_text[:200]}")
                return False
                
        except Exception as e:
            logger.error(f"联络员变更异常: {e}")
            self.take_screenshot(page, f"change_liaison_error_{reg_no}")
            return False
    
    # ==================== 联络员登录 ====================
    
    def login(self, page: Page, reg_no: str) -> bool:
        """联络员登录
        
        Args:
            page: 页面对象
            reg_no: 统一社会信用代码/注册号
        Returns:
            登录是否成功
        """
        logger.info(f"开始登录: {reg_no}")
        
        try:
            page.goto(config.LOGIN_URL, wait_until="networkidle")
            time.sleep(2)
            
            # 填入注册号
            page.fill('input[name="regno"]', reg_no)
            
            # 等待页面自动加载联络员信息
            time.sleep(2)
            
            # 图形验证码
            if not self.solve_captcha_with_retry(
                page,
                'img#captchaImg',       # 需调整
                'input[name="captcha"]'  # 需调整
            ):
                return False
            
            # 点击获取短信验证码
            page.click('button:has-text("获取验证码"), input[value="获取验证码"]')
            time.sleep(1)
            
            # 等待短信验证码
            sms_code = self.sms.wait_for_sms_code(
                config.NEW_LIAISON["phone"],
                purpose=f"登录-{reg_no}"
            )
            if not sms_code:
                return False
            
            page.fill('input[name="smsCode"]', sms_code)
            
            # 点击登录
            page.click('button:has-text("登录"), input[value="登录"]')
            time.sleep(3)
            
            # 判断登录结果
            self.take_screenshot(page, f"login_{reg_no}")
            
            # 检查是否进入年报页面（URL变化或页面内容变化）
            if "年报" in page.inner_text("body") or "企业基本信息" in page.inner_text("body"):
                logger.info(f"登录成功: {reg_no}")
                return True
            else:
                logger.warning(f"登录可能失败: {reg_no}")
                return False
                
        except Exception as e:
            logger.error(f"登录异常: {e}")
            self.take_screenshot(page, f"login_error_{reg_no}")
            return False
    
    # ==================== 填写年报 ====================
    
    def fill_annual_report(self, page: Page, reg_no: str, report_data: dict) -> bool:
        """填写年报表单
        
        Args:
            page: 页面对象（已登录状态）
            reg_no: 注册号
            report_data: 年报数据字典
        Returns:
            填写是否成功
        """
        logger.info(f"开始填写年报: {reg_no}")
        
        try:
            # ---- 企业基本信息 ----
            # 根据截图，登录后直接进入年报填写页面
            # 以下字段需要根据实际页面HTML调整选择器
            
            field_mapping = {
                # "Excel字段名": "页面input的name或selector"
                "企业联系电话": 'input[name="tel"]',
                "邮政编码": 'input[name="postalCode"]',
                "企业通信地址": 'input[name="address"]',
                "电子邮箱": 'input[name="email"]',
                # 以下字段需要根据年报Excel的实际字段补充
                # "从业人数": 'input[name="empNum"]',
                # "营业总收入": 'input[name="totalRevenue"]',
                # "利润总额": 'input[name="totalProfit"]',
                # "纳税总额": 'input[name="totalTax"]',
            }
            
            for excel_field, selector in field_mapping.items():
                value = report_data.get(excel_field, "")
                if value:
                    try:
                        page.fill(selector, value)
                        logger.debug(f"填入: {excel_field} = {value}")
                    except Exception as e:
                        logger.warning(f"字段填写失败: {excel_field} -> {e}")
            
            time.sleep(1)
            self.take_screenshot(page, f"fill_report_{reg_no}")
            
            # ---- 提交 ----
            # 注意：首次调试建议先不自动提交，改为手动确认
            # page.click('button:has-text("提交")')
            logger.info(f"年报填写完成（未自动提交）: {reg_no}")
            
            # 等待人工确认
            input(f"\n请检查填写内容，确认无误后按回车提交（企业: {reg_no}）...")
            
            # 提交
            page.click('button:has-text("提交"), input[value="提交"]')
            time.sleep(3)
            
            self.take_screenshot(page, f"submit_report_{reg_no}")
            
            if page.locator('text=成功').count() > 0:
                logger.info(f"年报提交成功: {reg_no}")
                return True
            else:
                logger.warning(f"年报提交结果不确定: {reg_no}")
                return False
                
        except Exception as e:
            logger.error(f"年报填写异常: {e}")
            self.take_screenshot(page, f"fill_report_error_{reg_no}")
            return False
    
    # ==================== 处理单个企业 ====================
    
    def process_enterprise(self, page: Page, enterprise: dict, 
                           report_data: dict, need_change_liaison: bool = True) -> dict:
        """处理单个企业的完整流程
        
        Args:
            page: 页面对象
            enterprise: 企业基本信息
            report_data: 该企业的年报数据
            need_change_liaison: 是否需要变更联络员
        Returns:
            处理结果字典
        """
        reg_no = enterprise.get("注册号/统一社会信用代码", enterprise.get("注册号", ""))
        name = enterprise.get("企业名称", "")
        
        result = {
            "企业名称": name,
            "注册号": reg_no,
            "联络员变更": "跳过",
            "登录": "未执行",
            "年报填写": "未执行",
            "时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        
        logger.info(f"\n{'='*60}")
        logger.info(f"开始处理: {name} ({reg_no})")
        logger.info(f"{'='*60}")
        
        # 步骤1：联络员变更
        if need_change_liaison:
            if self.change_liaison(page, enterprise):
                result["联络员变更"] = "成功"
            else:
                result["联络员变更"] = "失败"
                logger.error(f"联络员变更失败，跳过此企业: {reg_no}")
                return result
        
        # 步骤2：登录
        if self.login(page, reg_no):
            result["登录"] = "成功"
        else:
            result["登录"] = "失败"
            logger.error(f"登录失败，跳过此企业: {reg_no}")
            return result
        
        # 步骤3：填写年报
        if report_data:
            if self.fill_annual_report(page, reg_no, report_data):
                result["年报填写"] = "成功"
            else:
                result["年报填写"] = "失败"
        else:
            result["年报填写"] = "无数据"
            logger.warning(f"未找到年报数据: {reg_no}")
        
        return result
    
    # ==================== 主入口 ====================
    
    def run(self, start_index: int = 0, end_index: int = None, 
            need_change_liaison: bool = True):
        """运行批量处理
        
        Args:
            start_index: 起始行索引（从0开始）
            end_index: 结束行索引（None表示全部）
            need_change_liaison: 是否需要变更联络员
        """
        # 读取数据
        enterprises = read_enterprise_data(config.ENTERPRISE_EXCEL)
        report_data_map = read_annual_report_data(config.ANNUAL_REPORT_EXCEL)
        
        if end_index is None:
            end_index = len(enterprises)
        
        enterprises_to_process = enterprises[start_index:end_index]
        logger.info(f"本次处理 {len(enterprises_to_process)} 家企业 (索引 {start_index}-{end_index})")
        
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=config.HEADLESS,
                slow_mo=config.SLOW_MO,
            )
            context = browser.new_context(
                viewport={"width": 1280, "height": 900},
                locale="zh-CN",
            )
            page = context.new_page()
            page.set_default_timeout(config.TIMEOUT)
            
            for idx, enterprise in enumerate(enterprises_to_process):
                reg_no = enterprise.get("注册号/统一社会信用代码", 
                                       enterprise.get("注册号", ""))
                
                # 查找对应的年报数据
                report = report_data_map.get(reg_no, {})
                
                # 处理
                result = self.process_enterprise(
                    page, enterprise, report, need_change_liaison
                )
                self.results.append(result)
                
                # 打印进度
                total = len(enterprises_to_process)
                logger.info(f"进度: {idx+1}/{total} | "
                          f"联络员变更:{result['联络员变更']} "
                          f"登录:{result['登录']} "
                          f"年报:{result['年报填写']}")
                
                # 企业之间间隔，避免频率过高
                time.sleep(3)
            
            browser.close()
        
        # 保存结果
        self.save_results()
    
    def save_results(self):
        """保存处理结果"""
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        filepath = f"logs/results_{ts}.json"
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(self.results, f, ensure_ascii=False, indent=2)
        logger.info(f"结果已保存: {filepath}")
        
        # 打印汇总
        success = sum(1 for r in self.results if r["年报填写"] == "成功")
        failed = sum(1 for r in self.results if r["年报填写"] == "失败")
        skipped = sum(1 for r in self.results if r["年报填写"] in ("未执行", "无数据"))
        
        print(f"\n{'='*50}")
        print(f"处理完成！")
        print(f"  成功: {success}")
        print(f"  失败: {failed}")
        print(f"  跳过: {skipped}")
        print(f"  总计: {len(self.results)}")
        print(f"详细结果: {filepath}")
        print(f"{'='*50}")


def main():
    """单企业原型测试入口"""
    bot = AnnualReportBot()
    
    # 原型测试：只处理第1家企业
    bot.run(start_index=0, end_index=1, need_change_liaison=True)


if __name__ == "__main__":
    main()

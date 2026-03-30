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
    
    def _find_captcha_img(self, page: Page) -> str:
        """尝试多种选择器找到验证码图片"""
        selectors = [
            'img#vImg',
            'img[name="vImg"]',
            'img[id="vImg"]',
            'img.verifyImg',
            'img[src*="captcha"]',
            'img[src*="verify"]',
            'img[src*="vImg"]',
            'img[src*="code"]',
        ]
        for sel in selectors:
            try:
                if page.locator(sel).count() > 0:
                    logger.info(f"找到验证码图片: {sel}")
                    return sel
            except Exception:
                continue
        # 最后尝试用JS找所有img标签
        try:
            result = page.evaluate('''() => {
                const imgs = document.querySelectorAll('img');
                const info = [];
                for (const img of imgs) {
                    info.push({id: img.id, name: img.name, src: img.src.substring(0,80), className: img.className});
                }
                return info;
            }''')
            logger.info(f"页面上所有img标签: {result}")
        except Exception:
            pass
        return ""

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
            # 先打开登录页面
            page.goto(config.LOGIN_URL, wait_until="domcontentloaded")
            time.sleep(3)

            # 点击【联络员变更】链接，可能打开新标签页
            with page.context.expect_page() as new_page_info:
                page.click('a:has-text("联络员变更")')

            # 如果打开了新标签页，切换到新页面
            try:
                change_page = new_page_info.value
                change_page.wait_for_load_state("domcontentloaded")
                logger.info("联络员变更在新标签页打开")
            except Exception:
                # 没有新标签页，说明是在当前页面跳转
                change_page = page
                time.sleep(3)
                change_page.wait_for_load_state("domcontentloaded")
                logger.info("联络员变更在当前页面跳转")

            time.sleep(2)
            logger.info(f"当前页面URL: {change_page.url}")

            # ---- 填写表单 ----
            # 统一社会信用代码/注册号
            change_page.fill('input#regNo', reg_no)

            # 法定代表人姓名
            logger.info(f"填入法定代表人: {enterprise.get('法定代表人', '')}")
            change_page.fill('input[name="leRep"]', enterprise.get("法定代表人", ""))

            # 法定代表人证件号码
            logger.info(f"填入法定代表人证件号: {enterprise.get('身份证', '')[:4]}****")
            change_page.fill('input[name="certId"]', enterprise.get("身份证", ""))

            # ---- 新联络员信息（从Excel读取）----
            new_name = enterprise.get("新联络员姓名", "")
            new_id = enterprise.get("新联络员身份证", "")
            new_phone = enterprise.get("新联络员手机号", "")

            logger.info(f"填入新联络员姓名: {new_name}")
            change_page.fill('input[name="liaName_xin"]', new_name)

            # 联络员证件类型（下拉选择：中华人民共和国居民身份证，value="1"）
            logger.info("选择联络员证件类型: 中华人民共和国居民身份证")
            try:
                change_page.locator('select[name="certIdType_xin"]').scroll_into_view_if_needed()
                time.sleep(0.5)
                change_page.select_option('select[name="certIdType_xin"]', value="1")
                logger.info("证件类型选择完成（value=1）")
            except Exception as e1:
                logger.warning(f"select_option(value)失败: {e1}，尝试label方式")
                try:
                    change_page.select_option('select[name="certIdType_xin"]',
                                     label="中华人民共和国居民身份证")
                    logger.info("证件类型选择完成（label方式）")
                except Exception as e2:
                    logger.warning(f"label方式也失败: {e2}，尝试JS方式")
                    change_page.evaluate('''() => {
                        const sel = document.querySelector('select[name="certIdType_xin"]');
                        if (sel) { sel.value = "1"; sel.dispatchEvent(new Event("change", {bubbles:true})); }
                    }''')
                    logger.info("证件类型选择完成（JS方式）")
            time.sleep(0.5)

            # 新联络员证件号码
            logger.info(f"填入新联络员证件号: {new_id[:4]}****")
            change_page.fill('input[name="certId_xin"]', new_id)

            # 新联络员手机号码
            logger.info(f"填入新联络员手机号: {new_phone[:3]}****{new_phone[-3:]}")
            change_page.fill('input[name="mobileTel_xin"]', new_phone)

            logger.info("表单数据填入完成，开始处理验证码")

            # ---- 图形验证码 ----
            # 验证码图片可能是 img#vImg 或 img[name="vImg"] 或其他
            captcha_img_sel = self._find_captcha_img(change_page)
            if not captcha_img_sel:
                logger.error("找不到验证码图片元素")
                self.take_screenshot(change_page, f"no_captcha_{reg_no}")
                return False
            logger.info(f"验证码图片选择器: {captcha_img_sel}")
            if not self.solve_captcha_with_retry(
                change_page,
                captcha_img_sel,
                'input[name="verifyCodeTw"]'
            ):
                return False

            # ---- 短信验证码 ----
            # 点击获取验证码按钮
            change_page.click('input[value="获取验证码"], button:has-text("获取验证码")')
            time.sleep(1)

            # 等待人工输入短信验证码
            sms_code = self.sms.wait_for_sms_code(
                new_phone,
                purpose=f"联络员变更-{enterprise.get('企业名称', '')}"
            )
            if not sms_code:
                logger.error("未获取到短信验证码")
                return False

            change_page.fill('input[name="verifyCode"]', sms_code)

            # ---- 提交 ----
            change_page.click('input[value="保 存"], button:has-text("保 存")')
            time.sleep(3)

            # 判断是否成功
            self.take_screenshot(change_page, f"change_liaison_{reg_no}")
            
            # 检查页面是否有成功提示
            if change_page.locator('text=成功').count() > 0 or change_page.locator('text=变更成功').count() > 0:
                logger.info(f"联络员变更成功: {reg_no}")
                # 如果是新标签页，关闭它
                if change_page != page:
                    change_page.close()
                return True
            else:
                page_text = change_page.inner_text("body")
                logger.warning(f"联络员变更结果不确定: {page_text[:200]}")
                if change_page != page:
                    change_page.close()
                return False
                
        except Exception as e:
            logger.error(f"联络员变更异常: {e}")
            self.take_screenshot(page, f"change_liaison_error_{reg_no}")
            return False
    
    # ==================== 联络员登录 ====================
    
    def login(self, page: Page, reg_no: str, phone: str = "") -> bool:
        """联络员登录

        Args:
            page: 页面对象
            reg_no: 统一社会信用代码/注册号
            phone: 联络员手机号（用于提示短信验证码）
        Returns:
            登录是否成功
        """
        logger.info(f"开始登录: {reg_no}")

        try:
            page.goto(config.LOGIN_URL, wait_until="domcontentloaded")
            time.sleep(3)

            # 填入注册号
            page.fill('input#regNo', reg_no)

            # 等待页面自动加载联络员信息
            time.sleep(2)

            # 图形验证码
            captcha_img_sel = self._find_captcha_img(page)
            if not captcha_img_sel:
                logger.error("登录页找不到验证码图片")
                self.take_screenshot(page, f"no_captcha_login_{reg_no}")
                return False
            logger.info(f"登录页验证码图片选择器: {captcha_img_sel}")
            if not self.solve_captcha_with_retry(
                page,
                captcha_img_sel,
                'input[name="verifyCodeTw"]'
            ):
                return False

            # 点击获取短信验证码
            page.click('input[value="获取验证码"], button:has-text("获取验证码")')
            time.sleep(1)

            # 等待短信验证码
            sms_code = self.sms.wait_for_sms_code(
                phone,
                purpose=f"登录-{reg_no}"
            )
            if not sms_code:
                return False

            page.fill('input[name="verifyCode"]', sms_code)

            # 点击登录
            page.click('input[value="登录"], button:has-text("登录")')
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
        
        # 步骤2：登录（用新联络员手机号）
        phone = enterprise.get("新联络员手机号", "")
        if self.login(page, reg_no, phone):
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
                channel="msedge",
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

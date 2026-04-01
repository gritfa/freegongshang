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
from sms_handler import create_sms_handler


class AnnualReportBot:
    def __init__(self):
        self.captcha = CaptchaSolver()
        self.sms = create_sms_handler()
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
    
    def solve_captcha_with_retry(self, page, captcha_img_selector: str,
                                  captcha_input_selector: str) -> bool:
        """识别并填写图形验证码，支持重试。page可以是Page或Frame对象。"""
        for attempt in range(config.CAPTCHA_MAX_RETRY):
            try:
                # 等待验证码图片出现
                page.wait_for_selector(captcha_img_selector, timeout=10000)
                time.sleep(1)

                # 主要方式：Playwright截图
                code = None
                try:
                    img_el = page.locator(captcha_img_selector)
                    img_el.wait_for(timeout=10000)
                    time.sleep(0.5)
                    img_bytes = img_el.screenshot(timeout=15000)
                    code = self.captcha.solve_from_bytes(img_bytes)
                    logger.info(f"截图识别验证码: {code}")
                except Exception as ss_err:
                    logger.warning(f"截图方式失败: {ss_err}")
                    # 备选：用JS获取base64图片数据
                    try:
                        b64_data = page.evaluate('''() => {
                            var img = document.getElementById("vimg");
                            if (!img) return null;
                            var canvas = document.createElement("canvas");
                            canvas.width = img.naturalWidth || img.width;
                            canvas.height = img.naturalHeight || img.height;
                            var ctx = canvas.getContext("2d");
                            ctx.drawImage(img, 0, 0);
                            return canvas.toDataURL("image/png").split(",")[1];
                        }''')
                        if b64_data:
                            import base64
                            raw = base64.b64decode(b64_data)
                            code = self.captcha.solve_from_bytes(raw)
                            logger.info(f"JS canvas识别验证码: {code}")
                    except Exception as js_err:
                        logger.warning(f"JS canvas方式也失败: {js_err}")

                if not code or len(code) < 3:
                    logger.warning(f"验证码识别异常: {code}，刷新重试")
                    page.evaluate('() => { var img = document.getElementById("vimg"); if(img) img.click(); }')
                    time.sleep(1)
                    continue

                # 用JS直接填入验证码输入框（根据传入的selector查找）
                # 从selector中提取id，如 "input#verifyTxCode" -> "verifyTxCode"
                input_id = captcha_input_selector.split('#')[-1] if '#' in captcha_input_selector else ""

                # 先尝试用Playwright的type方法（更可靠）
                filled_by_type = False
                try:
                    input_el = page.locator(captcha_input_selector)
                    if input_el.count() > 0:
                        input_el.click(timeout=3000)
                        time.sleep(0.2)
                        input_el.fill('', timeout=3000)
                        input_el.type(code, delay=50, timeout=5000)
                        actual = input_el.input_value(timeout=3000)
                        logger.info(f"验证码Playwright type填入: 期望={code} 实际={actual}")
                        if actual == code:
                            filled_by_type = True
                except Exception as type_err:
                    logger.warning(f"验证码Playwright type失败: {type_err}")

                if filled_by_type:
                    logger.info(f"验证码填入成功(type方式) (第{attempt+1}次)")
                    return True

                # 备选：用JS填入
                fill_result = page.evaluate('''([code, inputId, selector]) => {
                    var el = null;
                    // 优先用id查找
                    if (inputId) {
                        el = document.getElementById(inputId);
                    }
                    // 备选：用querySelector
                    if (!el) {
                        el = document.querySelector(selector);
                    }
                    // 再备选：按name查找
                    if (!el) {
                        var inputs = document.querySelectorAll("input");
                        for (var i = 0; i < inputs.length; i++) {
                            if (inputs[i].id === inputId || inputs[i].name === inputId) {
                                el = inputs[i];
                                break;
                            }
                        }
                    }
                    // 再备选：按placeholder查找验证码输入框
                    if (!el) {
                        var inputs = document.querySelectorAll("input[type='text']");
                        for (var i = 0; i < inputs.length; i++) {
                            if (inputs[i].placeholder && inputs[i].placeholder.indexOf("验证码") >= 0) {
                                el = inputs[i];
                                break;
                            }
                        }
                    }
                    if (el) {
                        el.focus();
                        el.value = code;
                        el.dispatchEvent(new Event("input", {bubbles: true}));
                        el.dispatchEvent(new Event("change", {bubbles: true}));
                        return el.value;
                    }
                    // 返回调试信息
                    var allInputs = [];
                    document.querySelectorAll("input").forEach(function(inp) {
                        allInputs.push(inp.id + "/" + inp.name + "/" + inp.type);
                    });
                    return "NOT_FOUND|inputs:" + allInputs.join(",");
                }''', [code, input_id, captcha_input_selector])

                logger.info(f"验证码填入结果: 识别={code}, 填入后={fill_result} (第{attempt+1}次)")

                if fill_result == code:
                    return True
                elif fill_result == "NOT_FOUND":
                    logger.error("验证码输入框未找到！")
                else:
                    logger.warning(f"填入值不一致: {fill_result}")
                    return True  # 值已设置，继续流程

            except Exception as e:
                logger.error(f"验证码处理失败(第{attempt+1}次): {e}")
                time.sleep(1)

        logger.error("验证码多次识别失败")
        return False
    
    # ==================== 联络员变更 ====================
    
    def change_liaison(self, page: Page, enterprise: dict):
        """执行联络员变更

        Returns:
            成功时返回当前活动页面（可能是原page或变更后的页面），失败返回None
        """
        reg_no = enterprise.get("注册号/统一社会信用代码", enterprise.get("注册号", ""))
        logger.info(f"开始联络员变更: {enterprise.get('企业名称', '')} ({reg_no})")

        try:
            # 先打开登录页面
            page.goto(config.LOGIN_URL, wait_until="domcontentloaded")
            time.sleep(3)

            # 记录当前页面数量，用于检测是否打开了新标签页
            pages_before = len(page.context.pages)

            # 点击【联络员变更】链接
            page.click('a:has-text("联络员变更")')
            time.sleep(3)

            # 检查是否打开了新标签页
            all_pages = page.context.pages
            if len(all_pages) > pages_before:
                change_page = all_pages[-1]
                change_page.wait_for_load_state("domcontentloaded")
                logger.info("联络员变更在新标签页打开")
            else:
                change_page = page
                change_page.wait_for_load_state("domcontentloaded")
                logger.info("联络员变更在当前页面跳转")

            time.sleep(3)
            logger.info(f"当前页面URL: {change_page.url}")

            # 检查是否有iframe
            iframe_count = change_page.evaluate('document.querySelectorAll("iframe").length')
            logger.info(f"页面iframe数量: {iframe_count}")

            form_context = change_page  # 默认用页面本身
            if iframe_count > 0:
                # 获取iframe信息
                iframe_info = change_page.evaluate('''() => {
                    var iframes = document.querySelectorAll("iframe");
                    var result = [];
                    for (var i = 0; i < iframes.length; i++) {
                        result.push({id: iframes[i].id, name: iframes[i].name, src: iframes[i].src});
                    }
                    return result;
                }''')
                logger.info(f"iframe信息: {iframe_info}")

                # 尝试切换到第一个有内容的iframe
                frames = change_page.frames
                for frame in frames:
                    if frame != change_page.main_frame:
                        # 检查这个frame里是否有表单元素
                        try:
                            has_form = frame.evaluate('!!document.querySelector("input#regNo") || !!document.querySelector("input[name=\\"regNo\\"]")')
                            if has_form:
                                form_context = frame
                                logger.info(f"找到表单所在iframe: {frame.url}")
                                break
                        except:
                            pass

                if form_context == change_page:
                    logger.warning("未在iframe中找到表单，继续在主页面操作")

            # 等待表单元素加载完成
            logger.info("等待表单加载...")
            try:
                form_context.wait_for_selector('input#regNo', timeout=15000)
                logger.info("表单已加载")
            except Exception:
                logger.warning("等待表单超时，等5秒再试")
                time.sleep(5)

            time.sleep(2)

            # ---- 用JS填写所有表单字段（最可靠的方式）----
            new_name = enterprise.get("新联络员姓名", "")
            new_id = enterprise.get("新联络员身份证", "")
            new_phone = enterprise.get("新联络员手机号", "")

            # 先单独填注册号（填入后页面可能会AJAX加载其他字段）
            fields_to_fill = [
                ("regNo", reg_no, "注册号"),
            ]
            # 注册号以外的字段，在regNo填入后再处理
            other_fields = [
                ("leRep", enterprise.get("法定代表人", ""), "法定代表人"),
                ("certId", enterprise.get("身份证", ""), "法定代表人证件号"),
                ("liaName_xin", new_name, "新联络员姓名"),
                ("certId_xin", new_id, "新联络员证件号"),
                ("mobileTel_xin", new_phone, "新联络员手机号"),
            ]

            for field_name, field_value, field_label in fields_to_fill:
                if not field_value:
                    logger.warning(f"字段为空，跳过: {field_label}")
                    continue

                # 用JS查找元素（同时尝试id和name两种方式）
                el_info = form_context.evaluate(f'''() => {{
                    var el = document.getElementById("{field_name}");
                    if (!el) el = document.querySelector('input[name="{field_name}"]');
                    if (!el) return null;
                    return {{
                        tag: el.tagName,
                        type: el.type || "",
                        visible: !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length),
                        id: el.id || "",
                        name: el.name || ""
                    }};
                }}''')

                if not el_info:
                    # 如果regNo已填入，等待页面加载后再试一次
                    if field_name != "regNo":
                        logger.info(f"元素暂未找到，等待3秒后重试: {field_label} ({field_name})")
                        time.sleep(3)
                        el_info = form_context.evaluate(f'''() => {{
                            var el = document.getElementById("{field_name}");
                            if (!el) el = document.querySelector('input[name="{field_name}"]');
                            if (!el) return null;
                            return {{
                                tag: el.tagName,
                                type: el.type || "",
                                visible: !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length),
                                id: el.id || "",
                                name: el.name || ""
                            }};
                        }}''')
                    if not el_info:
                        # 打印页面上所有input元素的信息用于调试
                        all_inputs = form_context.evaluate('''() => {
                            var inputs = document.querySelectorAll("input");
                            var result = [];
                            inputs.forEach(function(inp) {
                                result.push(inp.id + "|" + inp.name + "|" + inp.type);
                            });
                            return result.join(", ");
                        }''')
                        logger.warning(f"元素不存在: {field_label} ({field_name})，页面所有input: {all_inputs}")
                        continue

                logger.info(f"找到元素: {field_label} ({field_name}) -> id={el_info.get('id')}, name={el_info.get('name')}, type={el_info.get('type')}, visible={el_info.get('visible')}")

                is_visible = el_info.get("visible", False)
                # 构建有效的CSS选择器（优先用id）
                if el_info.get("id"):
                    css_sel = f'#{el_info["id"]}'
                else:
                    css_sel = f'input[name="{el_info["name"]}"]'

                if is_visible:
                    # 可见元素：点击+键盘输入
                    try:
                        form_context.click(css_sel)
                        time.sleep(0.3)
                        form_context.fill(css_sel, "")
                        form_context.type(css_sel, str(field_value), delay=50)
                        time.sleep(0.5)
                        actual = form_context.input_value(css_sel)
                        if actual == str(field_value):
                            logger.info(f"填入成功(键盘): {field_label} = {str(field_value)[:6]}...")
                        else:
                            raise Exception(f"值不一致: {actual}")
                    except Exception as e:
                        logger.warning(f"键盘方式失败({e})，用JS填入")
                        # 键盘方式失败，用JS
                        form_context.evaluate(f'''() => {{
                            var el = document.getElementById("{field_name}") || document.querySelector('input[name="{field_name}"]');
                            if(el) {{ el.value = "{field_value}"; el.dispatchEvent(new Event("input",{{bubbles:true}})); el.dispatchEvent(new Event("change",{{bubbles:true}})); }}
                        }}''')
                        logger.info(f"填入成功(JS-可见): {field_label} = {str(field_value)[:6]}...")
                else:
                    # 隐藏元素：直接用JS设值
                    form_context.evaluate(f'''() => {{
                        var el = document.getElementById("{field_name}") || document.querySelector('input[name="{field_name}"]');
                        if(el) {{ el.value = "{field_value}"; el.dispatchEvent(new Event("input",{{bubbles:true}})); el.dispatchEvent(new Event("change",{{bubbles:true}})); }}
                    }}''')
                    # 验证
                    actual = form_context.evaluate(f'''() => {{
                        var el = document.getElementById("{field_name}") || document.querySelector('input[name="{field_name}"]');
                        return el ? el.value : "NOT_FOUND";
                    }}''')
                    if actual == str(field_value):
                        logger.info(f"填入成功(JS-隐藏): {field_label} = {str(field_value)[:6]}...")
                    else:
                        logger.warning(f"JS填入可能失败: {field_label} 期望={str(field_value)[:6]} 实际={str(actual)[:6] if actual else 'empty'}")
                time.sleep(0.5)

            # regNo填入后，用Tab键自然触发blur事件（比JS dispatch更可靠）
            logger.info("注册号已填入，按Tab键触发blur事件...")
            try:
                form_context.press('input#regNo', 'Tab')
            except Exception:
                pass
            time.sleep(1)
            # 再用JS补充触发事件
            form_context.evaluate('''() => {
                var el = document.getElementById("regNo") || document.querySelector('input[name="regNo"]');
                if (el) {
                    el.dispatchEvent(new Event("blur", {bubbles: true}));
                    el.dispatchEvent(new Event("change", {bubbles: true}));
                    // 尝试触发可能的onblur处理函数
                    if (el.onblur) el.onblur();
                    if (el.onchange) el.onchange();
                }
            }''')

            # 动态等待：检查_xin字段是否出现，最多等30秒
            # 同时在当前form_context、主页面、所有iframe中搜索
            xin_found = False
            for wait_i in range(15):
                time.sleep(2)

                # 先在当前form_context中查找
                try:
                    has_xin = form_context.evaluate('''() => {
                        return !!(document.getElementById("liaName_xin") ||
                                  document.querySelector('input[name="liaName_xin"]'));
                    }''')
                    if has_xin:
                        logger.info(f"AJAX加载完成，_xin字段在当前form_context中出现（等待{(wait_i+1)*2}秒）")
                        xin_found = True
                        break
                except Exception:
                    pass

                # 在主页面中查找
                try:
                    has_xin_main = change_page.evaluate('''() => {
                        return !!(document.getElementById("liaName_xin") ||
                                  document.querySelector('input[name="liaName_xin"]'));
                    }''')
                    if has_xin_main:
                        form_context = change_page
                        logger.info(f"_xin字段在主页面中找到，切换form_context（等待{(wait_i+1)*2}秒）")
                        xin_found = True
                        break
                except Exception:
                    pass

                # 在所有iframe中查找
                for frame in change_page.frames:
                    if frame == change_page.main_frame:
                        continue
                    try:
                        has_xin_frame = frame.evaluate('''() => {
                            return !!(document.getElementById("liaName_xin") ||
                                      document.querySelector('input[name="liaName_xin"]'));
                        }''')
                        if has_xin_frame:
                            form_context = frame
                            logger.info(f"_xin字段在iframe中找到，切换form_context: {frame.url}（等待{(wait_i+1)*2}秒）")
                            xin_found = True
                            break
                    except Exception:
                        pass
                if xin_found:
                    break

                # 也用更宽泛的选择器搜索（name包含"xin"的input）
                try:
                    xin_inputs = change_page.evaluate('''() => {
                        var inputs = document.querySelectorAll('input[name*="xin"]');
                        var result = [];
                        inputs.forEach(function(inp) {
                            result.push(inp.id + "|" + inp.name + "|" + inp.type);
                        });
                        return result.join(", ");
                    }''')
                    if xin_inputs:
                        logger.info(f"主页面中包含'xin'的input: {xin_inputs}")
                except Exception:
                    pass

                logger.info(f"等待AJAX加载中... ({(wait_i+1)*2}秒)")

            if not xin_found:
                # 最后尝试：打印所有页面和iframe中的全部input元素用于调试
                logger.warning("等待30秒后_xin字段仍未出现")
                try:
                    all_inputs_main = change_page.evaluate('''() => {
                        var inputs = document.querySelectorAll("input, select, textarea");
                        var result = [];
                        inputs.forEach(function(el) {
                            result.push(el.tagName + "#" + el.id + "|name=" + el.name + "|type=" + el.type);
                        });
                        return result.join("\\n");
                    }''')
                    logger.info(f"主页面所有表单元素:\\n{all_inputs_main}")
                except Exception:
                    pass
                for frame in change_page.frames:
                    if frame == change_page.main_frame:
                        continue
                    try:
                        all_inputs_frame = frame.evaluate('''() => {
                            var inputs = document.querySelectorAll("input, select, textarea");
                            var result = [];
                            inputs.forEach(function(el) {
                                result.push(el.tagName + "#" + el.id + "|name=" + el.name + "|type=" + el.type);
                            });
                            return result.join("\\n");
                        }''')
                        logger.info(f"iframe({frame.url})所有表单元素:\\n{all_inputs_frame}")
                    except Exception:
                        pass
                logger.warning("尝试继续填入...")

            time.sleep(1)

            # 填入其余字段
            for field_name, field_value, field_label in other_fields:
                if not field_value:
                    logger.warning(f"字段为空，跳过: {field_label}")
                    continue

                # 用JS查找元素（同时尝试id和name两种方式）
                el_info = form_context.evaluate(f'''() => {{
                    var el = document.getElementById("{field_name}");
                    if (!el) el = document.querySelector('input[name="{field_name}"]');
                    if (!el) return null;
                    return {{
                        tag: el.tagName,
                        type: el.type || "",
                        visible: !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length),
                        id: el.id || "",
                        name: el.name || ""
                    }};
                }}''')

                if not el_info:
                    logger.info(f"元素暂未找到，等待3秒后重试: {field_label} ({field_name})")
                    time.sleep(3)
                    el_info = form_context.evaluate(f'''() => {{
                        var el = document.getElementById("{field_name}");
                        if (!el) el = document.querySelector('input[name="{field_name}"]');
                        if (!el) return null;
                        return {{
                            tag: el.tagName,
                            type: el.type || "",
                            visible: !!(el.offsetWidth || el.offsetHeight || el.getClientRects().length),
                            id: el.id || "",
                            name: el.name || ""
                        }};
                    }}''')
                    if not el_info:
                        all_inputs = form_context.evaluate('''() => {
                            var inputs = document.querySelectorAll("input");
                            var result = [];
                            inputs.forEach(function(inp) {
                                result.push(inp.id + "|" + inp.name + "|" + inp.type);
                            });
                            return result.join(", ");
                        }''')
                        logger.warning(f"元素不存在: {field_label} ({field_name})，页面所有input: {all_inputs}")
                        continue

                logger.info(f"找到元素: {field_label} ({field_name}) -> id={el_info.get('id')}, name={el_info.get('name')}, type={el_info.get('type')}, visible={el_info.get('visible')}")

                is_visible = el_info.get("visible", False)
                if el_info.get("id"):
                    css_sel = f'#{el_info["id"]}'
                else:
                    css_sel = f'input[name="{el_info["name"]}"]'

                if is_visible:
                    try:
                        form_context.click(css_sel)
                        time.sleep(0.3)
                        form_context.fill(css_sel, "")
                        form_context.type(css_sel, str(field_value), delay=50)
                        time.sleep(0.5)
                        actual = form_context.input_value(css_sel)
                        if actual == str(field_value):
                            logger.info(f"填入成功(键盘): {field_label} = {str(field_value)[:6]}...")
                        else:
                            raise Exception(f"值不一致: {actual}")
                    except Exception as e:
                        logger.warning(f"键盘方式失败({e})，用JS填入")
                        form_context.evaluate(f'''() => {{
                            var el = document.getElementById("{field_name}") || document.querySelector('input[name="{field_name}"]');
                            if(el) {{ el.value = "{field_value}"; el.dispatchEvent(new Event("input",{{bubbles:true}})); el.dispatchEvent(new Event("change",{{bubbles:true}})); }}
                        }}''')
                        logger.info(f"填入成功(JS-可见): {field_label} = {str(field_value)[:6]}...")
                else:
                    form_context.evaluate(f'''() => {{
                        var el = document.getElementById("{field_name}") || document.querySelector('input[name="{field_name}"]');
                        if(el) {{ el.value = "{field_value}"; el.dispatchEvent(new Event("input",{{bubbles:true}})); el.dispatchEvent(new Event("change",{{bubbles:true}})); }}
                    }}''')
                    actual = form_context.evaluate(f'''() => {{
                        var el = document.getElementById("{field_name}") || document.querySelector('input[name="{field_name}"]');
                        return el ? el.value : "NOT_FOUND";
                    }}''')
                    if actual == str(field_value):
                        logger.info(f"填入成功(JS-隐藏): {field_label} = {str(field_value)[:6]}...")
                    else:
                        logger.warning(f"JS填入可能失败: {field_label} 期望={str(field_value)[:6]} 实际={str(actual)[:6] if actual else 'empty'}")
                time.sleep(0.5)

            # ---- 下拉框：选择中华人民共和国居民身份证 ----
            logger.info("选择联络员证件类型: 中华人民共和国居民身份证")
            # 先用Playwright原生select_option
            try:
                form_context.select_option('select#cerIdType_xin', value="1")
                selected = form_context.input_value('select#cerIdType_xin')
                logger.info(f"下拉框Playwright选择结果: value={selected}")
            except Exception as e:
                logger.warning(f"Playwright select_option失败: {e}，用JS方式")
            # 不管上面是否成功，都用JS再设置一次确保生效
            select_result = form_context.evaluate('''() => {
                var sel = document.getElementById("cerIdType_xin");
                if (!sel) {
                    var selects = document.querySelectorAll("select");
                    for (var i = 0; i < selects.length; i++) {
                        var n = (selects[i].name || "").toLowerCase();
                        if (n.indexOf("cerid") >= 0) { sel = selects[i]; break; }
                    }
                }
                if (!sel) return "SELECT_NOT_FOUND";
                sel.value = "1";
                sel.selectedIndex = 1;
                sel.dispatchEvent(new Event("change", {bubbles: true}));
                return "设置值=" + sel.value + " 文本=" + sel.options[sel.selectedIndex].text;
            }''')
            logger.info(f"下拉框JS设置结果: {select_result}")

            time.sleep(1)
            logger.info("表单数据填入完成，开始处理验证码")

            # ---- 图形验证码 ----
            if not self.solve_captcha_with_retry(
                form_context,
                'img#vimg',
                'input#verifyCodetw'
            ):
                return False

            # ---- 短信验证码 ----
            logger.info("点击获取验证码按钮")
            # 按钮是 a#butn 里面套了 <img>，直接用JS调用 onclick 函数最可靠
            try:
                form_context.evaluate('getCode2()')
                logger.info("获取验证码: JS getCode2() 调用成功")
            except Exception as e1:
                logger.warning(f"JS getCode2() 失败: {e1}，尝试点击按钮")
                try:
                    form_context.evaluate('document.getElementById("butn").click()')
                    logger.info("获取验证码: JS click butn 成功")
                except Exception as e2:
                    logger.error(f"获取验证码按钮全部失败: {e2}")
            time.sleep(2)

            sms_code = self.sms.wait_for_sms_code(
                new_phone,
                purpose=f"联络员变更-{enterprise.get('企业名称', '')}"
            )
            if not sms_code:
                logger.error("未获取到短信验证码")
                return False

            logger.info(f"准备填入短信验证码: {sms_code}")

            # 先用type方式填入短信验证码（和图形验证码一样的方式）
            try:
                form_context.click('input#verifyCode')
                time.sleep(0.3)
                form_context.fill('input#verifyCode', '')
                form_context.type('input#verifyCode', sms_code, delay=50)
                actual = form_context.input_value('input#verifyCode')
                logger.info(f"短信验证码type方式填入: 期望={sms_code} 实际={actual}")
                if actual != sms_code:
                    raise Exception("值不一致")
            except Exception as e:
                logger.warning(f"type方式失败({e})，用JS填入")
                # JS方式填入
                sms_fill_result = form_context.evaluate(f'''() => {{
                    var el = document.getElementById("verifyCode");
                    if (!el) {{
                        var inputs = document.querySelectorAll("input");
                        for (var i=0;i<inputs.length;i++) {{ if(inputs[i].name=="verifyCode") {{ el=inputs[i]; break; }} }}
                    }}
                    if (el) {{
                        el.focus();
                        el.value = "{sms_code}";
                        el.dispatchEvent(new Event("input", {{bubbles:true}}));
                        el.dispatchEvent(new Event("change", {{bubbles:true}}));
                        return "JS填入成功: " + el.value;
                    }}
                    return "NOT_FOUND";
                }}''')
                logger.info(f"短信验证码JS填入结果: {sms_fill_result}")

            # ---- 提交保存 ----
            logger.info("点击保存按钮 (a#subBtn)")
            save_clicked = False
            # 保存按钮是 <a type="button" id="subBtn"><img src="...add2.png"></a>
            try:
                change_page.click('a#subBtn', timeout=5000)
                save_clicked = True
                logger.info("保存按钮点击成功: a#subBtn")
            except Exception as e:
                logger.warning(f"a#subBtn点击失败: {e}")

            if not save_clicked:
                # 直接用JS点击
                logger.info("尝试用JS点击保存按钮")
                js_result = change_page.evaluate('''() => {
                    var btn = document.getElementById("subBtn");
                    if (btn) { btn.click(); return "clicked_subBtn"; }
                    // 兜底：找所有a标签里的img
                    var links = document.querySelectorAll("a[type='button']");
                    for (var i = 0; i < links.length; i++) {
                        var img = links[i].querySelector("img");
                        if (img && img.src && img.src.includes("add")) {
                            links[i].click();
                            return "clicked_a_img";
                        }
                    }
                    return "NOT_FOUND";
                }''')
                logger.info(f"JS保存按钮结果: {js_result}")
                if "NOT_FOUND" not in js_result:
                    save_clicked = True

            time.sleep(3)

            self.take_screenshot(change_page, f"change_liaison_{reg_no}")

            # 页面跳转了就说明保存成功（跳回登录页）
            current_url = change_page.url
            page_text = change_page.inner_text("body")
            if ("成功" in page_text or "变更成功" in page_text or
                "联络员登录" in page_text or "liaisonsLogin" in current_url or
                current_url != config.CHANGE_LIAISON_URL):
                logger.info(f"联络员变更成功（页面已跳转）: {reg_no}")
                # 返回当前活动页面（不关闭，因为可能已跳转到登录页）
                return change_page
            else:
                logger.warning(f"联络员变更结果不确定: {page_text[:200]}")
                return None

        except Exception as e:
            logger.error(f"联络员变更异常: {e}")
            self.take_screenshot(page, f"change_liaison_error_{reg_no}")
            return None
    
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
            # 检查当前页面是否已经在登录页
            current_url = page.url
            if "liaisonsLogin" not in current_url:
                logger.info(f"当前不在登录页({current_url})，跳转到登录页")
                page.goto(config.LOGIN_URL, wait_until="domcontentloaded")
                time.sleep(3)
            else:
                logger.info("已在登录页，直接操作")
                time.sleep(2)

            # 关闭蓝色弹窗（如果存在）
            try:
                close_btn = page.locator('div#divClose')
                if close_btn.count() > 0:
                    close_btn.click(timeout=3000)
                    logger.info("已关闭蓝色弹窗")
                    time.sleep(1)
            except Exception:
                # 也尝试JS方式关闭
                try:
                    page.evaluate('''() => {
                        var el = document.getElementById("divClose");
                        if(el) el.click();
                        // 也尝试隐藏整个弹窗容器
                        document.querySelectorAll('[class*="float"],[class*="notice"],[class*="popup"],[class*="tip"]').forEach(e => e.style.display="none");
                    }''')
                    logger.info("JS方式关闭弹窗")
                except Exception:
                    logger.warning("未找到弹窗或已关闭")
                time.sleep(0.5)

            # 关闭 layui-layer 弹窗（如果存在）
            try:
                page.evaluate('''() => {
                    // 关闭所有 layui-layer 弹窗
                    document.querySelectorAll("a.layui-layer-close").forEach(btn => btn.click());
                    // 隐藏遮罩层
                    document.querySelectorAll("div.layui-layer-shade").forEach(el => el.style.display = "none");
                    // 隐藏弹窗本身
                    document.querySelectorAll("div.layui-layer").forEach(el => el.style.display = "none");
                }''')
                logger.info("已处理 layui-layer 弹窗")
                time.sleep(1)
            except Exception:
                logger.debug("无 layui-layer 弹窗或已关闭")

            # 点击"联络员登录"标签页
            try:
                page.click('a#denglu-a2', timeout=5000)
                logger.info("已点击联络员登录标签")
                time.sleep(2)
            except Exception as e:
                logger.warning(f"点击联络员登录标签失败({e})，可能已在该标签页")

            # 关闭操作指引开关（联络员登录标签点击后执行）
            try:
                time.sleep(1)
                # 方法1: Playwright locator点击
                czybtn = page.locator('input#czybtn')
                if czybtn.count() > 0:
                    if czybtn.is_checked():
                        czybtn.click()
                        logger.info("操作指引已关闭（Playwright locator）")
                        time.sleep(1)
                    else:
                        logger.info("操作指引已经是关闭状态")
                else:
                    # 方法2: 在所有frame中查找
                    closed = False
                    for frame in page.frames:
                        try:
                            found = frame.evaluate('() => !!document.getElementById("czybtn")')
                            if found:
                                is_checked = frame.evaluate('document.getElementById("czybtn").checked')
                                if is_checked:
                                    frame.evaluate('document.getElementById("czybtn").click()')
                                    logger.info("操作指引已关闭（iframe）")
                                    time.sleep(1)
                                else:
                                    logger.info("操作指引已经是关闭状态（iframe）")
                                closed = True
                                break
                        except Exception:
                            continue
                    if not closed:
                        # 方法3: 用文本匹配找操作指引开关
                        try:
                            page.evaluate('''() => {
                                var inputs = document.querySelectorAll('input[type="checkbox"]');
                                for(var i=0; i<inputs.length; i++) {
                                    var el = inputs[i];
                                    var parent = el.parentElement;
                                    if(parent && parent.textContent && parent.textContent.includes("操作指引")) {
                                        if(el.checked) el.click();
                                    }
                                }
                            }''')
                            logger.info("操作指引: 用文本匹配方式尝试关闭")
                        except Exception:
                            pass
                        logger.info("未通过id找到操作指引开关，已用文本匹配尝试")
            except Exception as e:
                logger.warning(f"关闭操作指引失败: {e}，继续执行")

            # 勾选协议复选框
            try:
                checkbox = page.locator('input#czzybut')
                if checkbox.count() > 0 and not checkbox.is_checked():
                    checkbox.check()
                    logger.info("已勾选协议复选框")
            except Exception as e:
                logger.warning(f"勾选协议复选框失败({e})，继续")

            # 等待注册号输入框出现
            try:
                page.wait_for_selector('input#regNo', timeout=15000)
                logger.info("登录页表单已加载")
            except Exception:
                logger.warning("等待登录页表单超时，等3秒再试")
                time.sleep(3)

            # 填入注册号 — 用JS直接设值，不点击输入框（避免触发弹窗）
            page.evaluate(f'''() => {{
                var el = document.getElementById("regNo");
                if(el) {{
                    // 先去掉可能触发弹窗的事件
                    el.onclick = null;
                    el.onfocus = null;
                    el.removeAttribute("onclick");
                    el.removeAttribute("onfocus");
                    // 直接设值
                    el.value = "{reg_no}";
                    el.dispatchEvent(new Event("input", {{bubbles:true}}));
                    el.dispatchEvent(new Event("change", {{bubbles:true}}));
                    el.dispatchEvent(new Event("blur", {{bubbles:true}}));
                }}
            }}''')
            actual = page.evaluate('() => document.getElementById("regNo") ? document.getElementById("regNo").value : ""')
            logger.info(f"注册号JS填入: 期望={reg_no[:8]}... 实际={actual[:8] if actual else 'EMPTY'}...")

            # 等待页面自动加载联络员信息
            time.sleep(3)

            # 再次检查并关闭可能出现的 layui-layer 弹窗
            try:
                page.evaluate('''() => {
                    document.querySelectorAll("a.layui-layer-close").forEach(btn => btn.click());
                    document.querySelectorAll("div.layui-layer-shade").forEach(el => el.style.display = "none");
                    document.querySelectorAll("div.layui-layer").forEach(el => el.style.display = "none");
                }''')
            except Exception:
                pass
            time.sleep(0.5)

            # 图形验证码（登录页用 verifyTxCode，不是联络员变更页的 verifyCodetw）
            # 先诊断页面状态：列出所有input元素和iframe
            try:
                diag = page.evaluate('''() => {
                    var result = {inputs: [], iframes: [], url: location.href};
                    document.querySelectorAll("input").forEach(el => {
                        result.inputs.push({id: el.id, name: el.name, type: el.type, visible: el.offsetParent !== null});
                    });
                    document.querySelectorAll("iframe").forEach(el => {
                        result.iframes.push({id: el.id, name: el.name, src: el.src});
                    });
                    return result;
                }''')
                logger.info(f"登录页诊断 - URL: {diag.get('url', 'N/A')}")
                logger.info(f"登录页诊断 - 所有input: {json.dumps(diag.get('inputs', []), ensure_ascii=False)}")
                if diag.get('iframes'):
                    logger.info(f"登录页诊断 - iframe: {json.dumps(diag.get('iframes', []), ensure_ascii=False)}")
            except Exception as diag_err:
                logger.warning(f"登录页诊断失败: {diag_err}")

            # 检查verifyTxCode是否在iframe中
            captcha_page = page  # 默认在主页面
            try:
                has_input = page.evaluate('() => !!document.getElementById("verifyTxCode")')
                if not has_input:
                    logger.warning("主页面未找到verifyTxCode，检查iframe...")
                    for frame in page.frames:
                        try:
                            has_in_frame = frame.evaluate('() => !!document.getElementById("verifyTxCode")')
                            if has_in_frame:
                                logger.info(f"在iframe中找到verifyTxCode: {frame.url}")
                                captcha_page = frame
                                break
                        except Exception:
                            continue
                    if captcha_page == page:
                        # 仍未找到，等待更长时间后再试
                        logger.warning("iframe中也未找到，等待5秒后重试...")
                        time.sleep(5)
                        # 再次尝试点击联络员登录tab
                        try:
                            page.click('a#denglu-a2', timeout=3000)
                            time.sleep(2)
                        except Exception:
                            pass
                else:
                    logger.info("主页面已找到verifyTxCode")
            except Exception as e:
                logger.warning(f"检查verifyTxCode位置失败: {e}")

            # 等待验证码输入框出现
            try:
                captcha_page.wait_for_selector('input#verifyTxCode', timeout=10000)
                logger.info("验证码输入框已就绪")
            except Exception:
                logger.warning("等待verifyTxCode超时，继续尝试...")

            if not self.solve_captcha_with_retry(
                captcha_page,
                'img#vimg',
                'input#verifyTxCode'
            ):
                return False

            # 点击获取短信验证码 — 按钮onclick是hyzm()，name="butn"
            logger.info("登录页: 点击获取验证码")
            sms_btn_clicked = False

            # 方法1: 在captcha_page调用hyzm()
            try:
                captcha_page.evaluate('hyzm()')
                logger.info("登录页获取验证码: JS hyzm() 调用成功(captcha_page)")
                sms_btn_clicked = True
            except Exception as e1:
                logger.warning(f"captcha_page hyzm() 失败: {e1}")

            # 方法2: 在主页面page调用hyzm()
            if not sms_btn_clicked and captcha_page != page:
                try:
                    page.evaluate('hyzm()')
                    logger.info("登录页获取验证码: JS hyzm() 调用成功(page)")
                    sms_btn_clicked = True
                except Exception as e2:
                    logger.warning(f"page hyzm() 失败: {e2}")

            # 方法3: 在所有frame中尝试hyzm()
            if not sms_btn_clicked:
                for frame in page.frames:
                    try:
                        frame.evaluate('hyzm()')
                        logger.info(f"登录页获取验证码: JS hyzm() 调用成功(frame: {frame.url})")
                        sms_btn_clicked = True
                        break
                    except Exception:
                        continue

            # 方法4: 用Playwright直接点击按钮元素
            if not sms_btn_clicked:
                try:
                    btn = captcha_page.locator('a[name="butn"]')
                    if btn.count() > 0:
                        btn.click()
                        logger.info("登录页获取验证码: Playwright click a[name=butn] 成功")
                        sms_btn_clicked = True
                except Exception as e3:
                    logger.warning(f"Playwright click butn 失败: {e3}")

            # 方法5: getElementsByName
            if not sms_btn_clicked:
                try:
                    captcha_page.evaluate('document.getElementsByName("butn")[0].click()')
                    logger.info("登录页获取验证码: JS click butn 成功")
                    sms_btn_clicked = True
                except Exception as e4:
                    logger.error(f"登录页获取验证码按钮全部失败: {e4}")

            if not sms_btn_clicked:
                logger.error("所有获取验证码方法都失败了！")
            time.sleep(2)

            # 等待短信验证码
            sms_code = self.sms.wait_for_sms_code(
                phone,
                purpose=f"登录-{reg_no}"
            )
            if not sms_code:
                return False

            # 填入短信验证码（登录页用 vcode，不是联络员变更页的 verifyCode）
            logger.info(f"登录页: 准备填入短信验证码: {sms_code}")
            try:
                captcha_page.click('input#vcode')
                time.sleep(0.3)
                captcha_page.fill('input#vcode', '')
                captcha_page.type('input#vcode', sms_code, delay=50)
                actual = captcha_page.input_value('input#vcode')
                logger.info(f"登录页短信验证码type填入: 期望={sms_code} 实际={actual}")
                if actual != sms_code:
                    raise Exception("值不一致")
            except Exception as e:
                logger.warning(f"登录页短信验证码type失败({e})，用JS填入")
                captcha_page.evaluate(f'''() => {{
                    var el = document.getElementById("vcode");
                    if (!el) {{ var inputs = document.querySelectorAll("input"); for (var i=0;i<inputs.length;i++) {{ if(inputs[i].name=="vcode") {{ el=inputs[i]; break; }} }} }}
                    if (el) {{ el.focus(); el.value = "{sms_code}"; el.dispatchEvent(new Event("input", {{bubbles:true}})); el.dispatchEvent(new Event("change", {{bubbles:true}})); }}
                }}''')
                logger.info("登录页短信验证码JS填入完成")

            # 点击登录按钮 — "点击登陆"是<a>标签
            logger.info("登录页: 点击登录按钮")
            login_clicked = False

            # 方式1：用文字匹配<a>标签"点击登陆"
            for selector in ['a:has-text("点击登陆")', 'a:has-text("点击登录")', 'a:has-text("登陆")', 'a:has-text("登录")']:
                try:
                    btn = captcha_page.locator(selector).first
                    if btn.count() > 0:
                        btn.click(timeout=5000)
                        login_clicked = True
                        logger.info(f"登录按钮点击成功: {selector}")
                        break
                except Exception:
                    pass

            # 方式2：JS精确查找"点击登陆"链接
            if not login_clicked:
                js_result = captcha_page.evaluate('''() => {
                    var links = document.querySelectorAll("a");
                    for (var i = 0; i < links.length; i++) {
                        var txt = (links[i].textContent || "").trim();
                        if (txt === "点击登陆" || txt === "点击登录" || txt === "登陆" || txt === "登录") {
                            links[i].click();
                            return "clicked:" + txt;
                        }
                    }
                    return "NOT_FOUND";
                }''')
                logger.info(f"登录按钮JS查找结果: {js_result}")
                if "NOT_FOUND" not in js_result:
                    login_clicked = True

            if not login_clicked:
                logger.error("登录按钮全部方式失败！")

            time.sleep(5)

            # 判断登录结果
            self.take_screenshot(page, f"login_{reg_no}")

            # 检查是否进入年报页面（URL变化或页面内容变化）
            page_text = page.inner_text("body")
            if "年报" in page_text or "企业基本信息" in page_text or "填报" in page_text:
                logger.info(f"登录成功: {reg_no}")
                return True
            else:
                logger.warning(f"登录可能失败: {reg_no}, 页面内容前100字: {page_text[:100]}")
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
            active_page = self.change_liaison(page, enterprise)
            if active_page:
                result["联络员变更"] = "成功"
                # 使用变更后返回的活动页面（可能已跳转到登录页）
                page = active_page
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

"""短信验证码处理模块

支持两种模式：
1. 手动输入（半自动）
2. HTTP接收（全自动，配合安卓SmsForwarder使用）
"""
import re
import time
import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import unquote, unquote_plus
import json
from loguru import logger
import config


class SmsReceiver:
    """HTTP服务接收SmsForwarder转发的短信"""

    def __init__(self, port=5000):
        self.port = port
        self.latest_sms = {}  # {phone: {"code": "xxx", "time": timestamp, "content": "xxx"}}
        self.server = None
        self.thread = None

    def start(self):
        """启动HTTP服务（后台线程）"""
        handler = self._make_handler()
        self.server = HTTPServer(("0.0.0.0", self.port), handler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        logger.info(f"短信接收服务已启动: http://0.0.0.0:{self.port}/sms")
        logger.info(f"请在SmsForwarder中设置转发地址为: http://电脑IP:{self.port}/sms")

    def stop(self):
        if self.server:
            self.server.shutdown()

    def _make_handler(self):
        receiver = self

        class Handler(BaseHTTPRequestHandler):
            def do_POST(self):
                content_len = int(self.headers.get("Content-Length", 0))
                body = self.rfile.read(content_len).decode("utf-8", errors="replace")
                logger.info(f"收到短信转发: {body[:500]}")

                try:
                    data = json.loads(body)
                except Exception:
                    # SmsForwarder可能用form格式，需要URL解码
                    data = {}
                    for pair in body.split("&"):
                        if "=" in pair:
                            k, v = pair.split("=", 1)
                            data[unquote_plus(k)] = unquote_plus(v)

                # 提取短信内容（兼容SmsForwarder的多种格式）
                sms_content = (
                    data.get("content", "")
                    or data.get("sms_content", "")
                    or data.get("msg", "")
                    or data.get("text", "")
                    or data.get("sms_msg", "")
                    or body
                )
                # URL解码（确保中文和特殊字符正确）
                sms_content = unquote_plus(str(sms_content))
                logger.info(f"解码后短信内容: {sms_content[:200]}")

                phone_from = (
                    data.get("from", "")
                    or data.get("phone", "")
                    or data.get("sender", "")
                    or data.get("sim_info", "")
                    or ""
                )
                phone_from = unquote_plus(str(phone_from))

                # 提取验证码（4-8位数字）
                code = receiver._extract_code(sms_content)
                if code:
                    receiver.latest_sms["_latest"] = {
                        "code": code,
                        "time": time.time(),
                        "content": sms_content,
                        "from": phone_from,
                    }
                    logger.info(f"提取到验证码: {code} (来自: {phone_from})")
                else:
                    logger.warning(f"未能从短信中提取到验证码，短信内容: {sms_content[:200]}")

                self.send_response(200)
                self.send_header("Content-Type", "application/json")
                self.end_headers()
                self.wfile.write(json.dumps({"status": "ok", "code": code or ""}).encode())

            def do_GET(self):
                """GET请求返回状态页面"""
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                latest = receiver.latest_sms.get("_latest", {})
                html = f"""<html><body>
                <h2>短信验证码接收服务</h2>
                <p>状态: 运行中</p>
                <p>最新验证码: {latest.get('code', '暂无')}</p>
                <p>短信内容: {latest.get('content', '暂无')}</p>
                <p>SmsForwarder转发地址: POST http://电脑IP:{receiver.port}/sms</p>
                </body></html>"""
                self.wfile.write(html.encode())

            def log_message(self, format, *args):
                pass  # 不打印HTTP访问日志

        return Handler

    def _extract_code(self, text):
        """从短信内容提取验证码"""
        # 优先匹配"验证码"后面的数字
        patterns = [
            r"验证码[：:\s]*(\d{4,8})",
            r"校验码[：:\s]*(\d{4,8})",
            r"动态码[：:\s]*(\d{4,8})",
            r"code[：:\s]*(\d{4,8})",
            r"(\d{4,8})[（(（].*验证码",
            r"(\d{6})",  # 最后兜底：6位数字
            r"(\d{4})",  # 4位数字
        ]
        for pattern in patterns:
            m = re.search(pattern, text, re.IGNORECASE)
            if m:
                return m.group(1)
        return None

    def get_latest_code(self, timeout=120):
        """等待并获取最新验证码"""
        start = time.time()
        # 清除旧的验证码
        old_time = self.latest_sms.get("_latest", {}).get("time", 0)

        while time.time() - start < timeout:
            latest = self.latest_sms.get("_latest", {})
            if latest.get("time", 0) > old_time and latest.get("code"):
                return latest["code"]
            time.sleep(1)

        return None


class SmsHandler:
    """短信验证码处理 - 手动输入模式"""

    def wait_for_sms_code(self, phone, purpose="登录"):
        logger.info(f"等待短信验证码 - 手机号: {phone}, 用途: {purpose}")
        print(f"\n{'='*50}")
        print(f"请输入短信验证码")
        print(f"   手机号: {phone}")
        print(f"   用途: {purpose}")
        print(f"   超时时间: {config.SMS_WAIT_TIMEOUT}秒")
        print(f"{'='*50}")

        code = input("请输入验证码: ").strip()

        if not code:
            logger.warning("未输入验证码")
            return ""

        logger.info(f"收到验证码: {code}")
        return code


class AutoSmsHandler:
    """短信验证码处理 - HTTP自动接收模式（配合SmsForwarder）"""

    def __init__(self, port=5000):
        self.receiver = SmsReceiver(port=port)
        self.receiver.start()

    def wait_for_sms_code(self, phone, purpose="登录"):
        logger.info(f"自动等待短信验证码 - 手机号: {phone}, 用途: {purpose}")
        print(f"\n等待短信验证码（自动模式）...")
        print(f"   手机号: {phone}")
        print(f"   用途: {purpose}")
        print(f"   等待SmsForwarder转发短信...")

        code = self.receiver.get_latest_code(timeout=config.SMS_WAIT_TIMEOUT)

        if code:
            logger.info(f"自动获取到验证码: {code}")
            print(f"   自动获取到验证码: {code}")
            return code
        else:
            logger.warning("自动获取验证码超时，切换手动输入")
            print("   自动获取超时，请手动输入")
            code = input("请输入验证码: ").strip()
            return code

    def stop(self):
        self.receiver.stop()


def create_sms_handler():
    """根据配置创建对应的短信处理器"""
    if config.SMS_MODE == "http":
        return AutoSmsHandler(port=config.SMS_HTTP_PORT)
    else:
        return SmsHandler()

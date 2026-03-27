"""短信验证码处理模块"""
import time
from loguru import logger
import config


class SmsHandler:
    """短信验证码处理
    
    当前实现：弹窗等待人工输入（半自动模式）
    后续可扩展：对接短信转发接口实现全自动
    """
    
    def wait_for_sms_code(self, phone: str, purpose: str = "登录") -> str:
        """等待并获取短信验证码
        
        Args:
            phone: 接收验证码的手机号
            purpose: 用途说明（登录/联络员变更）
        Returns:
            用户输入的验证码
        """
        logger.info(f"等待短信验证码 - 手机号: {phone}, 用途: {purpose}")
        print(f"\n{'='*50}")
        print(f"📱 请输入短信验证码")
        print(f"   手机号: {phone}")
        print(f"   用途: {purpose}")
        print(f"   超时时间: {config.SMS_WAIT_TIMEOUT}秒")
        print(f"{'='*50}")
        
        # 简单的控制台输入方式
        code = input("请输入验证码: ").strip()
        
        if not code:
            logger.warning("未输入验证码")
            return ""
        
        logger.info(f"收到验证码: {code}")
        return code


class AutoSmsHandler(SmsHandler):
    """自动短信验证码处理（对接短信转发接口）
    
    TODO: 后续实现
    - 监听HTTP接口接收转发的短信
    - 从短信内容中提取验证码
    - 自动返回验证码
    """
    
    def __init__(self, api_url: str = ""):
        self.api_url = api_url
    
    def wait_for_sms_code(self, phone: str, purpose: str = "登录") -> str:
        """自动获取短信验证码"""
        logger.info(f"自动模式等待短信验证码 - 手机号: {phone}")
        # TODO: 实现自动获取逻辑
        # 1. 调用短信转发接口
        # 2. 轮询等待新短信
        # 3. 正则提取验证码
        raise NotImplementedError("自动短信获取功能待实现")

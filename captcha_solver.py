"""图形验证码识别模块"""
import ddddocr
from loguru import logger
from PIL import Image
import io


class CaptchaSolver:
    def __init__(self):
        self.ocr = ddddocr.DdddOcr(show_ad=False)
        logger.info("验证码识别引擎初始化完成")

    def solve_from_bytes(self, image_bytes: bytes) -> str:
        """从图片字节数据识别验证码"""
        result = self.ocr.classification(image_bytes)
        logger.info(f"验证码识别结果: {result}")
        return result

    def solve_from_element(self, page, selector: str) -> str:
        """从页面元素截图识别验证码
        
        Args:
            page: Playwright page对象
            selector: 验证码图片的CSS选择器
        Returns:
            识别出的验证码文本
        """
        element = page.locator(selector)
        image_bytes = element.screenshot()
        return self.solve_from_bytes(image_bytes)

    def solve_from_file(self, filepath: str) -> str:
        """从文件识别验证码"""
        with open(filepath, "rb") as f:
            return self.solve_from_bytes(f.read())

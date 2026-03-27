"""配置文件"""

# 网站地址
LOGIN_URL = "https://portal.scjgj.gz.gov.cn/aiceps/liaisonsLogin.html"
CHANGE_LIAISON_URL = "https://portal.scjgj.gz.gov.cn/aiceps/liaisonsChange.html"

# Excel文件路径
ENTERPRISE_EXCEL = "data/企业信息.xlsx"
ANNUAL_REPORT_EXCEL = "data/年报数据.xlsx"

# 新联络员信息（统一）
NEW_LIAISON = {
    "name": "",        # 新联络员姓名
    "id_type": "中华人民共和国居民身份证",
    "id_number": "",   # 新联络员身份证号
    "phone": "",       # 新联络员手机号（统一号码）
}

# 浏览器配置
HEADLESS = False
SLOW_MO = 500
TIMEOUT = 30000

# 验证码重试次数
CAPTCHA_MAX_RETRY = 5

# 短信验证码等待超时（秒）
SMS_WAIT_TIMEOUT = 120

# 截图保存目录
SCREENSHOT_DIR = "screenshots"

# 日志文件
LOG_FILE = "logs/report_{time}.log"

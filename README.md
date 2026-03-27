# 工商年报自动申报工具

广州市市场监督管理局 - 国家企业信用信息公示系统(广东) 年报自动填报工具

## 功能

1. **联络员变更** - 自动填写联络员变更表单
2. **联络员登录** - 自动登录（图形验证码自动识别 + 短信验证码人工输入）
3. **年报填写** - 从Excel读取数据自动填入年报表单
4. **结果记录** - 每家企业的处理结果记录到日志

## 环境要求

- Python 3.9+
- Windows 10/11（推荐）

## 安装

```bash
# 安装依赖
pip install -r requirements.txt

# 安装Playwright浏览器
playwright install chromium
```

## 配置

编辑 `config.py`：

1. 填写新联络员信息（统一手机号、姓名、身份证号）
2. 放置Excel数据文件到 `data/` 目录
3. 根据需要调整浏览器参数

## 使用

```bash
# 单企业测试
python annual_report_bot.py

# 批量处理（修改main函数的start_index和end_index）
```

## 文件结构

```
gsxt-annual-report/
├── annual_report_bot.py   # 主程序
├── captcha_solver.py      # 图形验证码识别
├── data_reader.py         # Excel数据读取
├── sms_handler.py         # 短信验证码处理
├── config.py              # 配置文件
├── requirements.txt       # 依赖包
├── data/                  # Excel数据文件
├── screenshots/           # 截图保存
└── logs/                  # 日志和结果
```

## 注意事项

- 页面选择器（CSS selector）需要根据实际网页HTML结构调整
- 首次运行建议使用有头模式（HEADLESS=False）观察执行过程
- 图形验证码识别率非100%，失败会自动重试
- 短信验证码需要人工输入（后续可对接短信转发接口）

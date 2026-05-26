'''
Before running, you need to run the following commands
运行前需执行以下命令：
pip install playwright openpyxl matplotlib
playwright install chromium
'''

# -------------------------- 1. Import Required Dependencies | 导入所需依赖库 --------------------------
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
import os
import re
from urllib.parse import urlparse, urlunparse
from datetime import datetime
import argparse
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
import matplotlib.pyplot as plt
import traceback

# -------------------------- 2. Global Configuration | 全局配置项 --------------------------
CONFIG = {
    # File Path Configuration | 文件路径配置（默认值，可通过命令行/交互覆盖）
    "input_excel_path": "",  # Input Excel with URL list | 输入URL的Excel文件路径
    "output_excel_path": "",  # Output Excel with results | 输出结果的Excel文件路径
    "input_sheet_name": "Feuil1",  # Input sheet name | 输入Sheet名称
    "summary_sheet_name": "output_url",  # Summary result sheet | 输出汇总Sheet名称
    "detail_sheet_name": "jump_chain",  # Jump detail sheet | 输出详情Sheet名称

    # Browser Runtime Configuration | 浏览器运行配置
    "headless_mode": True,  # 云服务器部署默认无头模式 | Headless mode (default True for cloud deployment)
    "timeout": 30 * 1000,  # Single page access timeout (ms) | 单页面访问超时时间（毫秒）
    "max_retry_times": 2,  # Max retry times for single URL | 单条URL失败重试次数
    "max_jump_count": 20,  # Max jump count, prevent infinite redirect loop | 最大跳转次数
    "wait_for_network_idle": 2000,  # Wait for network idle after page load (ms) | 页面加载后等待网络空闲时间
    "proxy_config": None,  # Proxy config: None=no proxy, format: {"server": "http://ip:port"} | 代理配置

    # Click Trigger Configuration | 点击触发配置
    "click_config": {
        "enable_click_function": True,  # Enable click simulation | 是否开启点击模拟功能
        "max_click_retry": 2,  # Max retry times for click | 点击最大重试次数
        "wait_after_click": 5 * 1000,  # Wait time after click (ms) | 点击后等待时间
        "capture_new_window": True,  # Capture jump in new window after click | 捕获新窗口跳转
        "target_keywords": {
            "auth_login": ["login", "signin", "sign-in", "secure access", "account",
                           "connexion", "se connecter", "accès sécurisé", "compte"],
            "espace_access": ["espace", "my espace", "access espace", "enter", "access",
                              "mon espace", "accéder à l'espace", "entrer", "accéder"]
        },
        "selector_priority": [
            "a[href*={kw}], button[id*={kw}], button[name*={kw}], a[class*={kw}]",
            "button:has-text('{kw}'), a:has-text('{kw}'), div[role='button']:has-text('{kw}')",
            "[role='button']:has-text('{kw}'), [role='link']:has-text('{kw}')"
        ]
    },

    # Cookie Consent Configuration | Cookie同意弹窗配置
    "cookie_config": {
        "enable_cookie_handle": True,
        "keywords": ["accept cookie", "accept all cookies", "accepter les cookies", "tout accepter", "ok", "agree"],
        "selectors": [
            "button:has-text('{kw}'), a:has-text('{kw}')",
            "[role='button']:has-text('{kw}'), div[class*='cookie'] button"
        ]
    },

    # URL Classification Rules (EN/FR) | URL分类规则（英法双语）
    "classify_rules": {
        # Main Category | 主分类
        "Authentication": {
            "keywords": ["login", "auth", "iam", "sso", "token", "signin", "authentification"],
            "page_features": ["username", "password", "login", "sign in",
                              "nom d'utilisateur", "mot de passe", "connexion"],
            # Sub Category | 细分分类
            "sub_class": {
                "SSO/Auth Service": ["sso", "iam", "oauth", "openid", "authentification unique"],
                "Login Page": ["login", "signin", "connexion", "se connecter"],
                "Token/Credential": ["token", "credential", "jwt", "bearer", "jeton"]
            }
        },
        "Espace": {
            "keywords": ["espace", "espresso", "station", "business", "portal"],
            "page_features": ["login required", "please log in",
                              "connexion requise", "veuillez vous connecter"],
            "sub_class": {
                "Station Espace": ["station", "espresso station", "station espresso"],
                "Business Espace": ["business", "entreprise", "professionnel"],
                "User Espace": ["mon espace", "my espace", "espace utilisateur"]
            }
        },
        "Backend/API": {
            "keywords": ["api", "gateway", "backend", "server", "interface", "rest"],
            "request_type": ["xhr", "fetch"],
            "sub_class": {
                "REST API": ["rest", "api/v1", "api/v2", "json", "xml"],
                "Gateway API": ["gateway", "proxy", "api gateway", "passerelle api"],
                "Internal API": ["internal", "backend", "serveur interne"]
            }
        },
        "Public front": {
            "default": True,
            "sub_class": {
                "Landing Page": ["home", "accueil", "landing", "welcome"],
                "Product Page": ["product", "produit", "service", "offre"],
                "Help/Support": ["help", "aide", "support", "assistance", "faq"]
            }
        },
        # Error Category | 错误分类
        "Error": {
            "sub_class": {
                "Connection Error": ["ERR_CONNECTION", "connection refused", "connexion refusée"],
                "Timeout Error": ["ERR_TIMED_OUT", "timeout", "délai expiré"],
                "Certificate Error": ["ERR_CERT", "certificate", "certificat invalide"],
                "4XX/5XX HTTP": ["404", "403", "500", "502", "http error", "erreur http"],
                "Auth Required": ["401", "unauthorized", "non autorisé", "authentification requise"]
            }
        }
    },

    # Log Configuration | 日志配置
    "log_config": {
        "log_file": "url_scan_{}.log".format(datetime.now().strftime("%Y%m%d_%H%M%S")),
        "log_level": "INFO"
    }
}

# -------------------------- 3. Initialization | 初始化配置 --------------------------
# Setup Logging | 配置日志
def setup_logging():
    """
    Initialize logging configuration (support cloud server deployment)
    初始化日志配置（适配云服务器部署）
    """
    logging.basicConfig(
        level=getattr(logging, CONFIG["log_config"]["log_level"]),
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(CONFIG["log_config"]["log_file"], encoding="utf-8"),
            logging.StreamHandler()  # Output to console | 同时输出到控制台
        ]
    )
    return logging.getLogger(__name__)

logger = setup_logging()

# Parse Command Line Args | 解析命令行参数
def parse_command_line_args():
    """
    Parse command line arguments for input/output paths
    解析命令行参数（输入/输出路径）
    """
    parser = argparse.ArgumentParser(
        description="URL Scan Tool | URL扫描工具",
        epilog="Example: python script.py -i URL_CORE.xlsx -o URL_CORE_result.xlsx"
    )
    parser.add_argument("-i", "--input", help="Input Excel file path | 输入Excel文件路径", type=str)
    parser.add_argument("-o", "--output", help="Output Excel file path | 输出Excel文件路径", type=str)
    args = parser.parse_args()

    # Interactive input if no args | 无参数时交互式输入
    if not args.input:
        args.input = input("Please enter input Excel file path | 请输入输入Excel文件路径: ").strip()
        while not args.input:
            args.input = input("Input path cannot be empty | 输入路径不能为空，请重新输入: ").strip()

    if not args.output:
        default_output = f"URL_SCAN_RESULT_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        args.output = input(f"Please enter output Excel file path (default: {default_output}) | 请输入输出Excel文件路径（默认：{default_output}）: ").strip() or default_output

    CONFIG["input_excel_path"] = args.input
    CONFIG["output_excel_path"] = args.output
    logger.info(f"Input path: {CONFIG['input_excel_path']} | 输入路径: {CONFIG['input_excel_path']}")
    logger.info(f"Output path: {CONFIG['output_excel_path']} | 输出路径: {CONFIG['output_excel_path']}")

# -------------------------- 4. Utility Functions | 工具函数模块 --------------------------
def standardize_url(url: str) -> str:
    """
    Standardize URL format, complete protocol, clean invalid characters
    功能：URL标准化处理，补全协议、清洗无效字符，确保URL可正常访问
    :param url: Raw input URL | 原始输入的URL字符串
    :return: Standardized valid URL, empty string if invalid | 标准化后的URL，无效URL返回空字符串
    """
    url = url.strip()
    if not url:
        return ""

    parsed = urlparse(url)
    if not parsed.scheme:
        url = "https://" + url
        parsed = urlparse(url)

    if not parsed.netloc or parsed.scheme not in ["http", "https"]:
        return ""

    return urlunparse(parsed).rstrip("/")

def is_valid_url(url: str) -> bool:
    """
    Validate if URL is legal and valid
    功能：校验URL是否合法有效
    :param url: URL to validate | 待校验的URL
    :return: True=valid, False=invalid | 有效返回True，无效返回False
    """
    if not url:
        return False
    parsed = urlparse(url)
    return all([parsed.scheme in ["http", "https"], parsed.netloc])

def handle_cookie_consent(page):
    """
    Auto handle cookie consent popup (EN/FR bilingual)
    自动处理Cookie同意弹窗（英法双语）
    :param page: Playwright page object | Playwright页面对象
    """
    if not CONFIG["cookie_config"]["enable_cookie_handle"]:
        return

    cookie_config = CONFIG["cookie_config"]
    for kw in cookie_config["keywords"]:
        for selector_template in cookie_config["selectors"]:
            selector = selector_template.format(kw=kw.lower())
            try:
                element = page.locator(selector).first
                if element.is_visible(timeout=1000) and element.is_enabled(timeout=1000):
                    element.click(timeout=2000)
                    logger.info(f"Clicked cookie consent button | 点击Cookie同意按钮: {selector}")
                    page.wait_for_timeout(1000)
                    return
            except Exception as e:
                continue
    logger.info("No cookie consent popup found | 未找到Cookie同意弹窗")

def extract_error_code(error_msg: str) -> tuple[str, str]:
    """
    Extract error code and error type from error message
    从错误信息中提取错误码和错误类型
    :param error_msg: Original error message | 原始错误信息
    :return: (error_code, error_type) | 错误码、错误类型
    """
    if not error_msg:
        return "", "Unknown Error | 未知错误"

    # Match Playwright/Chrome error codes | 匹配Playwright/Chrome错误码
    error_code_pattern = r"ERR_\w+"
    error_code = re.search(error_code_pattern, error_msg)
    error_code = error_code.group() if error_code else ""

    # Match HTTP status codes | 匹配HTTP状态码
    if not error_code:
        http_code_pattern = r"\b(4\d{2}|5\d{2}|3\d{2})\b"
        http_code = re.search(http_code_pattern, error_msg)
        error_code = http_code.group() if http_code else ""

    # Classify error type | 分类错误类型
    error_type = "Unknown Error | 未知错误"
    error_rules = CONFIG["classify_rules"]["Error"]["sub_class"]
    for type_name, keywords in error_rules.items():
        if any(kw in error_msg for kw in keywords) or (error_code and any(kw in error_code for kw in keywords)):
            error_type = type_name
            break

    return error_code, error_type

def classify_url(jump_info: dict) -> tuple[str, str]:
    """
    Classify URL (main + sub category) according to configured rules
    功能：根据配置规则对URL进行主分类+细分分类
    :param jump_info: Full info of single jump | 单条跳转的完整信息字典
    :return: (main_class, sub_class) | 主分类、细分分类
    """
    # If error exists | 存在错误时优先分类为Error
    if jump_info.get("error_msg"):
        _, error_type = extract_error_code(jump_info["error_msg"])
        return "Error", error_type

    url_lower = jump_info["jump_url"].lower()
    page_html = jump_info.get("page_html", "").lower()
    request_type = jump_info.get("request_type", "")

    # 1. Authentication | 认证模块
    auth_rule = CONFIG["classify_rules"]["Authentication"]
    if any(keyword in url_lower for keyword in auth_rule["keywords"]) or \
            any(feature in page_html for feature in auth_rule["page_features"]):
        # Get sub class | 匹配细分分类
        sub_class = "Other Authentication | 其他认证"
        for sub_name, sub_keywords in auth_rule["sub_class"].items():
            if any(kw in url_lower for kw in sub_keywords):
                sub_class = sub_name
                break
        return "Authentication", sub_class

    # 2. Espace | 业务空间
    espace_rule = CONFIG["classify_rules"]["Espace"]
    if any(keyword in url_lower for keyword in espace_rule["keywords"]) or \
            any(feature in page_html for feature in espace_rule["page_features"]):
        sub_class = "Other Espace | 其他业务空间"
        for sub_name, sub_keywords in espace_rule["sub_class"].items():
            if any(kw in url_lower for kw in sub_keywords):
                sub_class = sub_name
                break
        return "Espace", sub_class

    # 3. Backend/API | 后端接口
    api_rule = CONFIG["classify_rules"]["Backend/API"]
    if any(keyword in url_lower for keyword in api_rule["keywords"]) or \
            request_type in api_rule["request_type"]:
        sub_class = "Other API | 其他接口"
        for sub_name, sub_keywords in api_rule["sub_class"].items():
            if any(kw in url_lower for kw in sub_keywords):
                sub_class = sub_name
                break
        return "Backend/API", sub_class

    # 4. Public front | 公共前端
    sub_class = "Other Public Page | 其他公共页面"
    for sub_name, sub_keywords in CONFIG["classify_rules"]["Public front"]["sub_class"].items():
        if any(kw in url_lower for kw in sub_keywords):
            sub_class = sub_name
            break
    return "Public front", sub_class

def find_target_clickable_element(page) -> tuple[object, str]:
    """
    Find target clickable element according to EN/FR bilingual rules
    功能：根据英法双语规则，定位目标可点击元素
    :param page: Playwright page object | Playwright页面对象
    :return: (matched_element, element_description) | 匹配到的元素、元素描述信息
    """
    click_config = CONFIG["click_config"]
    all_keywords = click_config["target_keywords"]["auth_login"] + click_config["target_keywords"]["espace_access"]
    all_keywords = list(set([kw.lower() for kw in all_keywords]))

    for selector_template in click_config["selector_priority"]:
        for keyword in all_keywords:
            selector = selector_template.format(kw=keyword)
            try:
                element = page.locator(selector).first
                if element.is_visible(timeout=1000) and element.is_enabled(timeout=1000):
                    element_desc = f"Selector: {selector}, Keyword: {keyword}"
                    return element, element_desc
            except:
                continue
    return None, "No matched element | 未匹配到元素"

# -------------------------- 5. Core Scan Function | 核心扫描函数 --------------------------
def scan_single_url(raw_url: str) -> dict:
    """
    Scan single URL, capture jump chain, error info, classification
    扫描单个URL，捕获跳转链路、错误信息、分类结果
    :param raw_url: Raw URL string | 原始URL字符串
    :return: Scan result dict | 扫描结果字典
    """
    result = {
        "raw_url": raw_url,
        "standard_url": "",
        "jump_chain": [],
        "main_class": "",
        "sub_class": "",
        "error_msg": "",
        "error_code": "",
        "auth_required": False,  # Whether auth/certificate is required | 是否要求账号密码/证书
        "scan_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    # URL Standardization | URL标准化
    standard_url = standardize_url(raw_url)
    result["standard_url"] = standard_url
    if not is_valid_url(standard_url):
        result["error_msg"] = "Invalid URL format | 无效的URL格式"
        result["error_code"] = "INVALID_URL"
        result["main_class"], result["sub_class"] = classify_url(result)
        logger.error(f"Invalid URL | 无效URL: {raw_url}")
        return result

    # Retry mechanism | 重试机制
    retry_count = 0
    while retry_count <= CONFIG["max_retry_times"]:
        try:
            with sync_playwright() as p:
                # Launch browser | 启动浏览器
                browser_kwargs = {"headless": CONFIG["headless_mode"], "timeout": CONFIG["timeout"]}
                if CONFIG["proxy_config"]:
                    browser_kwargs["proxy"] = CONFIG["proxy_config"]

                browser = p.chromium.launch(**browser_kwargs)
                context = browser.new_context(ignore_https_errors=True)  # Ignore certificate errors | 忽略证书错误
                page = context.new_page()

                # Set timeout | 设置超时
                page.set_default_timeout(CONFIG["timeout"])

                try:
                    # Navigate to URL | 访问URL
                    response = page.goto(standard_url, wait_until="networkidle", timeout=CONFIG["timeout"])

                    # Handle cookie consent | 处理Cookie弹窗
                    handle_cookie_consent(page)

                    # Check auth required | 检查是否要求账号密码/证书
                    auth_indicators = ["authentication required", "certificate required", "nom d'utilisateur", "mot de passe",
                                       "authentification requise", "certificat requis"]
                    page_html = page.content().lower()
                    if any(indicator in page_html for indicator in auth_indicators):
                        result["auth_required"] = True
                        logger.info(f"Auth/certificate required | 要求账号密码/证书: {standard_url}")

                    # Capture jump chain | 捕获跳转链路
                    jump_count = 0
                    current_url = standard_url
                    while jump_count < CONFIG["max_jump_count"]:
                        result["jump_chain"].append(current_url)
                        # Check redirect | 检查重定向
                        if page.url == current_url and jump_count > 0:
                            break
                        current_url = page.url
                        jump_count += 1

                    # Click simulation (optional) | 点击模拟（可选）
                    if CONFIG["click_config"]["enable_click_function"]:
                        element, desc = find_target_clickable_element(page)
                        if element:
                            for _ in range(CONFIG["click_config"]["max_click_retry"]):
                                try:
                                    element.click()
                                    page.wait_for_timeout(CONFIG["click_config"]["wait_after_click"])
                                    # Capture new window | 捕获新窗口
                                    if CONFIG["click_config"]["capture_new_window"] and context.pages:
                                        for new_page in context.pages:
                                            if new_page.url not in result["jump_chain"]:
                                                result["jump_chain"].append(new_page.url)
                                    logger.info(f"Clicked element | 点击元素: {desc} for URL: {standard_url}")
                                    break
                                except Exception as e:
                                    logger.warning(f"Click failed | 点击失败: {e}, retry {_+1}")
                                    continue

                    # Get final page info | 获取最终页面信息
                    final_url = page.url
                    result["jump_chain"].append(final_url)
                    result["jump_chain"] = list(dict.fromkeys(result["jump_chain"]))  # Remove duplicates | 去重

                    # Classify URL | URL分类
                    jump_info = {
                        "jump_url": final_url,
                        "page_html": page_html,
                        "request_type": response.request.resource_type if response else ""
                    }
                    main_class, sub_class = classify_url(jump_info)
                    result["main_class"] = main_class
                    result["sub_class"] = sub_class

                    logger.info(f"Scan success | 扫描成功: {standard_url} -> {final_url}")
                    break

                except PlaywrightTimeoutError:
                    result["error_msg"] = f"Timeout ({CONFIG['timeout']/1000}s) | 超时（{CONFIG['timeout']/1000}秒）"
                    result["error_code"] = "ERR_TIMED_OUT"
                    logger.warning(f"Timeout | 超时: {standard_url}, retry {retry_count+1}")
                except Exception as e:
                    error_msg = str(e)
                    result["error_msg"] = error_msg
                    result["error_code"], _ = extract_error_code(error_msg)
                    logger.warning(f"Scan failed | 扫描失败: {standard_url}, error: {error_msg}, retry {retry_count+1}")
                finally:
                    # Ensure browser close | 确保浏览器关闭
                    context.close()
                    browser.close()

            retry_count += 1
        except Exception as e:
            result["error_msg"] = f"Browser launch failed | 浏览器启动失败: {str(e)}"
            logger.error(f"Browser error | 浏览器错误: {standard_url}, error: {str(e)}")
            break

    # Final classification for error | 错误最终分类
    if result["error_msg"] and not result["main_class"]:
        result["main_class"], result["sub_class"] = classify_url(result)

    return result

# -------------------------- 6. Excel Export & Statistics | Excel导出与统计 --------------------------
def export_to_excel(scan_results: list):
    """
    Export scan results to Excel (bilingual headers)
    导出扫描结果到Excel（双语表头）
    :param scan_results: List of scan results | 扫描结果列表
    """
    # Create workbook | 创建工作簿
    wb = Workbook()

    # Summary sheet | 汇总表
    summary_sheet = wb.active
    summary_sheet.title = CONFIG["summary_sheet_name"]
    summary_headers = [
        "Raw URL | 原始URL", "Standard URL | 标准化URL", "Main Category | 主分类",
        "Sub Category | 细分分类", "Error Message | 错误信息", "Error Code | 错误码",
        "Auth Required | 是否要求认证", "Scan Time | 扫描时间"
    ]
    summary_sheet.append(summary_headers)

    # Detail sheet (jump chain) | 详情表（跳转链路）
    detail_sheet = wb.create_sheet(title=CONFIG["detail_sheet_name"])
    detail_headers = ["Raw URL | 原始URL", "Jump Step | 跳转步骤", "Jump URL | 跳转URL"]
    detail_sheet.append(detail_headers)

    # Fill data | 填充数据
    for idx, result in enumerate(scan_results, start=2):
        # Summary sheet | 汇总表
        summary_sheet.append([
            result["raw_url"], result["standard_url"], result["main_class"],
            result["sub_class"], result["error_msg"], result["error_code"],
            "Yes | 是" if result["auth_required"] else "No | 否", result["scan_time"]
        ])

        # Detail sheet | 详情表
        for step, jump_url in enumerate(result["jump_chain"], start=1):
            detail_sheet.append([result["raw_url"], step, jump_url])

    # Style settings | 样式设置
    for sheet in [summary_sheet, detail_sheet]:
        for col in sheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            sheet.column_dimensions[column].width = adjusted_width

        # Header style | 表头样式
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")

    # Save Excel | 保存Excel
    try:
        wb.save(CONFIG["output_excel_path"])
        logger.info(f"Excel exported successfully | Excel导出成功: {CONFIG['output_excel_path']}")
    except Exception as e:
        logger.error(f"Excel export failed | Excel导出失败: {e}")

def generate_statistics(scan_results: list):
    """
    Generate scan statistics and charts (support cloud server deployment)
    生成扫描统计数据和图表（适配云服务器部署）
    :param scan_results: List of scan results | 扫描结果列表
    """
    # Statistics calculation | 统计计算
    total_count = len(scan_results)
    success_count = len([r for r in scan_results if not r["error_msg"]])
    error_count = total_count - success_count

    # Main category statistics | 主分类统计
    main_class_stats = {}
    sub_class_stats = {}
    error_code_stats = {}
    auth_required_count = len([r for r in scan_results if r["auth_required"]])

    for result in scan_results:
        # Main class | 主分类
        main_class = result["main_class"]
        main_class_stats[main_class] = main_class_stats.get(main_class, 0) + 1

        # Sub class | 细分分类
        sub_class = f"{main_class} - {result['sub_class']}"
        sub_class_stats[sub_class] = sub_class_stats.get(sub_class, 0) + 1

        # Error code | 错误码
        if result["error_code"]:
            error_code_stats[result["error_code"]] = error_code_stats.get(result["error_code"], 0) + 1

    # Print statistics | 打印统计信息
    logger.info("\n" + "="*50)
    logger.info("Scan Statistics | 扫描统计报告")
    logger.info("="*50)
    logger.info(f"Total URLs Scanned | 总扫描URL数: {total_count}")
    logger.info(f"Successfully Scanned | 扫描成功数: {success_count} ({success_count/total_count*100:.2f}%)")
    logger.info(f"Failed Scanned | 扫描失败数: {error_count} ({error_count/total_count*100:.2f}%)")
    logger.info(f"Auth Required URLs | 要求认证的URL数: {auth_required_count} ({auth_required_count/total_count*100:.2f}%)")
    logger.info("\nMain Category Distribution | 主分类分布:")
    for cls, count in main_class_stats.items():
        logger.info(f"  {cls}: {count} ({count/total_count*100:.2f}%)")

    logger.info("\nError Code Distribution | 错误码分布:")
    for code, count in error_code_stats.items():
        logger.info(f"  {code}: {count}")

    # Generate charts | 生成图表（适配云服务器，保存为文件而非弹窗）
    plt.rcParams["font.sans-serif"] = ["DejaVu Sans"]  # Support EN/FR | 支持英法文
    plt.rcParams["axes.unicode_minus"] = False

    # 1. Main category pie chart | 主分类饼图
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 8))

    labels = list(main_class_stats.keys())
    sizes = list(main_class_stats.values())
    colors = plt.cm.Set3(range(len(labels)))
    ax1.pie(sizes, labels=labels, autopct='%1.1f%%', colors=colors, startangle=90)
    ax1.set_title("Main Category Distribution | 主分类分布", fontsize=14)

    # 2. Error code bar chart | 错误码柱状图
    if error_code_stats:
        error_codes = list(error_code_stats.keys())
        error_counts = list(error_code_stats.values())
        ax2.bar(error_codes, error_counts, color=plt.cm.viridis(range(len(error_codes))))
        ax2.set_title("Error Code Distribution | 错误码分布", fontsize=14)
        ax2.set_xlabel("Error Code | 错误码")
        ax2.set_ylabel("Count | 数量")
        ax2.tick_params(axis='x', rotation=45)

    # Save chart | 保存图表
    chart_path = f"url_scan_chart_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    plt.tight_layout()
    plt.savefig(chart_path, dpi=300, bbox_inches='tight')
    logger.info(f"Statistics chart saved | 统计图表已保存: {chart_path}")

# -------------------------- 6. Main Function | 主函数 --------------------------
def main():
    """
    Main function: parse args -> load URLs -> multi-thread scan -> export -> statistics
    主函数：解析参数 -> 加载URL -> 多线程扫描 -> 导出结果 -> 生成统计
    """
    try:
        # Parse command line args | 解析命令行参数
        parse_command_line_args()

        # Load input Excel | 加载输入Excel
        if not os.path.exists(CONFIG["input_excel_path"]):
            logger.error(f"Input file not found | 输入文件不存在: {CONFIG['input_excel_path']}")
            return

        wb = load_workbook(CONFIG["input_excel_path"])
        if CONFIG["input_sheet_name"] not in wb.sheetnames:
            logger.error(f"Input sheet not found | 输入Sheet不存在: {CONFIG['input_sheet_name']}")
            return

        sheet = wb[CONFIG["input_sheet_name"]]
        # Assume first column is URL list | 假设第一列为URL列表
        url_list = [cell.value for cell in sheet['A'] if cell.value is not None and cell.value != "URL"]
        wb.close()

        if not url_list:
            logger.error("No URLs found in input file | 输入文件中未找到URL")
            return

        logger.info(f"Loaded {len(url_list)} URLs | 加载了{len(url_list)}个URL")

        # Multi-thread scan | 多线程扫描
        scan_results = []
        # Set thread count (adjust according to server resources) | 设置线程数（根据服务器资源调整）
        thread_count = min(10, len(url_list))  # Max 10 threads | 最大10线程
        logger.info(f"Start multi-thread scan with {thread_count} threads | 启动{thread_count}线程多线程扫描")

        with ThreadPoolExecutor(max_workers=thread_count) as executor:
            # Submit tasks | 提交任务
            future_to_url = {executor.submit(scan_single_url, url): url for url in url_list}
            # Get results | 获取结果
            for future in as_completed(future_to_url):
                url = future_to_url[future]
                try:
                    result = future.result()
                    scan_results.append(result)
                except Exception as e:
                    logger.error(f"Scan task failed | 扫描任务失败: {url}, error: {traceback.format_exc()}")
                    scan_results.append({
                        "raw_url": url,
                        "standard_url": "",
                        "jump_chain": [],
                        "main_class": "Error",
                        "sub_class": "Task Execution Error | 任务执行错误",
                        "error_msg": str(e),
                        "error_code": "TASK_ERROR",
                        "auth_required": False,
                        "scan_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    })

        # Export results to Excel | 导出结果到Excel
        export_to_excel(scan_results)

        # Generate statistics | 生成统计数据
        generate_statistics(scan_results)

        logger.info("Scan completed successfully | 扫描任务全部完成！")

    except Exception as e:
        logger.error(f"Main function error | 主函数执行错误: {traceback.format_exc()}")

if __name__ == "__main__":
    main()
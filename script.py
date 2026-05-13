'''
Before running, you need to run the following commands

pip install playwright openpyxl
playwright install chromium
'''

# -------------------------- 1. Import Required Dependencies | 导入所需依赖库 --------------------------
# Core library for browser automation, simulate user operations and capture jump chain
# 浏览器自动化核心库，用于模拟用户操作、捕获跳转链路
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
# Excel read/write library, for source URL loading and result export
# Excel读写库，用于读取源URL、写入处理结果
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
# System & utility libraries, for path handling, exception capture, timestamp recording
# 系统与工具库，用于路径处理、异常捕获、时间戳记录
import os
import re
from urllib.parse import urlparse, urlunparse
from datetime import datetime

# -------------------------- 2. Global Configuration | 全局配置项 --------------------------
# All configurable parameters are here, no need to modify core code | 所有可配置参数均在此处，无需修改核心代码
CONFIG = {
    # File Path Configuration | 文件路径配置
    "input_excel_path": "URL_CORE.xlsx",  # Input Excel with URL list | 输入URL的Excel文件路径
    "output_excel_path": "URL_CORE_result.xlsx",  # Output Excel with results | 输出结果的Excel文件路径
    "input_sheet_name": "Feuil1",  # Input sheet name | 输入Sheet名称
    "summary_sheet_name": "output_url",  # Summary result sheet | 输出汇总Sheet名称
    "detail_sheet_name": "jump_chain",  # Jump detail sheet | 输出详情Sheet名称

    # Browser Runtime Configuration | 浏览器运行配置
    "headless_mode": False,  # Headless mode: True=background run, False=show browser (for debug) | 无头模式
    "timeout": 30 * 1000,  # Single page access timeout (ms) | 单页面访问超时时间（毫秒）
    "max_retry_times": 2,  # Max retry times for single URL | 单条URL失败重试次数
    "max_jump_count": 20,  # Max jump count, prevent infinite redirect loop | 最大跳转次数，防止无限重定向
    "wait_for_network_idle": 2000,  # Wait for network idle after page load (ms) | 页面加载后等待网络空闲时间（毫秒）
    "proxy_config": None,  # Proxy config: None=no proxy, format: {"server": "http://ip:port"} | 代理配置

    # Click Trigger Configuration (NEW EXTENSION) | 点击触发配置（新增扩展）
    "click_config": {
        "enable_click_function": True,  # Enable click simulation | 是否开启点击模拟功能
        "max_click_retry": 2,  # Max retry times for click | 点击最大重试次数
        "wait_after_click": 5 * 1000,  # Wait time after click for jump execution (ms) | 点击后等待跳转执行时间（毫秒）
        "capture_new_window": True,  # Capture jump in new window after click | 是否捕获点击后新窗口的跳转
        # EN/FR Bilingual Element Positioning Keywords | 英法双语元素定位关键词
        "target_keywords": {
            "auth_login": ["login", "signin", "sign-in", "secure access", "account",
                           "connexion", "se connecter", "accès sécurisé", "compte"],
            "espace_access": ["espace", "my espace", "access espace", "enter", "access",
                              "mon espace", "accéder à l'espace", "entrer", "accéder"]
        },
        # Element Selector Priority (High to Low) | 元素选择器优先级（从高到低）
        "selector_priority": [
            # High Priority: Attribute match (most stable) | 高优先级：属性匹配（最稳定）
            "a[href*={kw}], button[id*={kw}], button[name*={kw}], a[class*={kw}]",
            # Medium Priority: Clickable element text match | 中优先级：可点击元素文本匹配
            "button:has-text('{kw}'), a:has-text('{kw}'), div[role='button']:has-text('{kw}')",
            # Low Priority: Role backup match | 低优先级：角色兜底匹配
            "[role='button']:has-text('{kw}'), [role='link']:has-text('{kw}')"
        ]
    },

    # URL Classification Rules (EN/FR Only, No Chinese) | URL分类规则（仅英法双语，无中文）
    "classify_rules": {
        # Authentication/Security Module | 认证/安全模块（最高优先级）
        "Authentication": {
            "keywords": ["login", "auth", "iam", "sso", "token", "signin", "authentification"],
            "page_features": ["username", "password", "login", "sign in",
                              "nom d'utilisateur", "mot de passe", "connexion"]
        },
        # Espace Business Page | Espace业务空间页
        "Espace": {
            "keywords": ["espace", "espresso", "station", "business", "portal"],
            "page_features": ["login required", "please log in",
                              "connexion requise", "veuillez vous connecter"]
        },
        # Backend/API Layer | 后端/API接口层
        "Backend/API": {
            "keywords": ["api", "gateway", "backend", "server", "interface", "rest"],
            "request_type": ["xhr", "fetch"]
        },
        # Public Front Page (Default) | 公共前端页（兜底默认）
        "Public front": {
            "default": True
        }
    }
}


# -------------------------- 3. Utility Functions | 工具函数模块 --------------------------
def standardize_url(url: str) -> str:
    """
    Standardize URL format, complete protocol, clean invalid characters
    功能：URL标准化处理，补全协议、清洗无效字符，确保URL可正常访问
    :param url: Raw input URL | 原始输入的URL字符串
    :return: Standardized valid URL, empty string if invalid | 标准化后的URL，无效URL返回空字符串
    """
    # Remove leading/trailing spaces and line breaks | 去除首尾空格、换行符
    url = url.strip()
    if not url:
        return ""

    # Complete http/https protocol | 补全http/https协议头
    parsed = urlparse(url)
    if not parsed.scheme:
        url = "https://" + url
        parsed = urlparse(url)

    # Validate legal domain and protocol | 校验域名是否合法，过滤非HTTP/HTTPS协议
    if not parsed.netloc or parsed.scheme not in ["http", "https"]:
        return ""

    # Standardize URL format, remove trailing slash | 标准化URL格式，去除末尾多余斜杠
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


def classify_url(jump_info: dict) -> str:
    """
    Classify URL according to configured rules (EN/FR only)
    功能：根据配置的分类规则，对单条跳转记录进行分类打标（仅英法双语）
    :param jump_info: Full info of single jump | 单条跳转的完整信息字典
    :return: Matched classification tag | 匹配到的分类标签
    """
    url_lower = jump_info["jump_url"].lower()
    page_html = jump_info.get("page_html", "").lower()
    request_type = jump_info.get("request_type", "")

    # 1. Priority 1: Match Authentication | 优先级1：匹配认证/安全模块
    auth_rule = CONFIG["classify_rules"]["Authentication"]
    if any(keyword in url_lower for keyword in auth_rule["keywords"]) or \
            any(feature in page_html for feature in auth_rule["page_features"]):
        return "Authentication"

    # 2. Priority 2: Match Espace | 优先级2：匹配Espace业务空间页
    espace_rule = CONFIG["classify_rules"]["Espace"]
    if any(keyword in url_lower for keyword in espace_rule["keywords"]) or \
            any(feature in page_html for feature in espace_rule["page_features"]):
        return "Espace"

    # 3. Priority 3: Match Backend/API | 优先级3：匹配Backend/API后端接口
    api_rule = CONFIG["classify_rules"]["Backend/API"]
    if any(keyword in url_lower for keyword in api_rule["keywords"]) or \
            request_type in api_rule["request_type"]:
        return "Backend/API"

    # 4. Default: Public front | 兜底：匹配公共前端页
    return "Public front"


# -------------------------- NEW EXTENDED FUNCTIONS FOR CLICK SIMULATION | 点击模拟新增扩展函数 --------------------------
def find_target_clickable_element(page) -> tuple[object, str]:
    """
    Find target clickable element according to EN/FR bilingual rules
    功能：根据英法双语规则，定位目标可点击元素
    :param page: Playwright page object | Playwright页面对象
    :return: (matched_element, element_description) | 匹配到的元素、元素描述信息
    """
    click_config = CONFIG["click_config"]
    # Merge all target keywords | 合并所有目标关键词
    all_keywords = click_config["target_keywords"]["auth_login"] + click_config["target_keywords"]["espace_access"]
    all_keywords = list(set([kw.lower() for kw in all_keywords]))

    # Traverse selector by priority | 按优先级遍历选择器
    for selector_template in click_config["selector_priority"]:
        for keyword in all_keywords:
            # Fill keyword into selector template | 填充关键词到选择器模板
            selector = selector_template.format(kw=keyword)
            try:
                # Find element, wait for 1s, no exception if not found | 查找元素，等待1秒，未找到不抛异常
                element = page.locator(selector).first
                if element.is_visible(timeout=1000) and element.is_enabled(timeout=1000):
                    element_desc = f"Selector: {selector}, Keyword: {keyword}"
                    return element, element_desc
            except:
                continue

    # No matched element found | 未找到匹配元素
    return None, "No matched clickable element found"


def simulate_user_click(page, element) -> tuple[bool, str]:
    """
    Simulate real user click action, handle scroll and exception
    功能：模拟真实用户点击动作，处理滚动、异常兜底
    :param page: Playwright page object | Playwright页面对象
    :param element: Target element to click | 待点击的目标元素
    :return: (click_success, message) | 点击是否成功、结果信息
    """
    click_config = CONFIG["click_config"]
    retry_times = 0
    click_success = False
    message = ""

    while retry_times <= click_config["max_click_retry"] and not click_success:
        try:
            # Scroll element into view, avoid being blocked | 滚动元素到可视区域，避免被遮挡
            element.scroll_into_view_if_needed(timeout=2000)
            # Wait for element to be clickable | 等待元素可点击
            element.wait_for(state="visible", timeout=2000)
            element.wait_for(state="enabled", timeout=2000)
            # Simulate real user click (not forced JS click) | 模拟真实用户点击（非强制JS点击）
            element.click(timeout=3000, force=False)
            click_success = True
            message = "Click executed successfully"
        except Exception as e:
            retry_times += 1
            message = f"Click failed, retry {retry_times} times, error: {str(e)}"

    return click_success, message


def init_output_excel() -> None:
    """
    Initialize output Excel file, create sheets with headers (added click-related columns)
    功能：初始化输出Excel文件，创建汇总表和详情表（新增点击相关列）
    """
    # Create new workbook | 新建工作簿
    wb = Workbook()

    # Create Summary Sheet | 创建汇总表
    summary_ws = wb.active
    summary_ws.title = CONFIG["summary_sheet_name"]
    # Summary Headers (added click-related columns) | 汇总表表头（新增点击相关列）
    # summary_headers = [
    #     "序号", "原始URL", "最终落地URL", "总跳转次数", "完整跳转路径",
    #     "主分类标签", "是否触发点击", "点击元素信息", "处理状态", "错误信息", "处理完成时间"
    # ]
    summary_headers = [
        "No./序号", "Original URL/原始URL", "Final Landing URL/最终落地URL", "Total Redirect Times/总跳转次数",
        "Complete Redirect Path/完整跳转路径",
        "Main Category Tag/主分类标签", "Click Triggered/是否触发点击", "Clicked Element Info/点击元素信息",
        "Processing Status/处理状态", "Error Message/错误信息", "Processing Completion Time/处理完成时间"
    ]
    summary_ws.append(summary_headers)

    # Create Detail Sheet | 创建详情表
    detail_ws = wb.create_sheet(CONFIG["detail_sheet_name"])
    # Detail Headers (added jump trigger type) | 详情表表头（新增跳转触发类型）
    # detail_headers = [
    #     "关联原始URL", "跳转序号", "跳转URL", "来源URL", "跳转类型",
    #     "HTTP状态码", "跳转触发方式", "分类标签", "跳转时间戳"
    # ]
    detail_headers = [
        "Associated Original URL/关联原始URL", "Redirect Serial No./跳转序号", "Redirect URL/跳转URL",
        "Source URL/来源URL", "Redirect Type/跳转类型",
        "HTTP Status Code/HTTP状态码", "Redirect Trigger Mode/跳转触发方式", "Category Tag/分类标签",
        "Redirect Timestamp/跳转时间戳"
    ]
    detail_ws.append(detail_headers)

    # Header Style Formatting | 表头格式美化
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    for ws in [summary_ws, detail_ws]:
        for cell in ws[1]:
            cell.font = bold_font
            cell.alignment = center_align
        ws.column_dimensions["A"].width = 8
        ws.column_dimensions["B"].width = 50
        ws.column_dimensions["C"].width = 50

    # Save initialized file | 保存初始化文件
    wb.save(CONFIG["output_excel_path"])
    print(f"✅ Output Excel initialized successfully: {CONFIG['output_excel_path']}")


def get_processed_urls() -> set:
    """
    Get already processed URLs for breakpoint resume
    功能：读取已处理完成的URL，实现断点续跑，避免重复处理
    :return: Set of successfully processed URLs | 已Success/处理成功的URL集合
    """
    if not os.path.exists(CONFIG["output_excel_path"]):
        return set()

    try:
        wb = load_workbook(CONFIG["output_excel_path"], read_only=True)
        summary_ws = wb[CONFIG["summary_sheet_name"]]
        processed_urls = set()
        # Read column 2 (raw URL) with success status | 读取第2列（原始URL），处理状态为成功的URL
        for row in summary_ws.iter_rows(min_row=2, values_only=True):
            raw_url = row[1]
            status = row[8]
            if raw_url and status == "Success/处理成功":
                processed_urls.add(raw_url)
        wb.close()
        return processed_urls
    except Exception as e:
        print(f"⚠️  Failed to read processed URLs, will reprocess all: {str(e)}")
        return set()


# -------------------------- 4. Core Business Module | 核心业务模块 --------------------------
def capture_url_jump_chain(context, raw_url: str) -> tuple[list, dict]:
    """
    Core function: Simulate user access, capture full jump chain (auto + click-triggered)
    功能：核心函数，模拟用户访问URL，捕获完整跳转链路（自动跳转+点击触发跳转）
    :param context: Playwright browser context | Playwright浏览器上下文
    :param raw_url: Standardized raw URL | 标准化后的原始URL
    :return: (jump_chain, summary_info) | 完整跳转链路列表、汇总信息字典
    """
    click_config = CONFIG["click_config"]
    # Initialize jump chain storage | 初始化跳转链路存储
    jump_chain = []
    jump_id = 1
    # Initialize click-related info | 初始化点击相关信息
    click_triggered = False
    click_element_desc = ""
    # Initialize summary info | 初始化汇总信息
    summary_info = {
        "raw_url": raw_url,
        "final_url": raw_url,
        "jump_count": 0,
        "jump_path": raw_url,
        "main_classify": "Public front",
        "click_triggered": "否",
        "click_element_desc": "",
        "status": "Success/处理成功",
        "error_msg": "",
        "finish_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    # -------------------------- Jump Event Listeners | 跳转事件监听器 --------------------------
    navigation_events = []
    current_trigger_type = "自动跳转"  # Default trigger type | 默认触发类型为自动跳转

    # Listen page navigation (capture client-side jump) | 监听页面导航事件（捕获客户端跳转）
    def on_framenavigated(frame):
        nonlocal jump_id, current_trigger_type
        if frame == page.main_frame:  # Only listen main frame, ignore iframe | 只监听主框架，忽略iframe
            jump_url = frame.url
            if not navigation_events or navigation_events[-1]["jump_url"] != jump_url:
                navigation_events.append({
                    "jump_id": jump_id,
                    "jump_url": jump_url,
                    "from_url": page.url if jump_id > 1 else "",
                    "jump_type": "客户端跳转/页面导航",
                    "status_code": 200,
                    "trigger_type": current_trigger_type,
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                })
                jump_id += 1

    # Listen network response (capture server 3xx redirect) | 监听网络响应事件（捕获服务器3xx重定向）
    def on_response(response):
        nonlocal jump_id, current_trigger_type
        if 300 <= response.status < 400:  # Capture 3xx redirect | 捕获3xx重定向
            redirect_url = response.headers.get("location", "")
            if redirect_url:
                redirect_url = standardize_url(redirect_url)
                if redirect_url:
                    navigation_events.append({
                        "jump_id": jump_id,
                        "jump_url": redirect_url,
                        "from_url": response.request.url,
                        "jump_type": "服务器3xx重定向",
                        "status_code": response.status,
                        "trigger_type": current_trigger_type,
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                    })
                    jump_id += 1

    # Listen new window opened after click | 监听点击后打开的新窗口
    new_page = None

    def on_page_opened(new_page_obj):
        nonlocal new_page
        new_page = new_page_obj
        print("🔍 New window opened after click, start capturing jump chain")

    # -------------------------- Step 1: Initialize Page & Access URL | 步骤1：初始化页面、访问URL --------------------------
    page = context.new_page()
    # Register listeners | 注册监听器
    page.on("framenavigated", on_framenavigated)
    page.on("response", on_response)
    if click_config["capture_new_window"]:
        context.on("page", on_page_opened)

    try:
        # Simulate user input URL and press enter | 模拟人工在地址栏输入URL并回车访问
        page.goto(
            raw_url,
            timeout=CONFIG["timeout"],
            wait_until="networkidle"  # Wait for network idle | 等待网络空闲
        )
        # Extra wait for delayed auto jump | 额外等待延迟自动跳转
        page.wait_for_timeout(CONFIG["wait_for_network_idle"])

        # -------------------------- Step 2: Click Simulation (NEW EXTENSION) | 步骤2：点击模拟（新增扩展） --------------------------
        if click_config["enable_click_function"]:
            # Find target clickable element | 定位目标可点击元素
            target_element, element_desc = find_target_clickable_element(page)
            click_element_desc = element_desc

            if target_element is not None:
                # Simulate user click | 模拟用户点击
                click_success, click_msg = simulate_user_click(page, target_element)
                if click_success:
                    click_triggered = True
                    current_trigger_type = "点击触发跳转"  # Update trigger type after click | 点击后更新触发类型
                    summary_info["click_triggered"] = "是"
                    print(f"✅ {click_msg}, element: {element_desc}")
                    # Wait for jump execution after click | 点击后等待跳转执行
                    page.wait_for_timeout(click_config["wait_after_click"])
                    # Wait for network idle | 等待网络空闲
                    try:
                        page.wait_for_load_state("networkidle", timeout=CONFIG["timeout"])
                    except:
                        pass
                    # Capture jump in new window | 捕获新窗口内的跳转
                    if new_page is not None and click_config["capture_new_window"]:
                        try:
                            new_page.wait_for_load_state("networkidle", timeout=CONFIG["timeout"])
                            # Add new window final URL to jump chain | 将新窗口最终URL加入跳转链路
                            final_new_url = new_page.url
                            if is_valid_url(final_new_url):
                                navigation_events.append({
                                    "jump_id": jump_id,
                                    "jump_url": final_new_url,
                                    "from_url": page.url,
                                    "jump_type": "新窗口页面跳转",
                                    "status_code": 200,
                                    "trigger_type": current_trigger_type,
                                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                                })
                                jump_id += 1
                        except Exception as e:
                            print(f"⚠️  Failed to capture new window jump: {str(e)}")
                else:
                    print(f"⚠️  {click_msg}, element: {element_desc}")
            else:
                print(f"ℹ️  {element_desc}, skip click simulation")

        # -------------------------- Step 3: Process Full Jump Chain | 步骤3：处理完整跳转链路 --------------------------
        # Get final landing URL | 获取最终落地URL
        final_url = new_page.url if (new_page is not None and click_config["capture_new_window"]) else page.url
        summary_info["final_url"] = final_url

        # Supplement final landing URL to chain | 补充最终落地页到跳转链路
        if not navigation_events or navigation_events[-1]["jump_url"] != final_url:
            navigation_events.append({
                "jump_id": jump_id,
                "jump_url": final_url,
                "from_url": navigation_events[-1]["jump_url"] if navigation_events else "",
                "jump_type": "最终落地页",
                "status_code": 200,
                "trigger_type": current_trigger_type,
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            })

        # Filter valid jumps, deduplicate, limit max jump count | 过滤有效跳转、去重、限制最大跳转次数
        valid_jumps = []
        seen_urls = set()
        for jump in sorted(navigation_events, key=lambda x: x["jump_id"]):
            jump_url = jump["jump_url"]
            if is_valid_url(jump_url) and jump_url not in seen_urls and len(valid_jumps) < CONFIG["max_jump_count"]:
                # Supplement page HTML for classification | 补充页面HTML用于分类
                try:
                    if jump["jump_type"] != "服务器3xx重定向":
                        temp_page = context.new_page()
                        temp_page.goto(jump_url, timeout=10000, wait_until="domcontentloaded")
                        jump["page_html"] = temp_page.content()[:5000]
                        temp_page.close()
                    else:
                        jump["page_html"] = ""
                except:
                    jump["page_html"] = ""

                # Classify URL | 分类打标
                jump["classify"] = classify_url(jump)
                valid_jumps.append(jump)
                seen_urls.add(jump_url)

        # Update summary info | 更新汇总信息
        summary_info["jump_count"] = len(valid_jumps)
        summary_info["jump_path"] = " → ".join([j["jump_url"] for j in valid_jumps])
        summary_info["click_element_desc"] = click_element_desc
        if valid_jumps:
            summary_info["main_classify"] = valid_jumps[-1]["classify"]
        jump_chain = valid_jumps

    # -------------------------- Exception Handling | 异常捕获处理 --------------------------
    except PlaywrightTimeoutError:
        summary_info["status"] = "Failure/处理失败"
        summary_info["error_msg"] = "Page access timeout, exceed max waiting time"
    except Exception as e:
        summary_info["status"] = "Failure/处理失败"
        summary_info["error_msg"] = f"Access exception: {str(e)}"
    finally:
        # Remove listeners, release resources | 移除监听器，释放资源
        page.remove_listener("framenavigated", on_framenavigated)
        page.remove_listener("response", on_response)
        context.remove_listener("page", on_page_opened)
        if new_page is not None:
            new_page.close()
        page.close()

    return jump_chain, summary_info


def write_result_to_excel(summary_info: dict, jump_chain: list) -> None:
    """
    Write single URL processing result to Excel file
    功能：将单条URL的处理结果写入Excel文件
    :param summary_info: Summary info dict | 汇总信息字典
    :param jump_chain: Jump chain list | 跳转链路列表
    """
    # Load output Excel | 加载输出Excel文件
    wb = load_workbook(CONFIG["output_excel_path"])
    summary_ws = wb[CONFIG["summary_sheet_name"]]
    detail_ws = wb[CONFIG["detail_sheet_name"]]

    # Write to Summary Sheet | 写入汇总表
    max_row = summary_ws.max_row
    serial_number = max_row
    summary_row = [
        serial_number,
        summary_info["raw_url"],
        summary_info["final_url"],
        summary_info["jump_count"],
        summary_info["jump_path"],
        summary_info["main_classify"],
        summary_info["click_triggered"],
        summary_info["click_element_desc"],
        summary_info["status"],
        summary_info["error_msg"],
        summary_info["finish_time"]
    ]
    summary_ws.append(summary_row)

    # Write to Detail Sheet | 写入详情表
    for jump in jump_chain:
        detail_row = [
            summary_info["raw_url"],
            jump["jump_id"],
            jump["jump_url"],
            jump["from_url"],
            jump["jump_type"],
            jump["status_code"],
            jump["trigger_type"],
            jump["classify"],
            jump["timestamp"]
        ]
        detail_ws.append(detail_row)

    # Save file | 保存文件
    wb.save(CONFIG["output_excel_path"])
    wb.close()


# -------------------------- 5. Main Function | 主函数，串联整个流程 --------------------------
def main():
    print("=" * 60)
    print("🚀 URL Jump Chain Capture & Classification Script (With Click Simulation)")
    print("🚀 URL跳转链路捕获与分类脚本（含点击模拟扩展）")
    print("=" * 60)

    # Step 1: Validate Input File | 步骤1：校验输入文件
    if not os.path.exists(CONFIG["input_excel_path"]):
        print(f"❌ Input file not found: {CONFIG['input_excel_path']}, please check the path")
        return

    # Initialize output file if not exists | 初始化输出文件（不存在则创建）
    if not os.path.exists(CONFIG["output_excel_path"]):
        init_output_excel()

    # Step 2: Read and Preprocess URL List | 步骤2：读取并预处理URL列表
    try:
        input_wb = load_workbook(CONFIG["input_excel_path"], read_only=True)
        input_ws = input_wb[CONFIG["input_sheet_name"]]
    except Exception as e:
        print(f"❌ Failed to read input Excel: {str(e)}, please check sheet name and file format")
        return

    # Read and standardize URLs | 读取URL并标准化
    raw_url_list = []
    for row in input_ws.iter_rows(min_row=1, values_only=True):
        raw_url = row[0]
        if raw_url:
            std_url = standardize_url(str(raw_url))
            if is_valid_url(std_url):
                raw_url_list.append(std_url)
    input_wb.close()

    # Deduplicate and filter processed URLs | 去重、过滤已处理URL
    raw_url_list = list(set(raw_url_list))
    processed_urls = get_processed_urls()
    pending_urls = [url for url in raw_url_list if url not in processed_urls]

    print(f"📊 Total valid URLs: {len(raw_url_list)}")
    print(f"✅ Already processed URLs: {len(processed_urls)}")
    print(f"⏳ Pending URLs: {len(pending_urls)}")
    if not pending_urls:
        print("🎉 All URLs have been processed, script exit")
        return

    # Step 3: Launch Browser and Process URLs | 步骤3：启动浏览器，批量处理URL
    with sync_playwright() as p:
        # Launch Chromium browser | 启动Chromium浏览器（已修复：移除了错误的参数）
        browser = p.chromium.launch(
            headless=CONFIG["headless_mode"],
            proxy=CONFIG["proxy_config"]
        )
        # Create browser context | 创建浏览器上下文（已修复：正确放置 ignore_https_errors）
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            ignore_https_errors=True
        )
        context.set_default_timeout(CONFIG["timeout"])

        # Batch process pending URLs | 批量处理待处理URL
        success_count = 0
        fail_count = 0
        for index, raw_url in enumerate(pending_urls, 1):
            print(f"\n--- Processing {index}/{len(pending_urls)}: {raw_url} ---")
            retry_times = 0
            handle_success = False

            # Retry mechanism | 失败重试机制
            while retry_times <= CONFIG["max_retry_times"] and not handle_success:
                try:
                    # Core: Capture full jump chain | 核心：捕获完整跳转链路
                    jump_chain, summary_info = capture_url_jump_chain(context, raw_url)
                    # Write result to Excel | 写入结果到Excel
                    write_result_to_excel(summary_info, jump_chain)

                    # Statistics | 处理结果统计
                    if summary_info["status"] == "Success/处理成功":
                        handle_success = True
                        success_count += 1
                        print(
                            f"✅ Process success | Jump count: {summary_info['jump_count']} | Main classify: {summary_info['main_classify']} | Click triggered: {summary_info['click_triggered']}")
                    else:
                        retry_times += 1
                        print(f"⚠️  Process failed, retry {retry_times} times | Error: {summary_info['error_msg']}")

                except Exception as e:
                    retry_times += 1
                    print(f"⚠️  Process exception, retry {retry_times} times | Exception: {str(e)}")

            # Record failed result after max retry | 重试多次仍失败，记录失败结果
            if not handle_success:
                fail_count += 1
                fail_summary = {
                    "raw_url": raw_url,
                    "final_url": "",
                    "jump_count": 0,
                    "jump_path": "",
                    "main_classify": "",
                    "click_triggered": "否",
                    "click_element_desc": "",
                    "status": "Failure/处理失败",
                    "error_msg": f"Failed after {CONFIG['max_retry_times']} retries",
                    "finish_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                write_result_to_excel(fail_summary, [])
                print(f"❌ Process failed, result recorded")

        # Close browser | 关闭浏览器
        context.close()
        browser.close()

    # Step 4: Final Statistics | 步骤4：处理完成，输出统计结果
    print("\n" + "=" * 60)
    print("🎉 Script execution completed!")
    print(f"📊 Processing Statistics: Success {success_count} | Failed {fail_count} | Total {len(pending_urls)}")
    print(f"📁 Result file saved to: {os.path.abspath(CONFIG['output_excel_path'])}")
    print("=" * 60)


# Script Entry | 脚本入口
if __name__ == "__main__":
    main()
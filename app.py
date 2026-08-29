"""
订单披露 - Streamlit 应用
通过飞书应用凭证实时读取电子表格数据，展示在投订单
"""
import streamlit as st
import requests
import json
import time
import base64
import os
from datetime import datetime, timedelta, timezone
import pandas as pd
import streamlit.components.v1 as components

# ===== 飞书应用凭证（从环境变量或 Streamlit secrets 读取）=====
try:
    APP_ID = st.secrets["FEISHU_APP_ID"]
    APP_SECRET = st.secrets["FEISHU_APP_SECRET"]
except Exception:
    APP_ID = os.environ.get("FEISHU_APP_ID", "")
    APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "")

# ===== 电子表格配置 =====
SPREADSHEET_TOKEN = "RsPys96vjhOATftKC7mcj5UGn4c"
SHEET_ID = "1d74a0"          # 订单明细
MATCH_SHEET_ID = "bPtGPo"    # 包名匹配
TOTAL_ROWS = 631
MATCH_TOTAL_ROWS = 628

# ===== 表头定义（与电子表格实际列顺序一致）=====
HEADERS = ['预算源', '任务类型', '配置号', '分端', '广告主', '包名', '接口文档', '产品',
           '渠道号', '合作价格', '需求量级', '上线时间', '下线时间', '上下线状态',
           '回传维度', '考核', '考核数值', 'RTA', '考核备注', '下载链接',
           '归属', '中台负责人', '是否打满', '是否披露', '']

# ===== Token 缓存 =====
_token_cache = {"token": None, "expire": 0}

def get_tenant_access_token():
    """获取 tenant_access_token"""
    if _token_cache["token"] and time.time() < _token_cache["expire"]:
        return _token_cache["token"]
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    resp = requests.post(url, json={"app_id": APP_ID, "app_secret": APP_SECRET})
    data = resp.json()
    if data.get("code") != 0:
        return None
    token = data["tenant_access_token"]
    _token_cache["token"] = token
    _token_cache["expire"] = time.time() + data.get("expire", 7200) - 300
    return token

def feishu_get(path, params=None):
    """GET 请求飞书 API"""
    token = get_tenant_access_token()
    if not token:
        return None
    url = f"https://open.feishu.cn/open-apis{path}"
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, params=params)
    return resp.json()

def read_sheet_range(sheet_id, start_row, end_row, col_start="A", col_end="X"):
    """读取电子表格数据"""
    from urllib.parse import quote
    range_str = f"{sheet_id}!{col_start}{start_row}:{col_end}{end_row}"
    encoded = quote(range_str, safe='')
    path = f"/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{encoded}"
    data = feishu_get(path, params={"valueRenderOption": "ToString"})
    if not data or data.get("code") != 0:
        return []
    return data.get("data", {}).get("valueRange", {}).get("values", [])

@st.cache_data(ttl=300, show_spinner=False)
def get_sheet_row_count(sheet_id):
    """动态获取工作表的实际行数"""
    path = f"/sheets/v3/spreadsheets/{SPREADSHEET_TOKEN}/sheets/{sheet_id}"
    data = feishu_get(path)
    if data and data.get("code") == 0:
        props = data.get("data", {}).get("sheet", {}).get("grid_properties", {})
        return props.get("rowCount", 1000)
    return 1000

def excel_serial_to_date(val):
    """Excel 序列号转日期"""
    if val is None or val == "":
        return ""
    try:
        serial = int(val)
        dt = datetime(1899, 12, 30) + timedelta(days=serial)
        return dt.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return str(val)

@st.cache_data(show_spinner=False)
def load_gif_base64(filename):
    """读取 GIF 文件并返回 base64 编码（缓存结果）"""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", filename)
    if os.path.exists(path):
        with open(path, "rb") as f:
            data = f.read()
        return base64.b64encode(data).decode()
    return None

@st.cache_data(ttl=300, show_spinner=False)
def fetch_active_orders():
    """从飞书电子表格读取在投订单（5分钟缓存）"""
    # 动态获取实际行数，避免数据被截断
    total_rows = get_sheet_row_count(SHEET_ID)
    match_rows = get_sheet_row_count(MATCH_SHEET_ID)

    all_rows = []
    row = 1
    batch_size = 100
    while row <= total_rows:
        end = min(row + batch_size - 1, total_rows)
        batch = read_sheet_range(SHEET_ID, row, end)
        all_rows.extend(batch)
        row = end + 1

    if not all_rows:
        return pd.DataFrame(), datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M")

    actual_head = [str(x).strip() if x is not None else "" for x in (all_rows[0] if all_rows else [])]
    # 根据 HEADERS 与实际表头行做精确匹配 → 得到每个列名在原始数据中的 index（鲁棒：飞书列顺序 / 加删列不再崩）
    col_idx = {}
    for h in HEADERS:
        h_s = str(h).strip()
        for i, hh in enumerate(actual_head):
            if hh == h_s:
                col_idx[h] = i
                break
    # 未匹配的列会缺失 col_idx，后续所有取列都用 .get(col, None) 做保护

    data_rows = all_rows[1:]

    # 构建包名、归属匹配表（匹配表：A列=产品, B列=包名, C列=归属）
    pkg_map = {}
    attr_map = {}
    row = 1
    while row <= match_rows:
        end = min(row + batch_size - 1, match_rows)
        batch = read_sheet_range(MATCH_SHEET_ID, row, end, "A", "C")
        for r in batch[1:] if row == 1 else batch:
            if len(r) >= 2 and r[0] and r[1]:
                pkg_map[str(r[0]).strip()] = str(r[1]).strip()
            if len(r) >= 3 and r[0] and r[2]:
                attr_map[str(r[0]).strip()] = str(r[2]).strip()
        row = end + 1

    STATUS_COL    = col_idx.get("上下线状态")
    DISCLOSE_COL  = col_idx.get("是否披露")
    PRODUCT_COL   = col_idx.get("产品")
    PACKAGE_COL   = col_idx.get("包名")
    ATTR_COL      = col_idx.get("归属")
    ONLINE_TIME_COL  = col_idx.get("上线时间")
    OFFLINE_TIME_COL = col_idx.get("下线时间")

    unmatched_pkg = []
    unmatched_attr = []

    active_orders = []
    for r in data_rows:
        def _get(col):
            return r[col] if col is not None and len(r) > col else None

        status   = _get(STATUS_COL)
        disclose = _get(DISCLOSE_COL)
        if status == "在投" and str(disclose).strip() == "1":
            product = str(r[PRODUCT_COL]).strip() if (PRODUCT_COL is not None and len(r) > PRODUCT_COL and r[PRODUCT_COL]) else ""
            pkg_val  = _get(PACKAGE_COL)
            attr_val = _get(ATTR_COL)
            pkg_was_vlookup = isinstance(pkg_val, str) and pkg_val.startswith("VLOOKUP")
            attr_was_vlookup = isinstance(attr_val, str) and attr_val.startswith("VLOOKUP")
            if pkg_was_vlookup:
                new_pkg = pkg_map.get(product, "")
                if PACKAGE_COL is not None:
                    r[PACKAGE_COL] = new_pkg
                if not new_pkg and product and product not in unmatched_pkg:
                    unmatched_pkg.append(product)
            if attr_was_vlookup:
                new_attr = attr_map.get(product, "")
                if ATTR_COL is not None:
                    r[ATTR_COL] = new_attr
                if not new_attr and product and product not in unmatched_attr:
                    unmatched_attr.append(product)
            if ONLINE_TIME_COL is not None and len(r) > ONLINE_TIME_COL:
                r[ONLINE_TIME_COL] = excel_serial_to_date(r[ONLINE_TIME_COL])
            if OFFLINE_TIME_COL is not None and len(r) > OFFLINE_TIME_COL:
                r[OFFLINE_TIME_COL] = excel_serial_to_date(r[OFFLINE_TIME_COL])
            active_orders.append(r)

    # 转为 DataFrame（列顺序严格按 HEADERS，缺的列自动填空，避免实际数据少列导致后续 display_cols KeyError）
    rows_data = []
    for order in active_orders:
        row_dict = {}
        for i, header in enumerate(HEADERS):
            idx = col_idx.get(header, i)
            row_dict[header] = order[idx] if idx is not None and idx < len(order) else ""
        rows_data.append(row_dict)

    warnings = {}
    if unmatched_pkg:
        warnings["包名未匹配（VLOOKUP 返回空，检查匹配表）"] = unmatched_pkg
    if unmatched_attr:
        warnings["归属未匹配（VLOOKUP 返回空，检查匹配表）"] = unmatched_attr
    # 表头缺失的列
    missing_cols = [h for h in HEADERS if h not in col_idx]
    if missing_cols:
        warnings["飞书表头缺少这些列（已按空值填充，建议核对列顺序/列名）"] = missing_cols

    fetch_meta = {"warnings": warnings, "col_idx": col_idx}

    return pd.DataFrame(rows_data), datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M"), fetch_meta

def main():
    # 页面配置
    st.set_page_config(
        page_title="订单LIST",
        page_icon="📊",
        layout="wide"
    )

    # 自定义 CSS：统计卡片彩色、自定义HTML表格（无竖线+颜色）
    st.markdown("""
    <style>
    /* 筛选区域背景色 */
    .filter-container {
        background-color: #f7f8fa;
        padding: 14px 20px;
        border-radius: 8px;
        margin-bottom: 10px;
    }
    /* 去除 Streamlit columns 之间的竖线 */
    div[data-testid="stColumn"] {
        border: none !important;
    }
    div[data-testid="stColumn"] > div:first-child {
        border: none !important;
        box-shadow: none !important;
    }
    /* 下拉框样式 */
    .stSelectbox select {
        border-radius: 6px;
    }
    /* 隐藏 Streamlit 默认 divider */
    .st-emotion-cache-1aj19jr {
        display: none;
    }
    section[data-testid="stSidebar"] {
        display: none !important;
    }
    /* ===== 彩色统计卡片 ===== */
    .stat-card {
        background: #ffffff;
        border: 1px solid #eef0f3;
        border-radius: 10px;
        padding: 16px 20px;
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    .stat-card:hover {
        box-shadow: 0 4px 14px rgba(0,0,0,0.06);
    }
    .stat-label {
        font-size: 13px;
        color: #8a8f99;
        margin-bottom: 6px;
        font-weight: 500;
    }
    .stat-value-green {
        font-size: 30px;
        font-weight: 700;
        color: #00b578;
        line-height: 1.1;
    }
    .stat-value-blue {
        font-size: 30px;
        font-weight: 700;
        color: #1677ff;
        line-height: 1.1;
    }
    .stat-value-orange {
        font-size: 30px;
        font-weight: 700;
        color: #ff8800;
        line-height: 1.1;
    }
    .stat-value-gray {
        font-size: 30px;
        font-weight: 700;
        color: #4e5969;
        line-height: 1.1;
    }
    /* ===== 自定义订单表格（仅横线，无竖线） ===== */
    .order-table {
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        font-size: 13.5px;
        color: #1d2129;
        background-color: #ffffff;
    }
    .order-table thead th {
        background-color: #ffffff;
        font-weight: 600;
        color: #4e5969;
        padding: 8px 12px;
        text-align: left;
        border: none;
        border-bottom: 1px solid #e5e6eb;
        position: sticky;
        top: 0;
        z-index: 1;
    }
    .order-table tbody td {
        padding: 9px 12px;
        border: none;
        border-bottom: 1px solid #f0f1f3;
        vertical-align: middle;
    }
    .order-table tbody tr:last-child td {
        border-bottom: none;
    }
    .order-table tbody tr:hover td {
        background-color: #f7f8fa;
    }
    /* 状态胶囊 */
    .status-tag-active {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 12px;
        background-color: #e8ffea;
        color: #00b578;
        font-size: 12px;
        font-weight: 600;
    }
    /* 合作价格（高亮色） */
    .price-cell {
        color: #ff7d00;
        font-weight: 600;
    }
    /* 备注列溢出省略 */
    .note-cell {
        max-width: 220px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        display: inline-block;
        width: 100%;
    }
    .note-cell:hover {
        white-space: normal;
        overflow: visible;
    }
    </style>
    """, unsafe_allow_html=True)

    # 标题（带小熊动图）
    bear_b64 = load_gif_base64("bear.gif")
    if bear_b64:
        col_img, col_txt = st.columns([1, 15])
        with col_img:
            components.html(
                f'<img src="data:image/gif;base64,{bear_b64}" style="width:50px;height:auto;border-radius:8px;">',
                height=60
            )
        with col_txt:
            st.markdown("# 订单LIST")
    else:
        st.markdown("# 订单LIST")
    st.caption("数据来源：飞书电子表格「订单明细」| 仅展示「在投」状态订单")

    # 读取数据（无凭证时使用 Demo 假数据，便于预览表格样式）
    if not APP_ID or not APP_SECRET:
        st.info("ℹ️ 未配置飞书凭证，当前使用 Demo 假数据预览界面样式。配置凭证后将自动读取真实数据。")
        st.code('FEISHU_APP_ID = "your_app_id"\nFEISHU_APP_SECRET = "your_app_secret"', language="toml")
        # 生成 Demo 假数据
        import random
        random.seed(42)
        demo_rows = []
        products = ["猿辅导", "快手", "快手极速版", "自如", "百度网盘", "学而思网校", "七猫免费小说", "英语天天练", "360文库", "百度地图", "懂车帝"]
        platforms = ["安卓", "iOS", "安卓&iOS"]
        configs = ["微博-猿辅导fid2232", "微博-快手拉新-已卸载组合-fid3537", "微博-快手极速版拉新-已卸载-fid3538",
                   "微博-自如fid373(ocpx)", "微博-百度网盘fid735(ocpx)", "微博-学而思网校fid752(ocpx)",
                   "喜马拉雅-七猫免费小说174", "喜马拉雅-英语天天练1481593", "微博-360文库新客下单3993",
                   "微博-百度地图fid4018", "微博-懂车帝fid4171"]
        sources = ["媒体", "直客", "信息流", "搜索", "官方"]
        task_types = ["拉新"]
        rtas = ["是", "无", "次留率50%", "次留率35%", "次留率45%", "下单率3.5%", "下单率5%", "16-24点,次留34%,下单5%", "下单率30%", "下单率35%"]
        callbacks = ["激活", "下单"]
        statuses = ["在投"]
        attributions = ["快手", "百度", "官方", "自营", "字节", "腾讯"]
        price_ranges = [(3, 10), (6.5, 12.5), (10, 20), (55, 60), (0.5, 5)]
        demands = [20, 50, 100, 150, 200, 300, 500, 1000, 1500, 2000, 3000, 5000]

        for i in range(120):
            product = random.choice(products)
            pkg = f"com.{product.lower().replace('&','').replace(' ','')}.browser" if product in ["猿辅导","百度网盘"] else \
                  f"com.smile.gifmaker" if product == "快手" else \
                  f"com.kuaishou.nebula" if product == "快手极速版" else \
                  f"com.ziroom.android" if product == "自如" else \
                  f"com.xueersi.online" if product == "学而思网校" else \
                  f"com.qimao.reader" if product == "七猫免费小说" else \
                  f"com.english.daily.practice" if product == "英语天天练" else \
                  f"com.qihoo.pluginbox.wenku" if product == "360文库" else \
                  f"com.baidu.BaiduMap" if product == "百度地图" else \
                  f"com.ss.android.auto" if product == "懂车帝" else f"com.example.app{i}"
            price_range = random.choice(price_ranges)
            price = round(random.uniform(*price_range), 1)
            demand = random.choice(demands)
            platform = random.choice(platforms)
            callback = random.choice(callbacks)
            rta = random.choice(rtas) if callback == "激活" else random.choice(["nan", "无"])
            note = "定向女性，25-50岁；导课率50%（进入APP点击免费领课）" if product == "学而思网校" else \
                   "付费率25%（注册后会有弹窗，0元领取）" if product == "猿辅导" else \
                   "次留率50%" if rta == "次留率50%" else \
                   "" if rta in ["nan","无"] else rta
            row = {}
            for h in HEADERS:
                if h == '预算源': row[h] = random.choice(sources)
                elif h == '任务类型': row[h] = random.choice(task_types)
                elif h == '配置号': row[h] = random.choice(configs)
                elif h == '分端': row[h] = platform
                elif h == '广告主': row[h] = random.choice(["甲方A","甲方B","甲方C",""])
                elif h == '包名': row[h] = pkg
                elif h == '接口文档': row[h] = product
                elif h == '产品': row[h] = product
                elif h == '渠道号': row[h] = f"ch_{i+1:04d}"
                elif h == '合作价格': row[h] = price
                elif h == '需求量级': row[h] = demand
                elif h == '上线时间': row[h] = ""
                elif h == '下线时间': row[h] = ""
                elif h == '上下线状态': row[h] = random.choice(statuses)
                elif h == '回传维度': row[h] = callback
                elif h == '考核': row[h] = ""
                elif h == '考核数值': row[h] = ""
                elif h == 'RTA': row[h] = rta
                elif h == '考核备注': row[h] = note
                elif h == '下载链接': row[h] = ""
                elif h == '归属': row[h] = random.choice(attributions)
                elif h == '中台负责人': row[h] = random.choice(["张三", "李四", "王五", "赵六", "钱七", "孙八", "周九", "吴十", ""])
                elif h == '是否打满': row[h] = ""
                elif h == '是否披露': row[h] = "1"
                else: row[h] = ""
            demo_rows.append(row)
        df = pd.DataFrame(demo_rows)
        sync_time = datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M")
        fetch_meta = {"warnings": {}, "col_idx": {h: i for i, h in enumerate(HEADERS)}}
    else:
        with st.spinner("正在从飞书读取最新数据..."):
            fetch_res = fetch_active_orders()
            # 兼容旧返回值：支持 (DataFrame, time) 新格式与单 DataFrame 旧格式
            if isinstance(fetch_res, tuple):
                if len(fetch_res) >= 3:
                    df, sync_time, fetch_meta = fetch_res
                else:
                    df, sync_time = fetch_res
                    fetch_meta = {"warnings": {}, "col_idx": {h: i for i, h in enumerate(HEADERS)}}
            else:
                df = fetch_res
                sync_time = datetime.now(timezone(timedelta(hours=8))).strftime("%Y-%m-%d %H:%M")
                fetch_meta = {"warnings": {}, "col_idx": {h: i for i, h in enumerate(HEADERS)}}

        if isinstance(df, pd.DataFrame) and df.empty:
            st.warning("未读取到在投订单数据")
            # 调试信息
            with st.expander("查看调试信息"):
                token = get_tenant_access_token()
                st.write(f"APP_ID: {APP_ID[:10]}..." if APP_ID else "APP_ID: 未设置")
                st.write(f"Token 获取: {'成功' if token else '失败'}")
                if token:
                    test_data = feishu_get(f"/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{SHEET_ID}!A1:X2",
                                           params={"valueRenderOption": "ToString"})
                    st.write(f"API 测试: {test_data}")
            return

    if not isinstance(df, pd.DataFrame):
        st.warning("数据读取失败")
        return

    # 数据质量警告（包名/归属匹配失败、表头缺失列）——用户第一时间发现漏填
    warnings = fetch_meta.get("warnings", {}) if isinstance(fetch_meta, dict) else {}
    if warnings:
        with st.expander("⚠️ 数据质量提示", expanded=False):
            for title, items in warnings.items():
                st.markdown(f"**{title}**：{len(items)} 条  \n" + "、".join(f"`{x}`" for x in items[:50]) + (" 等" if len(items) > 50 else ""))

    # 统计卡片（自定义彩色圆角卡片）
    products_count = df['产品'].dropna().nunique() if '产品' in df.columns else 0
    advertisers_count = df['广告主'].dropna().nunique() if '广告主' in df.columns else 0
    update_time = sync_time

    st.markdown(f"""
    <div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:12px;margin-bottom:16px;">
        <div class="stat-card">
            <div class="stat-label">在投订单数</div>
            <div class="stat-value-green">{len(df)}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">产品数</div>
            <div class="stat-value-blue">{products_count}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">广告主数</div>
            <div class="stat-value-orange">{advertisers_count}</div>
        </div>
        <div class="stat-card">
            <div class="stat-label">飞书同步时间<small style="color:#8a8f99;font-weight:400;margin-left:4px;">(北京时间)</small></div>
            <div class="stat-value-gray" style="font-size:22px;">{update_time}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 筛选区（9 个多选筛选 + 清除筛选，同一行显示）
    st.markdown('<div class="filter-container">', unsafe_allow_html=True)

    FILTER_KEYS = {
        '配置号':   'ms_config',
        '接口文档': 'ms_api_doc',
        '产品':     'ms_product',
        '回传维度': 'ms_callback',
        '归属':     'ms_attrib',
        'RTA':      'ms_rta',
        '分端':     'ms_platform',
        '预算源':   'ms_source',
        '任务类型': 'ms_task',
    }

    cols = st.columns([1, 1, 1, 1, 1, 1, 1, 1, 1, 0.85], gap="small")

    FILTER_COL_ORDER = ['配置号', '接口文档', '产品', '回传维度', '归属', 'RTA', '分端', '预算源', '任务类型']

    def get_cascaded_options(target_col):
        """级联：目标列的可选项 = 原始 df 先应用「其他 8 个已选筛选」后的唯一值"""
        if target_col not in df.columns:
            return []
        cascaded = df
        for col in FILTER_COL_ORDER:
            if col == target_col:
                continue
            key = FILTER_KEYS.get(col)
            if not key:
                continue
            vals = st.session_state.get(key, [])
            if vals and col in cascaded.columns:
                cascaded = cascaded[cascaded[col].astype(str).isin(vals)]
                if cascaded.empty:
                    break
        return sorted([str(x) for x in cascaded[target_col].dropna().astype(str).unique()])

    for i, col_name in enumerate(FILTER_COL_ORDER):
        with cols[i]:
            key = FILTER_KEYS[col_name]
            opts = get_cascaded_options(col_name)
            # 若之前已选的值在新 options 中不存在（被级联排除了）→ 裁剪并自动重跑，避免 warning
            cur = st.session_state.get(key, [])
            if isinstance(cur, list):
                cleaned = [v for v in cur if v in opts]
            else:
                cleaned = []
            if cleaned != cur:
                st.session_state[key] = cleaned
                st.rerun()
            st.multiselect(
                col_name,
                options=opts,
                default=cleaned,
                key=key,
                placeholder="全部"
            )
    with cols[9]:
        # 占位对齐筛选器 label 高度
        st.markdown("<div style='height:29px'></div>", unsafe_allow_html=True)
        if st.button("🗑 清除筛选", use_container_width=True, key="btn_clear_filters"):
            for _k in FILTER_KEYS.values():
                if _k in st.session_state:
                    del st.session_state[_k]
            st.rerun()

    # 应用筛选（多选：空列表=全部；有值时 isin）——统一从 session_state 取，与级联渲染解耦
    filtered_df = df.copy()
    for col_name, key in FILTER_KEYS.items():
        vals = st.session_state.get(key, [])
        if vals and col_name in filtered_df.columns:
            filtered_df = filtered_df[filtered_df[col_name].astype(str).isin(vals)]

    st.markdown('</div>', unsafe_allow_html=True)

    # 数据表格
    # 在投订单明细：icon + 标题 + 导出按钮（同一行）。无论 icon 是否加载，标题只渲染一次。
    table_b64 = load_gif_base64("table_icon.gif")

    # ====== 导出：优先 Excel（xlsx，带筛选后完整列），没 openpyxl 回退到 CSV（utf-8-sig Excel 直接打开不乱码）======
    import io
    _export_cols = [c for c in HEADERS if c in filtered_df.columns]
    _export_df = filtered_df[_export_cols] if not filtered_df.empty else pd.DataFrame(columns=_export_cols)
    _file_stamp = datetime.now(timezone(timedelta(hours=8))).strftime("%Y%m%d_%H%M")
    try:
        import openpyxl  # noqa: F401
        _buf = io.BytesIO()
        with pd.ExcelWriter(_buf, engine="openpyxl") as _w:
            _export_df.to_excel(_w, index=False, sheet_name="在投订单")
        _file_bytes = _buf.getvalue()
        _mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        _file_name = f"在投订单筛选结果_{len(_export_df)}条_{_file_stamp}.xlsx"
    except Exception:
        _file_bytes = _export_df.to_csv(index=False).encode("utf-8-sig")
        _mime = "text/csv"
        _file_name = f"在投订单筛选结果_{len(_export_df)}条_{_file_stamp}.csv"

    _title_html = f"### 在投订单明细（{len(filtered_df)} 条）  \n<small style=\"color:#8a8f99;font-weight:400;\">具体测试量级需找对应中台确认</small>"

    if table_b64:
        # 布局：[gif icon 1 : [ 标题 : 导出按钮 ] 25] → 标题只渲染一次
        col_img2, col_txt2 = st.columns([1, 25])
        with col_img2:
            components.html(
                f'<img src="data:image/gif;base64,{table_b64}" style="width:35px;height:auto;">',
                height=45,
            )
        with col_txt2:
            _t2, _exp = st.columns([5, 1], gap="small")
            with _t2:
                st.markdown(_title_html, unsafe_allow_html=True)
            with _exp:
                # 占位对齐标题与 label 高度差
                st.markdown("<div style='height:29px'></div>", unsafe_allow_html=True)
                st.download_button(
                    label="📥 导出筛选结果",
                    data=_file_bytes,
                    file_name=_file_name,
                    mime=_mime,
                    use_container_width=True,
                    key="btn_export",
                    disabled=len(_export_df) == 0,
                    help="导出当前筛选结果的全部列（包含表格里隐藏的广告主/渠道号/下载链接）。没有 openpyxl 时会自动回退到 utf-8-sig 的 CSV，Excel 双击直接打开不乱码。",
                )
    else:
        # 没有 gif：标题 + 导出按钮并排
        _t2, _exp = st.columns([5, 1], gap="small")
        with _t2:
            st.markdown(_title_html, unsafe_allow_html=True)
        with _exp:
            st.markdown("<div style='height:29px'></div>", unsafe_allow_html=True)
            st.download_button(
                label="📥 导出筛选结果",
                data=_file_bytes,
                file_name=_file_name,
                mime=_mime,
                use_container_width=True,
                key="btn_export",
                disabled=len(_export_df) == 0,
            )

    # 选择要显示的列（移除不需要展示的列）
    display_cols = [c for c in HEADERS if c not in ['上线时间', '下线时间', '考核', '考核数值', '广告主', '渠道号', '下载链接', '是否打满', '是否披露', '其他备注']]
    display_df = filtered_df[display_cols] if not filtered_df.empty else filtered_df

    if not display_df.empty:
        # ===== 用自定义 HTML 表格渲染（无竖线、有颜色、整洁）=====
        def escape_html(s):
            return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") if s is not None else ""

        # 构建表头
        thead_html = "<thead><tr>"
        for col in display_cols:
            thead_html += f"<th>{escape_html(col)}</th>"
        thead_html += "</tr></thead>"

        # 构建表体
        tbody_html = "<tbody>"
        note_cols = {"考核备注"}
        price_cols = {"合作价格"}
        status_cols = {"上下线状态"}

        for _, row in display_df.iterrows():
            tbody_html += "<tr>"
            for col in display_cols:
                raw_val = row.get(col, "")
                val = escape_html(raw_val) if raw_val is not None else ""
                cell_content = val
                css_class = ""

                if col in price_cols and val:
                    # 合作价格：橙色加粗
                    try:
                        float(val)
                        cell_content = f'<span class="price-cell">{val}</span>'
                    except ValueError:
                        pass
                elif col in status_cols:
                    # 状态：绿色胶囊
                    if val == "在投":
                        cell_content = f'<span class="status-tag-active">{val}</span>'
                elif col in note_cols:
                    # 备注列：超长省略，hover显示
                    if val:
                        cell_content = f'<span class="note-cell" title="{val}">{val}</span>'
                    else:
                        cell_content = ""

                tbody_html += f'<td class="{css_class}">{cell_content}</td>'
            tbody_html += "</tr>"
        tbody_html += "</tbody>"

        table_html = f'<table class="order-table">{thead_html}{tbody_html}</table>'

        # 用 components.html 渲染表格，设置足够高度以滚动
        # 注意：components.html 在独立 iframe 中运行，必须把 CSS 样式直接内嵌进去
        components.html(
            f"""
            <style>
            /* ===== 自定义订单表格（仅横线，无竖线）===== */
            * {{ box-sizing: border-box; }}
            body {{ margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, "PingFang SC", "Microsoft YaHei", sans-serif; }}
            .order-table {{
                width: 100%;
                border-collapse: collapse;
                font-size: 13.5px;
                color: #1d2129;
                background-color: #ffffff;
            }}
            .order-table thead th {{
                background-color: #ffffff;
                font-weight: 600;
                color: #4e5969;
                padding: 10px 12px;
                text-align: left;
                /* 仅保留底部横线（表头） */
                border: none;
                border-bottom: 1.5px solid #c9cdd4;
                position: sticky;
                top: 0;
                z-index: 1;
                /* 文字平铺不换行 */
                white-space: nowrap;
            }}
            .order-table tbody td {{
                padding: 10px 12px;
                /* 仅保留底部横线（表体），去掉左右竖线 */
                border: none;
                border-bottom: 1px solid #e5e6eb;
                vertical-align: middle;
                /* 文字平铺不换行 */
                white-space: nowrap;
            }}
            .order-table tbody tr:last-child td {{
                border-bottom: none;
            }}
            .order-table tbody tr:hover td {{
                background-color: #f7f8fa;
            }}
            /* 状态胶囊 */
            .status-tag-active {{
                display: inline-block;
                padding: 2px 10px;
                border-radius: 12px;
                background-color: #e8ffea;
                color: #00b578;
                font-size: 12px;
                font-weight: 600;
            }}
            /* 合作价格（高亮色） */
            .price-cell {{
                color: #ff7d00;
                font-weight: 600;
            }}
            /* 备注列溢出省略 */
            .note-cell {{
                max-width: 220px;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
                display: inline-block;
                width: 100%;
            }}
            .note-cell:hover {{
                white-space: normal;
                overflow: visible;
            }}
            </style>
            <div style="max-height:800px;overflow-y:auto;padding-right:8px;">
                {table_html}
            </div>
            """,
            height=820,
            scrolling=True
        )

        # 下载按钮
        csv = display_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            "📥 下载 CSV",
            csv,
            file_name=f"在投订单_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )

        # ===== 可视化图表 =====
        chart_b64 = load_gif_base64("chart_icon.gif")
        if chart_b64:
            col_img3, col_txt3 = st.columns([1, 20])
            with col_img3:
                st.markdown(f'<img src="data:image/gif;base64,{chart_b64}" style="height:38px;width:38px;margin-top:4px;" alt=""/>', unsafe_allow_html=True)
            with col_txt3:
                st.markdown("### 数据概览")
        else:
            st.markdown("### 数据概览")
        import altair as alt

        # 图表区：产品筛选（多选 + 清除），Top N 固定为 15，不再滑条
        _all_products = sorted([str(x) for x in filtered_df['产品'].dropna().astype(str).unique()]) if '产品' in filtered_df.columns else []
        _chart_c1, _chart_c2 = st.columns([3, 1], gap="small")
        with _chart_c1:
            chart_product_filter = st.multiselect(
                "🔍 图表按产品筛选（多选，空=展示全部）",
                options=_all_products,
                default=[],
                key="ms_chart_product",
                placeholder="全部产品"
            )
        with _chart_c2:
            st.markdown("<div style='height:29px'></div>", unsafe_allow_html=True)
            if st.button("清除产品筛选", use_container_width=True, key="btn_clear_chart_product"):
                if "ms_chart_product" in st.session_state:
                    del st.session_state["ms_chart_product"]
                st.rerun()

        # Top N 固定为 15；数据少时取 min(实际可展示数, 15)
        _n_docs = (
            int(filtered_df["接口文档"].dropna().astype(str).nunique())
            if (not filtered_df.empty and "接口文档" in filtered_df.columns)
            else 0
        )
        _n_prod = (
            int(filtered_df["产品"].dropna().astype(str).nunique())
            if (not filtered_df.empty and "产品" in filtered_df.columns)
            else 0
        )
        top_n = max(3, min(15, max(_n_docs, _n_prod, 3)))

        chart_df = filtered_df
        if chart_product_filter and '产品' in chart_df.columns:
            chart_df = chart_df[chart_df['产品'].astype(str).isin(chart_product_filter)]

        # 空态保护：筛选后无数据时跳过画图
        if chart_df.empty:
            st.info("当前筛选条件下无数据可供图表展示")
        else:
            # 图表 1：按接口文档统计订单数（Top 15，降序，水平条形图 + 数值 label）
            doc_counts = (
                chart_df.groupby("接口文档", dropna=False)
                .size()
                .reset_index(name="订单数")
                .sort_values("订单数", ascending=False)
                .head(top_n)
            )
            _max_doc = int(doc_counts["订单数"].max()) if not doc_counts.empty else 0
            base_bar = alt.Chart(doc_counts).encode(
                x=alt.X("订单数:Q", title="订单数量", axis=alt.Axis(grid=True, gridColor="#f2f3f5", domain=False, ticks=False),
                        scale=alt.Scale(domain=[0, _max_doc * 1.18]) if _max_doc > 0 else alt.Undefined),
                y=alt.Y(
                    "接口文档:N",
                    title=None,
                    sort=alt.EncodingSortField(field="订单数", op="sum", order="descending"),
                    axis=alt.Axis(domain=False, ticks=False, labelLimit=250),
                ),
            )
            bar_chart = base_bar.mark_bar(color="#6C8EAD", cornerRadiusTopRight=4, cornerRadiusBottomRight=4).encode(
                tooltip=["接口文档:N", "订单数:Q"],
            )
            text_chart = base_bar.mark_text(
                align="left", baseline="middle", dx=5, fontSize=12, color="#4e5969", fontWeight=600,
            ).encode(
                text=alt.Text("订单数:Q", format="d"),
            )
            chart1 = (
                (bar_chart + text_chart)
                .properties(height=max(320, 34 * len(doc_counts)))
                .configure_view(stroke=None)
                .configure_axis(
                    labelFont="Microsoft YaHei",
                    titleFont="Microsoft YaHei",
                    labelFontSize=12,
                    titleFontSize=13,
                    labelColor="#4e5969",
                    titleColor="#4e5969",
                )
                .configure_title(font="Microsoft YaHei")
                .configure_legend(titleFont="Microsoft YaHei", labelFont="Microsoft YaHei")
            )

            # 图表 2：Top 产品 × 预算源 堆叠柱状图（按产品总订单数 Top N + 每块标数值）
            prod_src_counts = (
                chart_df.groupby(["产品", "预算源"], dropna=False)
                .size()
                .reset_index(name="订单数")
            )
            top_products = (
                prod_src_counts.groupby("产品")["订单数"].sum()
                .sort_values(ascending=False).head(top_n).index.tolist()
            )
            prod_src_top = prod_src_counts[prod_src_counts["产品"].isin(top_products)].copy()
            # 注意：传给 Altair 的所有列保持纯 pandas dtypes（object(str) / int）。
            # 1) 严禁使用 pd.Categorical：其底层 codes 为 int，会触发 pyarrow 的
            #    "Expected bytes, got a 'int' object" ArrowTypeError。
            # 2) 严禁在 alt.X(sort=...) 里直接塞 Python list：Altair V6 的 sort 只接受
            #    bool / float / Mapping（如 EncodingSortField），否则会抛 SchemaValidationError：
            #    "'番茄免费小说' is an invalid value for `0`"。
            # 替代方案：新增一个纯 int 的排名字段，用 EncodingSortField 引用它。
            prod_src_top["产品"] = prod_src_top["产品"].astype(str)
            prod_src_top["预算源"] = prod_src_top["预算源"].astype(str)
            _prod_rank = {str(p): i for i, p in enumerate(top_products)}
            prod_src_top["_产品排序"] = (
                prod_src_top["产品"].map(_prod_rank).fillna(9999).astype(int)
            )

            _src_domains = sorted(prod_src_top["预算源"].dropna().astype(str).unique().tolist())
            # 低饱和度商务灰莫兰迪调色板（色相差异足够 + 整体柔和不刺眼）
            _src_colors = ["#6C8EAD", "#88A880", "#C68867", "#B5A26A", "#8F94A3", "#9F8AAB", "#B57279", "#6FA29A",
                           "#C48BA1", "#A89060", "#7FB3A2", "#B29EB8"]
            color_scale = alt.Scale(
                domain=_src_domains,
                range=_src_colors[: len(_src_domains)] if len(_src_domains) <= len(_src_colors) else None,
            )

            _total_max = int(prod_src_top.groupby("产品")["订单数"].sum().max()) if not prod_src_top.empty else 0
            base2 = alt.Chart(prod_src_top).encode(
                x=alt.X(
                    "产品:N",
                    title=None,
                    sort=alt.EncodingSortField(field="_产品排序", op="min", order="ascending"),
                    axis=alt.Axis(domain=False, ticks=False, labelAngle=-25, labelLimit=180),
                ),
                y=alt.Y("订单数:Q", title="订单数量",
                        axis=alt.Axis(grid=True, gridColor="#f2f3f5", domain=False, ticks=False),
                        scale=alt.Scale(domain=[0, _total_max * 1.22]) if _total_max > 0 else alt.Undefined),
                color=alt.Color("预算源:N", title="预算源", scale=color_scale, legend=alt.Legend(orient="bottom")),
            )
            bars2 = base2.mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
                tooltip=["产品:N", "预算源:N", "订单数:Q"],
            )
            # 堆叠块数值：订单数 >= 2 才标，避免 1 个单位的块字重叠看不清
            text2 = base2.mark_text(
                align="center", baseline="middle", fontSize=11, color="#ffffff",
                fontWeight=700, dy=-1,
            ).encode(
                text=alt.condition(
                    alt.datum["订单数"] >= 2,
                    alt.Text("订单数:Q", format="d"),
                    alt.value(""),
                ),
                order=alt.Order("预算源:N", sort="ascending"),
            )
            chart2 = (
                (bars2 + text2)
                .properties(height=390)
                .configure_view(stroke=None)
                .configure_axis(
                    labelFont="Microsoft YaHei",
                    titleFont="Microsoft YaHei",
                    labelFontSize=12,
                    titleFontSize=13,
                    labelColor="#4e5969",
                    titleColor="#4e5969",
                )
                .configure_title(font="Microsoft YaHei")
                .configure_legend(titleFont="Microsoft YaHei", labelFont="Microsoft YaHei")
            )

            c1, c2 = st.columns(2, gap="medium")
            with c1:
                st.markdown(f"**按接口文档 · 订单数 Top {len(doc_counts)}（降序，共 {len(chart_df)} 条）**", unsafe_allow_html=True)
                st.altair_chart(chart1, use_container_width=True, theme=None, key="alt_chart1")
            with c2:
                st.markdown(f"**按产品 × 预算源 分布 Top {len(top_products)}**", unsafe_allow_html=True)
                st.altair_chart(chart2, use_container_width=True, theme=None, key="alt_chart2")

    else:
        st.info("没有符合筛选条件的订单")

    # 底部信息
    st.divider()
    st.caption(f"共 {len(df)} 条在投订单 | 筛选后 {len(filtered_df)} 条 | 缓存有效期 5 分钟")

if __name__ == "__main__":
    main()

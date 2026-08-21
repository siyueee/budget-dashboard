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
from datetime import datetime, timedelta
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
           '回传维度', '考核', '考核数值', 'RTA', '考核备注', '其他备注',
           '下载链接', '归属', '是否打满', '是否披露']

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
        return pd.DataFrame()

    data_rows = all_rows[1:]

    # 构建包名匹配表
    pkg_map = {}
    row = 1
    while row <= match_rows:
        end = min(row + batch_size - 1, match_rows)
        batch = read_sheet_range(MATCH_SHEET_ID, row, end, "A", "B")
        for r in batch[1:] if row == 1 else batch:
            if len(r) >= 2 and r[0] and r[1]:
                pkg_map[str(r[0]).strip()] = str(r[1]).strip()
        row = end + 1

    STATUS_COL = 13     # 上下线状态
    DISCLOSE_COL = 23   # 是否披露
    PRODUCT_COL = 7     # 产品
    PACKAGE_COL = 5     # 包名
    ONLINE_TIME_COL = 11
    OFFLINE_TIME_COL = 12

    active_orders = []
    for r in data_rows:
        status = r[STATUS_COL] if len(r) > STATUS_COL else None
        disclose = r[DISCLOSE_COL] if len(r) > DISCLOSE_COL else None
        if status == "在投" and str(disclose).strip() == "1":
            product = str(r[PRODUCT_COL]).strip() if len(r) > PRODUCT_COL and r[PRODUCT_COL] else ""
            pkg_val = r[PACKAGE_COL] if len(r) > PACKAGE_COL else ""
            if isinstance(pkg_val, str) and pkg_val.startswith("VLOOKUP"):
                r[PACKAGE_COL] = pkg_map.get(product, "")
            if len(r) > ONLINE_TIME_COL:
                r[ONLINE_TIME_COL] = excel_serial_to_date(r[ONLINE_TIME_COL])
            if len(r) > OFFLINE_TIME_COL:
                r[OFFLINE_TIME_COL] = excel_serial_to_date(r[OFFLINE_TIME_COL])
            active_orders.append(r)

    # 转为 DataFrame
    rows_data = []
    for order in active_orders:
        row_dict = {}
        for i, header in enumerate(HEADERS):
            row_dict[header] = order[i] if i < len(order) else ""
        rows_data.append(row_dict)

    return pd.DataFrame(rows_data)

def main():
    # 页面配置
    st.set_page_config(
        page_title="订单披露",
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
    /* ===== 自定义订单表格 ===== */
    .order-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 13.5px;
        color: #1d2129;
    }
    .order-table thead th {
        background-color: #f7f8fa;
        font-weight: 600;
        color: #4e5969;
        padding: 10px 14px;
        text-align: left;
        border: none;
        border-bottom: 1px solid #e5e6eb;
        border-top: 1px solid #e5e6eb;
    }
    .order-table tbody td {
        padding: 11px 14px;
        border: none;
        border-bottom: 1px solid #f0f1f3;
        vertical-align: middle;
    }
    .order-table tbody tr:hover td {
        background-color: #fafbfc;
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
            st.markdown("# 订单披露")
    else:
        st.markdown("# 订单披露")
    st.caption("数据来源：飞书电子表格「订单明细」| 仅展示「在投」状态订单")

    # 检查凭证
    if not APP_ID or not APP_SECRET:
        st.error("⚠️ 未配置飞书应用凭证！请在 `.streamlit/secrets.toml` 中设置 FEISHU_APP_ID 和 FEISHU_APP_SECRET")
        st.code('FEISHU_APP_ID = "your_app_id"\nFEISHU_APP_SECRET = "your_app_secret"', language="toml")
        return

    # 读取数据
    with st.spinner("正在从飞书读取最新数据..."):
        df = fetch_active_orders()

    if df.empty:
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

    # 统计卡片（自定义彩色圆角卡片）
    products_count = df['产品'].dropna().nunique() if '产品' in df.columns else 0
    advertisers_count = df['广告主'].dropna().nunique() if '广告主' in df.columns else 0
    update_time = datetime.now().strftime("%Y-%m-%d %H:%M")

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
            <div class="stat-label">更新时间</div>
            <div class="stat-value-gray" style="font-size:22px;">{update_time}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 筛选区（9 个筛选项，同一行显示）
    st.markdown('<div class="filter-container">', unsafe_allow_html=True)

    col_a, col_b, col_c, col_d, col_e, col_f, col_g, col_h, col_i = st.columns(9)
    with col_a:
        config_options = ['全部'] + sorted([str(x) for x in df['配置号'].dropna().unique()]) if '配置号' in df.columns else ['全部']
        config_filter = st.selectbox("配置号", config_options)
    with col_b:
        api_doc_options = ['全部'] + sorted([str(x) for x in df['接口文档'].dropna().unique()]) if '接口文档' in df.columns else ['全部']
        api_doc_filter = st.selectbox("接口文档", api_doc_options)
    with col_c:
        product_options = ['全部'] + sorted([str(x) for x in df['产品'].dropna().unique()]) if '产品' in df.columns else ['全部']
        product_filter = st.selectbox("产品", product_options)
    with col_d:
        callback_options = ['全部'] + sorted([str(x) for x in df['回传维度'].dropna().unique()]) if '回传维度' in df.columns else ['全部']
        callback_filter = st.selectbox("回传维度", callback_options)
    with col_e:
        attrib_options = ['全部'] + sorted([str(x) for x in df['归属'].dropna().unique()]) if '归属' in df.columns else ['全部']
        attribution_filter = st.selectbox("归属", attrib_options)
    with col_f:
        rta_options = ['全部'] + sorted([str(x) for x in df['RTA'].dropna().unique()]) if 'RTA' in df.columns else ['全部']
        rta_filter = st.selectbox("RTA", rta_options)
    with col_g:
        platform_options = ['全部'] + sorted([str(x) for x in df['分端'].dropna().unique()]) if '分端' in df.columns else ['全部']
        platform_filter = st.selectbox("分端", platform_options)
    with col_h:
        source_options = ['全部'] + sorted([str(x) for x in df['预算源'].dropna().unique()]) if '预算源' in df.columns else ['全部']
        source_filter = st.selectbox("预算源", source_options)
    with col_i:
        task_type_options = ['全部'] + sorted([str(x) for x in df['任务类型'].dropna().unique()]) if '任务类型' in df.columns else ['全部']
        task_type_filter = st.selectbox("任务类型", task_type_options)

    # 应用筛选
    filtered_df = df.copy()
    if config_filter != '全部':
        filtered_df = filtered_df[filtered_df['配置号'].astype(str) == config_filter]
    if product_filter != '全部':
        filtered_df = filtered_df[filtered_df['产品'].astype(str) == product_filter]
    if api_doc_filter != '全部':
        filtered_df = filtered_df[filtered_df['接口文档'].astype(str) == api_doc_filter]
    if attribution_filter != '全部':
        filtered_df = filtered_df[filtered_df['归属'].astype(str) == attribution_filter]
    if rta_filter != '全部':
        filtered_df = filtered_df[filtered_df['RTA'].astype(str) == rta_filter]
    if callback_filter != '全部':
        filtered_df = filtered_df[filtered_df['回传维度'].astype(str) == callback_filter]
    if platform_filter != '全部':
        filtered_df = filtered_df[filtered_df['分端'].astype(str) == platform_filter]
    if source_filter != '全部':
        filtered_df = filtered_df[filtered_df['预算源'].astype(str) == source_filter]
    if task_type_filter != '全部':
        filtered_df = filtered_df[filtered_df['任务类型'].astype(str) == task_type_filter]

    st.markdown('</div>', unsafe_allow_html=True)

    # 数据表格
    # 在投订单明细标题（带动图）
    table_b64 = load_gif_base64("table_icon.gif")
    if table_b64:
        col_img2, col_txt2 = st.columns([1, 25])
        with col_img2:
            components.html(
                f'<img src="data:image/gif;base64,{table_b64}" style="width:35px;height:auto;">',
                height=45
            )
        with col_txt2:
            st.markdown(f"### 在投订单明细（{len(filtered_df)} 条）")
    else:
        st.markdown(f"### 在投订单明细（{len(filtered_df)} 条）")

    # 选择要显示的列（移除不需要展示的列）
    display_cols = [c for c in HEADERS if c not in ['上线时间', '下线时间', '考核', '考核数值', '广告主', '渠道号', '下载链接', '是否打满', '是否披露']]
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
        note_cols = {"考核备注", "其他备注"}
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
        components.html(
            f"""
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
    else:
        st.info("没有符合筛选条件的订单")

    # 底部信息
    st.divider()
    st.caption(f"共 {len(df)} 条在投订单 | 筛选后 {len(filtered_df)} 条 | 缓存有效期 5 分钟")

if __name__ == "__main__":
    main()

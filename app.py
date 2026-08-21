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

@st.cache_data(ttl=300, show_spinner=False)
def fetch_active_orders():
    """从飞书电子表格读取在投订单（5分钟缓存）"""
    all_rows = []
    row = 1
    batch_size = 100
    while row <= TOTAL_ROWS:
        end = min(row + batch_size - 1, TOTAL_ROWS)
        batch = read_sheet_range(SHEET_ID, row, end)
        all_rows.extend(batch)
        row = end + 1

    if not all_rows:
        return pd.DataFrame()

    data_rows = all_rows[1:]

    # 构建包名匹配表
    pkg_map = {}
    row = 1
    while row <= MATCH_TOTAL_ROWS:
        end = min(row + batch_size - 1, MATCH_TOTAL_ROWS)
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

    # 自定义 CSS：去除竖线、美化筛选区
    st.markdown("""
    <style>
    /* 筛选区域背景色 */
    .filter-container {
        background-color: #f0f2f6;
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
    /* 指标卡片下方分隔线 */
    hr {
        border: none;
        border-top: 1px solid #e0e0e0;
        margin: 1rem 0;
    }
    /* 隐藏 Streamlit 默认 divider */
    .st-emotion-cache-1aj19jr {
        display: none;
    }
    </style>
    """, unsafe_allow_html=True)

    # 标题（带小熊图标）
    icon_path = os.path.join(os.path.dirname(__file__), "assets", "bear.gif")
    icon_img = ""
    if os.path.exists(icon_path):
        with open(icon_path, "rb") as f:
            icon_b64 = base64.b64encode(f.read()).decode()
        icon_img = f'<img src="data:image/gif;base64,{icon_b64}" width="50" style="vertical-align:middle;margin-right:10px;border-radius:8px;">'
    st.markdown(f"# {icon_img}订单披露", unsafe_allow_html=True)
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

    # 统计卡片
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("在投订单数", len(df))
    with col2:
        products = df['产品'].dropna().nunique() if '产品' in df.columns else 0
        st.metric("产品数", products)
    with col3:
        advertisers = df['广告主'].dropna().nunique() if '广告主' in df.columns else 0
        st.metric("广告主数", advertisers)
    with col4:
        st.metric("数据更新时间", datetime.now().strftime("%H:%M:%S"))

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
    st.subheader(f"📋 在投订单明细（{len(filtered_df)} 条）")

    # 选择要显示的列（移除不需要展示的列）
    display_cols = [c for c in HEADERS if c not in ['上线时间', '下线时间', '考核', '考核数值', '广告主', '渠道号', '下载链接', '是否打满', '是否披露']]
    display_df = filtered_df[display_cols] if not filtered_df.empty else filtered_df

    if not display_df.empty:
        # 配置列宽和样式，表格加高，支持点击行选中
        event = st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            height=800,
            selection_mode="single-row",
            column_config={
                "考核备注": st.column_config.TextColumn(width="large"),
                "其他备注": st.column_config.TextColumn(width="large"),
                "包名": st.column_config.TextColumn(width="medium"),
                "配置号": st.column_config.TextColumn(width="medium"),
            }
        )

        # 点击行显示完整备注
        selected = event.selected_rows
        if selected is not None:
            if hasattr(selected, 'empty'):
                if not selected.empty:
                    st.markdown("**📌 行详情**")
                    row = selected.iloc[0]
                    detail_cols = st.columns(2)
                    with detail_cols[0]:
                        for col in display_cols[:len(display_cols)//2]:
                            val = row.get(col, "")
                            if val:
                                st.markdown(f"**{col}**: {val}")
                    with detail_cols[1]:
                        for col in display_cols[len(display_cols)//2:]:
                            val = row.get(col, "")
                            if val:
                                st.markdown(f"**{col}**: {val}")
            elif hasattr(selected, '__len__'):
                if len(selected) > 0:
                    st.markdown("**📌 行详情**")
                    sel_df = display_df.iloc[list(selected)] if len(selected) > 0 else None
                    if sel_df is not None and not sel_df.empty:
                        row = sel_df.iloc[0]
                        detail_cols = st.columns(2)
                        with detail_cols[0]:
                            for col in display_cols[:len(display_cols)//2]:
                                val = row.get(col, "")
                                if val:
                                    st.markdown(f"**{col}**: {val}")
                        with detail_cols[1]:
                            for col in display_cols[len(display_cols)//2:]:
                                val = row.get(col, "")
                                if val:
                                    st.markdown(f"**{col}**: {val}")

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

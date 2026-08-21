"""
订单披露 - Streamlit 应用
通过飞书应用凭证实时读取电子表格数据，展示在投订单
"""
import streamlit as st
import requests
import json
import time
from datetime import datetime, timedelta
import pandas as pd

# ===== 飞书应用凭证（从环境变量或 Streamlit secrets 读取）=====
import os
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
HEADERS = ['预算源', '任务类型', '配置号', '广告主', '包名', '接口文档', '产品',
           '渠道号', '合作价格', '需求量级', '上线时间', '下线时间', '上下线状态',
           '回传维度', '考核', '考核数值', 'RTA', '考核备注', '其他备注', '下载链接']

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

def read_sheet_range(sheet_id, start_row, end_row, col_start="A", col_end="T"):
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

    STATUS_COL = 12     # 上下线状态
    PRODUCT_COL = 6     # 产品
    PACKAGE_COL = 4     # 包名
    ONLINE_TIME_COL = 10
    OFFLINE_TIME_COL = 11

    active_orders = []
    for r in data_rows:
        status = r[STATUS_COL] if len(r) > STATUS_COL else None
        if status == "在投":
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

    # 标题
    st.title("📊 订单披露")
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
                test_data = feishu_get(f"/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/values/{SHEET_ID}!A1:T2",
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

    st.divider()

    # 筛选区
    st.subheader("🔍 筛选")
    col_a, col_b, col_c, col_d = st.columns(4)

    with col_a:
        config_filter = st.text_input("配置号", placeholder="输入配置号搜索...")
    with col_b:
        product_options = ['全部'] + sorted([str(x) for x in df['产品'].dropna().unique()]) if '产品' in df.columns else ['全部']
        product_filter = st.selectbox("产品", product_options)
    with col_c:
        api_doc_options = ['全部'] + sorted([str(x) for x in df['接口文档'].dropna().unique()]) if '接口文档' in df.columns else ['全部']
        api_doc_filter = st.selectbox("接口文档", api_doc_options)
    with col_d:
        source_options = ['全部'] + sorted([str(x) for x in df['预算源'].dropna().unique()]) if '预算源' in df.columns else ['全部']
        attribution_filter = st.selectbox("预算源", source_options)

    col_e, col_f, col_g = st.columns(3)
    with col_e:
        rta_options = ['全部'] + sorted([str(x) for x in df['RTA'].dropna().unique()]) if 'RTA' in df.columns else ['全部']
        rta_filter = st.selectbox("RTA", rta_options)
    with col_f:
        callback_options = ['全部'] + sorted([str(x) for x in df['回传维度'].dropna().unique()]) if '回传维度' in df.columns else ['全部']
        callback_filter = st.selectbox("回传维度", callback_options)
    with col_g:
        config_no_options = ['全部'] + sorted([str(x) for x in df['配置号'].dropna().unique()]) if '配置号' in df.columns else ['全部']
        config_no_filter = st.selectbox("配置号筛选", config_no_options)

    # 应用筛选
    filtered_df = df.copy()
    if config_filter:
        filtered_df = filtered_df[filtered_df['配置号'].astype(str).str.contains(config_filter, case=False, na=False)]
    if product_filter != '全部':
        filtered_df = filtered_df[filtered_df['产品'].astype(str) == product_filter]
    if api_doc_filter != '全部':
        filtered_df = filtered_df[filtered_df['接口文档'].astype(str) == api_doc_filter]
    if attribution_filter != '全部':
        filtered_df = filtered_df[filtered_df['预算源'].astype(str) == attribution_filter]
    if rta_filter != '全部':
        filtered_df = filtered_df[filtered_df['RTA'].astype(str) == rta_filter]
    if callback_filter != '全部':
        filtered_df = filtered_df[filtered_df['回传维度'].astype(str) == callback_filter]
    if config_no_filter != '全部':
        filtered_df = filtered_df[filtered_df['配置号'].astype(str) == config_no_filter]

    st.divider()

    # 数据表格
    st.subheader(f"📋 在投订单明细（{len(filtered_df)} 条）")

    # 选择要显示的列（移除不需要展示的列）
    display_cols = [c for c in HEADERS if c not in ['上线时间', '下线时间', '预算源', '任务类型']]
    display_df = filtered_df[display_cols] if not filtered_df.empty else filtered_df

    if not display_df.empty:
        # 配置列宽和样式
        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "考核备注": st.column_config.TextColumn(width="medium"),
                "其他备注": st.column_config.TextColumn(width="medium"),
            }
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

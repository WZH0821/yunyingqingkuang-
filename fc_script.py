import streamlit as st
import pandas as pd
import plotly.express as px
from typing import Dict, List, Optional, Tuple
import requests
from io import BytesIO

st.set_page_config(page_title="Dashboard", layout="wide")

# ============================================================
# 配置 - 你的GitHub信息
# ============================================================
GITHUB_USERNAME = "WZH0821"
GITHUB_REPO = "yunyingqingkuang-"
GITHUB_BRANCH = "main"
EXCEL_FILENAME_DATA1 = "data1.xlsx"
EXCEL_FILENAME_DATA2 = "data2.xlsx"

GITHUB_FILE_URL_DATA1 = f"https://raw.githubusercontent.com/{GITHUB_USERNAME}/{GITHUB_REPO}/{GITHUB_BRANCH}/{EXCEL_FILENAME_DATA1}"
GITHUB_FILE_URL_DATA2 = f"https://raw.githubusercontent.com/{GITHUB_USERNAME}/{GITHUB_REPO}/{GITHUB_BRANCH}/{EXCEL_FILENAME_DATA2}"

# ============================================================
# 配置常量
# ============================================================
METRIC_CONFIG = {
    '成交量': {
        'market_divide': 100000000, 'market_unit': '亿手', 'market_title': '（亿手）', 'market_yaxis': '成交量（亿手）',
        'company_divide': 10000, 'company_unit': '万手', 'company_title': '（万手）', 'company_yaxis': '成交量（万手）',
    },
    '成交额': {
        'market_divide': 10000, 'market_unit': '万亿元', 'market_title': '（万亿元）', 'market_yaxis': '成交额（万亿元）',
        'company_divide': 100000000, 'company_unit': '亿元', 'company_title': '（亿元）', 'company_yaxis': '成交额（亿元）',
    },
    '持仓量': {
        'market_divide': 1000000, 'market_unit': '百万手', 'market_title': '（百万手）', 'market_yaxis': '持仓量（百万手）',
        'company_divide': 10000, 'company_unit': '万手', 'company_title': '（万手）', 'company_yaxis': '持仓量（万手）',
    }
}

COLOR_MAP = {
    '本月': '#2E86C1', '上月': '#F39C12', '去年同期': '#28B463',
    '本季度': '#2E86C1', '上季度': '#F39C12',
    '今年': '#2E86C1', '去年': '#F39C12',
}

MONTH_NAMES = {1: '1月', 2: '2月', 3: '3月', 4: '4月', 5: '5月', 6: '6月',
               7: '7月', 8: '8月', 9: '9月', 10: '10月', 11: '11月', 12: '12月'}

# ============================================================
# GitHub数据加载函数
# ============================================================
@st.cache_data(ttl=3600)
def load_excel_from_github(url: str) -> Dict[str, pd.DataFrame]:
    """
    从GitHub URL加载Excel文件的所有sheet
    """
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        
        excel_data = BytesIO(response.content)
        excel_file = pd.ExcelFile(excel_data)
        
        sheets_dict = {}
        for sheet_name in excel_file.sheet_names:
            try:
                df = pd.read_excel(excel_file, sheet_name=sheet_name, header=0)
                sheets_dict[sheet_name] = clean_dataframe(df)
            except Exception as e:
                st.warning(f"读取sheet '{sheet_name}' 时出错: {e}")
                sheets_dict[sheet_name] = pd.DataFrame()
        
        return sheets_dict
    except requests.exceptions.RequestException as e:
        st.error(f"从GitHub下载文件失败: {e}")
        st.info(f"请检查URL是否正确: {url}")
        return {}
    except Exception as e:
        st.error(f"读取Excel文件失败: {e}")
        return {}

@st.cache_data(ttl=3600)
def load_data_from_github() -> Tuple[Dict[str, pd.DataFrame], Dict[str, pd.DataFrame]]:
    """
    从GitHub加载Data1和Data2
    """
    st.info("🔄 正在从GitHub加载数据...")
    
    data1_cache = {}
    data2_cache = {}
    
    # 加载Data1
    st.text(f"📁 加载 Data1: {GITHUB_FILE_URL_DATA1}")
    data1_raw = load_excel_from_github(GITHUB_FILE_URL_DATA1)
    
    # 映射Data1的sheet
    data1_sheets = {
        '成交量-市场': '成交量-市场',
        '成交量-公司': '成交量-公司',
        '成交额-市场': '成交额-市场',
        '成交额-公司': '成交额-公司',
        '持仓量-市场': '持仓量-市场',
        '持仓量-公司': '持仓量-公司',
        '资金对账表-月': '资金对账表-月',
    }
    for key, sheet in data1_sheets.items():
        data1_cache[key] = data1_raw.get(sheet, pd.DataFrame())
    
    # 加载Data2
    st.text(f"📁 加载 Data2: {GITHUB_FILE_URL_DATA2}")
    data2_raw = load_excel_from_github(GITHUB_FILE_URL_DATA2)
    
    # 映射Data2的sheet
    data2_sheets = {
        '上一年资金对账表-月': '上一年资金对账表-月',
        '交易统计表-月': '交易统计表-月',
        '上一年交易统计表-月': '上一年交易统计表-月',
        '投资者资料查询': '投资者资料查询',
        '活跃客户': '活跃客户',
        '市场权益': '市场权益'
    }
    for key, sheet in data2_sheets.items():
        data2_cache[key] = data2_raw.get(sheet, pd.DataFrame())
    
    st.success("✅ 数据加载完成！")
    return data1_cache, data2_cache

# ============================================================
# 工具函数
# ============================================================
def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    df = df.loc[:, ~df.columns.isna()]
    df = df.loc[:, df.columns != '']
    df = df.loc[:, ~df.columns.duplicated()]
    return df.dropna(axis=1, how='all')

def parse_month_column(col) -> Tuple[Optional[int], Optional[int]]:
    col_str = str(col)
    if len(col_str) == 6 and col_str.isdigit():
        year, month = int(col_str[:4]), int(col_str[4:6])
        if 1 <= month <= 12:
            return year, month
    return None, None

def get_month_columns(df: pd.DataFrame) -> List[str]:
    if df.empty:
        return []
    return [col for col in df.columns if parse_month_column(col)[0] is not None]

def safe_division(a, b, default=0):
    if b is None or b == 0:
        return default
    return a / b

def format_percent(value, decimals=2):
    if value is None or pd.isna(value):
        return '-'
    return f"{value:+.{decimals}f}%"

def get_metric_config(data_type: str) -> dict:
    return METRIC_CONFIG.get(data_type, METRIC_CONFIG['成交量'])

def safe_get_column(df: pd.DataFrame, col_names: List[str], default_idx: int = None) -> Optional[str]:
    if df.empty:
        return None
    for name in col_names:
        if name in df.columns:
            return name
    if default_idx is not None and len(df.columns) > default_idx:
        return df.columns[default_idx]
    return None

def normalize_trade_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    
    col_mapping = {}
    for col in df.columns:
        col_str = str(col).strip()
        if '月份' in col_str:
            col_mapping[col] = '月份'
        elif '部门' in col_str:
            col_mapping[col] = '部门'
        elif '投资者' in col_str or '客户代码' in col_str:
            col_mapping[col] = '投资者代码'
        elif '平仓盈亏' in col_str:
            col_mapping[col] = '平仓盈亏'
        elif '权利金收入' in col_str or '期权权利金收入' in col_str:
            col_mapping[col] = '期权权利金收入'
        elif '权利金支出' in col_str or '期权权利金支出' in col_str:
            col_mapping[col] = '期权权利金支出'
    
    if col_mapping:
        df = df.rename(columns=col_mapping)
    return df

def compute_period_comparison(df: pd.DataFrame, selected_cols: list, 
                              prev_cols: list, last_year_cols: list,
                              divide: float) -> Dict:
    result = {
        'current': df[selected_cols].sum().sum() / divide if selected_cols else 0,
        'prev': df[prev_cols].sum().sum() / divide if prev_cols else None,
        'last_year': df[last_year_cols].sum().sum() / divide if last_year_cols else None
    }
    result['mom'] = safe_division(result['current'] - result['prev'], result['prev']) * 100 if result['prev'] is not None else None
    result['yoy'] = safe_division(result['current'] - result['last_year'], result['last_year']) * 100 if result['last_year'] is not None else None
    return result

def build_comparison_table(df: pd.DataFrame, group_col: str, groups: list,
                           selected_cols: list, prev_cols: list, last_year_cols: list,
                           divide: float, current_label: str, prev_label: str, last_year_label: str) -> pd.DataFrame:
    rows = []
    for group in groups:
        group_df = df[df[group_col] == group]
        row = {group_col: group}
        current_val = group_df[selected_cols].sum().sum() / divide if selected_cols else 0
        row[current_label] = current_val
        
        if prev_cols:
            prev_val = group_df[prev_cols].sum().sum() / divide if prev_cols else None
            row[prev_label] = prev_val if prev_val is not None else None
            row['环比'] = safe_division(current_val - prev_val, prev_val) * 100 if prev_val else None
        else:
            row[prev_label] = None
            row['环比'] = None
        
        if last_year_cols and last_year_label:
            last_val = group_df[last_year_cols].sum().sum() / divide if last_year_cols else None
            row[last_year_label] = last_val if last_val is not None else None
            row['同比'] = safe_division(current_val - last_val, last_val) * 100 if last_val else None
        else:
            row[last_year_label] = None if last_year_label else None
            row['同比'] = None
        
        rows.append(row)
    return pd.DataFrame(rows)

def create_bar_chart(df: pd.DataFrame, x_col: str, y_col: str, color_col: str,
                     title: str, yaxis_title: str, color_discrete_map: dict):
    if df.empty:
        return None
    fig = px.bar(df, x=x_col, y=y_col, color=color_col, barmode='group',
                 title=title, labels={y_col: yaxis_title, x_col: x_col},
                 text_auto='.2f', color_discrete_map=color_discrete_map)
    fig.update_layout(
        title_font=dict(size=18, color='#1A5276'), font=dict(size=13),
        bargap=0.25, bargroupgap=0.15,
        plot_bgcolor='#F8F9F9', paper_bgcolor='white',
        legend_title_text='', yaxis=dict(tickformat='.2f', title=yaxis_title)
    )
    fig.update_traces(texttemplate='%{y:.2f}', textfont=dict(size=11, color='black', family='Arial Black'),
                      textposition='outside')
    return fig

def create_line_chart(df: pd.DataFrame, x: str, y: str, color: str,
                      title: str, xlabel: str = '', ylabel: str = '',
                      color_map: dict = None, text_format: str = '.2f'):
    if df.empty:
        return None
    fig = px.line(df, x=x, y=y, color=color, title=title,
                  labels={x: xlabel, y: ylabel},
                  markers=True, color_discrete_map=color_map)
    fig.update_layout(
        title_font=dict(size=14, color='#1A5276'), font=dict(size=11),
        plot_bgcolor='#F8F9F9', paper_bgcolor='white',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
        height=350
    )
    fig.update_traces(texttemplate=f'%{{y:{text_format}}}', textposition='top center',
                      textfont=dict(size=8), mode='lines+markers+text')
    return fig

# ============================================================
# 侧边栏 - 显示数据来源信息
# ============================================================
with st.sidebar:
    st.header("📊 数据看板")
    
    st.subheader("📁 数据来源")
    st.info(f"""
    **Data1 (市场数据 + 资金对账表)**
    - 仓库: {GITHUB_REPO}
    - 文件: {EXCEL_FILENAME_DATA1}
    
    **Data2 (业务数据)**
    - 仓库: {GITHUB_REPO}
    - 文件: {EXCEL_FILENAME_DATA2}
    """)
    
    st.divider()
    
    # 显示数据加载状态
    st.subheader("📌 数据状态")
    try:
        if 'data1_cache' in st.session_state and 'data2_cache' in st.session_state:
            data1_count = sum(1 for df in st.session_state.data1_cache.values() if not df.empty)
            data2_count = sum(1 for df in st.session_state.data2_cache.values() if not df.empty)
            st.success(f"✅ Data1: {data1_count}/7 个sheet已加载")
            st.success(f"✅ Data2: {data2_count}/6 个sheet已加载")
        else:
            st.warning("⏳ 数据加载中...")
    except:
        st.warning("⏳ 数据加载中...")
    
    st.divider()
    
    # 刷新按钮
    if st.button("🔄 刷新数据", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    
    st.caption(f"最后更新: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")

# ============================================================
# 主逻辑
# ============================================================
st.title("📊 交易数据看板")

# 加载数据
if 'data1_cache' not in st.session_state or 'data2_cache' not in st.session_state:
    try:
        data1_cache, data2_cache = load_data_from_github()
        # 存储到 session_state
        st.session_state.data1_cache = data1_cache
        st.session_state.data2_cache = data2_cache
    except Exception as e:
        st.error(f"❌ 加载数据失败: {e}")
        st.stop()
else:
    # 从 session_state 获取数据
    data1_cache = st.session_state.data1_cache
    data2_cache = st.session_state.data2_cache

# ============================================================
# 从缓存中提取变量
# ============================================================
# 从 Data1 提取
df_vol_market = data1_cache.get('成交量-市场', pd.DataFrame())
df_vol_company = data1_cache.get('成交量-公司', pd.DataFrame())
df_amt_market = data1_cache.get('成交额-市场', pd.DataFrame())
df_amt_company = data1_cache.get('成交额-公司', pd.DataFrame())
df_oi_market = data1_cache.get('持仓量-市场', pd.DataFrame())
df_oi_company = data1_cache.get('持仓量-公司', pd.DataFrame())
df_fund_current = data1_cache.get('资金对账表-月', pd.DataFrame())

# 从 Data2 提取
df_fund_last_year = data2_cache.get('上一年资金对账表-月', pd.DataFrame())
df_trade_stats = data2_cache.get('交易统计表-月', pd.DataFrame())
df_trade_last = data2_cache.get('上一年交易统计表-月', pd.DataFrame())
df_investor = data2_cache.get('投资者资料查询', pd.DataFrame())
df_active = data2_cache.get('活跃客户', pd.DataFrame())
df_market_equity = data2_cache.get('市场权益', pd.DataFrame())

# 标准化数据
df_trade_stats = normalize_trade_columns(df_trade_stats)
df_trade_last = normalize_trade_columns(df_trade_last)

# 合并所有可用数据用于数据筛选器
all_data = {**data1_cache, **data2_cache}
available_sheets = {k: v for k, v in all_data.items() if not v.empty}
if not available_sheets:
    st.error("❌ 没有可用的数据表")
    st.stop()

# ============================================================
# 数据筛选
# ============================================================
st.subheader("📋 数据筛选")
col_sheet, col_month_start, col_month_end = st.columns([2, 1.5, 1.5])

with col_sheet:
    selected_sheet = st.selectbox("选择数据表", options=list(available_sheets.keys()), key="sheet_selector")

df = available_sheets[selected_sheet].copy()
month_cols = get_month_columns(df)

month_filter_applied = False
selected_months = month_cols

if month_cols:
    month_cols_sorted = sorted(month_cols)
    month_labels = {col: f"{str(col)[:4]}年{str(col)[4:6]}月" for col in month_cols_sorted}

    with col_month_start:
        selected_month_start = st.selectbox("开始月份", options=month_cols_sorted,
                                            format_func=lambda x: month_labels.get(x, str(x)),
                                            index=0, key="month_start")
    with col_month_end:
        selected_month_end = st.selectbox("结束月份", options=month_cols_sorted,
                                          format_func=lambda x: month_labels.get(x, str(x)),
                                          index=len(month_cols_sorted) - 1, key="month_end")

    if selected_month_start and selected_month_end:
        start_idx = month_cols_sorted.index(selected_month_start)
        end_idx = month_cols_sorted.index(selected_month_end)
        if start_idx <= end_idx:
            selected_months = month_cols_sorted[start_idx:end_idx + 1]
            month_filter_applied = True
        else:
            st.warning("⚠️ 开始月份不能晚于结束月份")

numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
filter_cols = [col for col in df.columns if col not in numeric_cols]
selected_filters = {}

if filter_cols:
    st.markdown("**文本筛选条件**")
    cols = st.columns(3)
    for idx, col_name in enumerate(filter_cols):
        with cols[idx % 3]:
            unique_vals = df[col_name].dropna().unique().tolist()
            if unique_vals:
                selected = st.multiselect(f"{col_name}", options=unique_vals, default=[], key=f"filter_{col_name}")
                if selected:
                    selected_filters[col_name] = selected

filtered_df = df.copy()
for col, vals in selected_filters.items():
    filtered_df = filtered_df[filtered_df[col].isin(vals)]

if month_filter_applied and selected_months:
    non_month_cols = [col for col in filtered_df.columns if col not in month_cols]
    filtered_df = filtered_df[non_month_cols + selected_months]

st.success(f"✅ 当前查看: {selected_sheet}，共 {len(filtered_df)} 行，{len(filtered_df.columns)} 列")
df_display = filtered_df.reset_index(drop=True)
df_display.insert(0, '序号', range(1, len(df_display) + 1))

numeric_cols_display = df_display.select_dtypes(include=['number']).columns.tolist()
sum_row = {col: '' for col in df_display.columns}
sum_row['序号'] = ''
for col in numeric_cols_display:
    if col != '序号':
        sum_row[col] = df_display[col].sum()
sum_row[df_display.columns[1]] = '【总和】'
df_with_sum = pd.concat([df_display, pd.DataFrame([sum_row])], ignore_index=True)
st.dataframe(df_with_sum, use_container_width=True, height=400, hide_index=True)

# ============================================================
# 各交易所柱状图
# ============================================================
st.subheader("📊 各交易所情况（市场）")

df_detail_map = {'成交量': df_vol_market, '成交额': df_amt_market, '持仓量': df_oi_market}
df_company_map = {'成交量': df_vol_company, '成交额': df_amt_company, '持仓量': df_oi_company}

col_filter1, col_filter2 = st.columns(2)
with col_filter1:
    data_type = st.selectbox("选择数据类型", options=['成交量', '成交额', '持仓量'], key="data_type")

df_detail = df_detail_map.get(data_type, pd.DataFrame())
date_cols = get_month_columns(df_detail)

with col_filter2:
    if date_cols:
        date_cols_sorted = sorted(date_cols)
        selected_key = st.selectbox(
            "选择月份", 
            options=date_cols_sorted,
            format_func=lambda x: f"{str(x)[:4]}年{str(x)[4:6]}月",
            index=len(date_cols_sorted) - 1, 
            key="main_month_selector"
        )
    else:
        selected_key = None
        st.warning("⚠️ 未找到月份列")

metric_config = get_metric_config(data_type)
df_detail = df_detail_map.get(data_type, pd.DataFrame())
df_company_detail = df_company_map.get(data_type, pd.DataFrame())

if df_detail.empty or df_company_detail.empty:
    st.warning(f"⚠️ {data_type}数据为空")
elif selected_key is None:
    st.warning("⚠️ 请选择有效的月份")
else:
    st.success(f"✅ {data_type}数据加载成功！市场 {len(df_detail)} 行，公司 {len(df_company_detail)} 行")
    selected_label = f"{str(selected_key)[:4]}年{str(selected_key)[4:6]}月"

    year, month = parse_month_column(selected_key)
    prev_cols, last_year_cols = [], []
    if year and month:
        prev_key = f"{year if month > 1 else year - 1}{month - 1 if month > 1 else 12:02d}"
        prev_key = int(prev_key) if prev_key.isdigit() else prev_key
        prev_cols = [prev_key] if prev_key in df_detail.columns else []
        last_key = f"{year - 1}{month:02d}"
        last_key = int(last_key) if last_key.isdigit() else last_key
        last_year_cols = [last_key] if last_key in df_detail.columns else []

    selected_cols = [selected_key]
    exchanges = df_detail['交易所'].unique().tolist()
    
    exchange_df = build_comparison_table(df_detail, '交易所', exchanges, selected_cols, prev_cols, last_year_cols,
                                         metric_config['market_divide'], '本月', '上月', '去年同期')
    exchange_company_df = build_comparison_table(df_company_detail, '交易所', exchanges, selected_cols, prev_cols, last_year_cols,
                                                 metric_config['company_divide'], '本月', '上月', '去年同期')

    market_total = compute_period_comparison(df_detail, selected_cols, prev_cols, last_year_cols, metric_config['market_divide'])
    company_total = compute_period_comparison(df_company_detail, selected_cols, prev_cols, last_year_cols, metric_config['company_divide'])

    value_cols = ['本月']
    if prev_cols:
        value_cols.append('上月')
    if last_year_cols:
        value_cols.append('去年同期')

    for df_plot, suffix, divide, yaxis in [
        (exchange_df, '市场', metric_config['market_divide'], metric_config['market_yaxis']),
        (exchange_company_df, '公司', metric_config['company_divide'], metric_config['company_yaxis'])
    ]:
        melted = df_plot.melt(id_vars=['交易所'], value_vars=value_cols, var_name='期间', value_name=data_type)
        melted = melted.dropna(subset=[data_type])
        period_order = [p for p in ['上月', '本月', '去年同期'] if p in melted['期间'].unique()]
        melted['期间'] = pd.Categorical(melted['期间'], categories=period_order, ordered=True)
        melted = melted.sort_values('期间')

        if not melted.empty:
            fig = create_bar_chart(melted, '交易所', data_type, '期间',
                                   f'各交易所{data_type}对比（{suffix}）- {selected_label}',
                                   yaxis, COLOR_MAP)
            if fig:
                st.plotly_chart(fig, use_container_width=True)

            st.subheader(f"📊 交易所环比同比综合表（{suffix}）")
            total = market_total if suffix == '市场' else company_total
            unit_key = 'market_unit' if suffix == '市场' else 'company_unit'
            
            table_data = []
            table_data.append({
                '维度': '合计',
                f'{data_type}': total['current'],
                '环比': total['mom'],
                '同比': total['yoy']
            })
            for _, row in df_plot.iterrows():
                table_data.append({
                    '维度': row['交易所'],
                    f'{data_type}': row['本月'],
                    '环比': row.get('环比'),
                    '同比': row.get('同比')
                })
            
            df_table = pd.DataFrame(table_data)
            
            st.dataframe(
                df_table,
                use_container_width=True,
                hide_index=True,
                column_config={
                    '维度': st.column_config.TextColumn('维度'),
                    f'{data_type}': st.column_config.NumberColumn(
                        f'{data_type}（{metric_config[unit_key]}）',
                        format="%.2f"
                    ),
                    '环比': st.column_config.NumberColumn(
                        '环比（%）',
                        format="%+.2f%%"
                    ),
                    '同比': st.column_config.NumberColumn(
                        '同比（%）',
                        format="%+.2f%%"
                    )
                }
            )

# ============================================================
# 公司占市场比重（整体）
# ============================================================
st.subheader("📊 公司占市场比重（整体）")
try:
    date_cols_all = get_month_columns(df_vol_market)
    filtered_cols = [c for c in date_cols_all if parse_month_column(c)[0] and parse_month_column(c)[0] >= 2024]

    data_by_month = {}
    for col in filtered_cols:
        year, month = parse_month_column(col)
        if year and month:
            data_by_month.setdefault(month, {}).setdefault(year, {})
            vol_market = df_vol_market[col].sum()
            vol_company = df_vol_company[col].sum()
            amt_market = df_amt_market[col].sum()
            amt_company = df_amt_company[col].sum()
            oi_market = df_oi_market[col].sum()
            oi_company = df_oi_company[col].sum()

            data_by_month[month][year]['成交量'] = safe_division(vol_company, vol_market * 2) * 100
            data_by_month[month][year]['成交额'] = safe_division(amt_company / 100000000, amt_market * 2) * 100
            data_by_month[month][year]['持仓量'] = safe_division(oi_company, oi_market * 2) * 100

    all_years = sorted({y for month_data in data_by_month.values() for y in month_data.keys()})
    latest_year = max(all_years) if all_years else 2024

    plot_data = []
    for month in sorted(data_by_month.keys()):
        for year in sorted(data_by_month[month].keys()):
            for metric in ['成交量', '成交额', '持仓量']:
                value = data_by_month[month][year].get(metric, 0)
                if value > 0:
                    plot_data.append({
                        '月份': MONTH_NAMES.get(month, str(month)),
                        '年份': f"{year}年",
                        '指标': metric,
                        '占比（%）': value
                    })

    if plot_data:
        plot_df = pd.DataFrame(plot_data)
        selected_metric_global = st.selectbox("选择查看指标", options=['成交量', '成交额', '持仓量'],
                                              key="metric_selector_global")
        metric_df = plot_df[plot_df['指标'] == selected_metric_global]

        if not metric_df.empty:
            decimal_places = 4 if selected_metric_global == '成交额' else 3
            
            # 获取实际存在的年份
            existing_years = sorted(metric_df['年份'].unique())
            
            # 为每个年份分配不同的颜色（使用三种不同的颜色）
            color_palette = ['#2E86C1', '#F39C12', '#28B463']  # 蓝色、橙色、绿色
            
            # 如果年份少于3个，只使用前几个颜色
            color_map = {}
            for i, year in enumerate(existing_years):
                color_map[year] = color_palette[i % len(color_palette)]
            
            fig = px.line(metric_df, x='月份', y='占比（%）', color='年份',
                          title=f'公司{selected_metric_global}占市场比重',
                          labels={'月份': '月份', '占比（%）': '占比（%）', '年份': '年份'},
                          color_discrete_map=color_map,
                          category_orders={'月份': list(MONTH_NAMES.values())})
            
            fig.update_yaxes(tickformat=f'.{decimal_places}f', title_text='占比（%）')
            fig.update_traces(texttemplate=f'%{{y:.{decimal_places}f}}', textposition='top center',
                              textfont=dict(size=10), mode='lines+markers+text')
            st.plotly_chart(fig, use_container_width=True)

        with st.expander("📋 查看详细数据"):
            dec = 4 if selected_metric_global == '成交额' else 3
            table_data = []
            for month in sorted(data_by_month.keys()):
                row = {'月份': MONTH_NAMES.get(month, str(month))}
                for year in sorted(data_by_month[month].keys()):
                    row[f"{year}年{selected_metric_global}"] = f"{data_by_month[month][year].get(selected_metric_global, 0):.{dec}f}"
                table_data.append(row)
            table_df = pd.DataFrame(table_data)
            st.dataframe(table_df, use_container_width=True, hide_index=True)
    else:
        st.info("暂无公司占市场比重数据")
except Exception as e:
    st.warning(f"无法加载公司/市场对比数据: {e}")

# ============================================================
# 各交易所单独展示折线图
# ============================================================
st.subheader("📊 公司占市场比重 - 各交易所单独展示")

try:
    date_cols_all = get_month_columns(df_vol_market)
    filtered_cols = []
    for col in date_cols_all:
        year, _ = parse_month_column(col)
        if year and year >= 2024:
            filtered_cols.append(col)
    
    exchanges_list = ['上期所', '能源中心', '郑商所', '大商所', '中金所', '广期所']
    data_by_month_year_exchange = {}
    for col in filtered_cols:
        year, month = parse_month_column(col)
        if year and month:
            data_by_month_year_exchange.setdefault(month, {}).setdefault(year, {})
            for ex in exchanges_list:
                data_by_month_year_exchange[month][year].setdefault(ex, {})
                
                vol_market_ex = df_vol_market[df_vol_market['交易所'] == ex][col].sum() if '交易所' in df_vol_market.columns else 0
                vol_company_ex = df_vol_company[df_vol_company['交易所'] == ex][col].sum() if '交易所' in df_vol_company.columns else 0
                amt_market_ex = df_amt_market[df_amt_market['交易所'] == ex][col].sum() if '交易所' in df_amt_market.columns else 0
                amt_company_ex = df_amt_company[df_amt_company['交易所'] == ex][col].sum() if '交易所' in df_amt_company.columns else 0
                oi_market_ex = df_oi_market[df_oi_market['交易所'] == ex][col].sum() if '交易所' in df_oi_market.columns else 0
                oi_company_ex = df_oi_company[df_oi_company['交易所'] == ex][col].sum() if '交易所' in df_oi_company.columns else 0
                
                data_by_month_year_exchange[month][year][ex]['成交量'] = safe_division(vol_company_ex, vol_market_ex * 2) * 100
                data_by_month_year_exchange[month][year][ex]['成交额'] = safe_division(amt_company_ex / 100000000, amt_market_ex * 2) * 100
                data_by_month_year_exchange[month][year][ex]['持仓量'] = safe_division(oi_company_ex, oi_market_ex * 2) * 100
    
    all_years = sorted({y for month_data in data_by_month_year_exchange.values() for y in month_data.keys()})
    latest_year = max(all_years) if all_years else 2024
    
    plot_data_exchange = []
    for month in sorted(data_by_month_year_exchange.keys()):
        for year in sorted(data_by_month_year_exchange[month].keys()):
            for ex in exchanges_list:
                for metric in ['成交量', '成交额', '持仓量']:
                    value = data_by_month_year_exchange[month][year][ex].get(metric, 0)
                    if value > 0:
                        plot_data_exchange.append({
                            '月份': MONTH_NAMES.get(month, str(month)),
                            '年份': str(year) + '年',
                            '交易所': ex,
                            '指标': metric,
                            '占比（%）': value
                        })
    
    if plot_data_exchange:
        plot_df_exchange = pd.DataFrame(plot_data_exchange)
        selected_metric_exchange = st.selectbox(
            "选择查看指标",
            options=['成交量', '成交额', '持仓量'],
            key="metric_selector_exchange"
        )
        metric_df_exchange = plot_df_exchange[plot_df_exchange['指标'] == selected_metric_exchange]
        
        if not metric_df_exchange.empty:
            decimal_places = 4 if selected_metric_exchange == '成交额' else 3
            color_map = {
                str(latest_year - 2) + '年': '#2E86C1',
                str(latest_year - 1) + '年': '#F39C12',
                str(latest_year) + '年': '#28B463'
            }
            color_map = {k: v for k, v in color_map.items() if k in metric_df_exchange['年份'].unique()}
            
            for i, ex in enumerate(exchanges_list):
                ex_df = metric_df_exchange[metric_df_exchange['交易所'] == ex]
                if not ex_df.empty:
                    if i % 2 == 0:
                        cols_container = st.columns(2)
                    with cols_container[i % 2]:
                        fig = px.line(
                            ex_df,
                            x='月份',
                            y='占比（%）',
                            color='年份',
                            title=f'{ex}',
                            labels={'月份': '', '占比（%）': '', '年份': ''},
                            color_discrete_map=color_map,
                            category_orders={'月份': ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']}
                        )
                        fig.update_yaxes(tickformat=f'.{decimal_places}f', title_text='', tickfont=dict(size=9))
                        fig.update_xaxes(tickvals=['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'], tickfont=dict(size=8))
                        fig.update_traces(
                            texttemplate=f'%{{y:.{decimal_places}f}}',
                            textposition='top center',
                            textfont=dict(size=7),
                            mode='lines+markers+text',
                            marker=dict(size=6)
                        )
                        fig.update_layout(
                            title_font=dict(size=12),
                            legend=dict(font=dict(size=9), orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
                            height=280,
                            margin=dict(l=40, r=40, t=50, b=30),
                            plot_bgcolor='#F8F9F9',
                            paper_bgcolor='white'
                        )
                        st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
        else:
            st.info(f"暂无{selected_metric_exchange}数据")
        
        with st.expander("📋 查看各交易所详细数据"):
            decimal_display = 4 if selected_metric_exchange == '成交额' else 3
            for ex in exchanges_list:
                st.subheader(f"{ex}")
                table_data = []
                for month in sorted(data_by_month_year_exchange.keys()):
                    row = {'月份': MONTH_NAMES.get(month, str(month))}
                    for year in sorted(data_by_month_year_exchange[month].keys()):
                        val = data_by_month_year_exchange[month][year][ex].get(selected_metric_exchange, 0)
                        row[f"{year}年"] = f"{val:.{decimal_display}f}"
                    table_data.append(row)
                table_df = pd.DataFrame(table_data)
                st.dataframe(table_df, use_container_width=True, hide_index=True)
    else:
        st.info("暂无公司占市场比重数据")
except Exception as e:
    st.warning(f"无法加载公司/市场对比数据: {e}")

# ============================================================
# 能源化工板块折线图
# ============================================================
st.subheader("📊 能源化工板块 - 公司占市场比重")

try:
    if '板块' in df_vol_market.columns:
        df_vol_market_energy = df_vol_market[df_vol_market['板块'] == '能源化工']
        df_vol_company_energy = df_vol_company[df_vol_company['板块'] == '能源化工']
        df_amt_market_energy = df_amt_market[df_amt_market['板块'] == '能源化工']
        df_amt_company_energy = df_amt_company[df_amt_company['板块'] == '能源化工']
        df_oi_market_energy = df_oi_market[df_oi_market['板块'] == '能源化工']
        df_oi_company_energy = df_oi_company[df_oi_company['板块'] == '能源化工']
    else:
        energy_exchanges = ['上期所', '能源中心']
        df_vol_market_energy = df_vol_market[df_vol_market['交易所'].isin(energy_exchanges)] if '交易所' in df_vol_market.columns else pd.DataFrame()
        df_vol_company_energy = df_vol_company[df_vol_company['交易所'].isin(energy_exchanges)] if '交易所' in df_vol_company.columns else pd.DataFrame()
        df_amt_market_energy = df_amt_market[df_amt_market['交易所'].isin(energy_exchanges)] if '交易所' in df_amt_market.columns else pd.DataFrame()
        df_amt_company_energy = df_amt_company[df_amt_company['交易所'].isin(energy_exchanges)] if '交易所' in df_amt_company.columns else pd.DataFrame()
        df_oi_market_energy = df_oi_market[df_oi_market['交易所'].isin(energy_exchanges)] if '交易所' in df_oi_market.columns else pd.DataFrame()
        df_oi_company_energy = df_oi_company[df_oi_company['交易所'].isin(energy_exchanges)] if '交易所' in df_oi_company.columns else pd.DataFrame()
    
    date_cols_all = get_month_columns(df_vol_market)
    filtered_cols = []
    for col in date_cols_all:
        year, _ = parse_month_column(col)
        if year and year >= 2024:
            filtered_cols.append(col)
    
    data_by_month_energy = {}
    for col in filtered_cols:
        year, month = parse_month_column(col)
        if year and month:
            data_by_month_energy.setdefault(month, {}).setdefault(year, {})
            
            vol_market = df_vol_market_energy[col].sum() if not df_vol_market_energy.empty else 0
            vol_company = df_vol_company_energy[col].sum() if not df_vol_company_energy.empty else 0
            amt_market = df_amt_market_energy[col].sum() if not df_amt_market_energy.empty else 0
            amt_company = df_amt_company_energy[col].sum() if not df_amt_company_energy.empty else 0
            oi_market = df_oi_market_energy[col].sum() if not df_oi_market_energy.empty else 0
            oi_company = df_oi_company_energy[col].sum() if not df_oi_company_energy.empty else 0
            
            data_by_month_energy[month][year]['成交量'] = safe_division(vol_company, vol_market * 2) * 100
            data_by_month_energy[month][year]['成交额'] = safe_division(amt_company / 100000000, amt_market * 2) * 100
            data_by_month_energy[month][year]['持仓量'] = safe_division(oi_company, oi_market * 2) * 100
    
    all_years = sorted({y for month_data in data_by_month_energy.values() for y in month_data.keys()})
    latest_year = max(all_years) if all_years else 2024
    
    plot_data_energy = []
    for month in sorted(data_by_month_energy.keys()):
        for year in sorted(data_by_month_energy[month].keys()):
            for metric in ['成交量', '成交额', '持仓量']:
                value = data_by_month_energy[month][year].get(metric, 0)
                if value > 0:
                    plot_data_energy.append({
                        '月份': MONTH_NAMES.get(month, str(month)),
                        '年份': str(year) + '年',
                        '指标': metric,
                        '占比（%）': value
                    })
    
    if plot_data_energy:
        plot_df_energy = pd.DataFrame(plot_data_energy)
        selected_metric_energy = st.selectbox(
            "选择查看指标",
            options=['成交量', '成交额', '持仓量'],
            key="metric_selector_energy"
        )
        color_map_energy = {
            str(latest_year - 2) + '年': '#2E86C1',
            str(latest_year - 1) + '年': '#F39C12',
            str(latest_year) + '年': '#28B463'
        }
        metric_df_energy = plot_df_energy[plot_df_energy['指标'] == selected_metric_energy]
        color_map_energy = {k: v for k, v in color_map_energy.items() if k in metric_df_energy['年份'].unique()}
        
        if not metric_df_energy.empty:
            st.subheader(f"📈 能源化工板块 - {selected_metric_energy}公司占市场比重")
            decimal_places = 4 if selected_metric_energy == '成交额' else 3
            fig = px.line(
                metric_df_energy,
                x='月份',
                y='占比（%）',
                color='年份',
                title=f'能源化工板块 - {selected_metric_energy}公司占市场比重',
                labels={'月份': '月份', '占比（%）': '占比（%）', '年份': '年份'},
                color_discrete_map=color_map_energy,
                category_orders={'月份': ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月']}
            )
            fig.update_yaxes(tickformat=f'.{decimal_places}f', title_text='占比（%）')
            fig.update_xaxes(tickvals=['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'])
            fig.update_traces(
                texttemplate=f'%{{y:.{decimal_places}f}}',
                textposition='top center',
                textfont=dict(size=10),
                mode='lines+markers+text'
            )
            fig.update_layout(
                title_font=dict(size=16, color='#1A5276'),
                plot_bgcolor='#F8F9F9',
                paper_bgcolor='white',
                height=400,
                legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5)
            )
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(f"暂无能源化工板块{selected_metric_energy}数据")
        
        with st.expander("📋 查看详细数据"):
            decimal_display = 4 if selected_metric_energy == '成交额' else 3
            table_data = []
            for month in sorted(data_by_month_energy.keys()):
                row = {'月份': MONTH_NAMES.get(month, str(month))}
                for year in sorted(data_by_month_energy[month].keys()):
                    key = f"{year}年{selected_metric_energy}"
                    val = data_by_month_energy[month][year].get(selected_metric_energy, 0)
                    row[key] = f"{val:.{decimal_display}f}"
                table_data.append(row)
            table_df_energy = pd.DataFrame(table_data)
            st.dataframe(table_df_energy, use_container_width=True, hide_index=True)
    else:
        st.info("暂无能源化工板块公司占市场比重数据")
except Exception as e:
    st.warning(f"无法加载能源化工板块对比数据: {e}")

# ============================================================
# 各板块分析
# ============================================================
if 'data_type' in locals():
    st.subheader(f"📊 各板块{data_type}对比（市场）{metric_config['market_title']}")

    col_filter1, col_filter2, col_filter3 = st.columns(3)
    with col_filter1:
        group_data_type = st.radio("选择数据类型", options=['成交量', '成交额', '持仓量'],
                                   horizontal=True, key="group_data_type")
    group_config = get_metric_config(group_data_type)

    group_market_df = {
        '成交量': df_vol_market, '成交额': df_amt_market, '持仓量': df_oi_market
    }.get(group_data_type, pd.DataFrame())
    group_company_df = {
        '成交量': df_vol_company, '成交额': df_amt_company, '持仓量': df_oi_company
    }.get(group_data_type, pd.DataFrame())

    if group_market_df.empty:
        st.warning(f"⚠️ {group_data_type}板块数据为空")
    else:
        group_date_cols = get_month_columns(group_market_df)
        with col_filter2:
            time_dimension = st.radio("选择时间维度", options=['月度', '季度', '年度'],
                                      horizontal=True, key="time_dimension_group")

        if time_dimension == '月度':
            options = sorted(group_date_cols, reverse=True)
            label_func = lambda x: f"{str(x)[:4]}年{str(x)[4:6]}月"
            value_cols_map = None
        elif time_dimension == '季度':
            quarter_map = {}
            for col in group_date_cols:
                year, month = parse_month_column(col)
                if year and month:
                    q = f"Q{(month - 1) // 3 + 1}"
                    quarter_map.setdefault(f"{year}{q}", []).append(col)
            options = sorted(quarter_map.keys(), reverse=True)
            label_func = lambda x: f"{x[:4]}年{x[4:]}"
            value_cols_map = quarter_map
        else:
            year_map = {}
            for col in group_date_cols:
                year, _ = parse_month_column(col)
                if year:
                    year_map.setdefault(str(year), []).append(col)
            options = sorted(year_map.keys(), reverse=True)
            label_func = lambda x: f"{x}年"
            value_cols_map = year_map

        with col_filter3:
            selected_key = st.selectbox(f"选择{time_dimension}", options=options,
                                        format_func=label_func, key="time_selector_group")

        if time_dimension == '月度':
            selected_cols_group = [selected_key]
        else:
            selected_cols_group = value_cols_map.get(selected_key, [])
            if not selected_cols_group:
                selected_cols_group = [group_date_cols[-1]] if group_date_cols else []

        if time_dimension == '月度':
            current_label, prev_label, last_label = '本月', '上月', '去年同期'
        elif time_dimension == '季度':
            current_label, prev_label, last_label = '本季度', '上季度', '去年同期'
        else:
            current_label, prev_label, last_label = '今年', '去年', None

        prev_cols_group, last_cols_group = [], []
        if time_dimension == '月度' and selected_cols_group:
            year, month = parse_month_column(selected_cols_group[0])
            if year and month:
                prev_key = f"{year if month > 1 else year - 1}{month - 1 if month > 1 else 12:02d}"
                prev_key = int(prev_key) if prev_key.isdigit() else prev_key
                prev_cols_group = [prev_key] if prev_key in group_market_df.columns else []
                last_key = f"{year - 1}{month:02d}"
                last_key = int(last_key) if last_key.isdigit() else last_key
                last_cols_group = [last_key] if last_key in group_market_df.columns else []

        group_col = '板块' if '板块' in group_market_df.columns else '交易所'
        groups = group_market_df[group_col].unique().tolist()

        group_df = build_comparison_table(group_market_df, group_col, groups,
                                          selected_cols_group, prev_cols_group, last_cols_group,
                                          group_config['market_divide'], current_label, prev_label, last_label)
        group_company_df_plot = build_comparison_table(group_company_df, group_col, groups,
                                                       selected_cols_group, prev_cols_group, last_cols_group,
                                                       group_config['company_divide'], current_label, prev_label, last_label)

        value_cols_group = [current_label]
        if prev_cols_group:
            value_cols_group.append(prev_label)
        if last_cols_group and last_label:
            value_cols_group.append(last_label)

        for df_plot, suffix, divide, yaxis in [
            (group_df, '市场', group_config['market_divide'], group_config['market_yaxis']),
            (group_company_df_plot, '公司', group_config['company_divide'], group_config['company_yaxis'])
        ]:
            melted = df_plot.melt(id_vars=[group_col], value_vars=value_cols_group,
                                  var_name='期间', value_name=group_data_type)
            melted = melted.dropna(subset=[group_data_type])
            period_order = [p for p in [prev_label, current_label, last_label] if p and p in melted['期间'].unique()]
            melted['期间'] = pd.Categorical(melted['期间'], categories=period_order, ordered=True)
            melted = melted.sort_values('期间')

            if not melted.empty:
                fig = create_bar_chart(melted, group_col, group_data_type, '期间',
                                       f'各板块{group_data_type}对比（{suffix}）- {label_func(selected_key)}',
                                       yaxis, COLOR_MAP)
                if fig:
                    st.plotly_chart(fig, use_container_width=True)

            st.subheader(f"📊 板块环比同比综合表（{suffix}）")
            total = compute_period_comparison(group_market_df if suffix == '市场' else group_company_df,
                                              selected_cols_group, prev_cols_group, last_cols_group,
                                              divide)
            unit = group_config['market_unit' if suffix == '市场' else 'company_unit']
            
            compare_label1 = '环比' if time_dimension != '年度' else '同比'
            compare_label2 = '同比' if time_dimension != '年度' else '-'
            
            table_data = []
            total_row = {
                '维度': '合计',
                f'{group_data_type}': total['current'],
                'compare1': total['mom'] if time_dimension != '年度' else total['yoy'],
                'compare2': total['yoy'] if time_dimension != '年度' else None
            }
            table_data.append(total_row)
            
            for _, row in df_plot.iterrows():
                table_data.append({
                    '维度': row[group_col],
                    f'{group_data_type}': row[current_label],
                    'compare1': row.get('环比' if time_dimension != '年度' else '同比'),
                    'compare2': row.get('同比' if time_dimension != '年度' else None)
                })
            
            df_table = pd.DataFrame(table_data)
            
            st.dataframe(
                df_table,
                use_container_width=True,
                hide_index=True,
                column_config={
                    '维度': st.column_config.TextColumn('维度'),
                    f'{group_data_type}': st.column_config.NumberColumn(
                        f'{group_data_type}（{unit}）',
                        format="%.2f"
                    ),
                    'compare1': st.column_config.NumberColumn(
                        f'{compare_label1}（%）',
                        format="%+.2f%%"
                    ),
                    'compare2': st.column_config.NumberColumn(
                        f'{compare_label2}（%）',
                        format="%+.2f%%"
                    )
                }
            )

# ============================================================
# 市场各板块份额饼图 + 公司占有率柱状图（并列）
# ============================================================
st.subheader("📊 市场各板块分析")

# 获取日期列和标签
date_cols = get_month_columns(df_vol_market)
date_labels = {col: f"{str(col)[:4]}年{str(col)[4:6]}月" for col in sorted(date_cols)}

col1, col2 = st.columns(2)
with col1:
    pie_month_options = sorted(date_cols, reverse=True)
    pie_month_labels = {col: date_labels.get(col, str(col)) for col in pie_month_options}
    
    selected_pie_month = st.selectbox(
        "选择月份",
        options=pie_month_options,
        format_func=lambda x: pie_month_labels.get(x, str(x)),
        key="pie_month_selector"
    )
with col2:
    pie_data_type = st.radio(
        "选择指标",
        options=['成交量', '成交额', '持仓量'],
        horizontal=True,
        key="pie_data_type"
    )

# 根据饼图选择的数据类型加载对应的市场数据和公司数据
if pie_data_type == '成交量':
    pie_market_df = df_vol_market.copy()
    pie_company_df = df_vol_company.copy()
    pie_unit = '亿手'
    pie_title = '成交量'
elif pie_data_type == '成交额':
    pie_market_df = df_amt_market.copy()
    pie_company_df = df_amt_company.copy()
    pie_unit = '万亿元'
    pie_title = '成交额'
else:  # 持仓量
    pie_market_df = df_oi_market.copy()
    pie_company_df = df_oi_company.copy()
    pie_unit = '百万手'
    pie_title = '持仓量'

# 清理市场数据
pie_market_df = clean_dataframe(pie_market_df)
pie_company_df = clean_dataframe(pie_company_df)

# 检查选中的月份是否存在
if selected_pie_month in pie_market_df.columns and selected_pie_month in pie_company_df.columns:
    if '板块' in pie_market_df.columns:
        pie_market_grouped = pie_market_df.groupby('板块')[selected_pie_month].sum().reset_index()
        pie_market_grouped.columns = ['板块', '市场数值']
        
        pie_company_grouped = pie_company_df.groupby('板块')[selected_pie_month].sum().reset_index()
        pie_company_grouped.columns = ['板块', '公司数值']
        
        pie_merged = pd.merge(pie_market_grouped, pie_company_grouped, on='板块', how='outer').fillna(0)
        
        if pie_data_type == '成交额':
            pie_merged['公司占比（%）'] = ((pie_merged['公司数值'] / 100000000) / (pie_merged['市场数值'] * 2) * 100).round(4)
        else:
            pie_merged['公司占比（%）'] = (pie_merged['公司数值'] / (pie_merged['市场数值'] * 2) * 100).round(4)
        
        pie_merged['公司占比（%）'] = pie_merged.apply(
            lambda row: 0 if row['市场数值'] == 0 else row['公司占比（%）'], 
            axis=1
        )
        
        total_market = pie_merged['市场数值'].sum()
        
        if total_market > 0:
            pie_merged['市场占比（%）'] = (pie_merged['市场数值'] / total_market * 100).round(2)
            
            pie_data = pie_merged[['板块', '市场数值', '市场占比（%）']].copy()
            bar_data = pie_merged[['板块', '公司占比（%）']].copy()
            bar_data = bar_data.sort_values('公司占比（%）', ascending=True)
            
            col_left, col_right = st.columns([1, 1], gap="medium")
            
            with col_left:
                fig_pie = px.pie(
                    pie_data,
                    values='市场数值',
                    names='板块',
                    title=f'市场各板块{pie_title}份额',
                    hover_data={'市场数值': True, '市场占比（%）': True},
                    labels={'市场数值': f'{pie_title}（{pie_unit}）', '市场占比（%）': '占比（%）'},
                    hole=0.3
                )
                fig_pie.update_layout(
                    title_font=dict(size=16, color='#1A5276'),
                    font=dict(size=12),
                    legend=dict(orientation='v', yanchor='middle', y=0.5, xanchor='right', x=1.0),
                    height=450
                )
                fig_pie.update_traces(
                    textinfo='percent+label',
                    texttemplate='%{label}<br>%{percent:.2%}',
                    textfont=dict(size=11),
                    marker=dict(line=dict(color='white', width=2)),
                    pull=[0.03 if i == 0 else 0 for i in range(len(pie_data))]
                )
                st.plotly_chart(fig_pie, use_container_width=True, config={'displayModeBar': False})
            
            with col_right:
                bar_data_sorted = bar_data.sort_values('公司占比（%）', ascending=False)
                max_val = bar_data_sorted['公司占比（%）'].max()
                
                fig_bar = px.bar(
                    bar_data_sorted,
                    x='公司占比（%）',
                    y='板块',
                    orientation='h',
                    title=f'各板块{pie_title}公司占有率',
                    labels={'公司占比（%）': '公司占市场比重（%）', '板块': ''},
                    text='公司占比（%）',
                    color='公司占比（%）',
                    color_continuous_scale='Blues',
                    range_color=[0, max_val * 1.2 if max_val > 0 else 1]
                )
                fig_bar.update_layout(
                    title_font=dict(size=16, color='#1A5276'),
                    font=dict(size=12),
                    xaxis=dict(
                        title='公司占市场比重（%）',
                        tickformat='.4f',
                        range=[0, max_val * 1.3 if max_val > 0 else 1]
                    ),
                    yaxis=dict(title='', categoryorder='total ascending'),
                    height=450,
                    showlegend=False,
                    coloraxis_showscale=False,
                    plot_bgcolor='#F8F9F9',
                    paper_bgcolor='white'
                )
                fig_bar.update_traces(
                    texttemplate='%{x:.4f}%',
                    textposition='outside',
                    textfont=dict(size=11, color='black'),
                    marker=dict(line=dict(color='white', width=1))
                )
                st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})
            
            with st.expander("📋 查看详细数据"):
                display_df = pie_merged.copy()
                
                if pie_data_type == '成交量':
                    market_unit_display = '手'
                    company_unit_display = '手'
                elif pie_data_type == '成交额':
                    market_unit_display = '亿'
                    company_unit_display = '亿'
                    display_df['公司数值'] = display_df['公司数值'] / 100000000
                else:  # 持仓量
                    market_unit_display = '手'
                    company_unit_display = '手'
                
                display_cols = ['板块', '市场数值', '公司数值', '公司占比（%）', '市场占比（%）']
                display_df_final = display_df[display_cols].copy()
                display_df_final.columns = ['板块', f'市场数值（{market_unit_display}）', f'公司数值（{company_unit_display}）', '公司占比（%）', '市场占比（%）']
                
                st.dataframe(
                    display_df_final,
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "板块": st.column_config.TextColumn("板块"),
                        f"市场数值（{market_unit_display}）": st.column_config.NumberColumn(
                            f"市场数值（{market_unit_display}）",
                            format="%,.2f" if pie_data_type == '成交额' else "%,.0f"
                        ),
                        f"公司数值（{company_unit_display}）": st.column_config.NumberColumn(
                            f"公司数值（{company_unit_display}）",
                            format="%,.2f" if pie_data_type == '成交额' else "%,.0f"
                        ),
                        "公司占比（%）": st.column_config.NumberColumn(
                            "公司占比（%）",
                            format="%.4f%%"
                        ),
                        "市场占比（%）": st.column_config.NumberColumn(
                            "市场占比（%）",
                            format="%.2f%%"
                        )
                    }
                )
        else:
            st.warning("该月份市场数据为空，请选择其他月份")
    else:
        st.warning("数据中无'板块'列，请检查数据")
else:
    st.warning(f"选中的月份 {selected_pie_month} 不存在于数据中")

# ============================================================
# 期货品种分析-市场
# ============================================================
st.subheader("📊 期货品种分析-市场")

# 根据饼图选择的数据类型加载对应的市场数据和公司数据
if pie_data_type == '成交量':
    future_market_df = df_vol_market.copy()
    future_company_df = df_vol_company.copy()
    future_unit = '亿手'
    future_title = '成交量'
elif pie_data_type == '成交额':
    future_market_df = df_amt_market.copy()
    future_company_df = df_amt_company.copy()
    future_unit = '万亿元'
    future_title = '成交额'
else:  # 持仓量
    future_market_df = df_oi_market.copy()
    future_company_df = df_oi_company.copy()
    future_unit = '百万手'
    future_title = '持仓量'

# 清理数据
future_market_df = clean_dataframe(future_market_df)
future_company_df = clean_dataframe(future_company_df)

# ===== 筛选期货品种（产品类型 == '期货'） =====
if '产品类型' in future_market_df.columns:
    future_market_df = future_market_df[future_market_df['产品类型'] == '期货']
    future_company_df = future_company_df[future_company_df['产品类型'] == '期货']
else:
    st.warning("⚠️ 数据中无'产品类型'列，无法筛选期货品种")

# 检查选中的月份是否存在
if not future_market_df.empty:
    month_exists_in_market = selected_pie_month in future_market_df.columns
    month_exists_in_company = selected_pie_month in future_company_df.columns
    
    if not month_exists_in_market:
        st.warning(f"⚠️ 选中的月份 {selected_pie_month} 不存在于期货市场数据中")
        # 尝试找一个存在的月份
        available_months = [col for col in future_market_df.columns if isinstance(col, (int, float)) or (isinstance(col, str) and col.isdigit())]
        if available_months:
            fallback_month = available_months[0]
            st.info(f"📌 使用替代月份: {fallback_month}")
            actual_month = fallback_month
        else:
            st.warning("无可用月份数据")
            actual_month = None
    else:
        actual_month = selected_pie_month
    
    if actual_month is not None and actual_month in future_market_df.columns and actual_month in future_company_df.columns:
        # 按品种分组汇总
        if '产品名称' in future_market_df.columns:
            future_market_grouped = future_market_df.groupby('产品名称')[actual_month].sum().reset_index()
            future_market_grouped.columns = ['品种', '市场数值']
            future_company_grouped = future_company_df.groupby('产品名称')[actual_month].sum().reset_index()
            future_company_grouped.columns = ['品种', '公司数值']
        else:
            st.warning("⚠️ 未找到'产品名称'列")
            future_market_grouped = None
        
        if future_market_grouped is not None and not future_market_grouped.empty:
            future_merged = pd.merge(future_market_grouped, future_company_grouped, on='品种', how='outer').fillna(0)
            
            # 成交额处理：公司数值从元转亿元
            if pie_data_type == '成交额':
                future_merged['公司数值'] = future_merged['公司数值'] / 100000000
            
            # 计算公司占比
            future_merged['公司占比（%）'] = (future_merged['公司数值'] / (future_merged['市场数值'] * 2) * 100).round(4)
            future_merged['公司占比（%）'] = future_merged.apply(
                lambda row: 0 if row['市场数值'] == 0 else row['公司占比（%）'], 
                axis=1
            )
            
            # 按市场数值降序排列，取前十，其余归为"其他"
            future_merged_sorted = future_merged.sort_values('市场数值', ascending=False).reset_index(drop=True)
            
            if len(future_merged_sorted) > 10:
                top_10 = future_merged_sorted.head(10).copy()
                other_market = future_merged_sorted.iloc[10:]['市场数值'].sum()
                other_company = future_merged_sorted.iloc[10:]['公司数值'].sum()
                other_ratio = (other_company / (other_market * 2) * 100).round(4) if other_market != 0 else 0
                
                other_row = pd.DataFrame({
                    '品种': ['其他'],
                    '市场数值': [other_market],
                    '公司数值': [other_company],
                    '公司占比（%）': [other_ratio]
                })
                future_merged_top = pd.concat([top_10, other_row], ignore_index=True)
            else:
                future_merged_top = future_merged_sorted.copy()
            
            # 计算市场占比
            total_market_future = future_merged_top['市场数值'].sum()
            
            if total_market_future > 0:
                future_merged_top['市场占比（%）'] = (future_merged_top['市场数值'] / total_market_future * 100).round(2)
                
                pie_data_future = future_merged_top[['品种', '市场数值', '市场占比（%）']].copy()
                bar_data_future = future_merged_top[['品种', '公司占比（%）']].copy()
                bar_data_future = bar_data_future.sort_values('公司占比（%）', ascending=True)
                
                col_left_future, col_right_future = st.columns([1, 1], gap="medium")
                
                # ===== 左侧：饼图 =====
                with col_left_future:
                    fig_pie_future = px.pie(
                        pie_data_future,
                        values='市场数值',
                        names='品种',
                        title=f'期货品种{future_title}份额（TOP10+其他）',
                        hover_data={'市场数值': True, '市场占比（%）': True},
                        labels={'市场数值': f'{future_title}（{future_unit}）', '市场占比（%）': '占比（%）'},
                        hole=0.3
                    )
                    fig_pie_future.update_layout(
                        title_font=dict(size=16, color='#1A5276'),
                        font=dict(size=12),
                        legend=dict(orientation='v', yanchor='middle', y=0.5, xanchor='right', x=1.0),
                        height=450
                    )
                    fig_pie_future.update_traces(
                        textinfo='percent+label',
                        texttemplate='%{label}<br>%{percent:.2%}',
                        textfont=dict(size=10),
                        marker=dict(line=dict(color='white', width=2)),
                        pull=[0.03 if i == 0 else 0 for i in range(len(pie_data_future))]
                    )
                    st.plotly_chart(fig_pie_future, use_container_width=True, config={'displayModeBar': False})
                
                # ===== 右侧：横向柱状图 =====
                with col_right_future:
                    bar_data_sorted_future = bar_data_future.sort_values('公司占比（%）', ascending=False)
                    max_val_future = bar_data_sorted_future['公司占比（%）'].max()
                    
                    fig_bar_future = px.bar(
                        bar_data_sorted_future,
                        x='公司占比（%）',
                        y='品种',
                        orientation='h',
                        title=f'期货品种{future_title}公司占有率（TOP10+其他）',
                        labels={'公司占比（%）': '公司占市场比重（%）', '品种': ''},
                        text='公司占比（%）',
                        color='公司占比（%）',
                        color_continuous_scale='Blues',
                        range_color=[0, max_val_future * 1.2 if max_val_future > 0 else 1]
                    )
                    fig_bar_future.update_layout(
                        title_font=dict(size=16, color='#1A5276'),
                        font=dict(size=12),
                        xaxis=dict(
                            title='公司占市场比重（%）',
                            tickformat='.4f',
                            range=[0, max_val_future * 1.3 if max_val_future > 0 else 1]
                        ),
                        yaxis=dict(title='', categoryorder='total ascending'),
                        height=450,
                        showlegend=False,
                        coloraxis_showscale=False,
                        plot_bgcolor='#F8F9F9',
                        paper_bgcolor='white'
                    )
                    fig_bar_future.update_traces(
                        texttemplate='%{x:.4f}%',
                        textposition='outside',
                        textfont=dict(size=11, color='black'),
                        marker=dict(line=dict(color='white', width=1))
                    )
                    st.plotly_chart(fig_bar_future, use_container_width=True, config={'displayModeBar': False})
                
                # ===== 数据表格 =====
                with st.expander("📋 查看详细数据"):
                    display_df_future = future_merged_top.copy()
                    
                    if pie_data_type == '成交量':
                        market_unit_display_future = '手'
                        company_unit_display_future = '手'
                    elif pie_data_type == '成交额':
                        market_unit_display_future = '亿元'
                        company_unit_display_future = '亿元'
                    else:  # 持仓量
                        market_unit_display_future = '手'
                        company_unit_display_future = '手'
                    
                    display_cols_future = ['品种', '市场数值', '公司数值', '公司占比（%）', '市场占比（%）']
                    display_df_future_final = display_df_future[display_cols_future].copy()
                    display_df_future_final.columns = ['品种', f'市场数值（{market_unit_display_future}）', f'公司数值（{company_unit_display_future}）', '公司占比（%）', '市场占比（%）']
                    
                    st.dataframe(
                        display_df_future_final,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "品种": st.column_config.TextColumn("品种"),
                            f"市场数值（{market_unit_display_future}）": st.column_config.NumberColumn(
                                f"市场数值（{market_unit_display_future}）",
                                format="%,.2f" if pie_data_type == '成交额' else "%,.0f"
                            ),
                            f"公司数值（{company_unit_display_future}）": st.column_config.NumberColumn(
                                f"公司数值（{company_unit_display_future}）",
                                format="%,.2f" if pie_data_type == '成交额' else "%,.0f"
                            ),
                            "公司占比（%）": st.column_config.NumberColumn(
                                "公司占比（%）",
                                format="%.4f%%"
                            ),
                            "市场占比（%）": st.column_config.NumberColumn(
                                "市场占比（%）",
                                format="%.2f%%"
                            )
                        }
                    )
            else:
                st.warning("该月份期货市场总值为0，请选择其他月份")
        else:
            st.warning("期货品种数据为空，请检查数据")
    else:
        if actual_month is None:
            st.warning("无法找到可用的月份数据")
        else:
            st.warning(f"月份 {actual_month} 在期货数据中不存在")
else:
    st.info("📊 当前数据中没有期货品种（产品类型='期货'），跳过期货品种分析")

# ============================================================
# 期权品种分析-市场
# ============================================================
st.subheader("📊 期权品种分析-市场")

# 根据饼图选择的数据类型加载对应的市场数据和公司数据
if pie_data_type == '成交量':
    option_market_df = df_vol_market.copy()
    option_company_df = df_vol_company.copy()
    option_unit = '亿手'
    option_title = '成交量'
elif pie_data_type == '成交额':
    option_market_df = df_amt_market.copy()
    option_company_df = df_amt_company.copy()
    option_unit = '万亿元'
    option_title = '成交额'
else:  # 持仓量
    option_market_df = df_oi_market.copy()
    option_company_df = df_oi_company.copy()
    option_unit = '百万手'
    option_title = '持仓量'

# 筛选期权品种（产品类型 == '期货期权' 或 '现货期权'）
if '产品类型' in option_market_df.columns:
    option_market_df = option_market_df[option_market_df['产品类型'].isin(['期货期权', '现货期权'])]
    option_company_df = option_company_df[option_company_df['产品类型'].isin(['期货期权', '现货期权'])]

# 清理数据
option_market_df = clean_dataframe(option_market_df)
option_company_df = clean_dataframe(option_company_df)

# 检查选中的月份是否存在
if selected_pie_month in option_market_df.columns and selected_pie_month in option_company_df.columns:
    # 按品种分组汇总
    option_market_grouped = option_market_df.groupby('产品名称')[selected_pie_month].sum().reset_index()
    option_market_grouped.columns = ['品种', '市场数值']
    
    option_company_grouped = option_company_df.groupby('产品名称')[selected_pie_month].sum().reset_index()
    option_company_grouped.columns = ['品种', '公司数值']
    
    option_merged = pd.merge(option_market_grouped, option_company_grouped, on='品种', how='outer').fillna(0)
    
    if pie_data_type == '成交额':
        option_merged['公司数值'] = option_merged['公司数值'] / 100000000
    
    # 计算公司占比（四位小数）
    option_merged['公司占比（%）'] = (option_merged['公司数值'] / (option_merged['市场数值'] * 2) * 100).round(4)
    option_merged['公司占比（%）'] = option_merged.apply(
        lambda row: 0 if row['市场数值'] == 0 else row['公司占比（%）'], 
        axis=1
    )
    
    # 按市场数值降序排列，取前十，其余归为"其他"
    option_merged_sorted = option_merged.sort_values('市场数值', ascending=False).reset_index(drop=True)
    
    if len(option_merged_sorted) > 10:
        top_10 = option_merged_sorted.head(10).copy()
        other_market = option_merged_sorted.iloc[10:]['市场数值'].sum()
        other_company = option_merged_sorted.iloc[10:]['公司数值'].sum()
        other_ratio = (other_company / (other_market * 2) * 100).round(4) if other_market != 0 else 0
        
        other_row = pd.DataFrame({
            '品种': ['其他'],
            '市场数值': [other_market],
            '公司数值': [other_company],
            '公司占比（%）': [other_ratio]
        })
        option_merged_top = pd.concat([top_10, other_row], ignore_index=True)
    else:
        option_merged_top = option_merged_sorted.copy()
    
    # 计算市场占比
    total_market_option = option_merged_top['市场数值'].sum()
    
    if total_market_option > 0:
        option_merged_top['市场占比（%）'] = (option_merged_top['市场数值'] / total_market_option * 100).round(2)
        
        pie_data_option = option_merged_top[['品种', '市场数值', '市场占比（%）']].copy()
        bar_data_option = option_merged_top[['品种', '公司占比（%）']].copy()
        bar_data_option = bar_data_option.sort_values('公司占比（%）', ascending=True)
        
        col_left_option, col_right_option = st.columns([1, 1], gap="medium")
        
        with col_left_option:
            fig_pie_option = px.pie(
                pie_data_option,
                values='市场数值',
                names='品种',
                title=f'期权品种{option_title}份额（TOP10+其他）',
                hover_data={'市场数值': True, '市场占比（%）': True},
                labels={'市场数值': f'{option_title}（{option_unit}）', '市场占比（%）': '占比（%）'},
                hole=0.3
            )
            fig_pie_option.update_layout(
                title_font=dict(size=16, color='#1A5276'),
                font=dict(size=12),
                legend=dict(orientation='v', yanchor='middle', y=0.5, xanchor='right', x=1.0),
                height=450
            )
            fig_pie_option.update_traces(
                textinfo='percent+label',
                texttemplate='%{label}<br>%{percent:.2%}',
                textfont=dict(size=10),
                marker=dict(line=dict(color='white', width=2)),
                pull=[0.03 if i == 0 else 0 for i in range(len(pie_data_option))]
            )
            st.plotly_chart(fig_pie_option, use_container_width=True, config={'displayModeBar': False})
        
        with col_right_option:
            bar_data_sorted_option = bar_data_option.sort_values('公司占比（%）', ascending=False)
            max_val_option = bar_data_sorted_option['公司占比（%）'].max()
            
            fig_bar_option = px.bar(
                bar_data_sorted_option,
                x='公司占比（%）',
                y='品种',
                orientation='h',
                title=f'期权品种{option_title}公司占有率（TOP10+其他）',
                labels={'公司占比（%）': '公司占市场比重（%）', '品种': ''},
                text='公司占比（%）',
                color='公司占比（%）',
                color_continuous_scale='Blues',
                range_color=[0, max_val_option * 1.2 if max_val_option > 0 else 1]
            )
            fig_bar_option.update_layout(
                title_font=dict(size=16, color='#1A5276'),
                font=dict(size=12),
                xaxis=dict(
                    title='公司占市场比重（%）',
                    tickformat='.4f',
                    range=[0, max_val_option * 1.3 if max_val_option > 0 else 1]
                ),
                yaxis=dict(title='', categoryorder='total ascending'),
                height=450,
                showlegend=False,
                coloraxis_showscale=False,
                plot_bgcolor='#F8F9F9',
                paper_bgcolor='white'
            )
            fig_bar_option.update_traces(
                texttemplate='%{x:.4f}%',
                textposition='outside',
                textfont=dict(size=11, color='black'),
                marker=dict(line=dict(color='white', width=1))
            )
            st.plotly_chart(fig_bar_option, use_container_width=True, config={'displayModeBar': False})
        
        with st.expander("📋 查看详细数据"):
            display_df_option = option_merged_top.copy()
            
            if pie_data_type == '成交量':
                market_unit_display_option = '手'
                company_unit_display_option = '手'
            elif pie_data_type == '成交额':
                market_unit_display_option = '亿元'
                company_unit_display_option = '亿元'
            else:
                market_unit_display_option = '手'
                company_unit_display_option = '手'
            
            display_cols_option = ['品种', '市场数值', '公司数值', '公司占比（%）', '市场占比（%）']
            display_df_option_final = display_df_option[display_cols_option].copy()
            display_df_option_final.columns = ['品种', f'市场数值（{market_unit_display_option}）', f'公司数值（{company_unit_display_option}）', '公司占比（%）', '市场占比（%）']
            
            st.dataframe(
                display_df_option_final,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "品种": st.column_config.TextColumn("品种"),
                    f"市场数值（{market_unit_display_option}）": st.column_config.NumberColumn(
                        f"市场数值（{market_unit_display_option}）",
                        format="%,.2f" if pie_data_type == '成交额' else "%,.0f"
                    ),
                    f"公司数值（{company_unit_display_option}）": st.column_config.NumberColumn(
                        f"公司数值（{company_unit_display_option}）",
                        format="%,.2f" if pie_data_type == '成交额' else "%,.0f"
                    ),
                    "公司占比（%）": st.column_config.NumberColumn(
                        "公司占比（%）",
                        format="%.4f%%"
                    ),
                    "市场占比（%）": st.column_config.NumberColumn(
                        "市场占比（%）",
                        format="%.2f%%"
                    )
                }
            )
    else:
        st.warning("该月份期权数据为空，请选择其他月份")
else:
    st.warning(f"选中的月份 {selected_pie_month} 不存在于期权数据中")

# ============================================================
# 期货品种分析-公司
# ============================================================
st.subheader("📊 期货品种分析-公司")

# 根据饼图选择的数据类型加载对应的市场数据和公司数据
if pie_data_type == '成交量':
    company_market_df = df_vol_market.copy()
    company_company_df = df_vol_company.copy()
    company_unit = '手'
    company_title = '成交量'
elif pie_data_type == '成交额':
    company_market_df = df_amt_market.copy()
    company_company_df = df_amt_company.copy()
    company_unit = '元'
    company_title = '成交额'
else:  # 持仓量
    company_market_df = df_oi_market.copy()
    company_company_df = df_oi_company.copy()
    company_unit = '手'
    company_title = '持仓量'

# 筛选期货品种（产品类型 == '期货'）
if '产品类型' in company_market_df.columns:
    company_market_df = company_market_df[company_market_df['产品类型'] == '期货']
    company_company_df = company_company_df[company_company_df['产品类型'] == '期货']

# 清理数据
company_market_df = clean_dataframe(company_market_df)
company_company_df = clean_dataframe(company_company_df)

if not company_market_df.empty and selected_pie_month in company_market_df.columns and selected_pie_month in company_company_df.columns:
    company_market_grouped = company_market_df.groupby('产品名称')[selected_pie_month].sum().reset_index()
    company_market_grouped.columns = ['品种', '市场数值']
    
    company_company_grouped = company_company_df.groupby('产品名称')[selected_pie_month].sum().reset_index()
    company_company_grouped.columns = ['品种', '公司数值']
    
    company_merged = pd.merge(company_market_grouped, company_company_grouped, on='品种', how='outer').fillna(0)
    
    # 成交额处理：公司数值从元转亿元
    if pie_data_type == '成交额':
        company_merged['公司数值'] = company_merged['公司数值'] / 100000000
    
    # 计算公司占市场比重（用于右侧柱状图）
    company_merged['公司占市场比重（%）'] = (company_merged['公司数值'] / (company_merged['市场数值'] * 2) * 100).round(4)
    company_merged['公司占市场比重（%）'] = company_merged.apply(
        lambda row: 0 if row['市场数值'] == 0 else row['公司占市场比重（%）'], 
        axis=1
    )
    
    # ★★★ 按公司数值降序排列（饼图按公司份额排序）★★★
    company_merged_sorted = company_merged.sort_values('公司数值', ascending=False).reset_index(drop=True)
    
    if len(company_merged_sorted) > 10:
        top_10 = company_merged_sorted.head(10).copy()
        other_market = company_merged_sorted.iloc[10:]['市场数值'].sum()
        other_company = company_merged_sorted.iloc[10:]['公司数值'].sum()
        other_ratio = (other_company / (other_market * 2) * 100).round(4) if other_market != 0 else 0
        
        other_row = pd.DataFrame({
            '品种': ['其他'],
            '市场数值': [other_market],
            '公司数值': [other_company],
            '公司占市场比重（%）': [other_ratio]
        })
        company_merged_top = pd.concat([top_10, other_row], ignore_index=True)
    else:
        company_merged_top = company_merged_sorted.copy()
    
    total_market_company = company_merged_top['市场数值'].sum()
    
    if total_market_company > 0:
        # 计算公司内部占比（饼图用）
        total_company = company_merged_top['公司数值'].sum()
        company_merged_top['公司内部占比（%）'] = (company_merged_top['公司数值'] / total_company * 100).round(2) if total_company > 0 else 0
        
        # ===== 饼图和柱状图左右并列 =====
        col_left_company, col_right_company = st.columns([1, 1], gap="medium")
        
        # 左列：饼图（公司各品种份额）
        with col_left_company:
            pie_data_company = company_merged_top[['品种', '公司数值', '公司内部占比（%）']].copy()
            fig_pie_company = px.pie(
                pie_data_company,
                values='公司数值',
                names='品种',
                title=f'公司期货{company_title}份额（TOP10）',
                hover_data={'公司数值': True, '公司内部占比（%）': True},
                labels={'公司数值': f'{company_title}（{company_unit}）', '公司内部占比（%）': '占比（%）'},
                hole=0.3
            )
            fig_pie_company.update_layout(
                title_font=dict(size=16, color='#1A5276'),
                font=dict(size=12),
                legend=dict(orientation='v', yanchor='middle', y=0.5, xanchor='right', x=1.0),
                height=450
            )
            fig_pie_company.update_traces(
                textinfo='percent+label',
                texttemplate='%{label}<br>%{percent:.2%}',
                textfont=dict(size=10),
                marker=dict(line=dict(color='white', width=2)),
                pull=[0.03 if i == 0 else 0 for i in range(len(pie_data_company))]
            )
            st.plotly_chart(fig_pie_company, use_container_width=True, config={'displayModeBar': False})
        
        # 右列：柱状图（公司占市场比重）
        with col_right_company:
            bar_data_company = company_merged_top[['品种', '公司占市场比重（%）']].copy()
            bar_data_sorted_company = bar_data_company.sort_values('公司占市场比重（%）', ascending=False)
            max_val_company = bar_data_sorted_company['公司占市场比重（%）'].max()
            
            fig_bar_company = px.bar(
                bar_data_sorted_company,
                x='公司占市场比重（%）',
                y='品种',
                orientation='h',
                title=f'公司期货{company_title}占市场比重（TOP10）',
                labels={'公司占市场比重（%）': '公司占市场比重（%）', '品种': ''},
                text='公司占市场比重（%）',
                color='公司占市场比重（%）',
                color_continuous_scale='Blues',
                range_color=[0, max_val_company * 1.2 if max_val_company > 0 else 1]
            )
            fig_bar_company.update_layout(
                title_font=dict(size=16, color='#1A5276'),
                font=dict(size=12),
                xaxis=dict(
                    title='公司占市场比重（%）',
                    tickformat='.4f',
                    range=[0, max_val_company * 1.3 if max_val_company > 0 else 1]
                ),
                yaxis=dict(title='', categoryorder='total ascending'),
                height=450,
                showlegend=False,
                coloraxis_showscale=False,
                plot_bgcolor='#F8F9F9',
                paper_bgcolor='white'
            )
            fig_bar_company.update_traces(
                texttemplate='%{x:.4f}%',
                textposition='outside',
                textfont=dict(size=11, color='black'),
                marker=dict(line=dict(color='white', width=1))
            )
            st.plotly_chart(fig_bar_company, use_container_width=True, config={'displayModeBar': False})
        
        # 数据表格
        with st.expander("📋 查看详细数据"):
            display_df_company = company_merged_top.copy()
            # 确保表格按公司数值降序排列
            display_df_company = display_df_company.sort_values('公司数值', ascending=False)
            
            if pie_data_type == '成交量':
                market_unit_display_company = '手'
                company_unit_display_company = '手'
            elif pie_data_type == '成交额':
                market_unit_display_company = '亿元'
                company_unit_display_company = '亿元'
            else:
                market_unit_display_company = '手'
                company_unit_display_company = '手'
            
            display_cols_company = ['品种', '市场数值', '公司数值', '公司内部占比（%）', '公司占市场比重（%）']
            display_df_company_final = display_df_company[display_cols_company].copy()
            display_df_company_final.columns = ['品种', f'市场数值（{market_unit_display_company}）', f'公司数值（{company_unit_display_company}）', '公司内部占比（%）', '公司占市场比重（%）']
            
            st.dataframe(
                display_df_company_final,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "品种": st.column_config.TextColumn("品种"),
                    f"市场数值（{market_unit_display_company}）": st.column_config.NumberColumn(
                        f"市场数值（{market_unit_display_company}）",
                        format="%,.2f" if pie_data_type == '成交额' else "%,.0f"
                    ),
                    f"公司数值（{company_unit_display_company}）": st.column_config.NumberColumn(
                        f"公司数值（{company_unit_display_company}）",
                        format="%,.2f" if pie_data_type == '成交额' else "%,.0f"
                    ),
                    "公司内部占比（%）": st.column_config.NumberColumn(
                        "公司内部占比（%）",
                        format="%.2f%%"
                    ),
                    "公司占市场比重（%）": st.column_config.NumberColumn(
                        "公司占市场比重（%）",
                        format="%.4f%%"
                    )
                }
            )
    else:
        st.warning("该月份公司数据为空，请选择其他月份")
else:
    if company_market_df.empty:
        st.info("📊 当前数据中没有期货品种（产品类型='期货'），跳过期货品种分析")
    else:
        st.warning(f"选中的月份 {selected_pie_month} 不存在于公司数据中")

# ============================================================
# 期权品种分析-公司
# ============================================================
st.subheader("📊 期权品种分析-公司")

# 根据饼图选择的数据类型加载对应的市场数据和公司数据
if pie_data_type == '成交量':
    option_market_df = df_vol_market.copy()
    option_company_df = df_vol_company.copy()
    option_unit = '手'
    option_title = '成交量'
elif pie_data_type == '成交额':
    option_market_df = df_amt_market.copy()
    option_company_df = df_amt_company.copy()
    option_unit = '元'
    option_title = '成交额'
else:  # 持仓量
    option_market_df = df_oi_market.copy()
    option_company_df = df_oi_company.copy()
    option_unit = '手'
    option_title = '持仓量'

# 筛选期权品种（产品类型 == '期货期权' 或 '现货期权'）
if '产品类型' in option_market_df.columns:
    option_market_df = option_market_df[option_market_df['产品类型'].isin(['期货期权', '现货期权'])]
    option_company_df = option_company_df[option_company_df['产品类型'].isin(['期货期权', '现货期权'])]

# 清理数据
option_market_df = clean_dataframe(option_market_df)
option_company_df = clean_dataframe(option_company_df)

if not option_market_df.empty and selected_pie_month in option_market_df.columns and selected_pie_month in option_company_df.columns:
    option_market_grouped = option_market_df.groupby('产品名称')[selected_pie_month].sum().reset_index()
    option_market_grouped.columns = ['品种', '市场数值']
    
    option_company_grouped = option_company_df.groupby('产品名称')[selected_pie_month].sum().reset_index()
    option_company_grouped.columns = ['品种', '公司数值']
    
    option_merged = pd.merge(option_market_grouped, option_company_grouped, on='品种', how='outer').fillna(0)
    
    # 成交额处理：公司数值从元转亿元
    if pie_data_type == '成交额':
        option_merged['公司数值'] = option_merged['公司数值'] / 100000000
    
    # 计算公司占市场比重（用于右侧柱状图）
    option_merged['公司占市场比重（%）'] = (option_merged['公司数值'] / (option_merged['市场数值'] * 2) * 100).round(4)
    option_merged['公司占市场比重（%）'] = option_merged.apply(
        lambda row: 0 if row['市场数值'] == 0 else row['公司占市场比重（%）'], 
        axis=1
    )
    
    # ★★★ 按公司数值降序排列（饼图按公司份额排序）★★★
    option_merged_sorted = option_merged.sort_values('公司数值', ascending=False).reset_index(drop=True)
    
    if len(option_merged_sorted) > 10:
        top_10 = option_merged_sorted.head(10).copy()
        other_market = option_merged_sorted.iloc[10:]['市场数值'].sum()
        other_company = option_merged_sorted.iloc[10:]['公司数值'].sum()
        other_ratio = (other_company / (other_market * 2) * 100).round(4) if other_market != 0 else 0
        
        other_row = pd.DataFrame({
            '品种': ['其他'],
            '市场数值': [other_market],
            '公司数值': [other_company],
            '公司占市场比重（%）': [other_ratio]
        })
        option_merged_top = pd.concat([top_10, other_row], ignore_index=True)
    else:
        option_merged_top = option_merged_sorted.copy()
    
    total_market_option = option_merged_top['市场数值'].sum()
    
    if total_market_option > 0:
        # 计算公司内部占比（饼图用）
        total_company = option_merged_top['公司数值'].sum()
        option_merged_top['公司内部占比（%）'] = (option_merged_top['公司数值'] / total_company * 100).round(2) if total_company > 0 else 0
        
        # ===== 饼图和柱状图左右并列 =====
        col_left_option, col_right_option = st.columns([1, 1], gap="medium")
        
        # 左列：饼图（公司各品种份额）
        with col_left_option:
            pie_data_option = option_merged_top[['品种', '公司数值', '公司内部占比（%）']].copy()
            fig_pie_option = px.pie(
                pie_data_option,
                values='公司数值',
                names='品种',
                title=f'公司期权{option_title}份额（TOP10）',
                hover_data={'公司数值': True, '公司内部占比（%）': True},
                labels={'公司数值': f'{option_title}（{option_unit}）', '公司内部占比（%）': '占比（%）'},
                hole=0.3
            )
            fig_pie_option.update_layout(
                title_font=dict(size=16, color='#1A5276'),
                font=dict(size=12),
                legend=dict(orientation='v', yanchor='middle', y=0.5, xanchor='right', x=1.0),
                height=450
            )
            fig_pie_option.update_traces(
                textinfo='percent+label',
                texttemplate='%{label}<br>%{percent:.2%}',
                textfont=dict(size=10),
                marker=dict(line=dict(color='white', width=2)),
                pull=[0.03 if i == 0 else 0 for i in range(len(pie_data_option))]
            )
            st.plotly_chart(fig_pie_option, use_container_width=True, config={'displayModeBar': False})
        
        # 右列：柱状图（公司占市场比重）
        with col_right_option:
            bar_data_option = option_merged_top[['品种', '公司占市场比重（%）']].copy()
            bar_data_sorted_option = bar_data_option.sort_values('公司占市场比重（%）', ascending=False)
            max_val_option = bar_data_sorted_option['公司占市场比重（%）'].max()
            
            fig_bar_option = px.bar(
                bar_data_sorted_option,
                x='公司占市场比重（%）',
                y='品种',
                orientation='h',
                title=f'公司期权{option_title}占市场比重（TOP10）',
                labels={'公司占市场比重（%）': '公司占市场比重（%）', '品种': ''},
                text='公司占市场比重（%）',
                color='公司占市场比重（%）',
                color_continuous_scale='Blues',
                range_color=[0, max_val_option * 1.2 if max_val_option > 0 else 1]
            )
            fig_bar_option.update_layout(
                title_font=dict(size=16, color='#1A5276'),
                font=dict(size=12),
                xaxis=dict(
                    title='公司占市场比重（%）',
                    tickformat='.4f',
                    range=[0, max_val_option * 1.3 if max_val_option > 0 else 1]
                ),
                yaxis=dict(title='', categoryorder='total ascending'),
                height=450,
                showlegend=False,
                coloraxis_showscale=False,
                plot_bgcolor='#F8F9F9',
                paper_bgcolor='white'
            )
            fig_bar_option.update_traces(
                texttemplate='%{x:.4f}%',
                textposition='outside',
                textfont=dict(size=11, color='black'),
                marker=dict(line=dict(color='white', width=1))
            )
            st.plotly_chart(fig_bar_option, use_container_width=True, config={'displayModeBar': False})
        
        # 数据表格
        with st.expander("📋 查看详细数据"):
            display_df_option = option_merged_top.copy()
            # 确保表格按公司数值降序排列
            display_df_option = display_df_option.sort_values('公司数值', ascending=False)
            
            if pie_data_type == '成交量':
                market_unit_display_option = '手'
                company_unit_display_option = '手'
            elif pie_data_type == '成交额':
                market_unit_display_option = '亿元'
                company_unit_display_option = '亿元'
            else:
                market_unit_display_option = '手'
                company_unit_display_option = '手'
            
            display_cols_option = ['品种', '市场数值', '公司数值', '公司内部占比（%）', '公司占市场比重（%）']
            display_df_option_final = display_df_option[display_cols_option].copy()
            display_df_option_final.columns = ['品种', f'市场数值（{market_unit_display_option}）', f'公司数值（{company_unit_display_option}）', '公司内部占比（%）', '公司占市场比重（%）']
            
            st.dataframe(
                display_df_option_final,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "品种": st.column_config.TextColumn("品种"),
                    f"市场数值（{market_unit_display_option}）": st.column_config.NumberColumn(
                        f"市场数值（{market_unit_display_option}）",
                        format="%,.2f" if pie_data_type == '成交额' else "%,.0f"
                    ),
                    f"公司数值（{company_unit_display_option}）": st.column_config.NumberColumn(
                        f"公司数值（{company_unit_display_option}）",
                        format="%,.2f" if pie_data_type == '成交额' else "%,.0f"
                    ),
                    "公司内部占比（%）": st.column_config.NumberColumn(
                        "公司内部占比（%）",
                        format="%.2f%%"
                    ),
                    "公司占市场比重（%）": st.column_config.NumberColumn(
                        "公司占市场比重（%）",
                        format="%.4f%%"
                    )
                }
            )
    else:
        st.warning("该月份期权数据为空，请选择其他月份")
else:
    if option_market_df.empty:
        st.info("📊 当前数据中没有期权品种（产品类型='期货期权'或'现货期权'），跳过期权品种分析")
    else:
        st.warning(f"选中的月份 {selected_pie_month} 不存在于期权数据中")

# ============================================================
# 资金统计
# ============================================================
if not df_fund_current.empty:
    st.subheader("📊 运营总体情况")
    try:
        fund_month_col = safe_get_column(df_fund_current, ['月份'], 0)
        
        col_mapping = {
            '入金': safe_get_column(df_fund_current, ['入金', '入金金额'], 3),
            '出金': safe_get_column(df_fund_current, ['出金', '出金金额'], 4),
            '留存手续费': safe_get_column(df_fund_current, ['留存手续费', '手续费', '手续费留存'], 5),
            '期末权益': safe_get_column(df_fund_current, ['期末权益', '权益'], 7),
        }

        if col_mapping.get('期末权益') is None:
            st.warning("未找到期末权益列，请检查数据格式")
        else:
            all_months = []
            if fund_month_col:
                for m in df_fund_current[fund_month_col].dropna().unique():
                    try:
                        m_str = str(int(m)) if isinstance(m, (int, float)) else str(m)
                        if len(m_str) == 6 and m_str.isdigit():
                            all_months.append(m_str)
                    except:
                        pass
            all_months = sorted(set(all_months), reverse=True)

            selected_month = st.selectbox("选择月份", options=all_months,
                                          format_func=lambda x: f"{x[:4]}年{x[4:6]}月",
                                          key="fund_month_selector") if all_months else None

            fund_data = []
            
            if fund_month_col:
                for month_val in df_fund_current[fund_month_col].dropna().unique():
                    try:
                        month_str = str(int(month_val)) if isinstance(month_val, (int, float)) else str(month_val)
                        if len(month_str) != 6 or not month_str.isdigit():
                            continue
                        mask = df_fund_current[fund_month_col] == month_val
                        equity = df_fund_current.loc[mask, col_mapping['期末权益']].sum()
                        if pd.isna(equity) or equity == 0:
                            continue
                        net_deposit = 0
                        if col_mapping.get('入金') and col_mapping.get('出金'):
                            net_deposit = df_fund_current.loc[mask, col_mapping['入金']].sum() - df_fund_current.loc[mask, col_mapping['出金']].sum()
                        fee = df_fund_current.loc[mask, col_mapping['留存手续费']].sum() if col_mapping.get('留存手续费') else 0
                        
                        pnl = 0
                        if not df_trade_stats.empty:
                            trade_mask = df_trade_stats['月份'] == month_val
                            trade_month_df = df_trade_stats[trade_mask]
                            if not trade_month_df.empty:
                                pnl_col = safe_get_column(df_trade_stats, ['平仓盈亏'])
                                if pnl_col:
                                    pnl = trade_month_df[pnl_col].sum()
                                
                                option_income_col = safe_get_column(df_trade_stats, ['期权权利金收入', '权利金收入'])
                                option_expense_col = safe_get_column(df_trade_stats, ['期权权利金支出', '权利金支出'])
                                
                                if option_income_col and option_expense_col:
                                    pnl = pnl + trade_month_df[option_income_col].sum() - trade_month_df[option_expense_col].sum()
                        
                        fund_data.append({
                            '月份': month_str, 
                            '期末权益': equity, 
                            '净入金': net_deposit,
                            '留存手续费': fee, 
                            '平仓盈亏': pnl, 
                            '类型': '今年'
                        })
                    except Exception:
                        continue
            
            if not df_fund_last_year.empty:
                last_month_col = safe_get_column(df_fund_last_year, ['月份'])
                if last_month_col:
                    for month_val in df_fund_last_year[last_month_col].dropna().unique():
                        try:
                            month_str = str(int(month_val)) if isinstance(month_val, (int, float)) else str(month_val)
                            if len(month_str) != 6 or not month_str.isdigit():
                                continue
                            mask = df_fund_last_year[last_month_col] == month_val
                            last_equity_col = safe_get_column(df_fund_last_year, ['期末权益', '权益'])
                            if last_equity_col:
                                equity = df_fund_last_year.loc[mask, last_equity_col].sum()
                            else:
                                equity = 0
                            if pd.isna(equity) or equity == 0:
                                continue
                            net_deposit = 0
                            last_in_col = safe_get_column(df_fund_last_year, ['入金', '入金金额'])
                            last_out_col = safe_get_column(df_fund_last_year, ['出金', '出金金额'])
                            if last_in_col and last_out_col:
                                net_deposit = df_fund_last_year.loc[mask, last_in_col].sum() - df_fund_last_year.loc[mask, last_out_col].sum()
                            last_fee_col = safe_get_column(df_fund_last_year, ['留存手续费', '手续费', '手续费留存'])
                            fee = df_fund_last_year.loc[mask, last_fee_col].sum() if last_fee_col else 0
                            
                            pnl = 0
                            if not df_trade_last.empty:
                                last_trade_mask = df_trade_last['月份'] == month_val
                                last_trade_month_df = df_trade_last[last_trade_mask]
                                if not last_trade_month_df.empty:
                                    last_pnl_col = safe_get_column(df_trade_last, ['平仓盈亏'])
                                    if last_pnl_col:
                                        pnl = last_trade_month_df[last_pnl_col].sum()
                                    
                                    last_option_income_col = safe_get_column(df_trade_last, ['期权权利金收入', '权利金收入'])
                                    last_option_expense_col = safe_get_column(df_trade_last, ['期权权利金支出', '权利金支出'])
                                    
                                    if last_option_income_col and last_option_expense_col:
                                        pnl = pnl + last_trade_month_df[last_option_income_col].sum() - last_trade_month_df[last_option_expense_col].sum()
                            
                            fund_data.append({
                                '月份': month_str, 
                                '期末权益': equity, 
                                '净入金': net_deposit,
                                '留存手续费': fee, 
                                '平仓盈亏': pnl, 
                                '类型': '去年'
                            })
                        except Exception:
                            continue

            if fund_data:
                fund_df = pd.DataFrame(fund_data).sort_values('月份')
                fund_df['月份显示'] = fund_df['月份'].str[4:6]

                display_df = fund_df[['月份', '期末权益', '净入金', '留存手续费', '平仓盈亏']].copy()
                display_df['期末权益'] = (display_df['期末权益'] / 100000000).round(2)
                display_df['净入金'] = (display_df['净入金'] / 10000000).round(2)
                display_df['留存手续费'] = (display_df['留存手续费'] / 100000).round(2)
                display_df['平仓盈亏'] = (display_df['平仓盈亏'] / 1000000).round(2)
                display_df.columns = ['月份', '期末权益（亿元）', '净入金（千万）', '留存手续费（十万）', '平仓盈亏（百万）']
                st.dataframe(display_df.sort_values('月份', ascending=False), use_container_width=True, hide_index=True)

                fund_df_sorted = fund_df.sort_values('月份')
                fund_df_sorted['期末权益（亿元）'] = fund_df_sorted['期末权益'] / 100000000
                fund_df_sorted['净入金（千万）'] = fund_df_sorted['净入金'] / 10000000
                fund_df_sorted['留存手续费（十万）'] = fund_df_sorted['留存手续费'] / 100000
                fund_df_sorted['平仓盈亏（百万）'] = fund_df_sorted['平仓盈亏'] / 1000000

                color_map_fund = {'今年': '#2E86C1', '去年': '#F39C12'}
                fund_df_sorted['类型标签'] = fund_df_sorted['类型'].map({'今年': '今年', '去年': '去年'})
                color_map_fund_label = {'今年': '#2E86C1', '去年': '#F39C12'}

                target_data = fund_df[fund_df['类型'] == '今年']
                if selected_month:
                    target_data = target_data[target_data['月份'] <= selected_month]
                cumsum_data = {
                    '期末权益': target_data['期末权益'].sum() / 100000000 if not target_data.empty else 0,
                    '净入金': target_data['净入金'].sum() / 10000000 if not target_data.empty else 0,
                    '留存手续费': target_data['留存手续费'].sum() / 100000 if not target_data.empty else 0,
                    '平仓盈亏': target_data['平仓盈亏'].sum() / 1000000 if not target_data.empty else 0
                }

                current_equity, mom, yoy = None, None, None
                if selected_month:
                    current_data = fund_df[(fund_df['月份'] == selected_month) & (fund_df['类型'] == '今年')]
                    if current_data.empty:
                        current_data = fund_df[(fund_df['月份'] == selected_month) & (fund_df['类型'] == '去年')]
                    if not current_data.empty:
                        current_equity = current_data['期末权益'].iloc[0] / 100000000
                        month_num = int(selected_month[4:6])
                        year_num = int(selected_month[:4])
                        if month_num > 1:
                            prev_month = f"{year_num}{month_num - 1:02d}"
                        else:
                            prev_month = f"{year_num - 1}12"
                        prev_data = fund_df[(fund_df['月份'] == prev_month) & (fund_df['类型'] == '今年')]
                        if not prev_data.empty:
                            mom = safe_division(current_equity - prev_data['期末权益'].iloc[0] / 100000000,
                                                prev_data['期末权益'].iloc[0] / 100000000) * 100
                        last_year = f"{year_num - 1}{selected_month[4:6]}"
                        last_data = fund_df[(fund_df['月份'] == last_year) & (fund_df['类型'] == '去年')]
                        if not last_data.empty:
                            yoy = safe_division(current_equity - last_data['期末权益'].iloc[0] / 100000000,
                                                last_data['期末权益'].iloc[0] / 100000000) * 100

                charts = [
                    ('期末权益（亿元）', '期末权益', '期末权益（亿）', '', ''),
                    ('净入金（千万）', '净入金', '净入金（千万）', '', f"{cumsum_data['净入金']:+.2f}千万"),
                    ('留存手续费（十万）', '留存手续费', '留存手续费（十万）', '', f"{cumsum_data['留存手续费']:.2f}十万"),
                    ('平仓盈亏（百万）', '平仓盈亏', '平仓盈亏（百万）', '', f"{cumsum_data['平仓盈亏']:+.2f}百万")
                ]

                rows = st.columns(2)
                for idx, (col, title, ylabel, annotation, cumsum_text) in enumerate(charts):
                    with rows[idx % 2]:
                        fig = create_line_chart(fund_df_sorted, '月份显示', col, '类型标签',
                                                title, '月份', ylabel, color_map_fund_label, '.2f')
                        if fig:
                            if annotation:
                                fig.add_annotation(x=0.98, y=0.98, xref='paper', yref='paper',
                                                   text=annotation, showarrow=False, font=dict(size=10, color='#1A5276'),
                                                   bgcolor='rgba(255,255,255,0.85)', bordercolor='#1A5276',
                                                   borderwidth=1, borderpad=4, xanchor='right', yanchor='top')
                            if cumsum_text:
                                fig.add_annotation(x=0.98, y=0.88 if annotation else 0.98, xref='paper', yref='paper',
                                                   text=f'累计: {cumsum_text}', showarrow=False,
                                                   font=dict(size=11, color='#1A5276'),
                                                   bgcolor='rgba(255,255,255,0.85)', bordercolor='#1A5276',
                                                   borderwidth=1, borderpad=4, xanchor='right', yanchor='top')
                            st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

                st.subheader("📊 客户情况统计")
                try:
                    trade_customer_data = []
                    for df_trade, year_type in [(df_trade_stats, '今年'), (df_trade_last, '去年')]:
                        if df_trade.empty:
                            continue
                        
                        if '月份' not in df_trade.columns or '投资者代码' not in df_trade.columns:
                            continue
                        
                        for m in df_trade['月份'].dropna().unique():
                            try:
                                m_str = str(int(m)) if isinstance(m, (int, float)) else str(m)
                                if len(m_str) != 6 or not m_str.isdigit():
                                    continue
                                
                                mask = df_trade['月份'] == m
                                month_df = df_trade[mask]
                                if month_df.empty:
                                    continue
                                
                                total_investors = month_df['投资者代码'].nunique()
                                
                                pnl_col = safe_get_column(df_trade, ['平仓盈亏'])
                                option_income_col = safe_get_column(df_trade, ['期权权利金收入', '权利金收入'])
                                option_expense_col = safe_get_column(df_trade, ['期权权利金支出', '权利金支出'])
                                
                                if pnl_col and option_income_col and option_expense_col:
                                    month_df['盈亏'] = month_df[pnl_col] + month_df[option_income_col] - month_df[option_expense_col]
                                    profit_investors = month_df[month_df['盈亏'] > 0]['投资者代码'].nunique()
                                else:
                                    profit_investors = 0
                                
                                trade_customer_data.append({
                                    '月份': m_str, 
                                    '交易客户数': total_investors,
                                    '盈利客户数': profit_investors, 
                                    '类型': year_type
                                })
                            except Exception:
                                continue

                    if trade_customer_data:
                        customer_df = pd.DataFrame(trade_customer_data).sort_values('月份')
                        customer_df['月份显示'] = customer_df['月份'].astype(str).str[4:6]
                        customer_df['类型标签'] = customer_df['类型'].map({'今年': '今年', '去年': '去年'})

                        col_left, col_right = st.columns(2)
                        with col_left:
                            fig = create_line_chart(customer_df, '月份显示', '盈利客户数', '类型标签',
                                                    '盈利客户数（当月）', '月份', '客户数', color_map_fund_label, '.0f')
                            if fig:
                                fig.update_layout(legend_title_text='')
                                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
                        with col_right:
                            fig = create_line_chart(customer_df, '月份显示', '交易客户数', '类型标签',
                                                    '交易客户数（当月）', '月份', '客户数', color_map_fund_label, '.0f')
                            if fig:
                                fig.update_layout(legend_title_text='')
                                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
                except Exception as e:
                    st.warning(f"加载客户数据时出错: {e}")
    except Exception as e:
        st.warning(f"加载资金对账表时出错: {e}")

# ============================================================
# 部门交易客户数统计
# ============================================================
if not df_trade_stats.empty:
    st.subheader("📊 部门交易客户数统计")
    try:
        if '月份' not in df_trade_stats.columns or '部门' not in df_trade_stats.columns or '投资者代码' not in df_trade_stats.columns:
            st.warning("交易统计表缺少必要列（月份、部门、投资者代码）")
        else:
            all_months = []
            for m in df_trade_stats['月份'].dropna().unique():
                try:
                    m_str = str(int(m)) if isinstance(m, (int, float)) else str(m)
                    if len(m_str) == 6 and m_str.isdigit():
                        all_months.append(m_str)
                except:
                    pass
            all_months = sorted(set(all_months), reverse=True)

            if all_months:
                selected_month = st.selectbox("选择月份", options=all_months,
                                              format_func=lambda x: f"{x[:4]}年{x[4:6]}月",
                                              key="trade_month_selector")

                df_trade_stats['月份_str'] = df_trade_stats['月份'].apply(
                    lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x))
                filtered_trade = df_trade_stats[df_trade_stats['月份_str'] == selected_month]

                if not filtered_trade.empty:
                    result = filtered_trade.groupby('部门').agg(
                        有交易客户数=('投资者代码', 'nunique')
                    ).reset_index()

                    pnl_col = safe_get_column(df_trade_stats, ['平仓盈亏'])
                    option_income_col = safe_get_column(df_trade_stats, ['期权权利金收入', '权利金收入'])
                    option_expense_col = safe_get_column(df_trade_stats, ['期权权利金支出', '权利金支出'])
                    
                    if pnl_col and option_income_col and option_expense_col:
                        filtered_trade['盈亏'] = filtered_trade[pnl_col] + filtered_trade[option_income_col] - filtered_trade[option_expense_col]
                        
                        profit_mask = filtered_trade['盈亏'] > 0
                        profit_df = filtered_trade[profit_mask]
                        profit_count = profit_df.groupby('部门').agg(
                            盈利客户数=('投资者代码', 'nunique')
                        ).reset_index()
                        result = pd.merge(result, profit_count, on='部门', how='left').fillna(0)

                        pnl_df = filtered_trade.groupby('部门').apply(
                            lambda x: (x[pnl_col].sum() + x[option_income_col].sum() - x[option_expense_col].sum())
                        ).reset_index()
                        pnl_df.columns = ['部门', '平仓盈亏']
                        result = pd.merge(result, pnl_df, on='部门', how='left').fillna(0)

                    if not df_fund_current.empty:
                        fund_dept_col = safe_get_column(df_fund_current, ['部门', '部门名称'], 2)
                        if fund_dept_col:
                            e_col = safe_get_column(df_fund_current, ['入金', '入金金额'], 3)
                            f_col_fund = safe_get_column(df_fund_current, ['出金', '出金金额'], 4)
                            g_col_fund = safe_get_column(df_fund_current, ['留存手续费', '手续费', '手续费留存'], 5)
                            i_col = safe_get_column(df_fund_current, ['期末权益', '权益'], 7)
                            
                            fund_month_col = safe_get_column(df_fund_current, ['月份'], 0)
                            if fund_month_col:
                                df_fund_current['月份_str'] = df_fund_current[fund_month_col].apply(
                                    lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x))
                                filtered_fund = df_fund_current[df_fund_current['月份_str'] == selected_month]
                                
                                if not filtered_fund.empty and e_col and f_col_fund:
                                    filtered_fund['净入金'] = filtered_fund[e_col] - filtered_fund[f_col_fund]
                                    fund_by_dept = filtered_fund.groupby(fund_dept_col).agg({
                                        '净入金': 'sum'
                                    }).reset_index()
                                    fund_by_dept.columns = ['部门', '净入金']
                                    
                                    if g_col_fund:
                                        fee_by_dept = filtered_fund.groupby(fund_dept_col)[g_col_fund].sum().reset_index()
                                        fee_by_dept.columns = ['部门', '留存手续费']
                                        fund_by_dept = pd.merge(fund_by_dept, fee_by_dept, on='部门', how='left')
                                    
                                    if i_col:
                                        equity_by_dept = filtered_fund.groupby(fund_dept_col)[i_col].sum().reset_index()
                                        equity_by_dept.columns = ['部门', '期末权益']
                                        fund_by_dept = pd.merge(fund_by_dept, equity_by_dept, on='部门', how='left')
                                    
                                    result = pd.merge(result, fund_by_dept, on='部门', how='left').fillna(0)

                    if '盈利客户数' in result.columns:
                        result['盈利面'] = result.apply(
                            lambda row: safe_division(row['盈利客户数'], row['有交易客户数']) * 100, axis=1
                        )
                    else:
                        result['盈利面'] = 0
                    
                    result = result.sort_values('有交易客户数', ascending=False)

                    for col in ['期末权益', '平仓盈亏', '净入金', '留存手续费']:
                        if col in result.columns:
                            result[col] = result[col].apply(lambda x: f"{int(x):,}")
                    if '盈利面' in result.columns:
                        result['盈利面'] = result['盈利面'].apply(lambda x: f"{x:.2f}%")

                    st.subheader(f"📊 {selected_month[:4]}年{selected_month[4:6]}月 部门统计")
                    st.dataframe(result, use_container_width=True, hide_index=True)

                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("部门总数", len(result))
                    with col2:
                        st.metric("客户总数", int(result['有交易客户数'].sum()))
                else:
                    st.info(f"{selected_month} 无交易数据")
            else:
                st.info("暂无月份数据")
    except Exception as e:
        st.warning(f"加载部门统计时出错: {e}")
        st.exception(e)

# ============================================================
# 每月开户数统计
# ============================================================
if not df_investor.empty:
    st.subheader("📊 每月开户数统计（自2021年起）")
    try:
        investor_col = safe_get_column(df_investor, ['投资者代码', '投资者代码'], 0)
        date_col = safe_get_column(df_investor, ['开户日期'], 10)
        type_col = safe_get_column(df_investor, ['投资者类型', '投资者类型'], 5)
        
        if investor_col is None:
            st.warning("未找到'投资者代码'列（A列），请检查数据格式")
        elif date_col is None:
            st.warning("未找到'开户日期'列（K列），请检查数据格式")
        elif type_col is None:
            st.warning("未找到'投资者类型'列（F列），请检查数据格式")
        else:
            investor_df = df_investor[[investor_col, date_col, type_col]].copy()
            investor_df = investor_df.dropna(subset=[date_col])
            
            def extract_year_month(date_val):
                try:
                    date_str = str(int(date_val)) if isinstance(date_val, (int, float)) else str(date_val)
                    if len(date_str) == 8 and date_str.isdigit():
                        year = int(date_str[:4])
                        month = int(date_str[4:6])
                        if year >= 2021:
                            return year, f"{year}{month:02d}"
                except:
                    pass
                return None, None
            
            investor_df['年份'] = investor_df[date_col].apply(lambda x: extract_year_month(x)[0])
            investor_df['年月'] = investor_df[date_col].apply(lambda x: extract_year_month(x)[1])
            investor_df = investor_df.dropna(subset=['年月'])
            
            if investor_df.empty:
                st.info("自2021年起暂无开户数据")
            else:
                monthly_count = investor_df.groupby(['年份', '年月'])[investor_col].nunique().reset_index()
                monthly_count.columns = ['年份', '年月', '开户数']
                
                investor_unique = investor_df.drop_duplicates(subset=['年份', '年月', investor_col])
                type_count = investor_unique.groupby(['年份', '年月', type_col]).size().reset_index(name='数量')
                
                all_periods = monthly_count[['年份', '年月']].copy()
                
                for type_name in ['自然人', '法人', '特殊法人']:
                    type_data = type_count[type_count[type_col] == type_name][['年份', '年月', '数量']]
                    type_data = type_data.rename(columns={'数量': type_name})
                    all_periods = all_periods.merge(type_data, on=['年份', '年月'], how='left')
                    all_periods[type_name] = all_periods[type_name].fillna(0).astype(int)
                
                monthly_count = monthly_count.merge(all_periods, on=['年份', '年月'], how='left')
                monthly_count = monthly_count.sort_values('年月')
                monthly_count['年月显示'] = monthly_count['年月'].apply(
                    lambda x: f"{x[:4]}年{x[4:6]}月"
                )
                monthly_count['月份数字'] = monthly_count['年月'].apply(lambda x: int(x[4:6]))
                
                all_months = sorted(monthly_count['年月'].unique(), reverse=True)
                all_months_display = [f"{m[:4]}年{m[4:6]}月" for m in all_months]
                
                filter_col1, filter_col2 = st.columns([1, 3])
                with filter_col1:
                    selected_month_filter = st.selectbox(
                        "选择月份",
                        options=['全部'] + all_months_display,
                        key="investor_month_filter"
                    )
                
                if selected_month_filter != '全部':
                    selected_month = selected_month_filter.replace('年', '').replace('月', '')
                    filtered_monthly = monthly_count[monthly_count['年月'] == selected_month]
                else:
                    filtered_monthly = monthly_count
                
                display_cols = ['年月显示', '开户数', '自然人', '法人', '特殊法人']
                st.dataframe(
                    filtered_monthly[display_cols],
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "年月显示": "月份",
                        "开户数": "开户数",
                        "自然人": "自然人",
                        "法人": "法人",
                        "特殊法人": "特殊法人"
                    }
                )
                
                years_sorted = sorted(monthly_count['年份'].unique())
                last_two_years = years_sorted[-2:] if len(years_sorted) >= 2 else years_sorted
                
                plot_df = monthly_count[monthly_count['年份'].isin(last_two_years)].copy()
                
                if not plot_df.empty:
                    current_year = max(last_two_years)
                    last_year = min(last_two_years) if len(last_two_years) >= 2 else current_year
                    
                    current_year_data = plot_df[plot_df['年份'] == current_year]
                    last_year_data = plot_df[plot_df['年份'] == last_year]
                    
                    max_month_current = current_year_data['月份数字'].max() if not current_year_data.empty else 12
                    max_month_last = last_year_data['月份数字'].max() if not last_year_data.empty else 12
                    
                    x_max = max(max_month_current, 12)
                    
                    complete_months = []
                    for year in last_two_years:
                        year_data = plot_df[plot_df['年份'] == year]
                        if year == current_year:
                            year_max_month = max_month_current
                        else:
                            year_max_month = 12
                        
                        for month in range(1, year_max_month + 1):
                            month_data = year_data[year_data['月份数字'] == month]
                            if not month_data.empty:
                                complete_months.append({
                                    '年份': int(year),
                                    '年月': month_data.iloc[0]['年月'],
                                    '开户数': month_data.iloc[0]['开户数'],
                                    '年月显示': month_data.iloc[0]['年月显示'],
                                    '月份数字': month
                                })
                            else:
                                year_str = str(int(year))
                                month_str = f"{month:02d}"
                                complete_months.append({
                                    '年份': int(year),
                                    '年月': f"{year_str}{month_str}",
                                    '开户数': 0,
                                    '年月显示': f"{year_str}年{month_str}月",
                                    '月份数字': month
                                })
                    
                    complete_df = pd.DataFrame(complete_months)
                    complete_df['年份标签'] = complete_df['年份'].apply(lambda x: f"{int(x)}年")
                    complete_df = complete_df.sort_values('月份数字')
                    
                    color_map = {}
                    year_colors = ['#2E86C1', '#F39C12']
                    for i, year in enumerate(last_two_years):
                        color_map[f"{int(year)}年"] = year_colors[i % len(year_colors)]
                    
                    month_labels = [f"{i}月" for i in range(1, x_max + 1)]
                    
                    fig = px.line(
                        complete_df,
                        x='月份数字',
                        y='开户数',
                        color='年份标签',
                        title='每月开户数趋势（近两年对比）',
                        labels={'月份数字': '月份', '开户数': '开户数', '年份标签': '年份'},
                        markers=True,
                        color_discrete_map=color_map
                    )
                    
                    fig.update_layout(
                        xaxis=dict(
                            tickmode='array',
                            tickvals=list(range(1, x_max + 1)),
                            ticktext=month_labels
                        )
                    )
                    
                    fig.update_traces(
                        texttemplate='%{y:.0f}',
                        textposition='top center',
                        mode='lines+markers+text'
                    )
                    fig.update_layout(
                        title_font=dict(size=16, color='#1A5276'),
                        plot_bgcolor='#F8F9F9',
                        paper_bgcolor='white',
                        height=400,
                        legend=dict(
                            orientation='h',
                            yanchor='bottom',
                            y=1.02,
                            xanchor='center',
                            x=0.5
                        )
                    )
                    st.plotly_chart(fig, use_container_width=True, key="investor_trend_chart")
                
                col_pie1, col_pie2 = st.columns(2)
                
                with col_pie1:
                    if selected_month_filter != '全部':
                        current_month = selected_month_filter.replace('年', '').replace('月', '')
                        pie_investors_current = investor_unique[investor_unique['年月'] == current_month]
                        if not pie_investors_current.empty:
                            type_pie_data_current = pie_investors_current.groupby(type_col).size().reset_index(name='数量')
                            all_types = ['自然人', '法人', '特殊法人']
                            for t in all_types:
                                if t not in type_pie_data_current[type_col].values:
                                    type_pie_data_current = pd.concat([type_pie_data_current, pd.DataFrame({type_col: [t], '数量': [0]})], ignore_index=True)
                            
                            fig_pie_current = px.pie(
                                type_pie_data_current,
                                values='数量',
                                names=type_col,
                                title=f'{selected_month_filter} 新开户客户类型分布',
                                color=type_col,
                                color_discrete_map={
                                    '自然人': '#2E86C1',
                                    '法人': '#F39C12',
                                    '特殊法人': '#28B463'
                                }
                            )
                            fig_pie_current.update_traces(textposition='inside', textinfo='percent+label')
                            fig_pie_current.update_layout(
                                title_font=dict(size=14, color='#1A5276'),
                                height=380
                            )
                            st.plotly_chart(fig_pie_current, use_container_width=True, key="investor_pie_current")
                        else:
                            st.info(f"{selected_month_filter} 暂无数据")
                    else:
                        latest_year = max(monthly_count['年份'].unique())
                        latest_year_data = investor_unique[investor_unique['年份'] == latest_year]
                        if not latest_year_data.empty:
                            type_pie_data_latest = latest_year_data.groupby(type_col).size().reset_index(name='数量')
                            all_types = ['自然人', '法人', '特殊法人']
                            for t in all_types:
                                if t not in type_pie_data_latest[type_col].values:
                                    type_pie_data_latest = pd.concat([type_pie_data_latest, pd.DataFrame({type_col: [t], '数量': [0]})], ignore_index=True)
                            
                            fig_pie_latest = px.pie(
                                type_pie_data_latest,
                                values='数量',
                                names=type_col,
                                title=f'{int(latest_year)}年 新开户客户类型分布（汇总）',
                                color=type_col,
                                color_discrete_map={
                                    '自然人': '#2E86C1',
                                    '法人': '#F39C12',
                                    '特殊法人': '#28B463'
                                }
                            )
                            fig_pie_latest.update_traces(textposition='inside', textinfo='percent+label')
                            fig_pie_latest.update_layout(
                                title_font=dict(size=14, color='#1A5276'),
                                height=380
                            )
                            st.plotly_chart(fig_pie_latest, use_container_width=True, key="investor_pie_latest")
                        else:
                            st.info(f"{int(latest_year)}年 暂无数据")
                
                with col_pie2:
                    if selected_month_filter != '全部':
                        current_month = selected_month_filter.replace('年', '').replace('月', '')
                        current_year = int(current_month[:4])
                        current_month_num = int(current_month[4:6])
                        last_year_month = f"{current_year - 1}{current_month_num:02d}"
                        last_year_display = f"{current_year - 1}年{current_month_num:02d}月"
                        
                        pie_investors_last = investor_unique[investor_unique['年月'] == last_year_month]
                        if not pie_investors_last.empty:
                            type_pie_data_last = pie_investors_last.groupby(type_col).size().reset_index(name='数量')
                            all_types = ['自然人', '法人', '特殊法人']
                            for t in all_types:
                                if t not in type_pie_data_last[type_col].values:
                                    type_pie_data_last = pd.concat([type_pie_data_last, pd.DataFrame({type_col: [t], '数量': [0]})], ignore_index=True)
                            
                            fig_pie_last = px.pie(
                                type_pie_data_last,
                                values='数量',
                                names=type_col,
                                title=f'{last_year_display} 新开户客户类型分布（去年同期）',
                                color=type_col,
                                color_discrete_map={
                                    '自然人': '#2E86C1',
                                    '法人': '#F39C12',
                                    '特殊法人': '#28B463'
                                }
                            )
                            fig_pie_last.update_traces(textposition='inside', textinfo='percent+label')
                            fig_pie_last.update_layout(
                                title_font=dict(size=14, color='#1A5276'),
                                height=380
                            )
                            st.plotly_chart(fig_pie_last, use_container_width=True, key="investor_pie_last")
                        else:
                            st.info(f"{last_year_display} 暂无数据")
                    else:
                        if len(years_sorted) >= 2:
                            prev_year = years_sorted[-2]
                            prev_year_data = investor_unique[investor_unique['年份'] == prev_year]
                            if not prev_year_data.empty:
                                type_pie_data_prev = prev_year_data.groupby(type_col).size().reset_index(name='数量')
                                all_types = ['自然人', '法人', '特殊法人']
                                for t in all_types:
                                    if t not in type_pie_data_prev[type_col].values:
                                        type_pie_data_prev = pd.concat([type_pie_data_prev, pd.DataFrame({type_col: [t], '数量': [0]})], ignore_index=True)
                                
                                fig_pie_prev = px.pie(
                                    type_pie_data_prev,
                                    values='数量',
                                    names=type_col,
                                    title=f'{int(prev_year)}年 新开户客户类型分布（去年同期）',
                                    color=type_col,
                                    color_discrete_map={
                                        '自然人': '#2E86C1',
                                        '法人': '#F39C12',
                                        '特殊法人': '#28B463'
                                    }
                                )
                                fig_pie_prev.update_traces(textposition='inside', textinfo='percent+label')
                                fig_pie_prev.update_layout(
                                    title_font=dict(size=14, color='#1A5276'),
                                    height=380
                                )
                                st.plotly_chart(fig_pie_prev, use_container_width=True, key="investor_pie_prev")
                            else:
                                st.info(f"{int(prev_year)}年 暂无数据")
                        else:
                            st.info("暂无去年数据")
                                            
    except Exception as e:
        st.warning(f"加载开户数据时出错: {e}")
        st.exception(e)

# ============================================================
# 活跃客户统计
# ============================================================
if '活跃客户' in data2_cache and not data2_cache['活跃客户'].empty:
    st.subheader("📊 公司客户资金情况统计")
    try:
        df_active = data2_cache['活跃客户'].copy()
        df_active = clean_dataframe(df_active)
        
        month_col = safe_get_column(df_active, ['月份', '月份'], 0)
        investor_col = safe_get_column(df_active, ['投资者代码', '投资者代码'], 1)
        
        if month_col is None:
            st.warning("未找到'月份'列（A列），请检查数据格式")
        elif investor_col is None:
            st.warning("未找到'投资者代码'列（B列），请检查数据格式")
        else:
            active_df = df_active[[month_col, investor_col]].copy()
            active_df = active_df.dropna(subset=[month_col, investor_col])
            
            def format_month(val):
                try:
                    val_str = str(int(val)) if isinstance(val, (int, float)) else str(val)
                    if len(val_str) == 6 and val_str.isdigit():
                        return val_str
                except:
                    pass
                return None
            
            active_df['年月'] = active_df[month_col].apply(format_month)
            active_df = active_df.dropna(subset=['年月'])
            
            if active_df.empty:
                st.info("暂无活跃客户数据")
            else:
                monthly_active = active_df.groupby('年月')[investor_col].nunique().reset_index()
                monthly_active.columns = ['年月', '活跃客户数']
                monthly_active = monthly_active.sort_values('年月')
                monthly_active['年月显示'] = monthly_active['年月'].apply(
                    lambda x: f"{x[:4]}年{x[4:6]}月"
                )
                
                df_fund = data1_cache.get('资金对账表-月', pd.DataFrame())
                df_fund = clean_dataframe(df_fund)
                
                df_trade = data2_cache.get('交易统计表-月', pd.DataFrame())
                df_trade = clean_dataframe(df_trade)
                df_trade = normalize_trade_columns(df_trade)
                
                fund_month_col = safe_get_column(df_fund, ['月份', '月份'], 0)
                fund_investor_col = safe_get_column(df_fund, ['投资者代码', '投资者代码', '客户代码', '投资者'], 2)
                fund_equity_col = safe_get_column(df_fund, ['期末权益', '期末权益', '权益'], 7)
                fund_fee_col = safe_get_column(df_fund, ['留存手续费', '留存手续费', '手续费', '手续费留存'], 5)
                fund_inflow_col = safe_get_column(df_fund, ['入金', '入金', '入金金额'], 3)
                fund_outflow_col = safe_get_column(df_fund, ['出金', '出金', '出金金额'], 4)
                
                trade_month_col = safe_get_column(df_trade, ['月份', '月份'], 0)
                trade_investor_col = safe_get_column(df_trade, ['投资者代码', '投资者代码', '客户代码', '投资者'], 3)
                trade_pnl_col = safe_get_column(df_trade, ['平仓盈亏', '平仓盈亏'], 4)
                trade_option_income_col = safe_get_column(df_trade, ['期权权利金收入', '期权权利金收入', '权利金收入'], 5)
                trade_option_expense_col = safe_get_column(df_trade, ['期权权利金支出', '期权权利金支出', '权利金支出'], 6)
                
                all_months = monthly_active['年月显示'].tolist()
                
                filter_col1, filter_col2 = st.columns([1, 3])
                with filter_col1:
                    selected_month_filter = st.selectbox(
                        "选择月份",
                        options=['全部'] + all_months,
                        key="active_month_filter"
                    )
                
                if selected_month_filter != '全部':
                    selected_month_raw = selected_month_filter.replace('年', '').replace('月', '')
                    filtered_active = monthly_active[monthly_active['年月显示'] == selected_month_filter]
                    display_month = selected_month_filter
                else:
                    filtered_active = monthly_active
                    display_month = "全部月份"
                
                active_count = filtered_active['活跃客户数'].sum() if not filtered_active.empty else 0
                
                fund_available = not df_fund.empty and fund_equity_col is not None and fund_investor_col is not None
                
                if not fund_available:
                    st.warning("未找到资金对账表或期末权益列，无法计算资金区间分布")
                    labels = ['2万以下', '2万（含）-10万', '10万（含）-50万', '50万（含）-100万', 
                              '100万（含）-300万', '300万（含）-1000万', '1000万（含）-3000万', '3000万（含）以上']
                    table_data = {
                        '客户类型': ['合计', '占比（%）', '期末权益（亿）', '权益占比（%）', '留存手续费（十万）', '手续费占比（%）', '平仓盈亏（百万）', '净出入金（百万）'],
                    }
                    for label in labels:
                        table_data[label] = [0, '0.00%', '0.00', '0.00%', '0.00', '0.00%', '0.00', '0.00']
                    table_data['总计'] = [0, '100.00%', '0.00', '100.00%', '0.00', '100.00%', '0.00', '0.00']
                    df_table = pd.DataFrame(table_data)
                    
                    st.dataframe(
                        df_table,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "客户类型": st.column_config.TextColumn("客户类型", width="small"),
                            "2万以下": st.column_config.TextColumn("2万以下", width="small"),
                            "2万（含）-10万": st.column_config.TextColumn("2万（含）-10万", width="small"),
                            "10万（含）-50万": st.column_config.TextColumn("10万（含）-50万", width="small"),
                            "50万（含）-100万": st.column_config.TextColumn("50万（含）-100万", width="small"),
                            "100万（含）-300万": st.column_config.TextColumn("100万（含）-300万", width="small"),
                            "300万（含）-1000万": st.column_config.TextColumn("300万（含）-1000万", width="small"),
                            "1000万（含）-3000万": st.column_config.TextColumn("1000万（含）-3000万", width="small"),
                            "3000万（含）以上": st.column_config.TextColumn("3000万（含）以上", width="small"),
                            "总计": st.column_config.TextColumn("总计", width="small")
                        }
                    )
                else:
                    def format_fund_month(val):
                        try:
                            val_str = str(int(val)) if isinstance(val, (int, float)) else str(val)
                            if len(val_str) == 6 and val_str.isdigit():
                                return val_str
                        except:
                            pass
                        return None
                    
                    df_fund['年月'] = df_fund[fund_month_col].apply(format_fund_month) if fund_month_col is not None else None
                    if fund_month_col is not None:
                        df_fund = df_fund.dropna(subset=['年月', fund_equity_col, fund_investor_col])
                    
                    bins = [0, 20000, 100000, 500000, 1000000, 3000000, 10000000, 30000000, float('inf')]
                    labels = ['2万以下', '2万（含）-10万', '10万（含）-50万', '50万（含）-100万', 
                              '100万（含）-300万', '300万（含）-1000万', '1000万（含）-3000万', '3000万（含）以上']
                    
                    if selected_month_filter != '全部':
                        filtered_fund = df_fund[df_fund['年月'] == selected_month_raw]
                    else:
                        filtered_fund = df_fund
                    
                    if selected_month_filter != '全部':
                        active_investors = active_df[active_df['年月'] == selected_month_raw][investor_col].unique()
                    else:
                        active_investors = active_df[investor_col].unique()
                    
                    pnl_by_investor = {}
                    if not df_trade.empty and trade_month_col is not None and trade_investor_col is not None:
                        df_trade['年月'] = df_trade[trade_month_col].apply(format_fund_month)
                        df_trade = df_trade.dropna(subset=['年月', trade_investor_col])
                        
                        if selected_month_filter != '全部':
                            filtered_trade = df_trade[df_trade['年月'] == selected_month_raw]
                        else:
                            filtered_trade = df_trade
                        
                        filtered_trade = filtered_trade[filtered_trade[trade_investor_col].isin(active_investors)]
                        
                        if trade_pnl_col is not None and not filtered_trade.empty:
                            filtered_trade['平仓盈亏计算'] = 0
                            filtered_trade['平仓盈亏计算'] += filtered_trade[trade_pnl_col].fillna(0)
                            if trade_option_income_col is not None:
                                filtered_trade['平仓盈亏计算'] += filtered_trade[trade_option_income_col].fillna(0)
                            if trade_option_expense_col is not None:
                                filtered_trade['平仓盈亏计算'] -= filtered_trade[trade_option_expense_col].fillna(0)
                            
                            trade_pnl_summary = filtered_trade.groupby(trade_investor_col)['平仓盈亏计算'].sum().reset_index()
                            trade_pnl_summary.columns = [fund_investor_col, '平仓盈亏']
                            pnl_by_investor = dict(zip(trade_pnl_summary[fund_investor_col], trade_pnl_summary['平仓盈亏']))
                    
                    filtered_fund = filtered_fund[filtered_fund[fund_investor_col].isin(active_investors)]
                    
                    if not filtered_fund.empty:
                        agg_dict = {
                            fund_equity_col: 'max',
                        }
                        if fund_fee_col is not None:
                            agg_dict[fund_fee_col] = 'sum'
                        if fund_inflow_col is not None:
                            agg_dict[fund_inflow_col] = 'sum'
                        if fund_outflow_col is not None:
                            agg_dict[fund_outflow_col] = 'sum'
                        
                        fund_grouped = filtered_fund.groupby([fund_investor_col]).agg(agg_dict).reset_index()
                        
                        if fund_inflow_col is not None and fund_outflow_col is not None:
                            fund_grouped['净出入金'] = fund_grouped[fund_inflow_col] - fund_grouped[fund_outflow_col]
                        else:
                            fund_grouped['净出入金'] = 0
                        
                        fund_grouped['平仓盈亏'] = fund_grouped[fund_investor_col].map(pnl_by_investor).fillna(0)
                        
                        fund_grouped['资金区间'] = pd.cut(
                            fund_grouped[fund_equity_col], 
                            bins=bins, 
                            labels=labels, 
                            right=False
                        )
                        
                        agg_interval = {
                            '客户数': ('资金区间', 'size'),
                            '权益总和': (fund_equity_col, 'sum'),
                            '净出入金总和': ('净出入金', 'sum'),
                            '平仓盈亏总和': ('平仓盈亏', 'sum')
                        }
                        if fund_fee_col is not None:
                            agg_interval['手续费总和'] = (fund_fee_col, 'sum')
                        
                        interval_stats = fund_grouped.groupby('资金区间').agg(**agg_interval).reset_index()
                        
                        if fund_fee_col is None:
                            interval_stats['手续费总和'] = 0
                        
                        table_data = {'客户类型': ['合计', '占比（%）', '期末权益（亿）', '权益占比（%）', '留存手续费（十万）', '手续费占比（%）', '平仓盈亏（百万）', '净出入金（百万）']}
                        
                        interval_counts = {}
                        interval_equity = {}
                        interval_fee = {}
                        interval_pnl = {}
                        interval_netflow = {}
                        total_all = 0
                        total_equity = 0
                        total_fee = 0
                        total_pnl = 0
                        total_netflow = 0
                        
                        for label in labels:
                            row = interval_stats[interval_stats['资金区间'] == label]
                            if not row.empty:
                                count = int(row['客户数'].iloc[0])
                                equity = float(row['权益总和'].iloc[0])
                                fee = float(row['手续费总和'].iloc[0]) if fund_fee_col is not None else 0
                                pnl = float(row['平仓盈亏总和'].iloc[0])
                                netflow = float(row['净出入金总和'].iloc[0])
                            else:
                                count = 0
                                equity = 0
                                fee = 0
                                pnl = 0
                                netflow = 0
                            interval_counts[label] = count
                            interval_equity[label] = equity
                            interval_fee[label] = fee
                            interval_pnl[label] = pnl
                            interval_netflow[label] = netflow
                            total_all += count
                            total_equity += equity
                            total_fee += fee
                            total_pnl += pnl
                            total_netflow += netflow
                        
                        for label in labels:
                            equity_billion = interval_equity[label] / 100000000
                            fee_ten_thousand = interval_fee[label] / 100000
                            pnl_million = interval_pnl[label] / 1000000
                            netflow_million = interval_netflow[label] / 1000000
                            table_data[label] = [
                                interval_counts[label],
                                f"{interval_counts[label]/total_all*100:.2f}%" if total_all > 0 else "0.00%",
                                f"{equity_billion:.2f}",
                                f"{interval_equity[label]/total_equity*100:.2f}%" if total_equity > 0 else "0.00%",
                                f"{fee_ten_thousand:.2f}",
                                f"{interval_fee[label]/total_fee*100:.2f}%" if total_fee > 0 else "0.00%",
                                f"{pnl_million:.2f}",
                                f"{netflow_million:.2f}"
                            ]
                        table_data['总计'] = [
                            total_all,
                            "100.00%",
                            f"{total_equity/100000000:.2f}",
                            "100.00%",
                            f"{total_fee/100000:.2f}",
                            "100.00%",
                            f"{total_pnl/1000000:.2f}",
                            f"{total_netflow/1000000:.2f}"
                        ]
                        
                        df_table = pd.DataFrame(table_data)
                    else:
                        table_data = {
                            '客户类型': ['合计', '占比（%）', '期末权益（亿）', '权益占比（%）', '留存手续费（十万）', '手续费占比（%）', '平仓盈亏（百万）', '净出入金（百万）'],
                        }
                        for label in labels:
                            table_data[label] = [0, "0.00%", "0.00", "0.00%", "0.00", "0.00%", "0.00", "0.00"]
                        table_data['总计'] = [0, "100.00%", "0.00", "100.00%", "0.00", "100.00%", "0.00", "0.00"]
                        df_table = pd.DataFrame(table_data)
                    
                    st.dataframe(
                        df_table,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            "客户类型": st.column_config.TextColumn("客户类型", width="small"),
                            "2万以下": st.column_config.TextColumn("2万以下", width="small"),
                            "2万（含）-10万": st.column_config.TextColumn("2万（含）-10万", width="small"),
                            "10万（含）-50万": st.column_config.TextColumn("10万（含）-50万", width="small"),
                            "50万（含）-100万": st.column_config.TextColumn("50万（含）-100万", width="small"),
                            "100万（含）-300万": st.column_config.TextColumn("100万（含）-300万", width="small"),
                            "300万（含）-1000万": st.column_config.TextColumn("300万（含）-1000万", width="small"),
                            "1000万（含）-3000万": st.column_config.TextColumn("1000万（含）-3000万", width="small"),
                            "3000万（含）以上": st.column_config.TextColumn("3000万（含）以上", width="small"),
                            "总计": st.column_config.TextColumn("总计", width="small")
                        }
                    )
                
                st.caption(f"📅 当前显示: {display_month} | 活跃客户数: {active_count:,} 户")
                                                                                    
    except Exception as e:
        st.warning(f"加载活跃客户数据时出错: {e}")
        st.exception(e)

# ============================================================
# 市场权益 vs 公司权益对比（含中位数）
# ============================================================
if '市场权益' in data2_cache and not data2_cache['市场权益'].empty:
    st.subheader("📊 市场权益 vs 公司权益 vs 中位数对比")
    try:
        df_market_equity = data2_cache['市场权益'].copy()
        df_market_equity = clean_dataframe(df_market_equity)
        
        month_col = safe_get_column(df_market_equity, ['月份', '月份'], 0)
        equity_col = safe_get_column(df_market_equity, ['市场权益', '市场权益'], 1)
        median_col = safe_get_column(df_market_equity, ['中位数', '中位数（亿）', '权益中位数', '中位数权益'], 2)
        
        if month_col is None:
            st.warning("未找到'月份'列，请检查数据格式")
        elif equity_col is None:
            st.warning("未找到'市场权益'列，请检查数据格式")
        else:
            cols_to_extract = [month_col, equity_col]
            if median_col is not None:
                cols_to_extract.append(median_col)
            
            market_equity_df = df_market_equity[cols_to_extract].copy()
            market_equity_df = market_equity_df.dropna(subset=[month_col, equity_col])
            
            def format_month(val):
                try:
                    val_str = str(int(val)) if isinstance(val, (int, float)) else str(val)
                    if len(val_str) == 6 and val_str.isdigit():
                        return val_str
                except:
                    pass
                return None
            
            market_equity_df['年月'] = market_equity_df[month_col].apply(format_month)
            market_equity_df['市场权益'] = market_equity_df[equity_col]
            
            if median_col is not None:
                market_equity_df['中位数（亿元）'] = market_equity_df[median_col]
            
            market_equity_df = market_equity_df.dropna(subset=['年月'])
            
            if market_equity_df.empty:
                st.info("暂无有效的市场权益数据（月份格式需为202601这样的6位数字）")
            else:
                company_equity_list = []
                
                df_fund_current = data1_cache.get('资金对账表-月', pd.DataFrame())
                df_fund_current = clean_dataframe(df_fund_current)
                
                fund_month_col = safe_get_column(df_fund_current, ['月份', '月份'], 0)
                fund_equity_col = safe_get_column(df_fund_current, ['期末权益', '权益'], 7)
                
                if not df_fund_current.empty and fund_month_col is not None and fund_equity_col is not None:
                    def format_fund_month(val):
                        try:
                            val_str = str(int(val)) if isinstance(val, (int, float)) else str(val)
                            if len(val_str) == 6 and val_str.isdigit():
                                return val_str
                        except:
                            pass
                        return None
                    
                    df_fund_current['年月'] = df_fund_current[fund_month_col].apply(format_fund_month)
                    df_fund_current = df_fund_current.dropna(subset=['年月', fund_equity_col])
                    
                    if not df_fund_current.empty:
                        current_equity = df_fund_current.groupby('年月')[fund_equity_col].sum().reset_index()
                        current_equity.columns = ['年月', '公司权益']
                        current_equity['公司权益'] = current_equity['公司权益'] / 100000000
                        company_equity_list.append(current_equity)
                
                df_fund_last = data2_cache.get('上一年资金对账表-月', pd.DataFrame())
                df_fund_last = clean_dataframe(df_fund_last)
                
                last_fund_month_col = safe_get_column(df_fund_last, ['月份', '月份'], 0)
                last_fund_equity_col = safe_get_column(df_fund_last, ['期末权益', '权益'], 7)
                
                if not df_fund_last.empty and last_fund_month_col is not None and last_fund_equity_col is not None:
                    def format_fund_last_month(val):
                        try:
                            val_str = str(int(val)) if isinstance(val, (int, float)) else str(val)
                            if len(val_str) == 6 and val_str.isdigit():
                                return val_str
                        except:
                            pass
                        return None
                    
                    df_fund_last['年月'] = df_fund_last[last_fund_month_col].apply(format_fund_last_month)
                    df_fund_last = df_fund_last.dropna(subset=['年月', last_fund_equity_col])
                    
                    if not df_fund_last.empty:
                        last_equity = df_fund_last.groupby('年月')[last_fund_equity_col].sum().reset_index()
                        last_equity.columns = ['年月', '公司权益']
                        last_equity['公司权益'] = last_equity['公司权益'] / 100000000
                        company_equity_list.append(last_equity)
                
                if company_equity_list:
                    company_equity_data = pd.concat(company_equity_list, ignore_index=True).drop_duplicates(subset=['年月']).sort_values('年月')
                else:
                    company_equity_data = pd.DataFrame()
                
                if not company_equity_data.empty:
                    merged_df = pd.merge(market_equity_df, company_equity_data, on='年月', how='outer')
                else:
                    merged_df = market_equity_df.copy()
                    merged_df['公司权益'] = None
                
                if merged_df.empty:
                    st.info("暂无合并数据")
                else:
                    merged_df['市场权益（百亿元）'] = merged_df['市场权益'] / 100
                    merged_df['公司权益（亿元）'] = merged_df['公司权益']
                    
                    merged_df = merged_df.sort_values('年月')
                    chart_df = merged_df.tail(12).copy()
                    
                    chart_df['年月显示'] = chart_df['年月'].apply(
                        lambda x: f"{x[:4]}年{x[4:6]}月"
                    )
                    
                    plot_data = []
                    for _, row in chart_df.iterrows():
                        if pd.notna(row['市场权益（百亿元）']):
                            plot_data.append({
                                '月份': row['年月显示'],
                                '权益': row['市场权益（百亿元）'],
                                '类型': '市场权益（百亿元）'
                            })
                        if pd.notna(row['公司权益（亿元）']):
                            plot_data.append({
                                '月份': row['年月显示'],
                                '权益': row['公司权益（亿元）'],
                                '类型': '公司权益（亿元）'
                            })
                        if '中位数（亿元）' in row and pd.notna(row['中位数（亿元）']):
                            plot_data.append({
                                '月份': row['年月显示'],
                                '权益': row['中位数（亿元）'],
                                '类型': '中位数（亿元）'
                            })
                    
                    if plot_data:
                        plot_df = pd.DataFrame(plot_data)
                        
                        month_order = chart_df['年月显示'].tolist()
                        plot_df['月份'] = pd.Categorical(plot_df['月份'], categories=month_order, ordered=True)
                        plot_df = plot_df.sort_values('月份')
                        
                        fig = px.bar(
                            plot_df,
                            x='月份',
                            y='权益',
                            color='类型',
                            barmode='group',
                            title='市场权益（百亿元）vs 公司权益（亿元）vs 中位数（亿元）对比（最近12个月）',
                            labels={'权益': '权益', '月份': '月份'},
                            text_auto='.2f',
                            color_discrete_map={
                                '市场权益（百亿元）': '#2E86C1',
                                '公司权益（亿元）': '#F39C12',
                                '中位数（亿元）': '#28B463'
                            }
                        )
                        
                        fig.update_layout(
                            title_font=dict(size=18, color='#1A5276'),
                            font=dict(size=13),
                            bargap=0.25,
                            bargroupgap=0.15,
                            plot_bgcolor='#F8F9F9',
                            paper_bgcolor='white',
                            legend_title_text='',
                            yaxis=dict(tickformat='.2f', title='权益'),
                            xaxis=dict(title='月份')
                        )
                        
                        fig.update_traces(
                            texttemplate='%{y:.2f}',
                            textfont=dict(size=11, color='black', family='Arial Black'),
                            textposition='outside'
                        )
                        
                        st.plotly_chart(fig, use_container_width=True)
                        
                        with st.expander("📋 查看详细数据"):
                            display_cols = ['年月显示', '市场权益（百亿元）', '公司权益（亿元）']
                            if '中位数（亿元）' in chart_df.columns:
                                display_cols.append('中位数（亿元）')
                            display_df = chart_df[display_cols].copy()
                            display_df = display_df.sort_values('年月显示')
                            display_df.columns = ['月份'] + display_cols[1:]
                            
                            for col in display_df.columns[1:]:
                                display_df[col] = display_df[col].apply(
                                    lambda x: f"{x:.2f}" if pd.notna(x) else '-'
                                )
                            
                            st.dataframe(display_df, use_container_width=True, hide_index=True)
                    else:
                        st.info("没有可用于绘图的数据")
                        
    except Exception as e:
        st.warning(f"加载市场权益数据时出错: {e}")
        st.exception(e)

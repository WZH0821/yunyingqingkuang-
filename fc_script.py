import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import datetime
import requests

st.set_page_config(page_title="Dashboard", layout="wide")

# ============================================================
# 配置 - 你的GitHub信息
# ============================================================
GITHUB_USERNAME = "WZH0821"
GITHUB_REPO = "yunyingqingkuang-"
GITHUB_BRANCH = "main"
EXCEL_FILENAME = "data1.xlsx"

GITHUB_FILE_URL = f"https://raw.githubusercontent.com/{GITHUB_USERNAME}/{GITHUB_REPO}/{GITHUB_BRANCH}/{EXCEL_FILENAME}"

# ============================================================
# 常量定义
# ============================================================
COLOR_MAP = {
    '本月': '#2E86C1', '上月': '#F39C12', '去年同期': '#28B463',
    '本季度': '#2E86C1', '上季度': '#F39C12',
    '今年': '#2E86C1', '去年': '#F39C12',
}

MONTH_NAMES = {1: '1月', 2: '2月', 3: '3月', 4: '4月', 5: '5月', 6: '6月',
               7: '7月', 8: '8月', 9: '9月', 10: '10月', 11: '11月', 12: '12月'}

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

def parse_month_column(col):
    col_str = str(col)
    if len(col_str) == 6 and col_str.isdigit():
        year, month = int(col_str[:4]), int(col_str[4:6])
        if 1 <= month <= 12:
            return year, month
    return None, None

def get_month_columns(df: pd.DataFrame) -> list:
    if df.empty:
        return []
    return [col for col in df.columns if parse_month_column(col)[0] is not None]

def safe_division(a, b, default=0):
    return default if (b is None or b == 0) else a / b

def format_percent(value, decimals=2):
    if value is None or pd.isna(value):
        return '-'
    return f"{value:+.{decimals}f}%"

def get_metric_config(data_type: str) -> dict:
    METRIC_CONFIG = {
        '成交量': {
            'market_divide': 100000000, 'market_unit': '亿手', 'market_title': '（亿手）',
            'market_yaxis': '成交量（亿手）', 'company_divide': 10000, 'company_unit': '万手',
            'company_title': '（万手）', 'company_yaxis': '成交量（万手）', 'metric_name': '成交量'
        },
        '成交额': {
            'market_divide': 10000, 'market_unit': '万亿元', 'market_title': '（万亿元）',
            'market_yaxis': '成交额（万亿元）', 'company_divide': 100000000, 'company_unit': '亿元',
            'company_title': '（亿元）', 'company_yaxis': '成交额（亿元）', 'metric_name': '成交额'
        },
        '持仓量': {
            'market_divide': 1000000, 'market_unit': '百万手', 'market_title': '（百万手）',
            'market_yaxis': '持仓量（百万手）', 'company_divide': 10000, 'company_unit': '万手',
            'company_title': '（万手）', 'company_yaxis': '持仓量（万手）', 'metric_name': '持仓量'
        }
    }
    return METRIC_CONFIG.get(data_type, METRIC_CONFIG['成交量'])

def safe_get_column(df: pd.DataFrame, col_names: list, default_idx: int = None):
    if df.empty:
        return None
    for name in col_names:
        if name in df.columns:
            return name
    if default_idx is not None and len(df.columns) > default_idx:
        return df.columns[default_idx]
    return None

def compute_period_comparison(df: pd.DataFrame, selected_cols: list, prev_cols: list, 
                               last_year_cols: list, divide: float) -> dict:
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
                  labels={x: xlabel, y: ylabel, color: ''},
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
# 缓存数据加载函数 - 从GitHub加载
# ============================================================
@st.cache_data(ttl=3600)
def load_all_data_from_github():
    try:
        with st.spinner(f"📥 正在从GitHub下载 {EXCEL_FILENAME}..."):
            response = requests.get(GITHUB_FILE_URL, timeout=30)
            response.raise_for_status()
        
        excel_data = BytesIO(response.content)
        
        sheets_config = {
            'vol_market': '成交量-市场',
            'vol_company': '成交量-公司',
            'amt_market': '成交额-市场',
            'amt_company': '成交额-公司',
            'oi_market': '持仓量-市场',
            'oi_company': '持仓量-公司',
            'fund_current': '资金对账表-月',
            'fund_last_year': '上一年资金对账表-月',
            'trade_stats': '交易统计表-月'
        }
        
        data = {}
        for key, sheet_name in sheets_config.items():
            try:
                excel_data.seek(0)
                df = pd.read_excel(excel_data, sheet_name=sheet_name, header=0)
                df = clean_dataframe(df)
                data[key] = df
            except Exception as e:
                st.warning(f"⚠️ 加载sheet '{sheet_name}' 失败: {e}")
                data[key] = pd.DataFrame()
        
        return data
    
    except requests.exceptions.RequestException as e:
        st.error(f"❌ 从GitHub下载文件失败: {e}")
        st.info("💡 请检查：\n1. 文件是否已在仓库中\n2. 仓库是否为公开仓库\n3. 文件名是否正确")
        return None
    except Exception as e:
        st.error(f"❌ 读取Excel文件失败: {e}")
        return None

# ============================================================
# 加载数据
# ============================================================
data_dict = load_all_data_from_github()

if data_dict is None:
    st.error("❌ 数据加载失败，请检查配置")
    st.stop()

# 解包数据
df_vol_market = data_dict['vol_market']
df_vol_company = data_dict['vol_company']
df_amt_market = data_dict['amt_market']
df_amt_company = data_dict['amt_company']
df_oi_market = data_dict['oi_market']
df_oi_company = data_dict['oi_company']
df_fund_current = data_dict['fund_current']
df_fund_last_year = data_dict['fund_last_year']
df_trade_stats = data_dict['trade_stats']

# 检查核心数据
if df_vol_market.empty:
    st.error("❌ 核心数据加载失败，请检查Excel文件是否包含所有必需的sheet")
    st.stop()

st.success("✅ 数据从GitHub加载成功！")

# ============================================================
# 侧边栏 - 数据源信息
# ============================================================
with st.sidebar:
    st.header("📁 数据源")
    st.info(f"📊 数据来源: GitHub")
    st.caption(f"📁 仓库: {GITHUB_USERNAME}/{GITHUB_REPO}")
    st.caption(f"📄 文件: {EXCEL_FILENAME}")
    st.divider()
    st.caption(f"🔄 最后更新: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
    st.caption(f"💡 数据缓存: 1小时")
    
    if st.button("🔄 刷新数据"):
        st.cache_data.clear()
        st.rerun()

# ============================================================
# 数据筛选
# ============================================================
try:
    df = df_vol_market.copy()
    df = clean_dataframe(df)
    
    st.success(f"✅ 数据加载成功！共 {len(df)} 行，{len(df.columns)} 列")
    st.dataframe(df.head(10))
    
    with st.expander("📌 查看所有列名"):
        st.write(df.columns.tolist())
    
    # ============================================================
    # 筛选器
    # ============================================================
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    filter_cols = [col for col in df.columns if col not in numeric_cols]
    
    selected_filters = {}
    if filter_cols:
        cols = st.columns(3)
        for idx, col_name in enumerate(filter_cols):
            with cols[idx % 3]:
                unique_vals = df[col_name].dropna().unique().tolist()
                if unique_vals:
                    selected = st.multiselect(f"筛选 {col_name}", options=unique_vals, default=[], key=f"filter_{col_name}")
                    if selected:
                        selected_filters[col_name] = selected
    
    filtered_df = df.copy()
    for col, vals in selected_filters.items():
        if vals:
            filtered_df = filtered_df[filtered_df[col].isin(vals)]
    
    st.subheader("📊 筛选后数据")
    
    # ============================================================
    # 选择数据类型
    # ============================================================
    data_type = st.radio(
        "选择数据类型",
        options=['成交量', '成交额', '持仓量'],
        horizontal=True,
        key="data_type"
    )
    
    metric_config = get_metric_config(data_type)
    
    # 获取对应的数据
    df_detail = {'成交量': df_vol_market, '成交额': df_amt_market, '持仓量': df_oi_market}.get(data_type, pd.DataFrame())
    df_company_detail = {'成交量': df_vol_company, '成交额': df_amt_company, '持仓量': df_oi_company}.get(data_type, pd.DataFrame())
    
    df_detail = clean_dataframe(df_detail)
    df_company_detail = clean_dataframe(df_company_detail)
    
    st.success(f"✅ {data_type}数据加载成功！共 {len(df_detail)} 行")
    st.success(f"✅ {data_type}公司数据加载成功！共 {len(df_company_detail)} 行")
    
    st.subheader(f"📋 {data_type}数据预览")
    st.dataframe(df_detail.head(10))
    
    with st.expander("📌 查看所有列名"):
        st.write(df_detail.columns.tolist())
    
    # ============================================================
    # 获取日期列信息
    # ============================================================
    date_cols = get_month_columns(df_detail)
    
    date_labels = {}
    date_info = {}
    for col in date_cols:
        try:
            col_str = str(col)
            if len(col_str) == 6 and col_str.isdigit():
                year = int(col_str[:4])
                month = int(col_str[4:6])
                date_labels[col] = f"{year}年{month:02d}月"
                date_info[col] = {'year': year, 'month': month}
            else:
                date_labels[col] = str(col)
                date_info[col] = {'year': None, 'month': None}
        except:
            date_labels[col] = str(col)
            date_info[col] = {'year': None, 'month': None}
    
    # ============================================================
    # 各交易所柱状图（市场 + 公司）
    # ============================================================
    st.subheader(f"📊 各交易所{data_type}对比（市场）{metric_config['market_title']}")
    
    col1, col2 = st.columns(2)
    with col1:
        time_dimension_exchange = st.radio(
            "选择时间维度",
            options=['月度', '季度', '年度'],
            horizontal=True,
            key="time_dimension_exchange"
        )
    with col2:
        if time_dimension_exchange == '月度':
            options_exchange = sorted(date_cols, reverse=True)
            option_labels_exchange = {col: date_labels.get(col, str(col)) for col in options_exchange}
            value_cols_map_exchange = None
        elif time_dimension_exchange == '季度':
            quarter_map_exchange = {}
            for col in date_cols:
                info = date_info.get(col, {})
                year = info.get('year')
                month = info.get('month')
                if year is not None and month is not None:
                    if 1 <= month <= 3:
                        q = 'Q1'
                    elif 4 <= month <= 6:
                        q = 'Q2'
                    elif 7 <= month <= 9:
                        q = 'Q3'
                    else:
                        q = 'Q4'
                    key = f"{year}{q}"
                    if key not in quarter_map_exchange:
                        quarter_map_exchange[key] = []
                    quarter_map_exchange[key].append(col)
            options_exchange = sorted(quarter_map_exchange.keys(), reverse=True)
            option_labels_exchange = {k: f"{k[:4]}年{k[4:]}" for k in options_exchange}
            value_cols_map_exchange = quarter_map_exchange
        else:
            year_map_exchange = {}
            for col in date_cols:
                info = date_info.get(col, {})
                year = info.get('year')
                if year is not None:
                    key = str(year)
                    if key not in year_map_exchange:
                        year_map_exchange[key] = []
                    year_map_exchange[key].append(col)
            options_exchange = sorted(year_map_exchange.keys(), reverse=True)
            option_labels_exchange = {k: f"{k}年" for k in options_exchange}
            value_cols_map_exchange = year_map_exchange
        
        selected_key_exchange = st.selectbox(
            f"选择{time_dimension_exchange}",
            options=options_exchange,
            format_func=lambda x: option_labels_exchange.get(x, str(x)),
            key="time_selector_exchange"
        )
    
    if time_dimension_exchange == '月度':
        selected_cols_exchange = [selected_key_exchange]
        selected_label_exchange = option_labels_exchange.get(selected_key_exchange, str(selected_key_exchange))
    else:
        selected_cols_exchange = value_cols_map_exchange.get(selected_key_exchange, [])
        selected_label_exchange = option_labels_exchange.get(selected_key_exchange, str(selected_key_exchange))
    
    # 计算前后期
    if time_dimension_exchange == '月度':
        col_str = str(selected_key_exchange)
        if len(col_str) == 6 and col_str.isdigit():
            year = int(col_str[:4])
            month = int(col_str[4:6])
            prev_year = year
            prev_month = month - 1
            if prev_month == 0:
                prev_month = 12
                prev_year -= 1
            prev_key = f"{prev_year}{prev_month:02d}"
            prev_key = int(prev_key) if prev_key.isdigit() else prev_key
            prev_cols_exchange = [prev_key] if prev_key in df_detail.columns else []
            last_year_key = f"{year - 1}{month:02d}"
            last_year_key = int(last_year_key) if last_year_key.isdigit() else last_year_key
            last_year_cols_exchange = [last_year_key] if last_year_key in df_detail.columns else []
        else:
            prev_cols_exchange = []
            last_year_cols_exchange = []
    elif time_dimension_exchange == '季度':
        key_str = str(selected_key_exchange)
        year = int(key_str[:4])
        q = key_str[4:]
        q_num = {'Q1': 1, 'Q2': 2, 'Q3': 3, 'Q4': 4}[q]
        prev_q_num = q_num - 1
        prev_year = year
        if prev_q_num == 0:
            prev_q_num = 4
            prev_year -= 1
        prev_q = ['Q1', 'Q2', 'Q3', 'Q4'][prev_q_num - 1]
        prev_key = f"{prev_year}{prev_q}"
        prev_cols_exchange = value_cols_map_exchange.get(prev_key, []) if value_cols_map_exchange else []
        last_year_key = f"{year - 1}{q}"
        last_year_cols_exchange = value_cols_map_exchange.get(last_year_key, []) if value_cols_map_exchange else []
    else:
        prev_year = int(selected_key_exchange) - 1
        prev_key = str(prev_year)
        prev_cols_exchange = value_cols_map_exchange.get(prev_key, []) if value_cols_map_exchange else []
        last_year_cols_exchange = []
    
    # ===== 市场交易所数据 =====
    exchanges = df_detail['交易所'].unique().tolist() if '交易所' in df_detail.columns else []
    
    exchange_compare = []
    current_sum_exchange = df_detail[selected_cols_exchange].sum().sum() if selected_cols_exchange else 0
    total_current_exchange = current_sum_exchange / metric_config['market_divide']
    
    total_prev_exchange = None
    if prev_cols_exchange:
        prev_sum = df_detail[prev_cols_exchange].sum().sum() if prev_cols_exchange else 0
        total_prev_exchange = prev_sum / metric_config['market_divide']
    
    total_last_year_exchange = None
    if last_year_cols_exchange:
        last_sum = df_detail[last_year_cols_exchange].sum().sum() if last_year_cols_exchange else 0
        total_last_year_exchange = last_sum / metric_config['market_divide']
    
    for exchange in exchanges:
        exchange_df = df_detail[df_detail['交易所'] == exchange]
        row = {'交易所': exchange}
        current_val = exchange_df[selected_cols_exchange].sum().sum() / metric_config['market_divide'] if selected_cols_exchange else 0
        row['本月'] = current_val
        
        if prev_cols_exchange:
            prev_val = exchange_df[prev_cols_exchange].sum().sum() / metric_config['market_divide'] if prev_cols_exchange else 0
            row['上月'] = prev_val
            row['环比'] = safe_division(current_val - prev_val, prev_val) * 100 if prev_val else None
        else:
            row['上月'] = None
            row['环比'] = None
        
        if last_year_cols_exchange:
            last_val = exchange_df[last_year_cols_exchange].sum().sum() / metric_config['market_divide'] if last_year_cols_exchange else 0
            row['去年同期'] = last_val
            row['同比'] = safe_division(current_val - last_val, last_val) * 100 if last_val else None
        else:
            row['去年同期'] = None
            row['同比'] = None
        
        exchange_compare.append(row)
    
    exchange_df_plot = pd.DataFrame(exchange_compare)
    
    # ===== 公司交易所数据 =====
    exchange_compare_company = []
    current_sum_exchange_company = df_company_detail[selected_cols_exchange].sum().sum() if selected_cols_exchange else 0
    total_current_exchange_company = current_sum_exchange_company / metric_config['company_divide']
    
    total_prev_exchange_company = None
    if prev_cols_exchange:
        prev_sum = df_company_detail[prev_cols_exchange].sum().sum() if prev_cols_exchange else 0
        total_prev_exchange_company = prev_sum / metric_config['company_divide']
    
    total_last_year_exchange_company = None
    if last_year_cols_exchange:
        last_sum = df_company_detail[last_year_cols_exchange].sum().sum() if last_year_cols_exchange else 0
        total_last_year_exchange_company = last_sum / metric_config['company_divide']
    
    for exchange in exchanges:
        exchange_df = df_company_detail[df_company_detail['交易所'] == exchange]
        row = {'交易所': exchange}
        current_val = exchange_df[selected_cols_exchange].sum().sum() / metric_config['company_divide'] if selected_cols_exchange else 0
        row['本月'] = current_val
        
        if prev_cols_exchange:
            prev_val = exchange_df[prev_cols_exchange].sum().sum() / metric_config['company_divide'] if prev_cols_exchange else 0
            row['上月'] = prev_val
            row['环比'] = safe_division(current_val - prev_val, prev_val) * 100 if prev_val else None
        else:
            row['上月'] = None
            row['环比'] = None
        
        if last_year_cols_exchange:
            last_val = exchange_df[last_year_cols_exchange].sum().sum() / metric_config['company_divide'] if last_year_cols_exchange else 0
            row['去年同期'] = last_val
            row['同比'] = safe_division(current_val - last_val, last_val) * 100 if last_val else None
        else:
            row['去年同期'] = None
            row['同比'] = None
        
        exchange_compare_company.append(row)
    
    exchange_df_plot_company = pd.DataFrame(exchange_compare_company)
    
    # ===== 绘制柱状图 =====
    value_cols_exchange = ['本月']
    if prev_cols_exchange:
        value_cols_exchange.append('上月')
    if last_year_cols_exchange:
        value_cols_exchange.append('去年同期')
    
    if not exchange_df_plot.empty:
        exchange_melted = exchange_df_plot.melt(
            id_vars=['交易所'],
            value_vars=value_cols_exchange,
            var_name='期间',
            value_name=data_type
        )
        exchange_melted = exchange_melted.dropna(subset=[data_type])
        period_order = ['上月', '本月', '去年同期']
        exchange_melted['期间'] = pd.Categorical(
            exchange_melted['期间'],
            categories=[p for p in period_order if p in exchange_melted['期间'].unique()],
            ordered=True
        )
        exchange_melted = exchange_melted.sort_values('期间')
        
        if not exchange_melted.empty:
            fig = create_bar_chart(
                exchange_melted, '交易所', data_type, '期间',
                f'各交易所{data_type}对比（市场）- {selected_label_exchange}',
                metric_config['market_yaxis'], COLOR_MAP
            )
            if fig:
                st.plotly_chart(fig, use_container_width=True)
            
            # 综合表
            st.subheader("📊 交易所环比同比综合表（市场）")
            table_data = [{
                '维度': '市场',
                f'{data_type}（{metric_config["market_unit"]}）': f"{total_current_exchange:.2f}",
                '环比（%）': format_percent((total_current_exchange - total_prev_exchange) / total_prev_exchange * 100 if total_prev_exchange else None),
                '同比（%）': format_percent((total_current_exchange - total_last_year_exchange) / total_last_year_exchange * 100 if total_last_year_exchange else None)
            }]
            for _, row in exchange_df_plot.iterrows():
                table_data.append({
                    '维度': row['交易所'],
                    f'{data_type}（{metric_config["market_unit"]}）': f"{row['本月']:.2f}",
                    '环比（%）': format_percent(row.get('环比')),
                    '同比（%）': format_percent(row.get('同比'))
                })
            st.dataframe(pd.DataFrame(table_data), use_container_width=True, hide_index=True)
    
    # ===== 公司柱状图 =====
    st.subheader(f"📊 各交易所{data_type}对比（公司）{metric_config['company_title']}")
    
    if not exchange_df_plot_company.empty:
        exchange_melted_company = exchange_df_plot_company.melt(
            id_vars=['交易所'],
            value_vars=value_cols_exchange,
            var_name='期间',
            value_name=data_type
        )
        exchange_melted_company = exchange_melted_company.dropna(subset=[data_type])
        period_order = ['上月', '本月', '去年同期']
        exchange_melted_company['期间'] = pd.Categorical(
            exchange_melted_company['期间'],
            categories=[p for p in period_order if p in exchange_melted_company['期间'].unique()],
            ordered=True
        )
        exchange_melted_company = exchange_melted_company.sort_values('期间')
        
        if not exchange_melted_company.empty:
            fig = create_bar_chart(
                exchange_melted_company, '交易所', data_type, '期间',
                f'各交易所{data_type}对比（公司）- {selected_label_exchange}',
                metric_config['company_yaxis'], COLOR_MAP
            )
            if fig:
                st.plotly_chart(fig, use_container_width=True)
            
            st.subheader("📊 交易所环比同比综合表（公司）")
            table_data = [{
                '维度': '市场',
                f'{data_type}（{metric_config["company_unit"]}）': f"{total_current_exchange_company:.2f}",
                '环比（%）': format_percent((total_current_exchange_company - total_prev_exchange_company) / total_prev_exchange_company * 100 if total_prev_exchange_company else None),
                '同比（%）': format_percent((total_current_exchange_company - total_last_year_exchange_company) / total_last_year_exchange_company * 100 if total_last_year_exchange_company else None)
            }]
            for _, row in exchange_df_plot_company.iterrows():
                table_data.append({
                    '维度': row['交易所'],
                    f'{data_type}（{metric_config["company_unit"]}）': f"{row['本月']:.2f}",
                    '环比（%）': format_percent(row.get('环比')),
                    '同比（%）': format_percent(row.get('同比'))
                })
            st.dataframe(pd.DataFrame(table_data), use_container_width=True, hide_index=True)
    
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
                color_map = {f"{y}年": '#2E86C1' for y in [latest_year - 2, latest_year - 1, latest_year] if f"{y}年" in metric_df['年份'].unique()}
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
                table_df.index = table_df.index + 1
                st.dataframe(table_df, use_container_width=True)
        else:
            st.info("暂无公司占市场比重数据")
    except Exception as e:
        st.warning(f"无法加载公司/市场对比数据: {e}")

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
                table_data = [{
                    '维度': '市场',
                    f'{group_data_type}（{unit}）': f"{total['current']:.2f}",
                    ('环比' if time_dimension != '年度' else '同比') + '（%）': format_percent(total['mom'] if time_dimension != '年度' else total['yoy']),
                    ('同比' if time_dimension != '年度' else '-') + '（%）': format_percent(total['yoy'] if time_dimension != '年度' else None)
                }]
                for _, row in df_plot.iterrows():
                    table_data.append({
                        '维度': row[group_col],
                        f'{group_data_type}（{unit}）': f"{row[current_label]:.2f}",
                        ('环比' if time_dimension != '年度' else '同比') + '（%）': format_percent(row.get('环比' if time_dimension != '年度' else '同比')),
                        ('同比' if time_dimension != '年度' else '-') + '（%）': format_percent(row.get('同比' if time_dimension != '年度' else None))
                    })
                st.dataframe(pd.DataFrame(table_data), use_container_width=True, hide_index=True)

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
                '平仓盈亏': safe_get_column(df_fund_current, ['平仓盈亏'], 6),
                '期末权益': safe_get_column(df_fund_current, ['期末权益', '权益'], 7),
                '期权权利金收入': safe_get_column(df_fund_current, ['期权权利金收入', '权利金收入'], 9),
                '期权权利金支出': safe_get_column(df_fund_current, ['期权权利金支出', '权利金支出'], 10),
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
                            pnl = df_fund_current.loc[mask, col_mapping['平仓盈亏']].sum() if col_mapping.get('平仓盈亏') else 0
                            if col_mapping.get('期权权利金收入') and col_mapping.get('期权权利金支出'):
                                pnl = pnl + df_fund_current.loc[mask, col_mapping['期权权利金收入']].sum() - df_fund_current.loc[mask, col_mapping['期权权利金支出']].sum()
                            fund_data.append({'月份': month_str, '期末权益': equity, '净入金': net_deposit,
                                              '留存手续费': fee, '平仓盈亏': pnl, '类型': '今年'})
                        except Exception:
                            continue
                
                if not df_fund_last_year.empty:
                    last_month_col = safe_get_column(df_fund_last_year, ['月份'], 0)
                    if last_month_col:
                        for month_val in df_fund_last_year[last_month_col].dropna().unique():
                            try:
                                month_str = str(int(month_val)) if isinstance(month_val, (int, float)) else str(month_val)
                                if len(month_str) != 6 or not month_str.isdigit():
                                    continue
                                mask = df_fund_last_year[last_month_col] == month_val
                                last_equity_col = safe_get_column(df_fund_last_year, ['期末权益', '权益'], 7)
                                if last_equity_col:
                                    equity = df_fund_last_year.loc[mask, last_equity_col].sum()
                                else:
                                    equity = 0
                                if pd.isna(equity) or equity == 0:
                                    continue
                                net_deposit = 0
                                last_in_col = safe_get_column(df_fund_last_year, ['入金', '入金金额'], 3)
                                last_out_col = safe_get_column(df_fund_last_year, ['出金', '出金金额'], 4)
                                if last_in_col and last_out_col:
                                    net_deposit = df_fund_last_year.loc[mask, last_in_col].sum() - df_fund_last_year.loc[mask, last_out_col].sum()
                                last_fee_col = safe_get_column(df_fund_last_year, ['留存手续费', '手续费', '手续费留存'], 5)
                                fee = df_fund_last_year.loc[mask, last_fee_col].sum() if last_fee_col else 0
                                last_pnl_col = safe_get_column(df_fund_last_year, ['平仓盈亏'], 6)
                                pnl = df_fund_last_year.loc[mask, last_pnl_col].sum() if last_pnl_col else 0
                                fund_data.append({'月份': month_str, '期末权益': equity, '净入金': net_deposit,
                                                  '留存手续费': fee, '平仓盈亏': pnl, '类型': '去年'})
                            except Exception:
                                continue

                if fund_data:
                    fund_df = pd.DataFrame(fund_data).sort_values('月份')
                    fund_df['月份显示'] = fund_df['月份'].str[4:6]

                    display_df = fund_df[['月份', '期末权益', '净入金', '留存手续费', '平仓盈亏', '类型']].copy()
                    display_df['期末权益'] = (display_df['期末权益'] / 100000000).round(2)
                    display_df['净入金'] = (display_df['净入金'] / 10000000).round(2)
                    display_df['留存手续费'] = (display_df['留存手续费'] / 100000).round(2)
                    display_df['平仓盈亏'] = (display_df['平仓盈亏'] / 1000000).round(2)
                    display_df.columns = ['月份', '期末权益（亿元）', '净入金（千万）', '留存手续费（十万）', '平仓盈亏（百万）', '类型']
                    st.dataframe(display_df.sort_values('月份', ascending=False), use_container_width=True, hide_index=True)

                    fund_df_sorted = fund_df.sort_values('月份')
                    fund_df_sorted['期末权益（亿元）'] = fund_df_sorted['期末权益'] / 100000000
                    fund_df_sorted['净入金（千万）'] = fund_df_sorted['净入金'] / 10000000
                    fund_df_sorted['留存手续费（十万）'] = fund_df_sorted['留存手续费'] / 100000
                    fund_df_sorted['平仓盈亏（百万）'] = fund_df_sorted['平仓盈亏'] / 1000000

                    color_map_fund = {'今年': '#2E86C1', '去年': '#F39C12'}

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
                        ('期末权益（亿元）', '期末权益', '期末权益（亿）', f"当期: {current_equity:.2f}亿" if current_equity else "当期: -",
                         f"{cumsum_data['期末权益']:.2f}亿"),
                        ('净入金（千万）', '净入金', '净入金（千万）', '', f"{cumsum_data['净入金']:+.2f}千万"),
                        ('留存手续费（十万）', '留存手续费', '留存手续费（十万）', '', f"{cumsum_data['留存手续费']:.2f}十万"),
                        ('平仓盈亏（百万）', '平仓盈亏', '平仓盈亏（百万）', '', f"{cumsum_data['平仓盈亏']:+.2f}百万")
                    ]

                    rows = st.columns(2)
                    for idx, (col, title, ylabel, annotation, cumsum_text) in enumerate(charts):
                        with rows[idx % 2]:
                            fig = create_line_chart(fund_df_sorted, '月份显示', col, '类型',
                                                    title, '', ylabel, color_map_fund, '.2f')
                            if fig:
                                if annotation:
                                    fig.add_annotation(x=0.98, y=0.98, xref='paper', yref='paper',
                                                        text=annotation, showarrow=False, font=dict(size=10, color='#1A5276'),
                                                        bgcolor='rgba(255,255,255,0.85)', bordercolor='#1A5276',
                                                        borderwidth=1, borderpad=4, xanchor='right', yanchor='top')
                                fig.add_annotation(x=0.98, y=0.88 if annotation else 0.98, xref='paper', yref='paper',
                                                    text=f'累计: {cumsum_text}', showarrow=False,
                                                    font=dict(size=11, color='#1A5276'),
                                                    bgcolor='rgba(255,255,255,0.85)', bordercolor='#1A5276',
                                                    borderwidth=1, borderpad=4, xanchor='right', yanchor='top')
                                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})

                    st.subheader("📊 客户情况统计")
                    try:
                        trade_customer_data = []
                        for df_trade, year_type in [(df_trade_stats, '今年')]:
                            if df_trade.empty:
                                continue
                            month_col = safe_get_column(df_trade, ['月份'], 0)
                            investor_col = safe_get_column(df_trade, ['投资者代码', '客户代码', '投资者'], 3)
                            if not month_col or not investor_col:
                                continue
                            f_col = safe_get_column(df_trade, ['平仓盈亏'], 4)
                            g_col = safe_get_column(df_trade, ['期权权利金收入', '权利金收入'], 5)
                            h_col = safe_get_column(df_trade, ['期权权利金支出', '权利金支出'], 6)

                            for m in df_trade[month_col].dropna().unique():
                                try:
                                    m_str = str(int(m)) if isinstance(m, (int, float)) else str(m)
                                    if len(m_str) != 6 or not m_str.isdigit():
                                        continue
                                    mask = df_trade[month_col] == m
                                    month_df = df_trade[mask]
                                    if month_df.empty:
                                        continue
                                    total_investors = month_df[investor_col].nunique()
                                    if f_col and g_col and h_col:
                                        month_df['盈亏'] = month_df[f_col] + month_df[g_col] - month_df[h_col]
                                        profit_investors = month_df[month_df['盈亏'] > 0][investor_col].nunique()
                                    else:
                                        profit_investors = 0
                                    trade_customer_data.append({'月份': m_str, '交易客户数': total_investors,
                                                                '盈利客户数': profit_investors, '类型': year_type})
                                except Exception:
                                    continue

                        if trade_customer_data:
                            customer_df = pd.DataFrame(trade_customer_data).sort_values('月份')
                            customer_df['月份显示'] = customer_df['月份'].astype(str).str[4:6]

                            col_left, col_right = st.columns(2)
                            with col_left:
                                fig = create_line_chart(customer_df, '月份显示', '盈利客户数', '类型',
                                                        '盈利客户数（当月）', '', '客户数', color_map_fund, '.0f')
                                if fig:
                                    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
                            with col_right:
                                fig = create_line_chart(customer_df, '月份显示', '交易客户数', '类型',
                                                        '交易客户数（当月）', '', '客户数', color_map_fund, '.0f')
                                if fig:
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
            trade_month_col = safe_get_column(df_trade_stats, ['月份'], 0)
            trade_dept_col = safe_get_column(df_trade_stats, ['部门', '部门名称'], 1)
            trade_investor_col = safe_get_column(df_trade_stats, ['投资者代码', '客户代码', '投资者'], 3)
            
            if not trade_month_col or not trade_dept_col or not trade_investor_col:
                st.warning("交易统计表缺少必要列（月份、部门、投资者代码）")
            else:
                f_col = safe_get_column(df_trade_stats, ['平仓盈亏'], 4)
                g_col = safe_get_column(df_trade_stats, ['权利金收入', '期权权利金收入'], 5)
                h_col = safe_get_column(df_trade_stats, ['权利金支出', '期权权利金支出'], 6)

                all_months = []
                for m in df_trade_stats[trade_month_col].dropna().unique():
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

                    df_trade_stats['月份_str'] = df_trade_stats[trade_month_col].apply(
                        lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x))
                    filtered_trade = df_trade_stats[df_trade_stats['月份_str'] == selected_month]

                    if not filtered_trade.empty:
                        result = filtered_trade.groupby(trade_dept_col).agg(
                            有交易客户数=(trade_investor_col, 'nunique')
                        ).reset_index()
                        result.columns = ['部门', '有交易客户数']

                        if f_col and g_col and h_col:
                            filtered_trade['盈亏'] = filtered_trade[f_col] + filtered_trade[g_col] - filtered_trade[h_col]
                            profit_mask = filtered_trade['盈亏'] > 0
                            profit_df = filtered_trade[profit_mask]
                            profit_count = profit_df.groupby(trade_dept_col).agg(
                                盈利客户数=(trade_investor_col, 'nunique')
                            ).reset_index()
                            profit_count.columns = ['部门', '盈利客户数']
                            result = pd.merge(result, profit_count, on='部门', how='left').fillna(0)

                            pnl_df = filtered_trade.groupby(trade_dept_col).apply(
                                lambda x: (x[f_col].sum() + x[g_col].sum() - x[h_col].sum())
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

                        result['盈利面'] = result.apply(
                            lambda row: safe_division(row['盈利客户数'], row['有交易客户数']) * 100, axis=1
                        )
                        result = result.sort_values('有交易客户数', ascending=False)

                        if '期末权益' in result.columns:
                            result['期末权益'] = result['期末权益'].apply(lambda x: f"{int(x):,}")
                        if '平仓盈亏' in result.columns:
                            result['平仓盈亏'] = result['平仓盈亏'].apply(lambda x: f"{int(x):,}")
                        if '净入金' in result.columns:
                            result['净入金'] = result['净入金'].apply(lambda x: f"{int(x):,}")
                        if '留存手续费' in result.columns:
                            result['留存手续费'] = result['留存手续费'].apply(lambda x: f"{int(x):,}")
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

    st.success("✅ 所有分析已完成！")

except Exception as e:
    st.error(f"❌ 处理数据时出错: {e}")
    st.exception(e)

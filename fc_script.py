import streamlit as st
import pandas as pd
import plotly.express as px
from typing import Dict, List, Optional, Tuple
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
# 从GitHub加载数据的函数（带缓存）
# ============================================================
@st.cache_data
def load_data_from_github(url: str):
    """从GitHub加载Excel数据"""
    try:
        # 检查文件是否存在
        response = requests.head(url, timeout=10)
        if response.status_code != 200:
            st.error(f"❌ 无法从GitHub获取数据文件: {EXCEL_FILENAME}")
            st.info(f"请检查文件路径: {url}")
            return None
        
        # 加载Excel文件的所有sheet
        df_excel = pd.read_excel(url, sheet_name=None, header=0)
        return df_excel
    except requests.exceptions.RequestException as e:
        st.error(f"❌ 网络请求失败: {e}")
        return None
    except Exception as e:
        st.error(f"❌ 加载Excel文件失败: {e}")
        return None

# ============================================================
# 侧边栏
# ============================================================
with st.sidebar:
    st.header("📊 数据源")
    st.info(f"📁 数据文件: {EXCEL_FILENAME}")
    st.caption(f"📅 最后更新: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
    st.divider()
    
    # 添加刷新按钮
    if st.button("🔄 刷新数据"):
        st.cache_data.clear()
        st.rerun()
    
    st.divider()
    st.caption(f"📌 数据来源: GitHub")
    st.caption(f"🔗 {GITHUB_USERNAME}/{GITHUB_REPO}")

# ============================================================
# 主逻辑
# ============================================================
# 从GitHub加载数据
with st.spinner('📥 正在从GitHub加载数据文件...'):
    df_excel = load_data_from_github(GITHUB_FILE_URL)

if df_excel is None:
    st.error("❌ 无法加载数据，请检查网络连接和文件路径")
    st.stop()

try:
    sheets = {
        '成交量-市场': '成交量-市场', '成交量-公司': '成交量-公司',
        '成交额-市场': '成交额-市场', '成交额-公司': '成交额-公司',
        '持仓量-市场': '持仓量-市场', '持仓量-公司': '持仓量-公司',
        '资金对账表-月': '资金对账表-月', '上一年资金对账表-月': '上一年资金对账表-月',
        '交易统计表-月': '交易统计表-月', '上一年交易统计表-月': '上一年交易统计表-月',
        '投资者资料查询': '投资者资料查询', '活跃客户': '活跃客户',
        '市场权益': '市场权益'
    }
    
    data_cache = {}
    missing_sheets = []
    
    for key, sheet in sheets.items():
        try:
            if sheet in df_excel:
                data_cache[key] = clean_dataframe(df_excel[sheet])
            else:
                data_cache[key] = pd.DataFrame()
                missing_sheets.append(sheet)
        except Exception as e:
            data_cache[key] = pd.DataFrame()
            st.warning(f"⚠️ 加载sheet '{sheet}' 时出错: {e}")
    
    if missing_sheets:
        st.warning(f"⚠️ 以下Sheet不存在于文件中: {', '.join(missing_sheets)}")

    df_vol_market = data_cache['成交量-市场']
    df_vol_company = data_cache['成交量-公司']
    df_amt_market = data_cache['成交额-市场']
    df_amt_company = data_cache['成交额-公司']
    df_oi_market = data_cache['持仓量-市场']
    df_oi_company = data_cache['持仓量-公司']
    df_fund_current = data_cache['资金对账表-月']
    df_fund_last_year = data_cache['上一年资金对账表-月']
    df_trade_stats = data_cache['交易统计表-月']
    df_trade_last = data_cache['上一年交易统计表-月']
    df_investor = data_cache['投资者资料查询']
    df_active = data_cache['活跃客户']
    df_market_equity = data_cache['市场权益']

    df_trade_stats = normalize_trade_columns(df_trade_stats)
    df_trade_last = normalize_trade_columns(df_trade_last)

    available_sheets = {k: v for k, v in data_cache.items() if not v.empty}
    if not available_sheets:
        st.error("❌ 没有可用的数据表")
        st.stop()

    st.title("📊 交易数据看板")

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
                table_data = [{
                    '维度': '合计',
                    f'{data_type}（{metric_config[unit_key]}）': f"{total['current']:.2f}",
                    '环比（%）': format_percent(total['mom']),
                    '同比（%）': format_percent(total['yoy'])
                }]
                for _, row in df_plot.iterrows():
                    table_data.append({
                        '维度': row['交易所'],
                        f'{data_type}（{metric_config[unit_key]}）': f"{row['本月']:.2f}",
                        '环比（%）': format_percent(row.get('环比')),
                        '同比（%）': format_percent(row.get('同比'))
                    })
                st.dataframe(pd.DataFrame(table_data), use_container_width=True, hide_index=True)

    # ============================================================
    # 公司占市场比重
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
                
                compare_label1 = '环比' if time_dimension != '年度' else '同比'
                compare_label2 = '同比' if time_dimension != '年度' else '-'
                
                table_data = [{
                    '维度': '合计',
                    f'{group_data_type}（{unit}）': f"{total['current']:.2f}",
                    f'{compare_label1}（%）': format_percent(total['mom'] if time_dimension != '年度' else total['yoy']),
                    f'{compare_label2}（%）': format_percent(total['yoy'] if time_dimension != '年度' else None)
                }]
                for _, row in df_plot.iterrows():
                    table_data.append({
                        '维度': row[group_col],
                        f'{group_data_type}（{unit}）': f"{row[current_label]:.2f}",
                        f'{compare_label1}（%）': format_percent(row.get('环比' if time_dimension != '年度' else '同比')),
                        f'{compare_label2}（%）': format_percent(row.get('同比' if time_dimension != '年度' else None))
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
                            fig = create_line_chart(fund_df_sorted, '月份显示', col, '类型',
                                                    title, '月份', ylabel, color_map_fund, '.2f')
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

                            col_left, col_right = st.columns(2)
                            with col_left:
                                fig = create_line_chart(customer_df, '月份显示', '盈利客户数', '类型',
                                                        '盈利客户数（当月）', '月份', '客户数', color_map_fund, '.0f')
                                if fig:
                                    st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
                            with col_right:
                                fig = create_line_chart(customer_df, '月份显示', '交易客户数', '类型',
                                                        '交易客户数（当月）', '月份', '客户数', color_map_fund, '.0f')
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
    # 每月开户数统计（带月份筛选器）
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
                    
                    if selected_month_filter != '全部':
                        current_month = selected_month_filter.replace('年', '').replace('月', '')
                        current_year = int(current_month[:4])
                        current_month_num = int(current_month[4:6])
                        last_year_month = f"{current_year - 1}{current_month_num:02d}"
                        last_year_display = f"{current_year - 1}年{current_month_num:02d}月"
                    else:
                        latest_year = max(monthly_count['年份'].unique())
                        current_month = None
                        last_year_month = None
                        last_year_display = None
                    
                    col_pie1, col_pie2 = st.columns(2)
                    
                    with col_pie1:
                        if selected_month_filter != '全部':
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
                        if selected_month_filter != '全部' and last_year_month is not None:
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
                        elif selected_month_filter == '全部':
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
                        else:
                            st.info("暂无去年同期数据")
                                            
        except Exception as e:
            st.warning(f"加载开户数据时出错: {e}")
            st.exception(e)

    # ============================================================
    # 活跃客户统计
    # ============================================================
    if '活跃客户' in data_cache and not data_cache['活跃客户'].empty:
        st.subheader("📊 公司客户资金情况统计")
        try:
            df_active = data_cache['活跃客户'].copy()
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
                    
                    df_fund = data_cache.get('资金对账表-月', pd.DataFrame())
                    df_fund = clean_dataframe(df_fund)
                    
                    df_trade = data_cache.get('交易统计表-月', pd.DataFrame())
                    df_trade = clean_dataframe(df_trade)
                    df_trade = normalize_trade_columns(df_trade)
                    
                    fund_month_col = safe_get_column(df_fund, ['月份', '月份'], 0)
                    fund_investor_col = safe_get_column(df_fund, ['投资者代码', '客户代码', '投资者'], 2)
                    fund_equity_col = safe_get_column(df_fund, ['期末权益', '权益'], 7)
                    fund_fee_col = safe_get_column(df_fund, ['留存手续费', '手续费', '手续费留存'], 5)
                    fund_inflow_col = safe_get_column(df_fund, ['入金', '入金金额'], 3)
                    fund_outflow_col = safe_get_column(df_fund, ['出金', '出金金额'], 4)
                    
                    trade_month_col = safe_get_column(df_trade, ['月份', '月份'], 0)
                    trade_investor_col = safe_get_column(df_trade, ['投资者代码', '客户代码', '投资者'], 3)
                    trade_pnl_col = safe_get_column(df_trade, ['平仓盈亏'], 4)
                    trade_option_income_col = safe_get_column(df_trade, ['期权权利金收入', '权利金收入'], 5)
                    trade_option_expense_col = safe_get_column(df_trade, ['期权权利金支出', '权利金支出'], 6)
                    
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
                    
                    if df_fund.empty or fund_equity_col is None or fund_investor_col is None:
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
                        
                        df_fund['年月'] = df_fund[fund_month_col].apply(format_fund_month)
                        df_fund = df_fund.dropna(subset=['年月', fund_equity_col, fund_investor_col])
                        
                        bins = [0, 20000, 100000, 500000, 1000000, 3000000, 10000000, 30000000, float('inf')]
                        labels = ['2万以下', '2万（含）-10万', '10万（含）-50万', '50万（含）-100万', 
                                  '100万（含）-300万', '300万（含）-1000万', '1000万（含）-3000万', '3000万（含）以上']
                        
                        if selected_month_filter != '全部':
                            filtered_fund = df_fund[df_fund['年月'] == selected_month_raw]
                        else:
                            filtered_fund = df_fund
                        
                        pnl_by_investor = {}
                        if not df_trade.empty and trade_month_col is not None and trade_investor_col is not None:
                            df_trade['年月'] = df_trade[trade_month_col].apply(format_fund_month)
                            df_trade = df_trade.dropna(subset=['年月', trade_investor_col])
                            
                            if selected_month_filter != '全部':
                                filtered_trade = df_trade[df_trade['年月'] == selected_month_raw]
                            else:
                                filtered_trade = df_trade
                            
                            if trade_pnl_col is not None:
                                filtered_trade['平仓盈亏计算'] = 0
                                filtered_trade['平仓盈亏计算'] += filtered_trade[trade_pnl_col].fillna(0)
                                if trade_option_income_col is not None:
                                    filtered_trade['平仓盈亏计算'] += filtered_trade[trade_option_income_col].fillna(0)
                                if trade_option_expense_col is not None:
                                    filtered_trade['平仓盈亏计算'] -= filtered_trade[trade_option_expense_col].fillna(0)
                                
                                trade_pnl_summary = filtered_trade.groupby(trade_investor_col)['平仓盈亏计算'].sum().reset_index()
                                trade_pnl_summary.columns = [fund_investor_col, '平仓盈亏']
                                pnl_by_investor = dict(zip(trade_pnl_summary[fund_investor_col], trade_pnl_summary['平仓盈亏']))
                        
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
    # 市场权益 vs 公司权益对比
    # ============================================================
    if '市场权益' in data_cache and not data_cache['市场权益'].empty:
        st.subheader("📊 市场权益 vs 公司权益对比")
        try:
            df_market_equity = data_cache['市场权益'].copy()
            df_market_equity = clean_dataframe(df_market_equity)
            
            month_col = safe_get_column(df_market_equity, ['月份', '月份'], 0)
            equity_col = safe_get_column(df_market_equity, ['市场权益', '市场权益'], 1)
            
            if month_col is None:
                st.warning("未找到'月份'列，请检查数据格式")
            elif equity_col is None:
                st.warning("未找到'市场权益'列，请检查数据格式")
            else:
                market_equity_df = df_market_equity[[month_col, equity_col]].copy()
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
                market_equity_df = market_equity_df.dropna(subset=['年月'])
                
                if market_equity_df.empty:
                    st.info("暂无有效的市场权益数据（月份格式需为202601这样的6位数字）")
                else:
                    company_equity_list = []
                    
                    df_fund_current = data_cache.get('资金对账表-月', pd.DataFrame())
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
                    
                    df_fund_last = data_cache.get('上一年资金对账表-月', pd.DataFrame())
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
                                title='市场权益（百亿元）vs 公司权益（亿元）对比（最近12个月）',
                                labels={'权益': '权益', '月份': '月份'},
                                text_auto='.2f',
                                color_discrete_map={
                                    '市场权益（百亿元）': '#2E86C1',
                                    '公司权益（亿元）': '#F39C12'
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
                                display_df = chart_df[['年月显示', '市场权益（百亿元）', '公司权益（亿元）']].copy()
                                display_df = display_df.sort_values('年月显示')
                                display_df.columns = ['月份', '市场权益（百亿元）', '公司权益（亿元）']
                                
                                for col in ['市场权益（百亿元）', '公司权益（亿元）']:
                                    display_df[col] = display_df[col].apply(
                                        lambda x: f"{x:.2f}" if pd.notna(x) else '-'
                                    )
                                
                                st.dataframe(display_df, use_container_width=True, hide_index=True)
                        else:
                            st.info("没有可用于绘图的数据")
                            
        except Exception as e:
            st.warning(f"加载市场权益数据时出错: {e}")
            st.exception(e)

except Exception as e:
    st.error(f"❌ 处理数据时出错: {e}")
    st.exception(e)

import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
import datetime
import requests

st.set_page_config(page_title="Dashboard", layout="wide")

st.title("📊 交易数据看板")

# ============================================================
# 配置 - 你的GitHub信息
# ============================================================
GITHUB_USERNAME = "WZH0821"
GITHUB_REPO = "yunyingqingkuang-"
GITHUB_BRANCH = "main"
EXCEL_FILENAME = "data1.xlsx"

# 构建GitHub原始文件URL
GITHUB_FILE_URL = f"https://raw.githubusercontent.com/{GITHUB_USERNAME}/{GITHUB_REPO}/{GITHUB_BRANCH}/{EXCEL_FILENAME}"

# ============================================================
# 缓存数据加载函数
# ============================================================
@st.cache_data(ttl=3600)
def load_all_data_from_github():
    """
    从GitHub仓库加载所有Excel数据
    """
    try:
        # 1. 从GitHub下载文件
        with st.spinner(f"📥 正在从GitHub下载 {EXCEL_FILENAME}..."):
            response = requests.get(GITHUB_FILE_URL, timeout=30)
            response.raise_for_status()
        
        # 2. 将下载的内容转为BytesIO对象
        excel_data = BytesIO(response.content)
        
        # 3. 定义需要加载的所有sheet名称
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
        
        # 4. 加载所有sheet
        data = {}
        for key, sheet_name in sheets_config.items():
            try:
                excel_data.seek(0)
                df = pd.read_excel(excel_data, sheet_name=sheet_name, header=0)
                # 清理数据
                df = df.loc[:, ~df.columns.isna()]
                df = df.loc[:, df.columns != '']
                df = df.loc[:, ~df.columns.duplicated()]
                df = df.dropna(axis=1, how='all')
                data[key] = df
            except Exception as e:
                st.warning(f"⚠️ 加载sheet '{sheet_name}' 失败: {e}")
                data[key] = None
        
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
if df_vol_market is None:
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
# 以下是你原有的数据分析代码（完全保持不变）
# ============================================================
try:
    df = df_vol_market.copy()
    df = df.loc[:, ~df.columns.isna()]
    df = df.loc[:, df.columns != '']
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.dropna(axis=1, how='all')
    
    st.success(f"✅ 数据加载成功！共 {len(df)} 行，{len(df.columns)} 列")
    
    st.dataframe(df.head(10))
    
    with st.expander("📌 查看所有列名"):
        st.write(df.columns.tolist())
    
    # ============================================================
    # 2. 筛选器
    # ============================================================
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    filter_cols = [col for col in df.columns if col not in numeric_cols]
    
    selected_filters = {}
    cols_per_row = 3
    col_count = len(filter_cols)
    
    if col_count > 0:
        rows = (col_count + cols_per_row - 1) // cols_per_row
        for row in range(rows):
            cols = st.columns(cols_per_row)
            for i in range(cols_per_row):
                idx = row * cols_per_row + i
                if idx < col_count:
                    col_name = filter_cols[idx]
                    with cols[i]:
                        unique_vals = df[col_name].dropna().unique().tolist()
                        if len(unique_vals) > 0:
                            selected = st.multiselect(
                                f"筛选 {col_name}",
                                options=unique_vals,
                                default=[],
                                key=f"filter_{col_name}"
                            )
                            if selected:
                                selected_filters[col_name] = selected
    
    if selected_filters:
        filtered_df = df.copy()
        for col, vals in selected_filters.items():
            if vals:
                filtered_df = filtered_df[filtered_df[col].isin(vals)]
    else:
        filtered_df = df.copy()
    
    st.subheader("📊 筛选后数据")
    
    # ============================================================
    # 3. 选择数据类型（成交量、成交额、持仓量）
    # ============================================================
    data_type = st.radio(
        "选择数据类型",
        options=['成交量', '成交额', '持仓量'],
        horizontal=True,
        key="data_type"
    )
    
    # 根据数据类型配置
    if data_type == '成交量':
        market_sheet = '成交量-市场'
        company_sheet = '成交量-公司'
        market_divide = 100000000
        market_unit = '亿手'
        market_title = '（亿手）'
        market_yaxis = '成交量（亿手）'
        company_divide = 10000
        company_unit = '万手'
        company_title = '（万手）'
        company_yaxis = '成交量（万手）'
        metric_name = '成交量'
    elif data_type == '成交额':
        market_sheet = '成交额-市场'
        company_sheet = '成交额-公司'
        market_divide = 10000
        market_unit = '万亿元'
        market_title = '（万亿元）'
        market_yaxis = '成交额（万亿元）'
        company_divide = 100000000
        company_unit = '亿元'
        company_title = '（亿元）'
        company_yaxis = '成交额（亿元）'
        metric_name = '成交额'
    else:  # 持仓量
        market_sheet = '持仓量-市场'
        company_sheet = '持仓量-公司'
        market_divide = 1000000
        market_unit = '百万手'
        market_title = '（百万手）'
        market_yaxis = '持仓量（百万手）'
        company_divide = 10000
        company_unit = '万手'
        company_title = '（万手）'
        company_yaxis = '持仓量（万手）'
        metric_name = '持仓量'
    
    # 加载市场数据和公司数据
    df_detail = df_vol_market if data_type == '成交量' else (df_amt_market if data_type == '成交额' else df_oi_market)
    df_company_detail = df_vol_company if data_type == '成交量' else (df_amt_company if data_type == '成交额' else df_oi_company)
    
    # 清理数据
    df_detail = df_detail.loc[:, ~df_detail.columns.isna()]
    df_detail = df_detail.loc[:, df_detail.columns != '']
    df_detail = df_detail.loc[:, ~df_detail.columns.duplicated()]
    df_detail = df_detail.dropna(axis=1, how='all')
    
    df_company_detail = df_company_detail.loc[:, ~df_company_detail.columns.isna()]
    df_company_detail = df_company_detail.loc[:, df_company_detail.columns != '']
    df_company_detail = df_company_detail.loc[:, ~df_company_detail.columns.duplicated()]
    df_company_detail = df_company_detail.dropna(axis=1, how='all')
    
    st.success(f"✅ {data_type}数据加载成功！共 {len(df_detail)} 行，{len(df_detail.columns)} 列")
    st.success(f"✅ {data_type}公司数据加载成功！共 {len(df_company_detail)} 行，{len(df_company_detail.columns)} 列")
    
    st.subheader(f"📋 {data_type}数据预览")
    st.dataframe(df_detail.head(10))
    
    with st.expander("📌 查看所有列名"):
        st.write(df_detail.columns.tolist())
    
    # ============================================================
    # 获取日期列信息
    # ============================================================
    date_cols = df_detail.columns[6:].tolist()
    
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
        except Exception as e:
            date_labels[col] = str(col)
            date_info[col] = {'year': None, 'month': None}
    
    month_names = {1:'1月',2:'2月',3:'3月',4:'4月',5:'5月',6:'6月',7:'7月',8:'8月',9:'9月',10:'10月',11:'11月',12:'12月'}
    
    # ============================================================
    # 4. 各交易所柱状图（市场 + 公司）
    # ============================================================
    st.subheader(f"📊 各交易所{metric_name}对比（市场）{market_title}")
    
    # ---- 交易所独立筛选器 ----
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
    
    # 获取选中的列
    if time_dimension_exchange == '月度':
        selected_cols_exchange = [selected_key_exchange]
        selected_label_exchange = option_labels_exchange.get(selected_key_exchange, str(selected_key_exchange))
    else:
        selected_cols_exchange = value_cols_map_exchange.get(selected_key_exchange, [])
        selected_label_exchange = option_labels_exchange.get(selected_key_exchange, str(selected_key_exchange))
    
    # 计算上期（交易所）
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
        else:
            prev_cols_exchange = []
        if len(col_str) == 6 and col_str.isdigit():
            last_year_key = f"{year - 1}{month:02d}"
            last_year_key = int(last_year_key) if last_year_key.isdigit() else last_year_key
            last_year_cols_exchange = [last_year_key] if last_year_key in df_detail.columns else []
        else:
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
        prev_cols_exchange = value_cols_map_exchange.get(prev_key, []) if 'value_cols_map_exchange' in dir() else []
        last_year_key = f"{year - 1}{q}"
        last_year_cols_exchange = value_cols_map_exchange.get(last_year_key, []) if 'value_cols_map_exchange' in dir() else []
    else:
        prev_year = int(selected_key_exchange) - 1
        prev_key = str(prev_year)
        prev_cols_exchange = value_cols_map_exchange.get(prev_key, []) if 'value_cols_map_exchange' in dir() else []
        last_year_cols_exchange = []
    
    # ===== 市场交易所数据 =====
    exchanges = df_detail['交易所'].unique().tolist()
    exchange_compare = []
    
    current_sum_exchange = 0
    for col in selected_cols_exchange:
        current_sum_exchange += df_detail[col].sum()
    total_current_exchange = current_sum_exchange / market_divide
    
    total_prev_exchange = None
    if prev_cols_exchange:
        prev_sum = 0
        for col in prev_cols_exchange:
            if col in df_detail.columns:
                prev_sum += df_detail[col].sum()
        total_prev_exchange = prev_sum / market_divide
    
    total_last_year_exchange = None
    if last_year_cols_exchange:
        last_sum = 0
        for col in last_year_cols_exchange:
            if col in df_detail.columns:
                last_sum += df_detail[col].sum()
        total_last_year_exchange = last_sum / market_divide
    
    market_mom_exchange = None
    market_yoy_exchange = None
    if total_prev_exchange is not None and total_prev_exchange != 0:
        market_mom_exchange = (total_current_exchange - total_prev_exchange) / total_prev_exchange * 100
    if total_last_year_exchange is not None and total_last_year_exchange != 0:
        market_yoy_exchange = (total_current_exchange - total_last_year_exchange) / total_last_year_exchange * 100
    
    for exchange in exchanges:
        exchange_df = df_detail[df_detail['交易所'] == exchange]
        row = {'交易所': exchange}
        current_val = 0
        for col in selected_cols_exchange:
            current_val += exchange_df[col].sum()
        current_val = current_val / market_divide
        row['本月'] = current_val
        
        if prev_cols_exchange:
            prev_val = 0
            for col in prev_cols_exchange:
                if col in df_detail.columns:
                    prev_val += exchange_df[col].sum()
            prev_val = prev_val / market_divide
            row['上月'] = prev_val
            if prev_val != 0:
                row['环比'] = (current_val - prev_val) / prev_val * 100
            else:
                row['环比'] = None
        else:
            row['上月'] = None
            row['环比'] = None
        
        if last_year_cols_exchange:
            last_val = 0
            for col in last_year_cols_exchange:
                if col in df_detail.columns:
                    last_val += exchange_df[col].sum()
            last_val = last_val / market_divide
            row['上年同期'] = last_val
            if last_val != 0:
                row['同比'] = (current_val - last_val) / last_val * 100
            else:
                row['同比'] = None
        else:
            row['上年同期'] = None
            row['同比'] = None
        
        exchange_compare.append(row)
    
    exchange_df_plot = pd.DataFrame(exchange_compare)
    
    # ===== 公司交易所数据 =====
    exchange_compare_company = []
    
    current_sum_exchange_company = 0
    for col in selected_cols_exchange:
        current_sum_exchange_company += df_company_detail[col].sum()
    total_current_exchange_company = current_sum_exchange_company / company_divide
    
    total_prev_exchange_company = None
    if prev_cols_exchange:
        prev_sum = 0
        for col in prev_cols_exchange:
            if col in df_company_detail.columns:
                prev_sum += df_company_detail[col].sum()
        total_prev_exchange_company = prev_sum / company_divide
    
    total_last_year_exchange_company = None
    if last_year_cols_exchange:
        last_sum = 0
        for col in last_year_cols_exchange:
            if col in df_company_detail.columns:
                last_sum += df_company_detail[col].sum()
        total_last_year_exchange_company = last_sum / company_divide
    
    market_mom_exchange_company = None
    market_yoy_exchange_company = None
    if total_prev_exchange_company is not None and total_prev_exchange_company != 0:
        market_mom_exchange_company = (total_current_exchange_company - total_prev_exchange_company) / total_prev_exchange_company * 100
    if total_last_year_exchange_company is not None and total_last_year_exchange_company != 0:
        market_yoy_exchange_company = (total_current_exchange_company - total_last_year_exchange_company) / total_last_year_exchange_company * 100
    
    for exchange in exchanges:
        exchange_df = df_company_detail[df_company_detail['交易所'] == exchange]
        row = {'交易所': exchange}
        current_val = 0
        for col in selected_cols_exchange:
            current_val += exchange_df[col].sum()
        current_val = current_val / company_divide
        row['本月'] = current_val
        
        if prev_cols_exchange:
            prev_val = 0
            for col in prev_cols_exchange:
                if col in df_company_detail.columns:
                    prev_val += exchange_df[col].sum()
            prev_val = prev_val / company_divide
            row['上月'] = prev_val
            if prev_val != 0:
                row['环比'] = (current_val - prev_val) / prev_val * 100
            else:
                row['环比'] = None
        else:
            row['上月'] = None
            row['环比'] = None
        
        if last_year_cols_exchange:
            last_val = 0
            for col in last_year_cols_exchange:
                if col in df_company_detail.columns:
                    last_val += exchange_df[col].sum()
            last_val = last_val / company_divide
            row['上年同期'] = last_val
            if last_val != 0:
                row['同比'] = (current_val - last_val) / last_val * 100
            else:
                row['同比'] = None
        else:
            row['上年同期'] = None
            row['同比'] = None
        
        exchange_compare_company.append(row)
    
    exchange_df_plot_company = pd.DataFrame(exchange_compare_company)
    
    # ===== 市场交易所柱状图 =====
    value_cols_exchange = ['本月']
    if prev_cols_exchange:
        value_cols_exchange.append('上月')
    if last_year_cols_exchange:
        value_cols_exchange.append('上年同期')
    
    exchange_melted = exchange_df_plot.melt(
        id_vars=['交易所'],
        value_vars=value_cols_exchange,
        var_name='期间',
        value_name=metric_name
    )
    exchange_melted = exchange_melted.dropna(subset=[metric_name])
    
    period_order = ['上月', '本月', '上年同期']
    exchange_melted['期间'] = pd.Categorical(
        exchange_melted['期间'],
        categories=[p for p in period_order if p in exchange_melted['期间'].unique()],
        ordered=True
    )
    exchange_melted = exchange_melted.sort_values('期间')
    
    if not exchange_melted.empty:
        fig_exchange = px.bar(
            exchange_melted,
            x='交易所',
            y=metric_name,
            color='期间',
            barmode='group',
            title=f'各交易所{metric_name}对比（市场）- {selected_label_exchange}',
            labels={metric_name: market_yaxis, '交易所': '交易所'},
            text_auto='.2f',
            color_discrete_map={
                '本月': '#2E86C1',
                '上月': '#F39C12',
                '上年同期': '#28B463'
            }
        )
        fig_exchange.update_layout(
            title_font=dict(size=18, color='#1A5276'),
            font=dict(size=13),
            bargap=0.25,
            bargroupgap=0.15,
            plot_bgcolor='#F8F9F9',
            paper_bgcolor='white',
            legend_title_text='',
            yaxis=dict(tickformat='.2f', title=market_yaxis)
        )
        fig_exchange.update_traces(
            texttemplate='%{y:.2f}',
            textfont=dict(size=11, color='black', family='Arial Black'),
            textposition='outside'
        )
        st.plotly_chart(fig_exchange, use_container_width=True)
        
        # 市场交易所环比同比综合表
        st.subheader("📊 交易所环比同比综合表（市场）")
        table_data_exchange = []
        table_data_exchange.append({
            '维度': '市场',
            f'{metric_name}（{market_unit}）': f"{total_current_exchange:.2f}",
            '环比（%）': f"{market_mom_exchange:+.2f}" if market_mom_exchange is not None else '-',
            '同比（%）': f"{market_yoy_exchange:+.2f}" if market_yoy_exchange is not None else '-'
        })
        for _, row in exchange_df_plot.iterrows():
            mom = row.get('环比')
            yoy = row.get('同比')
            table_data_exchange.append({
                '维度': row['交易所'],
                f'{metric_name}（{market_unit}）': f"{row['本月']:.2f}",
                '环比（%）': f"{mom:+.2f}" if mom is not None and not pd.isna(mom) else '-',
                '同比（%）': f"{yoy:+.2f}" if yoy is not None and not pd.isna(yoy) else '-'
            })
        table_df_exchange = pd.DataFrame(table_data_exchange)
        st.dataframe(table_df_exchange, use_container_width=True, hide_index=True)
    
    # ===== 公司交易所柱状图 =====
    st.subheader(f"📊 各交易所{metric_name}对比（公司）{company_title}")
    
    exchange_melted_company = exchange_df_plot_company.melt(
        id_vars=['交易所'],
        value_vars=value_cols_exchange,
        var_name='期间',
        value_name=metric_name
    )
    exchange_melted_company = exchange_melted_company.dropna(subset=[metric_name])
    
    exchange_melted_company['期间'] = pd.Categorical(
        exchange_melted_company['期间'],
        categories=[p for p in period_order if p in exchange_melted_company['期间'].unique()],
        ordered=True
    )
    exchange_melted_company = exchange_melted_company.sort_values('期间')
    
    if not exchange_melted_company.empty:
        fig_exchange_company = px.bar(
            exchange_melted_company,
            x='交易所',
            y=metric_name,
            color='期间',
            barmode='group',
            title=f'各交易所{metric_name}对比（公司）- {selected_label_exchange}',
            labels={metric_name: company_yaxis, '交易所': '交易所'},
            text_auto='.2f',
            color_discrete_map={
                '本月': '#2E86C1',
                '上月': '#F39C12',
                '上年同期': '#28B463'
            }
        )
        fig_exchange_company.update_layout(
            title_font=dict(size=18, color='#1A5276'),
            font=dict(size=13),
            bargap=0.25,
            bargroupgap=0.15,
            plot_bgcolor='#F8F9F9',
            paper_bgcolor='white',
            legend_title_text='',
            yaxis=dict(tickformat='.2f', title=company_yaxis)
        )
        fig_exchange_company.update_traces(
            texttemplate='%{y:.2f}',
            textfont=dict(size=11, color='black', family='Arial Black'),
            textposition='outside'
        )
        st.plotly_chart(fig_exchange_company, use_container_width=True)
        
        # 公司交易所环比同比综合表
        st.subheader("📊 交易所环比同比综合表（公司）")
        table_data_exchange_company = []
        table_data_exchange_company.append({
            '维度': '市场',
            f'{metric_name}（{company_unit}）': f"{total_current_exchange_company:.2f}",
            '环比（%）': f"{market_mom_exchange_company:+.2f}" if market_mom_exchange_company is not None else '-',
            '同比（%）': f"{market_yoy_exchange_company:+.2f}" if market_yoy_exchange_company is not None else '-'
        })
        for _, row in exchange_df_plot_company.iterrows():
            mom = row.get('环比')
            yoy = row.get('同比')
            table_data_exchange_company.append({
                '维度': row['交易所'],
                f'{metric_name}（{company_unit}）': f"{row['本月']:.2f}",
                '环比（%）': f"{mom:+.2f}" if mom is not None and not pd.isna(mom) else '-',
                '同比（%）': f"{yoy:+.2f}" if yoy is not None and not pd.isna(yoy) else '-'
            })
        table_df_exchange_company = pd.DataFrame(table_data_exchange_company)
        st.dataframe(table_df_exchange_company, use_container_width=True, hide_index=True)
    
    # ============================================================
    # 5. 公司占市场比重折线图（整体）
    # ============================================================
    st.subheader("📊 公司占市场比重（整体）")
    
    try:
        date_cols_all = df_vol_market.columns[6:].tolist()
        filtered_cols = []
        for col in date_cols_all:
            col_str = str(col)
            if len(col_str) == 6 and col_str.isdigit():
                year = int(col_str[:4])
                if year >= 2024:
                    filtered_cols.append(col)
        
        data_by_month = {}
        for col in filtered_cols:
            col_str = str(col)
            if len(col_str) == 6 and col_str.isdigit():
                year = int(col_str[:4])
                month = int(col_str[4:6])
                if month not in data_by_month:
                    data_by_month[month] = {}
                if year not in data_by_month[month]:
                    data_by_month[month][year] = {}
                
                vol_market = df_vol_market[col].sum()
                vol_company = df_vol_company[col].sum()
                amt_market = df_amt_market[col].sum()
                amt_company = df_amt_company[col].sum()
                oi_market = df_oi_market[col].sum()
                oi_company = df_oi_company[col].sum()
                
                data_by_month[month][year]['成交量'] = (vol_company / (vol_market * 2) * 100) if vol_market != 0 else 0
                data_by_month[month][year]['成交额'] = (amt_company / 100000000 / (amt_market * 2) * 100) if amt_market != 0 else 0
                data_by_month[month][year]['持仓量'] = (oi_company / (oi_market * 2) * 100) if oi_market != 0 else 0
        
        all_years = sorted(set([y for month_data in data_by_month.values() for y in month_data.keys()]))
        latest_year = max(all_years) if all_years else 2024
        
        plot_data = []
        for month in sorted(data_by_month.keys()):
            for year in sorted(data_by_month[month].keys()):
                for metric in ['成交量', '成交额', '持仓量']:
                    value = data_by_month[month][year].get(metric, 0)
                    if value > 0:
                        plot_data.append({
                            '月份': month_names.get(month, str(month)),
                            '年份': str(year) + '年',
                            '指标': metric,
                            '占比（%）': value
                        })
        
        if plot_data:
            plot_df = pd.DataFrame(plot_data)
            selected_metric_global = st.selectbox(
                "选择查看指标",
                options=['成交量', '成交额', '持仓量'],
                key="metric_selector_global"
            )
            color_map = {
                str(latest_year - 2) + '年': '#2E86C1',
                str(latest_year - 1) + '年': '#F39C12',
                str(latest_year) + '年': '#28B463'
            }
            metric_df = plot_df[plot_df['指标'] == selected_metric_global]
            
            if not metric_df.empty:
                st.subheader(f"📈 {selected_metric_global}公司占市场比重")
                if selected_metric_global == '成交额':
                    decimal_places = 4
                else:
                    decimal_places = 3
                fig = px.line(
                    metric_df,
                    x='月份',
                    y='占比（%）',
                    color='年份',
                    title=f'{selected_metric_global}公司占市场比重',
                    labels={'月份': '月份', '占比（%）': '占比（%）', '年份': '年份'},
                    color_discrete_map=color_map,
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
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info(f"暂无{selected_metric_global}数据")
            
            with st.expander("📋 查看详细数据"):
                if selected_metric_global == '成交额':
                    dec = 4
                else:
                    dec = 3
                table_data = []
                for month in sorted(data_by_month.keys()):
                    row = {'月份': month_names.get(month, str(month))}
                    for year in sorted(data_by_month[month].keys()):
                        key = f"{year}年{selected_metric_global}"
                        val = data_by_month[month][year].get(selected_metric_global, 0)
                        row[key] = f"{val:.{dec}f}"
                    table_data.append(row)
                table_df = pd.DataFrame(table_data)
                st.dataframe(table_df, use_container_width=True)
        else:
            st.info("暂无公司占市场比重数据")
    except Exception as e:
        st.warning(f"无法加载公司/市场对比数据: {e}")
    
    # ============================================================
    # 6. 各交易所单独展示折线图
    # ============================================================
    st.subheader("📊 公司占市场比重 - 各交易所单独展示")
    
    try:
        date_cols_all = df_vol_market.columns[6:].tolist()
        filtered_cols = []
        for col in date_cols_all:
            col_str = str(col)
            if len(col_str) == 6 and col_str.isdigit():
                year = int(col_str[:4])
                if year >= 2024:
                    filtered_cols.append(col)
        
        exchanges_list = ['上期所', '能源中心', '郑商所', '大商所', '中金所', '广期所']
        data_by_month_year_exchange = {}
        for col in filtered_cols:
            col_str = str(col)
            if len(col_str) == 6 and col_str.isdigit():
                year = int(col_str[:4])
                month = int(col_str[4:6])
                if month not in data_by_month_year_exchange:
                    data_by_month_year_exchange[month] = {}
                if year not in data_by_month_year_exchange[month]:
                    data_by_month_year_exchange[month][year] = {}
                if '上期所' not in data_by_month_year_exchange[month][year]:
                    for ex in exchanges_list:
                        data_by_month_year_exchange[month][year][ex] = {}
                for ex in exchanges_list:
                    vol_market_ex = df_vol_market[df_vol_market['交易所'] == ex][col].sum()
                    vol_company_ex = df_vol_company[df_vol_company['交易所'] == ex][col].sum()
                    amt_market_ex = df_amt_market[df_amt_market['交易所'] == ex][col].sum()
                    amt_company_ex = df_amt_company[df_amt_company['交易所'] == ex][col].sum()
                    oi_market_ex = df_oi_market[df_oi_market['交易所'] == ex][col].sum()
                    oi_company_ex = df_oi_company[df_oi_company['交易所'] == ex][col].sum()
                    data_by_month_year_exchange[month][year][ex]['成交量'] = (vol_company_ex / (vol_market_ex * 2) * 100) if vol_market_ex != 0 else 0
                    data_by_month_year_exchange[month][year][ex]['成交额'] = (amt_company_ex / 100000000 / (amt_market_ex * 2) * 100) if amt_market_ex != 0 else 0
                    data_by_month_year_exchange[month][year][ex]['持仓量'] = (oi_company_ex / (oi_market_ex * 2) * 100) if oi_market_ex != 0 else 0
        
        all_years = sorted(set([y for month_data in data_by_month_year_exchange.values() for y in month_data.keys()]))
        latest_year = max(all_years) if all_years else 2024
        
        plot_data = []
        for month in sorted(data_by_month_year_exchange.keys()):
            for year in sorted(data_by_month_year_exchange[month].keys()):
                for ex in exchanges_list:
                    for metric in ['成交量', '成交额', '持仓量']:
                        value = data_by_month_year_exchange[month][year][ex].get(metric, 0)
                        if value > 0:
                            plot_data.append({
                                '月份': month_names.get(month, str(month)),
                                '年份': str(year) + '年',
                                '交易所': ex,
                                '指标': metric,
                                '占比（%）': value
                            })
        
        if plot_data:
            plot_df = pd.DataFrame(plot_data)
            selected_metric_exchange = st.selectbox(
                "选择查看指标",
                options=['成交量', '成交额', '持仓量'],
                key="metric_selector_exchange"
            )
            metric_df = plot_df[plot_df['指标'] == selected_metric_exchange]
            
            if not metric_df.empty:
                if selected_metric_exchange == '成交额':
                    decimal_places = 4
                else:
                    decimal_places = 3
                color_map = {
                    str(latest_year - 2) + '年': '#2E86C1',
                    str(latest_year - 1) + '年': '#F39C12',
                    str(latest_year) + '年': '#28B463'
                }
                for i, ex in enumerate(exchanges_list):
                    ex_df = metric_df[metric_df['交易所'] == ex]
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
                                margin=dict(l=40, r=40, t=50, b=30)
                            )
                            st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': False})
            else:
                st.info(f"暂无{selected_metric_exchange}数据")
            
            with st.expander("📋 查看详细数据"):
                if selected_metric_exchange == '成交额':
                    dec = 4
                else:
                    dec = 3
                for ex in exchanges_list:
                    st.subheader(f"{ex}")
                    table_data = []
                    for month in sorted(data_by_month_year_exchange.keys()):
                        row = {'月份': month_names.get(month, str(month))}
                        for year in sorted(data_by_month_year_exchange[month].keys()):
                            val = data_by_month_year_exchange[month][year][ex].get(selected_metric_exchange, 0)
                            row[f"{year}年"] = f"{val:.{dec}f}"
                        table_data.append(row)
                    table_df = pd.DataFrame(table_data)
                    st.dataframe(table_df, use_container_width=True)
        else:
            st.info("暂无公司占市场比重数据")
    except Exception as e:
        st.warning(f"无法加载公司/市场对比数据: {e}")
    
    # ============================================================
    # 7. 能源化工板块折线图
    # ============================================================
    st.subheader("📊 能源化工板块 - 公司占市场比重")
    
    try:
        df_vol_market_energy = df_vol_market[df_vol_market['板块'] == '能源化工']
        df_vol_company_energy = df_vol_company[df_vol_company['板块'] == '能源化工']
        df_amt_market_energy = df_amt_market[df_amt_market['板块'] == '能源化工']
        df_amt_company_energy = df_amt_company[df_amt_company['板块'] == '能源化工']
        df_oi_market_energy = df_oi_market[df_oi_market['板块'] == '能源化工']
        df_oi_company_energy = df_oi_company[df_oi_company['板块'] == '能源化工']
        
        date_cols_all = df_vol_market.columns[6:].tolist()
        filtered_cols = []
        for col in date_cols_all:
            col_str = str(col)
            if len(col_str) == 6 and col_str.isdigit():
                year = int(col_str[:4])
                if year >= 2024:
                    filtered_cols.append(col)
        
        data_by_month_energy = {}
        for col in filtered_cols:
            col_str = str(col)
            if len(col_str) == 6 and col_str.isdigit():
                year = int(col_str[:4])
                month = int(col_str[4:6])
                if month not in data_by_month_energy:
                    data_by_month_energy[month] = {}
                if year not in data_by_month_energy[month]:
                    data_by_month_energy[month][year] = {}
                
                vol_market = df_vol_market_energy[col].sum()
                vol_company = df_vol_company_energy[col].sum()
                amt_market = df_amt_market_energy[col].sum()
                amt_company = df_amt_company_energy[col].sum()
                oi_market = df_oi_market_energy[col].sum()
                oi_company = df_oi_company_energy[col].sum()
                
                data_by_month_energy[month][year]['成交量'] = (vol_company / (vol_market * 2) * 100) if vol_market != 0 else 0
                data_by_month_energy[month][year]['成交额'] = (amt_company / 100000000 / (amt_market * 2) * 100) if amt_market != 0 else 0
                data_by_month_energy[month][year]['持仓量'] = (oi_company / (oi_market * 2) * 100) if oi_market != 0 else 0
        
        all_years = sorted(set([y for month_data in data_by_month_energy.values() for y in month_data.keys()]))
        latest_year = max(all_years) if all_years else 2024
        
        plot_data_energy = []
        for month in sorted(data_by_month_energy.keys()):
            for year in sorted(data_by_month_energy[month].keys()):
                for metric in ['成交量', '成交额', '持仓量']:
                    value = data_by_month_energy[month][year].get(metric, 0)
                    if value > 0:
                        plot_data_energy.append({
                            '月份': month_names.get(month, str(month)),
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
            color_map = {
                str(latest_year - 2) + '年': '#2E86C1',
                str(latest_year - 1) + '年': '#F39C12',
                str(latest_year) + '年': '#28B463'
            }
            metric_df_energy = plot_df_energy[plot_df_energy['指标'] == selected_metric_energy]
            
            if not metric_df_energy.empty:
                st.subheader(f"📈 能源化工板块 - {selected_metric_energy}公司占市场比重")
                if selected_metric_energy == '成交额':
                    decimal_places = 4
                else:
                    decimal_places = 3
                fig = px.line(
                    metric_df_energy,
                    x='月份',
                    y='占比（%）',
                    color='年份',
                    title=f'能源化工板块 - {selected_metric_energy}公司占市场比重',
                    labels={'月份': '月份', '占比（%）': '占比（%）', '年份': '年份'},
                    color_discrete_map=color_map,
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
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info(f"暂无能源化工板块{selected_metric_energy}数据")
            
            with st.expander("📋 查看详细数据"):
                if selected_metric_energy == '成交额':
                    dec = 4
                else:
                    dec = 3
                table_data = []
                for month in sorted(data_by_month_energy.keys()):
                    row = {'月份': month_names.get(month, str(month))}
                    for year in sorted(data_by_month_energy[month].keys()):
                        key = f"{year}年{selected_metric_energy}"
                        val = data_by_month_energy[month][year].get(selected_metric_energy, 0)
                        row[key] = f"{val:.{dec}f}"
                    table_data.append(row)
                table_df_energy = pd.DataFrame(table_data)
                st.dataframe(table_df_energy, use_container_width=True)
        else:
            st.info("暂无能源化工板块公司占市场比重数据")
    except Exception as e:
        st.warning(f"无法加载能源化工板块对比数据: {e}")
    
    # ============================================================
    # 8. 各板块柱状图（市场 + 公司）
    # ============================================================
    
    # ---- 各板块独立数据类型筛选器 ----
    group_data_type = st.radio(
        "选择数据类型",
        options=['成交量', '成交额', '持仓量'],
        horizontal=True,
        key="group_data_type"
    )
    
    # 根据独立筛选器加载对应的数据
    if group_data_type == '成交量':
        group_market_df = df_vol_market.copy()
        group_company_df = df_vol_company.copy()
        group_market_divide = 100000000
        group_company_divide = 10000
        group_market_unit = '亿手'
        group_company_unit = '万手'
        group_market_title = '（亿手）'
        group_company_title = '（万手）'
        group_market_yaxis = '成交量（亿手）'
        group_company_yaxis = '成交量（万手）'
        group_metric_name = '成交量'
    elif group_data_type == '成交额':
        group_market_df = df_amt_market.copy()
        group_company_df = df_amt_company.copy()
        group_market_divide = 10000
        group_company_divide = 1
        group_market_unit = '万亿元'
        group_company_unit = '亿元'
        group_market_title = '（万亿元）'
        group_company_title = '（亿元）'
        group_market_yaxis = '成交额（万亿元）'
        group_company_yaxis = '成交额（亿元）'
        group_metric_name = '成交额'
    else:  # 持仓量
        group_market_df = df_oi_market.copy()
        group_company_df = df_oi_company.copy()
        group_market_divide = 100000
        group_company_divide = 10000
        group_market_unit = '百万手'
        group_company_unit = '万手'
        group_market_title = '（百万手）'
        group_company_title = '（万手）'
        group_market_yaxis = '持仓量（百万手）'
        group_company_yaxis = '持仓量（万手）'
        group_metric_name = '持仓量'
    
    # 清理数据
    group_market_df = group_market_df.loc[:, ~group_market_df.columns.isna()]
    group_market_df = group_market_df.loc[:, group_market_df.columns != '']
    group_market_df = group_market_df.loc[:, ~group_market_df.columns.duplicated()]
    group_market_df = group_market_df.dropna(axis=1, how='all')
    
    group_company_df = group_company_df.loc[:, ~group_company_df.columns.isna()]
    group_company_df = group_company_df.loc[:, group_company_df.columns != '']
    group_company_df = group_company_df.loc[:, ~group_company_df.columns.duplicated()]
    group_company_df = group_company_df.dropna(axis=1, how='all')
    
    st.subheader(f"📊 各板块{group_metric_name}对比（市场）{group_market_title}")
    
    # ---- 板块独立筛选器 ----
    col1, col2 = st.columns(2)
    with col1:
        time_dimension_group = st.radio(
            "选择时间维度",
            options=['月度', '季度', '年度'],
            horizontal=True,
            key="time_dimension_group"
        )
    with col2:
        if time_dimension_group == '月度':
            options_group = sorted(date_cols, reverse=True)
            option_labels_group = {col: date_labels.get(col, str(col)) for col in options_group}
            value_cols_map_group = None
        elif time_dimension_group == '季度':
            quarter_map_group = {}
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
                    if key not in quarter_map_group:
                        quarter_map_group[key] = []
                    quarter_map_group[key].append(col)
            options_group = sorted(quarter_map_group.keys(), reverse=True)
            option_labels_group = {k: f"{k[:4]}年{k[4:]}" for k in options_group}
            value_cols_map_group = quarter_map_group
        else:
            year_map_group = {}
            for col in date_cols:
                info = date_info.get(col, {})
                year = info.get('year')
                if year is not None:
                    key = str(year)
                    if key not in year_map_group:
                        year_map_group[key] = []
                    year_map_group[key].append(col)
            options_group = sorted(year_map_group.keys(), reverse=True)
            option_labels_group = {k: f"{k}年" for k in options_group}
            value_cols_map_group = year_map_group
        
        selected_key_group = st.selectbox(
            f"选择{time_dimension_group}",
            options=options_group,
            format_func=lambda x: option_labels_group.get(x, str(x)),
            key="time_selector_group"
        )
    
    # 获取选中的列
    if time_dimension_group == '月度':
        selected_cols_group = [selected_key_group]
        selected_label_group = option_labels_group.get(selected_key_group, str(selected_key_group))
    else:
        selected_cols_group = value_cols_map_group.get(selected_key_group, [])
        selected_label_group = option_labels_group.get(selected_key_group, str(selected_key_group))
    
    # ===== 计算上期（环比）和上年同期（同比） =====
    if time_dimension_group == '月度':
        col_str = str(selected_key_group)
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
            prev_cols_group = [prev_key] if prev_key in group_market_df.columns else []
            last_year_key = f"{year - 1}{month:02d}"
            last_year_key = int(last_year_key) if last_year_key.isdigit() else last_year_key
            last_year_cols_group = [last_year_key] if last_year_key in group_market_df.columns else []
        else:
            prev_cols_group = []
            last_year_cols_group = []
    elif time_dimension_group == '季度':
        key_str = str(selected_key_group)
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
        prev_cols_group = value_cols_map_group.get(prev_key, []) if value_cols_map_group is not None else []
        last_year_key = f"{year - 1}{q}"
        last_year_cols_group = value_cols_map_group.get(last_year_key, []) if value_cols_map_group is not None else []
    else:
        year_num = int(selected_key_group)
        prev_key = str(year_num - 1)
        prev_cols_group = value_cols_map_group.get(prev_key, []) if value_cols_map_group is not None else []
        last_year_cols_group = []
    
    # 获取所有板块
    if '板块' in group_market_df.columns:
        groups = group_market_df['板块'].unique().tolist()
        group_col = '板块'
    else:
        groups = group_market_df['交易所'].unique().tolist()
        group_col = '交易所'
    
    # ===== 市场板块数据 =====
    group_compare = []
    for group in groups:
        group_df = group_market_df[group_market_df[group_col] == group]
        row = {group_col: group}
        
        current_val = 0
        for col in selected_cols_group:
            if col in group_df.columns:
                current_val += group_df[col].sum()
        current_val = current_val / group_market_divide
        row['本月'] = current_val
        
        if prev_cols_group:
            prev_val = 0
            for col in prev_cols_group:
                if col in group_df.columns:
                    prev_val += group_df[col].sum()
            prev_val = prev_val / group_market_divide
            row['上月'] = prev_val
            if prev_val != 0:
                row['环比'] = (current_val - prev_val) / prev_val * 100
            else:
                row['环比'] = None
        else:
            row['上月'] = None
            row['环比'] = None
        
        if last_year_cols_group:
            last_val = 0
            for col in last_year_cols_group:
                if col in group_df.columns:
                    last_val += group_df[col].sum()
            last_val = last_val / group_market_divide
            row['上年同期'] = last_val
            if last_val != 0:
                row['同比'] = (current_val - last_val) / last_val * 100
            else:
                row['同比'] = None
        else:
            row['上年同期'] = None
            row['同比'] = None
        
        group_compare.append(row)
    
    group_df_plot = pd.DataFrame(group_compare)
    
    # ===== 公司板块数据 =====
    group_compare_company = []
    for group in groups:
        group_df = group_company_df[group_company_df[group_col] == group]
        row = {group_col: group}
        
        current_val = 0
        for col in selected_cols_group:
            if col in group_df.columns:
                current_val += group_df[col].sum()
        current_val = current_val / group_company_divide
        row['本月'] = current_val
        
        if prev_cols_group:
            prev_val = 0
            for col in prev_cols_group:
                if col in group_df.columns:
                    prev_val += group_df[col].sum()
            prev_val = prev_val / group_company_divide
            row['上月'] = prev_val
            if prev_val != 0:
                row['环比'] = (current_val - prev_val) / prev_val * 100
            else:
                row['环比'] = None
        else:
            row['上月'] = None
            row['环比'] = None
        
        if last_year_cols_group:
            last_val = 0
            for col in last_year_cols_group:
                if col in group_df.columns:
                    last_val += group_df[col].sum()
            last_val = last_val / group_company_divide
            row['上年同期'] = last_val
            if last_val != 0:
                row['同比'] = (current_val - last_val) / last_val * 100
            else:
                row['同比'] = None
        else:
            row['上年同期'] = None
            row['同比'] = None
        
        group_compare_company.append(row)
    
    group_df_plot_company = pd.DataFrame(group_compare_company)
    
    # ===== 市场板块柱状图 =====
    value_cols_group = ['本月']
    if prev_cols_group:
        value_cols_group.append('上月')
    if last_year_cols_group:
        value_cols_group.append('上年同期')
    
    group_melted = group_df_plot.melt(
        id_vars=[group_col],
        value_vars=value_cols_group,
        var_name='期间',
        value_name=group_metric_name
    )
    group_melted = group_melted.dropna(subset=[group_metric_name])
    
    period_order = ['上月', '本月', '上年同期']
    group_melted['期间'] = pd.Categorical(
        group_melted['期间'],
        categories=[p for p in period_order if p in group_melted['期间'].unique()],
        ordered=True
    )
    group_melted = group_melted.sort_values('期间')
    
    if not group_melted.empty:
        fig_group = px.bar(
            group_melted,
            x=group_col,
            y=group_metric_name,
            color='期间',
            barmode='group',
            title=f'各板块{group_metric_name}对比（市场）- {selected_label_group}',
            labels={group_metric_name: group_market_yaxis, group_col: '板块'},
            text_auto='.2f',
            color_discrete_map={
                '本月': '#2E86C1',
                '上月': '#F39C12',
                '上年同期': '#28B463'
            }
        )
        fig_group.update_layout(
            title_font=dict(size=18, color='#1A5276'),
            font=dict(size=13),
            bargap=0.25,
            bargroupgap=0.15,
            plot_bgcolor='#F8F9F9',
            paper_bgcolor='white',
            legend_title_text='',
            yaxis=dict(tickformat='.2f', title=group_market_yaxis)
        )
        fig_group.update_traces(
            texttemplate='%{y:.2f}',
            textfont=dict(size=11, color='black', family='Arial Black'),
            textposition='outside'
        )
        st.plotly_chart(fig_group, use_container_width=True)
        
        # 市场板块环比同比综合表
        st.subheader(f"📊 板块环比同比综合表（市场）")
        
        current_sum_group = 0
        for col in selected_cols_group:
            if col in group_market_df.columns:
                current_sum_group += group_market_df[col].sum()
        total_current_group = current_sum_group / group_market_divide
        
        total_prev_group = None
        if prev_cols_group:
            prev_sum = 0
            for col in prev_cols_group:
                if col in group_market_df.columns:
                    prev_sum += group_market_df[col].sum()
            total_prev_group = prev_sum / group_market_divide
        
        total_last_year_group = None
        if last_year_cols_group:
            last_sum = 0
            for col in last_year_cols_group:
                if col in group_market_df.columns:
                    last_sum += group_market_df[col].sum()
            total_last_year_group = last_sum / group_market_divide
        
        table_data_group = []
        table_data_group.append({
            '维度': '市场',
            f'{group_metric_name}（{group_market_unit}）': f"{total_current_group:.2f}",
            '环比（%）': f"{(total_current_group - total_prev_group) / total_prev_group * 100:+.2f}" if total_prev_group is not None and total_prev_group != 0 else '-',
            '同比（%）': f"{(total_current_group - total_last_year_group) / total_last_year_group * 100:+.2f}" if total_last_year_group is not None and total_last_year_group != 0 else '-'
        })
        for _, row in group_df_plot.iterrows():
            mom = row.get('环比')
            yoy = row.get('同比')
            table_data_group.append({
                '维度': row[group_col],
                f'{group_metric_name}（{group_market_unit}）': f"{row['本月']:.2f}",
                '环比（%）': f"{mom:+.2f}" if mom is not None and not pd.isna(mom) else '-',
                '同比（%）': f"{yoy:+.2f}" if yoy is not None and not pd.isna(yoy) else '-'
            })
        table_df_group = pd.DataFrame(table_data_group)
        st.dataframe(table_df_group, use_container_width=True, hide_index=True)
    
    # ===== 公司板块柱状图 =====
    st.subheader(f"📊 各板块{group_metric_name}对比（公司）{group_company_title}")
    
    group_melted_company = group_df_plot_company.melt(
        id_vars=[group_col],
        value_vars=value_cols_group,
        var_name='期间',
        value_name=group_metric_name
    )
    group_melted_company = group_melted_company.dropna(subset=[group_metric_name])
    
    group_melted_company['期间'] = pd.Categorical(
        group_melted_company['期间'],
        categories=[p for p in period_order if p in group_melted_company['期间'].unique()],
        ordered=True
    )
    group_melted_company = group_melted_company.sort_values('期间')
    
    if not group_melted_company.empty:
        fig_group_company = px.bar(
            group_melted_company,
            x=group_col,
            y=group_metric_name,
            color='期间',
            barmode='group',
            title=f'各板块{group_metric_name}对比（公司）- {selected_label_group}',
            labels={group_metric_name: group_company_yaxis, group_col: '板块'},
            text_auto='.2f',
            color_discrete_map={
                '本月': '#2E86C1',
                '上月': '#F39C12',
                '上年同期': '#28B463'
            }
        )
        fig_group_company.update_layout(
            title_font=dict(size=18, color='#1A5276'),
            font=dict(size=13),
            bargap=0.25,
            bargroupgap=0.15,
            plot_bgcolor='#F8F9F9',
            paper_bgcolor='white',
            legend_title_text='',
            yaxis=dict(tickformat='.2f', title=group_company_yaxis)
        )
        fig_group_company.update_traces(
            texttemplate='%{y:.2f}',
            textfont=dict(size=11, color='black', family='Arial Black'),
            textposition='outside'
        )
        st.plotly_chart(fig_group_company, use_container_width=True)
        
        # 公司板块环比同比综合表
        st.subheader(f"📊 板块环比同比综合表（公司）")
        
        current_sum_group_company = 0
        for col in selected_cols_group:
            if col in group_company_df.columns:
                current_sum_group_company += group_company_df[col].sum()
        total_current_group_company = current_sum_group_company / group_company_divide
        
        total_prev_group_company = None
        if prev_cols_group:
            prev_sum = 0
            for col in prev_cols_group:
                if col in group_company_df.columns:
                    prev_sum += group_company_df[col].sum()
            total_prev_group_company = prev_sum / group_company_divide
        
        total_last_year_group_company = None
        if last_year_cols_group:
            last_sum = 0
            for col in last_year_cols_group:
                if col in group_company_df.columns:
                    last_sum += group_company_df[col].sum()
            total_last_year_group_company = last_sum / group_company_divide
        
        table_data_group_company = []
        table_data_group_company.append({
            '维度': '市场',
            f'{group_metric_name}（{group_company_unit}）': f"{total_current_group_company:.2f}",
            '环比（%）': f"{(total_current_group_company - total_prev_group_company) / total_prev_group_company * 100:+.2f}" if total_prev_group_company is not None and total_prev_group_company != 0 else '-',
            '同比（%）': f"{(total_current_group_company - total_last_year_group_company) / total_last_year_group_company * 100:+.2f}" if total_last_year_group_company is not None and total_last_year_group_company != 0 else '-'
        })
        for _, row in group_df_plot_company.iterrows():
            mom = row.get('环比')
            yoy = row.get('同比')
            table_data_group_company.append({
                '维度': row[group_col],
                f'{group_metric_name}（{group_company_unit}）': f"{row['本月']:.2f}",
                '环比（%）': f"{mom:+.2f}" if mom is not None and not pd.isna(mom) else '-',
                '同比（%）': f"{yoy:+.2f}" if yoy is not None and not pd.isna(yoy) else '-'
            })
        table_df_group_company = pd.DataFrame(table_data_group_company)
        st.dataframe(table_df_group_company, use_container_width=True, hide_index=True)
    else:
        st.warning("没有可用的板块数据来绘制图表")
    
    # ============================================================
    # 9. 市场各板块份额饼图 + 公司占有率柱状图（并列）
    # ============================================================
    st.subheader("📊 市场各板块分析")
    
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
    else:
        pie_market_df = df_oi_market.copy()
        pie_company_df = df_oi_company.copy()
        pie_unit = '百万手'
        pie_title = '持仓量'
    
    pie_market_df = pie_market_df.loc[:, ~pie_market_df.columns.isna()]
    pie_market_df = pie_market_df.loc[:, pie_market_df.columns != '']
    pie_market_df = pie_market_df.loc[:, ~pie_market_df.columns.duplicated()]
    pie_market_df = pie_market_df.dropna(axis=1, how='all')
    
    pie_company_df = pie_company_df.loc[:, ~pie_company_df.columns.isna()]
    pie_company_df = pie_company_df.loc[:, pie_company_df.columns != '']
    pie_company_df = pie_company_df.loc[:, ~pie_company_df.columns.duplicated()]
    pie_company_df = pie_company_df.dropna(axis=1, how='all')
    
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
                        display_df['市场数值'] = display_df['市场数值'].apply(lambda x: f"{x:,.0f}")
                        display_df['公司数值'] = display_df['公司数值'].apply(lambda x: f"{x:,.0f}")
                    elif pie_data_type == '成交额':
                        market_unit_display = '亿'
                        company_unit_display = '亿'
                        display_df['市场数值'] = display_df['市场数值'].apply(lambda x: f"{x:,.2f}")
                        display_df['公司数值'] = display_df['公司数值'].apply(lambda x: f"{(x / 100000000):,.2f}")
                    else:
                        market_unit_display = '手'
                        company_unit_display = '手'
                        display_df['市场数值'] = display_df['市场数值'].apply(lambda x: f"{x:,.0f}")
                        display_df['公司数值'] = display_df['公司数值'].apply(lambda x: f"{x:,.0f}")
                    
                    display_df['市场占比（%）'] = display_df['市场占比（%）'].apply(lambda x: f"{x:.2f}%")
                    display_df['公司占比（%）'] = display_df['公司占比（%）'].apply(lambda x: f"{x:.4f}%")
                    display_df.columns = ['板块', f'市场数值（{market_unit_display}）', f'公司数值（{company_unit_display}）', '公司占比（%）', '市场占比（%）']
                    st.dataframe(display_df, use_container_width=True, hide_index=True)
            else:
                st.warning("该月份市场数据为空，请选择其他月份")
        else:
            st.warning("数据中无'板块'列，请检查数据")
    else:
        st.warning(f"选中的月份 {selected_pie_month} 不存在于数据中")
    
    # ============================================================
    # 导出报告
    # ============================================================
    st.markdown("---")
    st.subheader("📥 导出报告")
    
    def generate_excel_report():
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary_data = {
                '指标': [
                    '报告生成时间',
                    '数据类型',
                    '分析期间-交易所',
                    f'{metric_name}总计（市场）-交易所',
                    f'{metric_name}总计（公司）-交易所',
                    '分析期间-板块',
                    f'{metric_name}总计（市场）-板块',
                    f'{metric_name}总计（公司）-板块'
                ],
                '数值': [
                    datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    data_type,
                    selected_label_exchange,
                    f"{total_current_exchange:.2f}",
                    f"{total_current_exchange_company:.2f}",
                    selected_label_group,
                    f"{total_current_group:.2f}",
                    f"{total_current_group_company:.2f}"
                ]
            }
            summary_df = pd.DataFrame(summary_data)
            summary_df.to_excel(writer, sheet_name='报告摘要', index=False)
            
            if not exchange_df_plot.empty:
                exchange_df_plot.to_excel(writer, sheet_name=f'{metric_name}交易所对比-市场', index=False)
                table_df_exchange.to_excel(writer, sheet_name='交易所环比同比综合表-市场', index=False)
            
            if not exchange_df_plot_company.empty:
                exchange_df_plot_company.to_excel(writer, sheet_name=f'{metric_name}交易所对比-公司', index=False)
                table_df_exchange_company.to_excel(writer, sheet_name='交易所环比同比综合表-公司', index=False)
            
            if not group_df_plot.empty:
                group_df_plot.to_excel(writer, sheet_name=f'{metric_name}板块对比-市场', index=False)
                table_df_group.to_excel(writer, sheet_name='板块环比同比综合表-市场', index=False)
            
            if not group_df_plot_company.empty:
                group_df_plot_company.to_excel(writer, sheet_name=f'{metric_name}板块对比-公司', index=False)
                table_df_group_company.to_excel(writer, sheet_name='板块环比同比综合表-公司', index=False)
            
            filtered_df.head(100).to_excel(writer, sheet_name='筛选后数据', index=False)
        return output.getvalue()
    
    excel_data = generate_excel_report()
    st.download_button(
        label=f"📥 下载{data_type}报告",
        data=excel_data,
        file_name=f"{data_type}报告_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
            
except Exception as e:
    st.error(f"❌ 处理数据时出错: {e}")
    st.exception(e)

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

# 在侧边栏显示数据源信息
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
# 以下是你原有的数据分析代码
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
    
    # ============================================================
    # 10. 期货品种分析（补充）
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
    future_market_df = future_market_df.loc[:, ~future_market_df.columns.isna()]
    future_market_df = future_market_df.loc[:, future_market_df.columns != '']
    future_market_df = future_market_df.loc[:, ~future_market_df.columns.duplicated()]
    future_market_df = future_market_df.dropna(axis=1, how='all')
    
    future_company_df = future_company_df.loc[:, ~future_company_df.columns.isna()]
    future_company_df = future_company_df.loc[:, future_company_df.columns != '']
    future_company_df = future_company_df.loc[:, ~future_company_df.columns.duplicated()]
    future_company_df = future_company_df.dropna(axis=1, how='all')
    
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
                            display_df_future['市场数值'] = display_df_future['市场数值'].apply(lambda x: f"{x:,.0f}")
                            display_df_future['公司数值'] = display_df_future['公司数值'].apply(lambda x: f"{x:,.0f}")
                        elif pie_data_type == '成交额':
                            market_unit_display_future = '亿元'
                            company_unit_display_future = '亿元'
                            display_df_future['市场数值'] = display_df_future['市场数值'].apply(lambda x: f"{x:,.2f}")
                            display_df_future['公司数值'] = display_df_future['公司数值'].apply(lambda x: f"{x:,.2f}")
                        else:  # 持仓量
                            market_unit_display_future = '手'
                            company_unit_display_future = '手'
                            display_df_future['市场数值'] = display_df_future['市场数值'].apply(lambda x: f"{x:,.0f}")
                            display_df_future['公司数值'] = display_df_future['公司数值'].apply(lambda x: f"{x:,.0f}")
                        
                        display_df_future['市场占比（%）'] = display_df_future['市场占比（%）'].apply(lambda x: f"{x:.2f}%")
                        display_df_future['公司占比（%）'] = display_df_future['公司占比（%）'].apply(lambda x: f"{x:.4f}%")
                        display_df_future.columns = ['品种', f'市场数值（{market_unit_display_future}）', f'公司数值（{company_unit_display_future}）', '公司占比（%）', '市场占比（%）']
                        st.dataframe(display_df_future, use_container_width=True, hide_index=True)
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
    # 11. 期权品种分析
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
    option_market_df = option_market_df.loc[:, ~option_market_df.columns.isna()]
    option_market_df = option_market_df.loc[:, option_market_df.columns != '']
    option_market_df = option_market_df.loc[:, ~option_market_df.columns.duplicated()]
    option_market_df = option_market_df.dropna(axis=1, how='all')
    
    option_company_df = option_company_df.loc[:, ~option_company_df.columns.isna()]
    option_company_df = option_company_df.loc[:, option_company_df.columns != '']
    option_company_df = option_company_df.loc[:, ~option_company_df.columns.duplicated()]
    option_company_df = option_company_df.dropna(axis=1, how='all')
    
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
                    display_df_option['市场数值'] = display_df_option['市场数值'].apply(lambda x: f"{x:,.0f}")
                    display_df_option['公司数值'] = display_df_option['公司数值'].apply(lambda x: f"{x:,.0f}")
                elif pie_data_type == '成交额':
                    market_unit_display_option = '亿元'
                    company_unit_display_option = '亿元'
                    display_df_option['市场数值'] = display_df_option['市场数值'].apply(lambda x: f"{x:,.2f}")
                    display_df_option['公司数值'] = display_df_option['公司数值'].apply(lambda x: f"{x:,.2f}")
                else:
                    market_unit_display_option = '手'
                    company_unit_display_option = '手'
                    display_df_option['市场数值'] = display_df_option['市场数值'].apply(lambda x: f"{x:,.0f}")
                    display_df_option['公司数值'] = display_df_option['公司数值'].apply(lambda x: f"{x:,.0f}")
                
                display_df_option['市场占比（%）'] = display_df_option['市场占比（%）'].apply(lambda x: f"{x:.2f}%")
                display_df_option['公司占比（%）'] = display_df_option['公司占比（%）'].apply(lambda x: f"{x:.4f}%")
                display_df_option.columns = ['品种', f'市场数值（{market_unit_display_option}）', f'公司数值（{company_unit_display_option}）', '公司占比（%）', '市场占比（%）']
                st.dataframe(display_df_option, use_container_width=True, hide_index=True)
        else:
            st.warning("该月份期权数据为空，请选择其他月份")
    else:
        st.warning(f"选中的月份 {selected_pie_month} 不存在于期权数据中")

    # ============================================================
    # 12. 公司情况 - 各品种份额饼图 + 公司占有率柱状图（并列）
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
    company_market_df = company_market_df.loc[:, ~company_market_df.columns.isna()]
    company_market_df = company_market_df.loc[:, company_market_df.columns != '']
    company_market_df = company_market_df.loc[:, ~company_market_df.columns.duplicated()]
    company_market_df = company_market_df.dropna(axis=1, how='all')
    
    company_company_df = company_company_df.loc[:, ~company_company_df.columns.isna()]
    company_company_df = company_company_df.loc[:, company_company_df.columns != '']
    company_company_df = company_company_df.loc[:, ~company_company_df.columns.duplicated()]
    company_company_df = company_company_df.dropna(axis=1, how='all')
    
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
        
        # 按公司数值降序排列（饼图按公司份额排序）
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
                display_df_company = display_df_company.sort_values('公司数值', ascending=False)
                
                if pie_data_type == '成交量':
                    market_unit_display_company = '手'
                    company_unit_display_company = '手'
                    display_df_company['市场数值'] = display_df_company['市场数值'].apply(lambda x: f"{x:,.0f}")
                    display_df_company['公司数值'] = display_df_company['公司数值'].apply(lambda x: f"{x:,.0f}")
                elif pie_data_type == '成交额':
                    market_unit_display_company = '亿元'
                    company_unit_display_company = '亿元'
                    display_df_company['市场数值'] = display_df_company['市场数值'].apply(lambda x: f"{x:,.2f}")
                    display_df_company['公司数值'] = display_df_company['公司数值'].apply(lambda x: f"{x:,.2f}")
                else:
                    market_unit_display_company = '手'
                    company_unit_display_company = '手'
                    display_df_company['市场数值'] = display_df_company['市场数值'].apply(lambda x: f"{x:,.0f}")
                    display_df_company['公司数值'] = display_df_company['公司数值'].apply(lambda x: f"{x:,.0f}")
                
                display_df_company['公司内部占比（%）'] = display_df_company['公司内部占比（%）'].apply(lambda x: f"{x:.2f}%")
                display_df_company['公司占市场比重（%）'] = display_df_company['公司占市场比重（%）'].apply(lambda x: f"{x:.4f}%")
                display_df_company.columns = ['品种', f'市场数值（{market_unit_display_company}）', f'公司数值（{company_unit_display_company}）', '公司内部占比（%）', '公司占市场比重（%）']
                st.dataframe(display_df_company, use_container_width=True, hide_index=True)
        else:
            st.warning("该月份公司数据为空，请选择其他月份")
    else:
        if company_market_df.empty:
            st.info("📊 当前数据中没有期货品种（产品类型='期货'），跳过期货品种分析")
        else:
            st.warning(f"选中的月份 {selected_pie_month} 不存在于公司数据中")

    # ============================================================
    # 13. 期权品种分析-公司
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
    option_market_df = option_market_df.loc[:, ~option_market_df.columns.isna()]
    option_market_df = option_market_df.loc[:, option_market_df.columns != '']
    option_market_df = option_market_df.loc[:, ~option_market_df.columns.duplicated()]
    option_market_df = option_market_df.dropna(axis=1, how='all')
    
    option_company_df = option_company_df.loc[:, ~option_company_df.columns.isna()]
    option_company_df = option_company_df.loc[:, option_company_df.columns != '']
    option_company_df = option_company_df.loc[:, ~option_company_df.columns.duplicated()]
    option_company_df = option_company_df.dropna(axis=1, how='all')
    
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
        
        # 按公司数值降序排列（饼图按公司份额排序）
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
                st.plotly_chart(fig_pie_option, use_container_width=True, config={'displayModeBar': False}, key="option_pie_chart")
            
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
                st.plotly_chart(fig_bar_option, use_container_width=True, config={'displayModeBar': False}, key="option_bar_chart")
            
            # 数据表格
            with st.expander("📋 查看详细数据"):
                display_df_option = option_merged_top.copy()
                display_df_option = display_df_option.sort_values('公司数值', ascending=False)
                
                if pie_data_type == '成交量':
                    market_unit_display_option = '手'
                    company_unit_display_option = '手'
                    display_df_option['市场数值'] = display_df_option['市场数值'].apply(lambda x: f"{x:,.0f}")
                    display_df_option['公司数值'] = display_df_option['公司数值'].apply(lambda x: f"{x:,.0f}")
                elif pie_data_type == '成交额':
                    market_unit_display_option = '亿元'
                    company_unit_display_option = '亿元'
                    display_df_option['市场数值'] = display_df_option['市场数值'].apply(lambda x: f"{x:,.2f}")
                    display_df_option['公司数值'] = display_df_option['公司数值'].apply(lambda x: f"{x:,.2f}")
                else:
                    market_unit_display_option = '手'
                    company_unit_display_option = '手'
                    display_df_option['市场数值'] = display_df_option['市场数值'].apply(lambda x: f"{x:,.0f}")
                    display_df_option['公司数值'] = display_df_option['公司数值'].apply(lambda x: f"{x:,.0f}")
                
                display_df_option['公司内部占比（%）'] = display_df_option['公司内部占比（%）'].apply(lambda x: f"{x:.2f}%")
                display_df_option['公司占市场比重（%）'] = display_df_option['公司占市场比重（%）'].apply(lambda x: f"{x:.4f}%")
                display_df_option.columns = ['品种', f'市场数值（{market_unit_display_option}）', f'公司数值（{company_unit_display_option}）', '公司内部占比（%）', '公司占市场比重（%）']
                st.dataframe(display_df_option, use_container_width=True, hide_index=True)
        else:
            st.warning("该月份期权数据为空，请选择其他月份")
    else:
        if option_market_df.empty:
            st.info("📊 当前数据中没有期权品种（产品类型='期货期权'或'现货期权'），跳过期权品种分析")
        else:
            st.warning(f"选中的月份 {selected_pie_month} 不存在于期权数据中")

    # ============================================================
    # 14. 资金统计（今年 vs 去年）
    # ============================================================
    try:
        # 加载资金对账表
        df_fund_current = data_dict['fund_current']
        df_fund_last_year = data_dict['fund_last_year']
        
        # 清理资金对账表数据
        df_fund_current = df_fund_current.loc[:, ~df_fund_current.columns.isna()]
        df_fund_current = df_fund_current.loc[:, df_fund_current.columns != '']
        df_fund_current = df_fund_current.loc[:, ~df_fund_current.columns.duplicated()]
        df_fund_current = df_fund_current.dropna(axis=1, how='all')
        
        df_fund_last_year = df_fund_last_year.loc[:, ~df_fund_last_year.columns.isna()]
        df_fund_last_year = df_fund_last_year.loc[:, df_fund_last_year.columns != '']
        df_fund_last_year = df_fund_last_year.loc[:, ~df_fund_last_year.columns.duplicated()]
        df_fund_last_year = df_fund_last_year.dropna(axis=1, how='all')
        
        st.subheader("📊运营总体情况")
        
        # ---- 月份筛选器 ----
        all_months = []
        if not df_fund_current.empty:
            current_months_list = df_fund_current.iloc[:, 0].dropna().unique()
            for m in current_months_list:
                m_str = str(int(m)) if isinstance(m, (int, float)) else str(m)
                if len(m_str) == 6 and m_str.isdigit():
                    all_months.append(m_str)
        if not df_fund_last_year.empty:
            last_months_list = df_fund_last_year.iloc[:, 0].dropna().unique()
            for m in last_months_list:
                m_str = str(int(m)) if isinstance(m, (int, float)) else str(m)
                if len(m_str) == 6 and m_str.isdigit():
                    all_months.append(m_str)
        
        all_months = sorted(set(all_months), reverse=True)
        
        if all_months:
            selected_month = st.selectbox(
                "选择月份",
                options=all_months,
                format_func=lambda x: f"{x[:4]}年{x[4:6]}月",
                key="fund_month_selector"
            )
        else:
            selected_month = None
            st.info("暂无月份数据")
        
        # 获取各列
        d_col = df_fund_current.columns[3] if len(df_fund_current.columns) > 3 else None
        e_col = df_fund_current.columns[4] if len(df_fund_current.columns) > 4 else None
        f_col = df_fund_current.columns[5] if len(df_fund_current.columns) > 5 else None
        g_col = df_fund_current.columns[6] if len(df_fund_current.columns) > 6 else None
        h_col = '期末权益' if '期末权益' in df_fund_current.columns else df_fund_current.columns[7] if len(df_fund_current.columns) > 7 else None
        j_col = df_fund_current.columns[9] if len(df_fund_current.columns) > 9 else None
        k_col = df_fund_current.columns[10] if len(df_fund_current.columns) > 10 else None
        
        if h_col is not None:
            # ---- 今年数据 ----
            current_data = []
            current_months = df_fund_current.iloc[:, 0].dropna().unique()
            for month_val in current_months:
                month_str = str(int(month_val)) if isinstance(month_val, (int, float)) else str(month_val)
                if len(month_str) == 6 and month_str.isdigit():
                    month_mask = df_fund_current.iloc[:, 0] == month_val
                    equity_sum = df_fund_current.loc[month_mask, h_col].sum()
                    if d_col is not None and e_col is not None:
                        net_deposit_sum = df_fund_current.loc[month_mask, d_col].sum() - df_fund_current.loc[month_mask, e_col].sum()
                    else:
                        net_deposit_sum = None
                    if f_col is not None:
                        fee_sum = df_fund_current.loc[month_mask, f_col].sum()
                    else:
                        fee_sum = None
                    if g_col is not None and j_col is not None and k_col is not None:
                        pnl_sum = (df_fund_current.loc[month_mask, g_col].sum() + 
                                  df_fund_current.loc[month_mask, j_col].sum() - 
                                  df_fund_current.loc[month_mask, k_col].sum())
                    else:
                        pnl_sum = None
                    if pd.notna(equity_sum) and equity_sum != 0:
                        current_data.append({
                            '月份': month_str,
                            '期末权益': equity_sum,
                            '净入金': net_deposit_sum if net_deposit_sum is not None else 0,
                            '留存手续费': fee_sum if fee_sum is not None else 0,
                            '平仓盈亏': pnl_sum if pnl_sum is not None else 0,
                            '类型': '今年'
                        })
            
            # ---- 去年数据 ----
            last_year_data = []
            last_year_months = df_fund_last_year.iloc[:, 0].dropna().unique()
            for month_val in last_year_months:
                month_str = str(int(month_val)) if isinstance(month_val, (int, float)) else str(month_val)
                if len(month_str) == 6 and month_str.isdigit():
                    month_mask = df_fund_last_year.iloc[:, 0] == month_val
                    equity_sum = df_fund_last_year.loc[month_mask, h_col].sum()
                    if d_col is not None and e_col is not None:
                        net_deposit_sum = df_fund_last_year.loc[month_mask, d_col].sum() - df_fund_last_year.loc[month_mask, e_col].sum()
                    else:
                        net_deposit_sum = None
                    if f_col is not None:
                        fee_sum = df_fund_last_year.loc[month_mask, f_col].sum()
                    else:
                        fee_sum = None
                    if g_col is not None and j_col is not None and k_col is not None:
                        pnl_sum = (df_fund_last_year.loc[month_mask, g_col].sum() + 
                                  df_fund_last_year.loc[month_mask, j_col].sum() - 
                                  df_fund_last_year.loc[month_mask, k_col].sum())
                    else:
                        pnl_sum = None
                    if pd.notna(equity_sum) and equity_sum != 0:
                        last_year_data.append({
                            '月份': month_str,
                            '期末权益': equity_sum,
                            '净入金': net_deposit_sum if net_deposit_sum is not None else 0,
                            '留存手续费': fee_sum if fee_sum is not None else 0,
                            '平仓盈亏': pnl_sum if pnl_sum is not None else 0,
                            '类型': '去年'
                        })
            
            all_fund_data = current_data + last_year_data
            
            if all_fund_data:
                fund_df = pd.DataFrame(all_fund_data)
                
                # ===== 表格 =====
                display_df = fund_df[['月份', '期末权益', '净入金', '留存手续费', '平仓盈亏', '类型']].copy()
                current_display = display_df[display_df['类型'] == '今年'][['月份', '期末权益', '净入金', '留存手续费', '平仓盈亏']].copy()
                current_display['类型'] = '今年'
                last_display = display_df[display_df['类型'] == '去年'][['月份', '期末权益', '净入金', '留存手续费', '平仓盈亏']].copy()
                last_display['类型'] = '去年'
                
                merged_display = pd.concat([current_display, last_display], ignore_index=True)
                merged_display = merged_display.sort_values('月份', ascending=False)
                merged_display = merged_display[['月份', '期末权益', '净入金', '留存手续费', '平仓盈亏']]
                
                merged_display['期末权益'] = (merged_display['期末权益'] / 100000000).round(2)
                merged_display['净入金'] = (merged_display['净入金'] / 10000000).round(2)
                merged_display['留存手续费'] = (merged_display['留存手续费'] / 100000).round(2)
                merged_display['平仓盈亏'] = (merged_display['平仓盈亏'] / 1000000).round(2)
                merged_display.columns = ['月份', '期末权益（亿元）', '净入金（千万）', '留存手续费（十万）', '平仓盈亏（百万）']
                
                st.dataframe(merged_display, use_container_width=True, hide_index=True)
                
                # ============================================================
                # 累计值计算
                # ============================================================
                if selected_month:
                    selected_year = int(str(selected_month)[:4])
                    current_year = datetime.datetime.now().year
                    
                    if selected_year == current_year:
                        target_data = fund_df[fund_df['类型'] == '今年'].copy()
                    else:
                        target_data = fund_df[fund_df['类型'] == '去年'].copy()
                else:
                    target_data = fund_df[fund_df['类型'] == '今年'].copy()
                
                if not target_data.empty:
                    target_data['月份'] = target_data['月份'].astype(str)
                    current_equity_cumsum = target_data['期末权益'].sum() / 100000000
                    current_deposit_cumsum = target_data['净入金'].sum() / 10000000
                    current_fee_cumsum = target_data['留存手续费'].sum() / 100000
                    current_pnl_cumsum = target_data['平仓盈亏'].sum() / 1000000
                    
                    if selected_month:
                        selected_month_str = str(selected_month)
                        target_filtered = target_data[target_data['月份'] <= selected_month_str]
                        if not target_filtered.empty:
                            current_equity_cumsum = target_filtered['期末权益'].sum() / 100000000
                            current_deposit_cumsum = target_filtered['净入金'].sum() / 10000000
                            current_fee_cumsum = target_filtered['留存手续费'].sum() / 100000
                            current_pnl_cumsum = target_filtered['平仓盈亏'].sum() / 1000000
                else:
                    current_equity_cumsum = 0
                    current_deposit_cumsum = 0
                    current_fee_cumsum = 0
                    current_pnl_cumsum = 0
                
                # ---- 计算期末权益的当期值、环比和同比 ----
                current_equity_value = None
                current_type = "无数据"
                mom_change = None
                yoy_change = None
                
                if selected_month:
                    current_month_data = fund_df[(fund_df['月份'] == selected_month) & (fund_df['类型'] == '今年')]
                    if current_month_data.empty:
                        current_month_data = fund_df[(fund_df['月份'] == selected_month) & (fund_df['类型'] == '去年')]
                    
                    month_num = int(selected_month[4:6])
                    year_num = int(selected_month[:4])
                    prev_month_str = f"{year_num}{month_num - 1:02d}" if month_num > 1 else f"{year_num - 1}12"
                    prev_month_data = fund_df[(fund_df['月份'] == prev_month_str) & (fund_df['类型'] == '今年')]
                    
                    if not current_month_data.empty:
                        current_equity_value = current_month_data['期末权益'].iloc[0] / 100000000
                        if not fund_df[(fund_df['月份'] == selected_month) & (fund_df['类型'] == '今年')].empty:
                            current_type = "今年"
                        else:
                            current_type = "去年"
                    
                    if not prev_month_data.empty:
                        prev_equity_value = prev_month_data['期末权益'].iloc[0] / 100000000
                    else:
                        prev_equity_value = None
                    
                    if current_equity_value is not None and prev_equity_value is not None and prev_equity_value != 0:
                        mom_change = (current_equity_value - prev_equity_value) / prev_equity_value * 100
                    
                    if current_equity_value is not None:
                        selected_year = int(selected_month[:4])
                        selected_month_num = selected_month[4:6]
                        last_year_month_str = f"{selected_year - 1}{selected_month_num}"
                        last_year_month_data = fund_df[(fund_df['月份'] == last_year_month_str) & (fund_df['类型'] == '去年')]
                        if not last_year_month_data.empty:
                            last_year_equity_value = last_year_month_data['期末权益'].iloc[0] / 100000000
                            if last_year_equity_value != 0:
                                yoy_change = (current_equity_value - last_year_equity_value) / last_year_equity_value * 100
                
                # ===== 四个折线图 2x2 =====
                fund_df_sorted = fund_df.sort_values('月份')
                fund_df_sorted['月份显示'] = fund_df_sorted['月份'].str[4:6]
                fund_df_sorted['期末权益（亿元）'] = (fund_df_sorted['期末权益'] / 100000000).round(2)
                fund_df_sorted['净入金（千万）'] = (fund_df_sorted['净入金'] / 10000000).round(2)
                fund_df_sorted['留存手续费（十万）'] = (fund_df_sorted['留存手续费'] / 100000).round(2)
                fund_df_sorted['平仓盈亏（百万）'] = (fund_df_sorted['平仓盈亏'] / 1000000).round(2)
                
                row1_col1, row1_col2 = st.columns(2)
                row2_col1, row2_col2 = st.columns(2)
                
                color_map = {'今年': '#2E86C1', '去年': '#F39C12'}
                
                # 图1：期末权益
                with row1_col1:
                    equity_annotation_text = f"{current_type}当期: {current_equity_value:.2f}亿" if current_equity_value is not None else "当期: -"
                    if mom_change is not None:
                        equity_annotation_text += f"<br>环比: {mom_change:+.2f}%"
                    else:
                        equity_annotation_text += "<br>环比: -"
                    if yoy_change is not None:
                        equity_annotation_text += f"<br>同比: {yoy_change:+.2f}%"
                    else:
                        equity_annotation_text += "<br>同比: -"
                    
                    fig_equity = px.line(
                        fund_df_sorted,
                        x='月份显示',
                        y='期末权益（亿元）',
                        color='类型',
                        title='期末权益',
                        labels={'月份显示': '月份', '期末权益（亿元）': '亿元', '类型': ''},
                        markers=True,
                        color_discrete_map=color_map
                    )
                    fig_equity.update_layout(
                        title_font=dict(size=14, color='#1A5276'),
                        font=dict(size=11),
                        plot_bgcolor='#F8F9F9',
                        paper_bgcolor='white',
                        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
                        height=350,
                        annotations=[
                            dict(
                                x=0.98,
                                y=0.98,
                                xref='paper',
                                yref='paper',
                                text=equity_annotation_text,
                                showarrow=False,
                                font=dict(size=10, color='#1A5276'),
                                bgcolor='rgba(255,255,255,0.85)',
                                bordercolor='#1A5276',
                                borderwidth=1,
                                borderpad=4,
                                xanchor='right',
                                yanchor='top'
                            )
                        ]
                    )
                    fig_equity.update_traces(
                        texttemplate='%{y:.2f}',
                        textposition='top center',
                        textfont=dict(size=8),
                        mode='lines+markers+text'
                    )
                    st.plotly_chart(fig_equity, use_container_width=True, config={'displayModeBar': False})
                
                # 图2：净入金
                with row1_col2:
                    fig_deposit = px.line(
                        fund_df_sorted,
                        x='月份显示',
                        y='净入金（千万）',
                        color='类型',
                        title='净入金',
                        labels={'月份显示': '月份', '净入金（千万）': '千万', '类型': ''},
                        markers=True,
                        color_discrete_map=color_map
                    )
                    fig_deposit.update_layout(
                        title_font=dict(size=14, color='#1A5276'),
                        font=dict(size=11),
                        plot_bgcolor='#F8F9F9',
                        paper_bgcolor='white',
                        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
                        height=350,
                        annotations=[
                            dict(
                                x=0.98,
                                y=0.98,
                                xref='paper',
                                yref='paper',
                                text=f'累计: {current_deposit_cumsum:+.2f}千万',
                                showarrow=False,
                                font=dict(size=11, color='#1A5276'),
                                bgcolor='rgba(255,255,255,0.85)',
                                bordercolor='#1A5276',
                                borderwidth=1,
                                borderpad=4,
                                xanchor='right',
                                yanchor='top'
                            )
                        ]
                    )
                    fig_deposit.update_traces(
                        texttemplate='%{y:.2f}',
                        textposition='top center',
                        textfont=dict(size=8),
                        mode='lines+markers+text'
                    )
                    st.plotly_chart(fig_deposit, use_container_width=True, config={'displayModeBar': False})
                
                # 图3：留存手续费
                with row2_col1:
                    fig_fee = px.line(
                        fund_df_sorted,
                        x='月份显示',
                        y='留存手续费（十万）',
                        color='类型',
                        title='留存手续费',
                        labels={'月份显示': '月份', '留存手续费（十万）': '十万', '类型': ''},
                        markers=True,
                        color_discrete_map=color_map
                    )
                    fig_fee.update_layout(
                        title_font=dict(size=14, color='#1A5276'),
                        font=dict(size=11),
                        plot_bgcolor='#F8F9F9',
                        paper_bgcolor='white',
                        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
                        height=350,
                        annotations=[
                            dict(
                                x=0.98,
                                y=0.98,
                                xref='paper',
                                yref='paper',
                                text=f'累计: {current_fee_cumsum:.2f}十万',
                                showarrow=False,
                                font=dict(size=11, color='#1A5276'),
                                bgcolor='rgba(255,255,255,0.85)',
                                bordercolor='#1A5276',
                                borderwidth=1,
                                borderpad=4,
                                xanchor='right',
                                yanchor='top'
                            )
                        ]
                    )
                    fig_fee.update_traces(
                        texttemplate='%{y:.2f}',
                        textposition='top center',
                        textfont=dict(size=8),
                        mode='lines+markers+text'
                    )
                    st.plotly_chart(fig_fee, use_container_width=True, config={'displayModeBar': False})
                
                # 图4：平仓盈亏
                with row2_col2:
                    fig_pnl = px.line(
                        fund_df_sorted,
                        x='月份显示',
                        y='平仓盈亏（百万）',
                        color='类型',
                        title='平仓盈亏',
                        labels={'月份显示': '月份', '平仓盈亏（百万）': '百万', '类型': ''},
                        markers=True,
                        color_discrete_map=color_map
                    )
                    fig_pnl.update_layout(
                        title_font=dict(size=14, color='#1A5276'),
                        font=dict(size=11),
                        plot_bgcolor='#F8F9F9',
                        paper_bgcolor='white',
                        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
                        height=350,
                        annotations=[
                            dict(
                                x=0.98,
                                y=0.98,
                                xref='paper',
                                yref='paper',
                                text=f'累计: {current_pnl_cumsum:+.2f}百万',
                                showarrow=False,
                                font=dict(size=11, color='#1A5276'),
                                bgcolor='rgba(255,255,255,0.85)',
                                bordercolor='#1A5276',
                                borderwidth=1,
                                borderpad=4,
                                xanchor='right',
                                yanchor='top'
                            )
                        ]
                    )
                    fig_pnl.update_traces(
                        texttemplate='%{y:.2f}',
                        textposition='top center',
                        textfont=dict(size=8),
                        mode='lines+markers+text'
                    )
                    st.plotly_chart(fig_pnl, use_container_width=True, config={'displayModeBar': False})
                
                # ============================================================
                # 第5-6个图：盈利客户数 + 交易客户数（1x2 并排）
                # ============================================================
                st.subheader("📊 客户情况统计")
                
                # ---- 获取交易统计表列名 ----
                trade_month_col = '月份' if '月份' in df_trade_stats.columns else df_trade_stats.columns[0] if not df_trade_stats.empty else '月份'
                trade_investor_col = '投资者代码' if '投资者代码' in df_trade_stats.columns else df_trade_stats.columns[3] if not df_trade_stats.empty else '投资者代码'
                trade_f_col = '平仓盈亏' if '平仓盈亏' in df_trade_stats.columns else df_trade_stats.columns[5] if len(df_trade_stats.columns) > 5 else None
                trade_g_col = '权利金收入' if '权利金收入' in df_trade_stats.columns else df_trade_stats.columns[6] if len(df_trade_stats.columns) > 6 else None
                trade_h_col = '权利金支出' if '权利金支出' in df_trade_stats.columns else df_trade_stats.columns[7] if len(df_trade_stats.columns) > 7 else None
                
                # ---- 按月份统计交易客户数和盈利客户数（从交易统计表-月） ----
                customer_data = []
                
                if not df_trade_stats.empty:
                    # 获取所有月份
                    all_trade_months = []
                    for m in df_trade_stats[trade_month_col].dropna().unique():
                        m_str = str(int(m)) if isinstance(m, (int, float)) else str(m)
                        if len(m_str) == 6 and m_str.isdigit():
                            all_trade_months.append(m_str)
                    all_trade_months = sorted(set(all_trade_months))
                    
                    for month_str in all_trade_months:
                        # 筛选该月份数据
                        month_mask = df_trade_stats[trade_month_col].apply(
                            lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x)
                        ) == month_str
                        month_df = df_trade_stats[month_mask]
                        
                        if not month_df.empty:
                            # 交易客户数：该月所有投资者代码去重
                            total_investors = month_df[trade_investor_col].nunique()
                            
                            # 盈利客户数：f+g-h > 0 的投资者代码去重
                            if trade_f_col is not None and trade_g_col is not None and trade_h_col is not None:
                                month_df['平仓盈亏_计算'] = month_df[trade_f_col] + month_df[trade_g_col] - month_df[trade_h_col]
                                profit_df = month_df[month_df['平仓盈亏_计算'] > 0]
                                profit_investors = profit_df[trade_investor_col].nunique()
                            else:
                                profit_investors = 0
                            
                            # 判断今年还是去年
                            year = int(month_str[:4])
                            current_year = datetime.datetime.now().year
                            type_label = '今年' if year == current_year else '去年'
                            
                            customer_data.append({
                                '月份': month_str,
                                '交易客户数': total_investors,
                                '盈利客户数': profit_investors,
                                '类型': type_label
                            })
                
                if customer_data:
                    customer_df = pd.DataFrame(customer_data)
                    customer_df = customer_df.sort_values('月份')
                    customer_df['月份显示'] = customer_df['月份'].astype(str).str[4:6]
                    
                    # ---- 计算累计值（只计算今年的累计） ----
                    if selected_month:
                        selected_month_str = str(selected_month)
                        current_customer = customer_df[(customer_df['类型'] == '今年') & (customer_df['月份'] <= selected_month_str)]
                        current_total_cumsum = current_customer['交易客户数'].sum() if not current_customer.empty else 0
                        current_profit_cumsum = current_customer['盈利客户数'].sum() if not current_customer.empty else 0
                    else:
                        current_customer = customer_df[customer_df['类型'] == '今年']
                        current_total_cumsum = current_customer['交易客户数'].sum() if not current_customer.empty else 0
                        current_profit_cumsum = current_customer['盈利客户数'].sum() if not current_customer.empty else 0
                    
                    # ---- 只保留今年的数据用于折线图 ----
                    customer_df = customer_df[customer_df['类型'] == '今年']
                    
                    # ===== 1x2 布局 =====
                    col_left, col_right = st.columns(2)
                    
                    # 左列：盈利客户数
                    with col_left:
                        profit_plot_df = customer_df[['月份显示', '盈利客户数']].copy()
                        profit_plot_df = profit_plot_df.rename(columns={'盈利客户数': '客户数'})
                        
                        fig_profit = px.line(
                            profit_plot_df,
                            x='月份显示',
                            y='客户数',
                            title='盈利客户数',
                            labels={'月份显示': '月份', '客户数': '客户数'},
                            markers=True,
                            color_discrete_map={'今年': '#2E86C1'}
                        )
                        fig_profit.update_traces(line=dict(color='#2E86C1'), marker=dict(color='#2E86C1'))
                        fig_profit.update_layout(
                            title_font=dict(size=14, color='#1A5276'),
                            font=dict(size=11),
                            plot_bgcolor='#F8F9F9',
                            paper_bgcolor='white',
                            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
                            height=350,
                            annotations=[
                                dict(
                                    x=0.98,
                                    y=0.98,
                                    xref='paper',
                                    yref='paper',
                                    text=f'累计: {current_profit_cumsum}',
                                    showarrow=False,
                                    font=dict(size=11, color='#1A5276'),
                                    bgcolor='rgba(255,255,255,0.85)',
                                    bordercolor='#1A5276',
                                    borderwidth=1,
                                    borderpad=4,
                                    xanchor='right',
                                    yanchor='top'
                                )
                            ]
                        )
                        fig_profit.update_traces(
                            texttemplate='%{y:.0f}',
                            textposition='top center',
                            textfont=dict(size=8),
                            mode='lines+markers+text'
                        )
                        st.plotly_chart(fig_profit, use_container_width=True, config={'displayModeBar': False})
                    
                    # 右列：交易客户数
                    with col_right:
                        trade_plot_df = customer_df[['月份显示', '交易客户数']].copy()
                        trade_plot_df = trade_plot_df.rename(columns={'交易客户数': '客户数'})
                        
                        fig_total = px.line(
                            trade_plot_df,
                            x='月份显示',
                            y='客户数',
                            title='交易客户数',
                            labels={'月份显示': '月份', '客户数': '客户数'},
                            markers=True,
                            color_discrete_map={'今年': '#2E86C1'}
                        )
                        fig_total.update_traces(line=dict(color='#2E86C1'), marker=dict(color='#2E86C1'))
                        fig_total.update_layout(
                            title_font=dict(size=14, color='#1A5276'),
                            font=dict(size=11),
                            plot_bgcolor='#F8F9F9',
                            paper_bgcolor='white',
                            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5),
                            height=350,
                            annotations=[
                                dict(
                                    x=0.98,
                                    y=0.98,
                                    xref='paper',
                                    yref='paper',
                                    text=f'累计: {current_total_cumsum}',
                                    showarrow=False,
                                    font=dict(size=11, color='#1A5276'),
                                    bgcolor='rgba(255,255,255,0.85)',
                                    bordercolor='#1A5276',
                                    borderwidth=1,
                                    borderpad=4,
                                    xanchor='right',
                                    yanchor='top'
                                )
                            ]
                        )
                        fig_total.update_traces(
                            texttemplate='%{y:.0f}',
                            textposition='top center',
                            textfont=dict(size=8),
                            mode='lines+markers+text'
                        )
                        st.plotly_chart(fig_total, use_container_width=True, config={'displayModeBar': False})
                else:
                    st.info("暂无客户数据")
            else:
                st.info("暂无资金对账数据")
        else:
            st.warning("未找到期末权益列（H列），请检查数据格式")
    except Exception as e:
        st.warning(f"加载资金对账表时出错: {e}")

    # ============================================================
    # 15. 交易统计表-月 - 部门交易客户数
    # ============================================================
    st.subheader("📊 部门交易客户数统计")
    
    try:
        # 加载交易统计表
        df_trade_stats = data_dict['trade_stats']
        
        # 清理数据
        df_trade_stats = df_trade_stats.loc[:, ~df_trade_stats.columns.isna()]
        df_trade_stats = df_trade_stats.loc[:, df_trade_stats.columns != '']
        df_trade_stats = df_trade_stats.loc[:, ~df_trade_stats.columns.duplicated()]
        df_trade_stats = df_trade_stats.dropna(axis=1, how='all')
        
        # 加载资金对账表-月
        df_fund_current = data_dict['fund_current']
        df_fund_current = df_fund_current.loc[:, ~df_fund_current.columns.isna()]
        df_fund_current = df_fund_current.loc[:, df_fund_current.columns != '']
        df_fund_current = df_fund_current.loc[:, ~df_fund_current.columns.duplicated()]
        df_fund_current = df_fund_current.dropna(axis=1, how='all')
        
        # ---- 获取交易统计表列名 ----
        trade_month_col = '月份'
        trade_dept_col = '部门'
        trade_investor_col = '投资者代码'
        trade_f_col = '平仓盈亏' if '平仓盈亏' in df_trade_stats.columns else df_trade_stats.columns[5]
        trade_g_col = '权利金收入' if '权利金收入' in df_trade_stats.columns else df_trade_stats.columns[6]
        trade_h_col = '权利金支出' if '权利金支出' in df_trade_stats.columns else df_trade_stats.columns[7]
        
        # ---- 获取资金对账表列名 ----
        fund_month_col = '月份'
        if '部门' in df_fund_current.columns:
            fund_dept_col = '部门'
        elif '部门名称' in df_fund_current.columns:
            fund_dept_col = '部门名称'
        else:
            fund_dept_col = df_fund_current.columns[2]
        
        fund_e_col = '入金' if '入金' in df_fund_current.columns else df_fund_current.columns[3]
        fund_f_col = '出金' if '出金' in df_fund_current.columns else df_fund_current.columns[4]
        fund_g_col = '留存手续费' if '留存手续费' in df_fund_current.columns else df_fund_current.columns[5]
        fund_i_col = '期末权益' if '期末权益' in df_fund_current.columns else df_fund_current.columns[8]
        
        # ---- 月份筛选器 ----
        all_months = []
        if not df_trade_stats.empty:
            for m in df_trade_stats[trade_month_col].dropna().unique():
                m_str = str(int(m)) if isinstance(m, (int, float)) else str(m)
                if len(m_str) == 6 and m_str.isdigit():
                    all_months.append(m_str)
        all_months = sorted(set(all_months), reverse=True)
        
        if not all_months:
            st.info("暂无交易统计表数据")
        else:
            selected_month = st.selectbox(
                "选择月份",
                options=all_months,
                format_func=lambda x: f"{x[:4]}年{x[4:6]}月",
                key="trade_month_selector"
            )
            
            # ---- 筛选交易统计表数据 ----
            df_trade_stats['月份_str'] = df_trade_stats[trade_month_col].apply(
                lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x)
            )
            filtered_trade = df_trade_stats[df_trade_stats['月份_str'] == selected_month].copy()
            
            if filtered_trade.empty:
                st.info(f"{selected_month} 无交易数据")
            else:
                # ---- 筛选资金对账表数据 ----
                df_fund_current['月份_str'] = df_fund_current[fund_month_col].apply(
                    lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x)
                )
                filtered_fund = df_fund_current[df_fund_current['月份_str'] == selected_month].copy()
                
                # ---- 按部门汇总资金对账表数据 ----
                if not filtered_fund.empty:
                    filtered_fund['净入金'] = filtered_fund[fund_e_col] - filtered_fund[fund_f_col]
                    filtered_fund['留存手续费'] = filtered_fund[fund_g_col]
                    filtered_fund['期末权益'] = filtered_fund[fund_i_col]
                    
                    fund_by_dept = filtered_fund.groupby(fund_dept_col).agg({
                        '净入金': 'sum',
                        '留存手续费': 'sum',
                        '期末权益': 'sum'
                    }).reset_index()
                    fund_by_dept.columns = ['部门', '净入金', '留存手续费', '期末权益']
                else:
                    fund_by_dept = pd.DataFrame(columns=['部门', '净入金', '留存手续费', '期末权益'])
                
                # ---- 按部门统计交易统计表数据 ----
                # 1. 有交易客户数
                result_df = filtered_trade.groupby(trade_dept_col).agg(
                    有交易客户数=(trade_investor_col, 'nunique')
                ).reset_index()
                result_df.columns = ['部门', '有交易客户数']
                
                # 2. 盈利客户数：f+g-h > 0 的投资者代码去重个数
                filtered_trade['平仓盈亏_计算'] = filtered_trade[trade_f_col] + filtered_trade[trade_g_col] - filtered_trade[trade_h_col]
                profit_df = filtered_trade[filtered_trade['平仓盈亏_计算'] > 0]
                profit_count_df = profit_df.groupby(trade_dept_col).agg(
                    盈利客户数=(trade_investor_col, 'nunique')
                ).reset_index()
                profit_count_df.columns = ['部门', '盈利客户数']
                result_df = pd.merge(result_df, profit_count_df, on='部门', how='left').fillna(0)
                
                # 3. 平仓盈亏 = F列(平仓盈亏) + G列(权利金收入) - H列(权利金支出)
                pnl_df = filtered_trade.groupby(trade_dept_col).apply(
                    lambda x: (x[trade_f_col].sum() + x[trade_g_col].sum() - x[trade_h_col].sum())
                ).reset_index()
                pnl_df.columns = ['部门', '平仓盈亏']
                result_df = pd.merge(result_df, pnl_df, on='部门', how='left')
                
                # ---- 合并资金对账表数据 ----
                result_df = pd.merge(result_df, fund_by_dept, on='部门', how='left').fillna(0)
                
                # ---- 计算盈利面 ----
                result_df['盈利面'] = (result_df['盈利客户数'] / result_df['有交易客户数'] * 100).round(2)
                
                # ---- 添加排名 ----
                result_df['期末权益排名'] = result_df['期末权益'].rank(method='min', ascending=False).astype(int)
                result_df['平仓盈亏排名'] = result_df['平仓盈亏'].rank(method='min', ascending=False).astype(int)
                result_df['净入金排名'] = result_df['净入金'].rank(method='min', ascending=False).astype(int)
                result_df['留存手续费排名'] = result_df['留存手续费'].rank(method='min', ascending=False).astype(int)
                
                # ---- 重新排列列顺序 ----
                result_df = result_df[[
                    '部门',
                    '有交易客户数',
                    '盈利客户数',
                    '盈利面',
                    '期末权益',
                    '期末权益排名',
                    '平仓盈亏',
                    '平仓盈亏排名',
                    '净入金',
                    '净入金排名',
                    '留存手续费',
                    '留存手续费排名'
                ]]
                
                # ---- 排序 ----
                result_df = result_df.sort_values('有交易客户数', ascending=False)
                
                # ---- 格式化 ----
                result_df['期末权益'] = result_df['期末权益'].apply(lambda x: f"{int(x):,}")
                result_df['平仓盈亏'] = result_df['平仓盈亏'].apply(lambda x: f"{int(x):,}")
                result_df['净入金'] = result_df['净入金'].apply(lambda x: f"{int(x):,}")
                result_df['留存手续费'] = result_df['留存手续费'].apply(lambda x: f"{int(x):,}")
                result_df['盈利面'] = result_df['盈利面'].apply(lambda x: f"{x:.2f}%")
                
                # ---- 显示结果 ----
                st.subheader(f"📊 {selected_month[:4]}年{selected_month[4:6]}月 部门统计")
                st.dataframe(result_df, use_container_width=True, hide_index=True)
                
                # ---- 汇总信息 ----
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("部门总数", len(result_df))
                with col2:
                    st.metric("客户总数（去重）", int(result_df['有交易客户数'].sum()))
    except Exception as e:
        st.warning(f"加载数据时出错: {e}")

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

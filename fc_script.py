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
EXCEL_FILENAME = "data.xlsx"

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
# 原有数据加载函数（保留作为备选）
# ============================================================
@st.cache_data
def load_data(uploaded_file, sheet_name):
    if uploaded_file is not None:
        if uploaded_file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=0)
        else:
            df = pd.read_csv(uploaded_file)
        return df
    return None

# ============================================================
# 加载数据
# ============================================================
data_dict = load_all_data_from_github()

if data_dict is None:
    # 如果GitHub加载失败，尝试手动上传
    st.warning("⚠️ GitHub加载失败，请手动上传文件")
    
    with st.sidebar:
        st.header("📁 手动上传")
        uploaded_file = st.file_uploader(
            "上传 Excel 或 CSV 文件",
            type=['xlsx', 'xls', 'csv']
        )
    
    if uploaded_file is not None:
        try:
            df_vol_market = load_data(uploaded_file, '成交量-市场')
            df_vol_company = load_data(uploaded_file, '成交量-公司')
            df_amt_market = load_data(uploaded_file, '成交额-市场')
            df_amt_company = load_data(uploaded_file, '成交额-公司')
            df_oi_market = load_data(uploaded_file, '持仓量-市场')
            df_oi_company = load_data(uploaded_file, '持仓量-公司')
            df_fund_current = load_data(uploaded_file, '资金对账表-月')
            df_fund_last_year = load_data(uploaded_file, '上一年资金对账表-月')
            df_trade_stats = load_data(uploaded_file, '交易统计表-月')
            st.success("✅ 数据手动加载成功！")
        except Exception as e:
            st.error(f"❌ 加载数据失败: {e}")
            st.stop()
    else:
        st.info("👈 请上传数据文件")
        st.stop()
else:
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
                ],
                '数值': [
                    datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    data_type,
                    selected_label_exchange,
                    f"{total_current_exchange:.2f}",
                    f"{total_current_exchange_company:.2f}",
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

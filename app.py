#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
货代报价查询系统 - 网页版 v3
作者: 强子 (OpenClaw)
更新: 2026-02-03

v3修复:
    1) 同一Sheet内多行匹配不再漏掉
    2) 时效列优先找全局时效，解决显示"-"的问题
"""

import streamlit as st
import pandas as pd
import os
from typing import List, Dict, Optional

# ============================================================
# 配置
# ============================================================

SKIP_SHEETS = ["首推王牌渠道", "目录", "新增网点报价栏", "附加费查询栏", "查询有效性"]

REGION_MAPPING = {
    "华南": "华南", "深圳": "华南", "广州": "华南",
    "华东": "华东", "上海": "华东", "江苏": "华东", 
    "苏州": "华东", "宁波": "华东", "浙江": "华东", "杭州": "华东",
    "青岛": "青岛", "山东": "青岛",
    "福建": "福建", "厦门": "福建", "福州": "福州",
    "天津": "天津", "北京": "天津",
}

REGIONS = ["华东", "华南", "青岛", "福建", "福州", "天津"]

# ============================================================
# 核心查询函数 v3
# ============================================================

def find_time_columns(df: pd.DataFrame) -> Dict:
    """
    查找全局时效列（不分区域的）
    """
    result = {'global_time_col': None, 'global_dw_col': None}
    
    header_row = 3
    if header_row >= len(df):
        return result
    
    for col_idx in range(len(df.columns)):
        header_cell = df.iloc[header_row, col_idx]
        if pd.isna(header_cell):
            continue
        header_str = str(header_cell)
        
        if "全程时效" in header_str and result['global_time_col'] is None:
            result['global_time_col'] = col_idx
        elif "DW" in header_str and "送达" in header_str and result['global_dw_col'] is None:
            result['global_dw_col'] = col_idx
    
    return result


def find_region_price_column(df: pd.DataFrame, target_region: str, tax_type: str) -> Optional[int]:
    """
    找到目标区域的价格列
    """
    header_row = 3
    region_row = 4
    unit_row = 5
    
    if region_row >= len(df) or header_row >= len(df):
        return None
    
    target_unit = "KG" if tax_type == "含税" else "CBM"
    
    for col_idx in range(len(df.columns)):
        header_cell = df.iloc[header_row, col_idx]
        header_str = str(header_cell) if pd.notna(header_cell) else ""
        
        if tax_type not in header_str:
            continue
        
        region_cell = df.iloc[region_row, col_idx]
        if pd.isna(region_cell):
            continue
        region_str = str(region_cell)
        
        region_matched = False
        if target_region == "华东":
            region_matched = any(x in region_str for x in ["华东", "上海", "宁波", "苏州"])
        elif target_region == "华南":
            region_matched = region_str == "华南" or (target_region in region_str and "华东" not in region_str)
        else:
            region_matched = target_region in region_str
        
        if not region_matched:
            continue
        
        unit_cell = df.iloc[unit_row, col_idx] if unit_row < len(df) else None
        unit_str = str(unit_cell) if pd.notna(unit_cell) else ""
        
        if target_unit in unit_str:
            return col_idx
    
    return None


def find_region_time_columns(df: pd.DataFrame, price_col: int) -> Dict:
    """
    在价格列附近找时效列
    """
    result = {'time_col': None, 'dw_col': None}
    header_row = 3
    
    if price_col is None or header_row >= len(df):
        return result
    
    for offset in range(1, 5):
        check_col = price_col + offset
        if check_col >= len(df.columns):
            break
        
        header_cell = df.iloc[header_row, check_col]
        header_str = str(header_cell) if pd.notna(header_cell) else ""
        
        if "全程时效" in header_str and result['time_col'] is None:
            result['time_col'] = check_col
        elif "DW" in header_str and result['dw_col'] is None:
            result['dw_col'] = check_col
    
    return result


def query_prices(df_dict: dict, warehouse_code: str, region: str, tax_type: str) -> List[Dict]:
    """
    查询指定仓库在所有渠道的价格
    v3: 不再break，搜索sheet内所有匹配行
    """
    normalized_region = REGION_MAPPING.get(region, region)
    if normalized_region in ["上海", "江苏", "苏州", "宁波", "浙江", "杭州"]:
        normalized_region = "华东"
    
    results = []
    target_code = warehouse_code.strip().upper()
    
    for sheet_name, df in df_dict.items():
        if sheet_name in SKIP_SHEETS:
            continue
        
        try:
            # 找全局时效列
            global_time = find_time_columns(df)
            
            # 找区域价格列
            price_col = find_region_price_column(df, normalized_region, tax_type)
            if price_col is None:
                continue
            
            # 找区域时效列
            region_time = find_region_time_columns(df, price_col)
            
            # 确定使用哪个时效列：优先全局，其次区域
            time_col = global_time['global_time_col'] or region_time['time_col']
            dw_col = global_time['global_dw_col'] or region_time['dw_col']
            
            data_start = 6
            
            # 遍历所有行，不break
            for row_idx in range(data_start, len(df)):
                warehouse_cell = df.iloc[row_idx, 2]
                
                if pd.isna(warehouse_cell):
                    continue
                
                warehouse_str = str(warehouse_cell).strip().upper()
                
                if warehouse_str == target_code or target_code in warehouse_str:
                    price = df.iloc[row_idx, price_col]
                    time_val = df.iloc[row_idx, time_col] if time_col and time_col < len(df.columns) else None
                    dw_val = df.iloc[row_idx, dw_col] if dw_col and dw_col < len(df.columns) else None
                    
                    channel = df.iloc[row_idx, 1]
                    if pd.isna(channel):
                        channel = sheet_name
                    
                    results.append({
                        '渠道': str(channel) if pd.notna(channel) else sheet_name,
                        '全程时效': str(time_val) if pd.notna(time_val) else '-',
                        'DW送达': str(dw_val) if pd.notna(dw_val) else '-',
                        '价格': price if pd.notna(price) else '-',
                        '渠道分类': sheet_name,
                        '仓库': warehouse_str,
                        '区域': normalized_region,
                        '税种': tax_type
                    })
                    # v3: 不再break
        
        except Exception as e:
            continue
    
    def sort_key(x):
        try:
            return float(x['价格'])
        except:
            return float('inf')
    
    results.sort(key=sort_key)
    return results


def get_all_warehouses(df_dict: dict) -> List[str]:
    """从报价表中提取所有仓库代码"""
    warehouses = set()
    
    for sheet_name, df in df_dict.items():
        if sheet_name in SKIP_SHEETS:
            continue
        
        try:
            for row_idx in range(5, min(100, len(df))):
                cell = df.iloc[row_idx, 2]
                if pd.notna(cell):
                    val = str(cell).strip()
                    if any(c.isalpha() for c in val) and len(val) <= 10:
                        if "起收量" not in val and "邮编" not in val:
                            warehouses.add(val.upper())
        except:
            continue
    
    return sorted(list(warehouses))


# ============================================================
# Streamlit UI
# ============================================================

st.set_page_config(
    page_title="货代报价查询",
    page_icon="📦",
    layout="wide"
)

st.title("📦 货代报价查询系统")
st.caption("v3 - 修复多渠道匹配和时效显示")
st.markdown("---")

# 上传文件或使用默认文件
uploaded_file = st.file_uploader("上传报价表 (Excel)", type=['xlsx', 'xls'])

default_file = "data/报价表.xlsx"
has_default = os.path.exists(default_file)

if uploaded_file is not None:
    @st.cache_data
    def load_excel(file):
        return pd.read_excel(file, sheet_name=None, header=None)
    
    with st.spinner("正在加载报价表..."):
        df_dict = load_excel(uploaded_file)
    st.success(f"✅ 已加载: {uploaded_file.name}")

elif has_default:
    @st.cache_data
    def load_default():
        return pd.read_excel(default_file, sheet_name=None, header=None)
    
    with st.spinner("正在加载默认报价表..."):
        df_dict = load_default()
    st.info("📋 使用默认报价表")

else:
    st.warning("⚠️ 请上传报价表 Excel 文件")
    st.stop()

warehouses = get_all_warehouses(df_dict)

# 查询界面
col1, col2, col3 = st.columns(3)

with col1:
    warehouse_input = st.text_input("🏭 仓库代码", placeholder="输入如 TEB6, ONT8...")

with col2:
    region = st.selectbox("📍 提货区域", REGIONS)

with col3:
    tax_type = st.radio("💰 税种", ["含税", "自税"], horizontal=True)

# 查询按钮
if st.button("🔍 查询价格", type="primary", use_container_width=True):
    if not warehouse_input:
        st.error("请输入仓库代码")
    else:
        with st.spinner("正在查询..."):
            results = query_prices(df_dict, warehouse_input, region, tax_type)
        
        if results:
            st.markdown("---")
            st.subheader(f"📊 查询结果")
            st.markdown(f"**仓库:** {warehouse_input.upper()} | **区域:** {region} | **税种:** {tax_type}")
            
            best = results[0]
            if best['价格'] != '-':
                st.success(f"💡 **推荐:** {best['渠道']} — 价格 ¥{best['价格']}/kg, 时效 {best['全程时效']}")
            
            df_result = pd.DataFrame(results)[['渠道', '全程时效', 'DW送达', '价格', '渠道分类']]
            df_result.index = range(1, len(df_result) + 1)
            st.dataframe(df_result, use_container_width=True)
            
            csv = df_result.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                "📥 下载结果 (CSV)",
                csv,
                f"报价查询_{warehouse_input}_{region}_{tax_type}.csv",
                "text/csv"
            )
        else:
            st.warning(f"❌ 未找到 {warehouse_input} 在 {region} 区域的报价")

st.markdown("---")
st.caption("Made with ❤️ by 强子 (OpenClaw) | v3 修复版")

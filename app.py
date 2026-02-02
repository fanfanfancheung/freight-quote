#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
货代报价查询系统 - 网页版 v2
作者: 强子 (OpenClaw)
更新: 2026-02-02 - 修复时效列按区域选择的问题
"""

import streamlit as st
import pandas as pd
import os
from typing import List, Dict

# ============================================================
# 配置
# ============================================================

SKIP_SHEETS = ["首推王牌渠道", "目录", "新增网点报价栏", "附加费查询栏"]

REGION_MAPPING = {
    "华南": "华南", "深圳": "华南", "广州": "华南",
    "华东": "华东", "上海": "华东", "江苏": "华东", 
    "苏州": "华东", "宁波": "华东", "浙江": "华东", "杭州": "华东",
    "青岛": "青岛", "山东": "青岛",
    "福建": "福建", "厦门": "福建", "福州": "福州",
    "天津": "天津", "北京": "天津",
}

REGIONS = ["华东", "华南", "青岛", "福建", "天津"]

# ============================================================
# 核心查询函数
# ============================================================

def find_region_columns(df: pd.DataFrame, target_region: str) -> Dict:
    """
    分析Sheet结构，找出目标区域的含税、自税、时效列位置
    
    v2修复: 时效列现在根据区域正确选择，而不是取全局第一个
    """
    result = {
        'region_row': None,
        'tax_included_col': None,
        'tax_excluded_col': None,
        'time_col': None,
        'dw_time_col': None,
        'data_start_row': 6
    }
    
    # 标题行和区域行
    header_row = 3
    region_row = 4
    
    if region_row >= len(df):
        return result
    
    result['region_row'] = region_row
    
    # 遍历列，找目标区域的所有相关列
    for col_idx in range(len(df.columns)):
        region_cell = df.iloc[region_row, col_idx]
        
        if pd.isna(region_cell):
            continue
        
        region_str = str(region_cell)
        
        # 检查是否匹配目标区域
        region_matched = False
        if target_region == "华东":
            region_matched = "华东" in region_str or "上海" in region_str or "宁波" in region_str or "苏州" in region_str
        elif target_region == "华南":
            region_matched = region_str == "华南" or (target_region in region_str and "华东" not in region_str)
        else:
            region_matched = target_region in region_str
        
        if region_matched:
            # 检查这一列的标题（第3行）来确定是什么类型的列
            header_cell = df.iloc[header_row, col_idx] if header_row < len(df) else None
            header_str = str(header_cell) if pd.notna(header_cell) else ""
            
            # 检查第5行的标记（KG+ 或 CBM+）
            unit_cell = df.iloc[5, col_idx] if 5 < len(df) else None
            unit_str = str(unit_cell) if pd.notna(unit_cell) else ""
            
            # 判断列类型
            if "含税" in header_str:
                if "KG" in unit_str and result['tax_included_col'] is None:
                    result['tax_included_col'] = col_idx
            elif "自税" in header_str:
                if "CBM" in unit_str and result['tax_excluded_col'] is None:
                    result['tax_excluded_col'] = col_idx
    
    # 关键修复：时效列必须在价格列的【后面】
    if result['tax_included_col'] is not None or result['tax_excluded_col'] is not None:
        price_col_ref = result['tax_included_col'] or result['tax_excluded_col']
        
        # 时效列在价格列后面1-3列的范围内查找
        for offset in [2, 1, 3]:
            check_col = price_col_ref + offset
            if check_col < len(df.columns):
                header_cell = df.iloc[header_row, check_col]
                header_str = str(header_cell) if pd.notna(header_cell) else ""
                
                if "全程时效" in header_str and result['time_col'] is None:
                    result['time_col'] = check_col
                elif "DW" in header_str and result['dw_time_col'] is None:
                    result['dw_time_col'] = check_col
        
        # DW送达通常在全程时效后面一列
        if result['time_col'] is not None and result['dw_time_col'] is None:
            dw_col = result['time_col'] + 1
            if dw_col < len(df.columns):
                header_cell = df.iloc[header_row, dw_col]
                if pd.notna(header_cell) and "DW" in str(header_cell):
                    result['dw_time_col'] = dw_col
    
    # 查找数据起始行
    for row_idx in range(5, 10):
        if row_idx >= len(df):
            continue
        cell = df.iloc[row_idx, 2]
        if pd.notna(cell) and isinstance(cell, str):
            if any(c.isalpha() for c in str(cell)) and len(str(cell)) <= 15:
                if "起收量" not in str(cell) and "邮编" not in str(cell):
                    result['data_start_row'] = row_idx
                    break
    
    return result


def query_prices(df_dict: dict, warehouse_code: str, region: str, tax_type: str) -> List[Dict]:
    """查询指定仓库在所有渠道的价格"""
    normalized_region = REGION_MAPPING.get(region, region)
    if "华东" in normalized_region or region in ["上海", "江苏", "苏州", "宁波", "浙江", "杭州"]:
        normalized_region = "华东"
    
    results = []
    
    for sheet_name, df in df_dict.items():
        if sheet_name in SKIP_SHEETS:
            continue
        
        try:
            structure = find_region_columns(df, normalized_region)
            
            if structure['region_row'] is None:
                continue
            
            if tax_type == "含税":
                price_col = structure['tax_included_col']
            else:
                price_col = structure['tax_excluded_col']
            
            if price_col is None:
                continue
            
            time_col = structure['time_col']
            dw_time_col = structure['dw_time_col']
            data_start = structure['data_start_row']
            
            for row_idx in range(data_start, len(df)):
                warehouse_cell = df.iloc[row_idx, 2]
                
                if pd.isna(warehouse_cell):
                    continue
                
                warehouse_str = str(warehouse_cell).strip().upper()
                target_code = warehouse_code.strip().upper()
                
                if warehouse_str == target_code or target_code in warehouse_str:
                    price = df.iloc[row_idx, price_col]
                    time_val = df.iloc[row_idx, time_col] if time_col else None
                    dw_time_val = df.iloc[row_idx, dw_time_col] if dw_time_col else None
                    
                    channel = df.iloc[row_idx, 1]
                    if pd.isna(channel):
                        channel = sheet_name
                    
                    # 组合时效显示
                    time_display = str(time_val) if pd.notna(time_val) else '-'
                    if pd.notna(dw_time_val):
                        time_display += f" (DW: {dw_time_val})"
                    
                    results.append({
                        '渠道': str(channel) if pd.notna(channel) else sheet_name,
                        '渠道分类': sheet_name,
                        '时效': time_display,
                        '全程时效': str(time_val) if pd.notna(time_val) else '-',
                        'DW送达': str(dw_time_val) if pd.notna(dw_time_val) else '-',
                        '价格': price if pd.notna(price) else '-',
                        '仓库': warehouse_str,
                        '区域': normalized_region,
                        '税种': tax_type
                    })
                    break
        
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
st.markdown("---")

# 上传文件或使用默认文件
uploaded_file = st.file_uploader("上传报价表 (Excel)", type=['xlsx', 'xls'])

# 检查是否有默认报价表
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

# 获取仓库列表
warehouses = get_all_warehouses(df_dict)

# 查询界面
col1, col2, col3 = st.columns(3)

with col1:
    warehouse_input = st.text_input("🏭 仓库代码", placeholder="输入如 ONT8, BOS7...")
    if not warehouse_input and warehouses:
        warehouse_input = st.selectbox("或选择仓库", [""] + warehouses)

with col2:
    region = st.selectbox("📍 提货区域", REGIONS)

with col3:
    tax_type = st.radio("💰 税种", ["含税", "自税"], horizontal=True)

# 查询按钮
if st.button("🔍 查询价格", type="primary", use_container_width=True):
    if not warehouse_input:
        st.error("请输入或选择仓库代码")
    else:
        with st.spinner("正在查询..."):
            results = query_prices(df_dict, warehouse_input, region, tax_type)
        
        if results:
            st.markdown("---")
            st.subheader(f"📊 查询结果")
            st.markdown(f"**仓库:** {warehouse_input.upper()} | **区域:** {region} | **税种:** {tax_type}")
            
            # 显示最优推荐
            best = results[0]
            if best['价格'] != '-':
                st.success(f"💡 **推荐:** {best['渠道']} — 价格 ¥{best['价格']}/kg, 时效 {best['全程时效']}")
            
            # 显示完整表格（新增DW送达列）
            df_result = pd.DataFrame(results)[['渠道', '全程时效', 'DW送达', '价格', '渠道分类']]
            df_result.index = range(1, len(df_result) + 1)
            st.dataframe(df_result, use_container_width=True)
            
            # 下载按钮
            csv = df_result.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                "📥 下载结果 (CSV)",
                csv,
                f"报价查询_{warehouse_input}_{region}_{tax_type}.csv",
                "text/csv"
            )
        else:
            st.warning(f"❌ 未找到 {warehouse_input} 在 {region} 区域的报价")

# 页脚
st.markdown("---")
st.caption("Made with ❤️ by 强子 (OpenClaw) | v2.0 - 修复时效按区域显示 | 如需更新报价表，直接上传新文件即可")

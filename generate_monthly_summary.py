# -*- coding: utf-8 -*-
"""
生成（汇总）2026.4月各街道问题汇总.xlsx
严格遵循已有参考模板的格式: D:\桌面\实习\月报表\（汇总）2026.4月各街道问题汇总.xlsx

布局说明（转置格式）:
  - A列: 序号
  - B列: 分类名称（合并单元格）
  - C列: 问题指标名称
  - D~R列: 15个街道的数据
  - 数据来源: 源台账文件的居民小区胡同/社会单位/餐饮单位/中转站sheet
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import datetime
import os, warnings
warnings.filterwarnings('ignore')

SOURCE_FILE = r"D:\桌面\实习\月报表\1.2026-2027年早高峰时段各街道问题汇总台账（4.21开始含早晚非）(2).xlsx"
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "（汇总）2026.4月各街道问题汇总.xlsx")

BASE_DATE = datetime.datetime(1899, 12, 30)
START_DATE = datetime.datetime(2026, 4, 20)
END_DATE = datetime.datetime(2026, 5, 19, 23, 59, 59)
START_SERIAL = (START_DATE - BASE_DATE).days
END_SERIAL = (END_DATE - BASE_DATE).days

ALL_STREETS = [
    "德胜街道", "什刹海街道", "西长安街街道", "大栅栏街道",
    "天桥街道", "新街口街道", "金融街街道", "椿树街道",
    "陶然亭街道", "展览路街道", "月坛街道", "广内街道",
    "牛街街道", "白纸坊街道", "广外街道"
]

# ============================================================
# 数据加载与预计算
# ============================================================
def load_and_filter(sheet_name, header_row):
    """读取sheet并筛选日期范围"""
    df = pd.read_excel(SOURCE_FILE, sheet_name=sheet_name, header=header_row)
    
    def is_in_range(val):
        if pd.isna(val):
            return False
        if isinstance(val, (int, float)):
            return START_SERIAL <= val <= END_SERIAL
        if isinstance(val, datetime.datetime):
            return START_DATE <= val <= END_DATE
        return False
    
    mask = df.iloc[:, 0].apply(is_in_range)
    return df[mask].copy()

def get_street_sum(df, col_idx_excel):
    """按街道求和（0-indexed列索引）"""
    street_col = df.columns[1]  # 街道列
    data_col = df.columns[col_idx_excel]
    df_temp = df[[street_col, data_col]].copy()
    df_temp[data_col] = pd.to_numeric(df_temp[data_col], errors='coerce').fillna(0)
    result = df_temp.groupby(street_col)[data_col].sum()
    result.index = result.index.astype(str)
    return [int(result.get(s, 0)) for s in ALL_STREETS]

def get_street_count(df):
    """按街道统计记录数"""
    street_col = df.columns[1]
    counts = df.groupby(street_col).size()
    counts.index = counts.index.astype(str)
    return [int(counts.get(s, 0)) for s in ALL_STREETS]

def get_rectified_stats(df, col_idx_excel):
    """统计已整改/未整改数"""
    street_col = df.columns[1]
    rect_col = df.columns[col_idx_excel]
    df_temp = df[[street_col, rect_col]].copy()
    df_temp[rect_col] = pd.to_numeric(df_temp[rect_col], errors='coerce')
    
    rectified = df_temp[df_temp[rect_col] == 0].groupby(street_col).size()
    unrectified = df_temp[df_temp[rect_col] == 1].groupby(street_col).size()
    rectified.index = rectified.index.astype(str)
    unrectified.index = unrectified.index.astype(str)
    
    rect_arr = [int(rectified.get(s, 0)) for s in ALL_STREETS]
    unrect_arr = [int(unrectified.get(s, 0)) for s in ALL_STREETS]
    
    rate_arr = []
    for r, u in zip(rect_arr, unrect_arr):
        total = r + u
        rate_arr.append(round(r / total, 4) if total > 0 else 1.0)
    
    return rect_arr, unrect_arr, rate_arr

def calc_rate(numerator_col, denominator_col, df):
    """计算比率"""
    street_col = df.columns[1]
    num = pd.to_numeric(df[df.columns[numerator_col]], errors='coerce').fillna(0)
    den = pd.to_numeric(df[df.columns[denominator_col]], errors='coerce').fillna(0)
    df_temp = pd.DataFrame({street_col: df[street_col], "num": num, "den": den})
    result = df_temp.groupby(street_col).sum()
    rates = []
    for s in ALL_STREETS:
        if s in result.index:
            n = result.loc[s, "num"]
            d = result.loc[s, "den"]
            rates.append(round(1 - n/d, 4) if d > 0 else 1.0)
        else:
            rates.append(1.0)
    return rates

print("加载数据...")
df_resident = load_and_filter("居民小区胡同", 1)
df_social = load_and_filter("社会单位", 0)
df_food = load_and_filter("餐饮单位", 0)
df_transfer = load_and_filter("中转站", 0)   # 本月无数据
print(f"加载完成: 居民小区={len(df_resident)}, 社会单位={len(df_social)}, 餐饮单位={len(df_food)}, 中转站={len(df_transfer)}")

# ---- 预计算所有需要的数据 ----
# 居民小区数据
rd = {}  # resident data
for col_idx, name in [(5, "col6"), (6, "col7"), (7, "col8"), (8, "col9"),
                       (10, "col11"), (11, "col12"), (12, "col13"), (13, "col14"),
                       (16, "col17"), (17, "col18"), (18, "col19"), (19, "col20"),
                       (20, "col21"), (21, "col22"), (22, "col23"), (23, "col24"),
                       (24, "col25"), (28, "col29"), (30, "col31"), (32, "col33"),
                       (34, "col35"), (35, "col36"), (36, "col37"), (37, "col38"),
                       (40, "col41"), (41, "col42"), (43, "col44"), (44, "col45"),
                       (46, "col47"), (50, "col51"), (51, "col52")]:
    rd[name] = get_street_sum(df_resident, col_idx)

# 特殊列 - 桶站组数/容器总数/居民投放数（月报告列）
rd["col32"] = get_street_sum(df_resident, 31)  # 桶站组数
rd["col43"] = get_street_sum(df_resident, 42)  # 容器总数
rd["col46"] = get_street_sum(df_resident, 45)  # 居民投放数

# 检查次数
check_count_resident = get_street_count(df_resident)

# 整改统计
rd["rectified"], rd["unrectified"], rd["rectify_rate"] = get_rectified_stats(df_resident, 59)

# 投放准确率: 1 - 居民自主投放不准确/居民投放数
resident_input = rd["col46"]
resident_wrong = rd["col47"]
rd["accuracy"] = [round(1 - w/i, 4) if i > 0 else 1.0 for w, i in zip(resident_wrong, resident_input)]

# 桶内分类纯净率: 1 - 桶内分类不纯净/检查容器总数
rd["purity"] = calc_rate(43, 42, df_resident)  # col44/col43

# 值守率: 1 - 无人值守记录数/总记录数
unattended_count = []
street_col_r = df_resident.columns[1]
for s in ALL_STREETS:
    subset = df_resident[df_resident[street_col_r] == s]
    total = len(subset)
    if total > 0:
        unat = pd.to_numeric(subset[subset.columns[32]], errors='coerce').notna().sum()  # 列33无人值守
        unattended_count.append(round(1 - unat/total, 4))
    else:
        unattended_count.append(1.0)
rd["attendance_rate"] = unattended_count

# 社会单位数据
sd = {}
for col_idx, name in [(3, "col4"), (4, "col5"), (5, "col6"), (6, "col7"),
                       (7, "col8"), (8, "col9"), (9, "col10"), (10, "col11"),
                       (11, "col12"), (12, "col13"), (13, "col14"), (14, "col15"),
                       (15, "col16"), (16, "col17"), (17, "col18"), (18, "col19"),
                       (19, "col20"), (20, "col21"), (21, "col22"), (22, "col23"),
                       (23, "col24"), (24, "col25"), (25, "col26"), (26, "col27"),
                       (27, "col28"), (28, "col29"), (29, "col30"), (30, "col31"),
                       (31, "col32"), (32, "col33"), (33, "col34"), (34, "col35"), (35, "col36")]:
    sd[name] = get_street_sum(df_social, col_idx)

# 社会单位检查问题总数
sd_keys = [k for k in sd.keys()]
sd["total"] = [sum(sd[k][j] for k in sd_keys) for j in range(15)]

# 餐饮单位数据
fd = {}
for col_idx, name in [(3, "col4"), (4, "col5"), (5, "col6"), (6, "col7"),
                       (7, "col8"), (8, "col9"), (9, "col10"), (10, "col11"),
                       (11, "col12"), (12, "col13"), (13, "col14"), (14, "col15"),
                       (15, "col16"), (16, "col17"), (17, "col18"), (18, "col19"),
                       (19, "col20"), (20, "col21"), (21, "col22"), (22, "col23"),
                       (23, "col24"), (24, "col25"), (25, "col26"), (26, "col27"),
                       (27, "col28"), (28, "col29"), (29, "col30"), (30, "col31")]:
    fd[name] = get_street_sum(df_food, col_idx)

fd_keys = [k for k in fd.keys()]
fd["total"] = [sum(fd[k][j] for k in fd_keys) for j in range(15)]

# 中转站（本月无数据，全0）
transfer_zero = [0] * 15

print("数据预计算完成")

# ============================================================
# 模板行定义（完全参照参考模板）
# ============================================================
# 每行: (A序号, B分类, C指标名, 数据获取函数(街道索引->值))
# 如果是合计行，data_func=None，通过后续行自动计算

class RowDef:
    def __init__(self, seq, category, indicator, data_func=None, is_total=False, is_rate=False):
        self.seq = seq        # A列序号
        self.category = category  # B列分类
        self.indicator = indicator  # C列指标名
        self.data_func = data_func  # 函数: i(0-14) -> value
        self.is_total = is_total
        self.is_rate = is_rate

def row_defs_builder():
    rows = []
    
    # ===== 分类设施建设达标情况 (Row2-Row22) =====
    rows.append(RowDef(1, "分类设施建设达标情况", "合计", is_total=True))
    rows.append(RowDef(2, None, "无宣传氛围", lambda i: rd["col6"][i]))
    rows.append(RowDef(3, None, "无小区公示牌", lambda i: rd["col7"][i]))
    rows.append(RowDef(4, None, "小区公示牌设置不合格", lambda i: rd["col8"][i]))
    rows.append(RowDef(5, None, "未设置大件、装修垃圾投放点", lambda i: rd["col9"][i]))
    rows.append(RowDef(6, None, "大件、装修垃圾投放点无公示牌或公示牌不合格", lambda i: rd["col11"][i]))
    rows.append(RowDef(7, None, "无桶站公示牌或桶站公示牌设置不合格", lambda i: rd["col12"][i]))
    rows.append(RowDef(8, None, "厨余、其他垃圾和可回收物桶未成组配备", lambda i: rd["col13"][i]))
    rows.append(RowDef(9, None, "桶站公示牌破损", lambda i: rd["col14"][i]))
    rows.append(RowDef(10, None, "高峰时段未开盖", lambda i: rd["col17"][i]))
    rows.append(RowDef(11, None, "无便利性措施或便利性措施损坏", lambda i: rd["col18"][i]))
    rows.append(RowDef(12, None, "四分类桶设置不齐全", lambda i: rd["col19"][i]))
    rows.append(RowDef(13, None, "无标识、标识不符、脱落、破损等容器", lambda i: rd["col20"][i]))
    rows.append(RowDef(14, None, "容器颜色不符", lambda i: rd["col21"][i]))
    rows.append(RowDef(15, None, "容器不符、缺盖或破损", lambda i: rd["col22"][i]))
    rows.append(RowDef(16, None, "无防雨棚", lambda i: rd["col23"][i]))
    rows.append(RowDef(17, None, "容器脏污", lambda i: rd["col24"][i]))
    rows.append(RowDef(18, None, "未铺设地垫", lambda i: 0))
    rows.append(RowDef(19, None, "地垫破损", lambda i: 0))
    rows.append(RowDef(20, None, "无灭蝇措施", lambda i: rd["col29"][i]))
    rows.append(RowDef(21, None, "站外设桶", lambda i: rd["col45"][i]))
    
    # ===== 中转站 (Row23-Row25) =====
    rows.append(RowDef(22, "中转站", "称重系统损坏", lambda i: transfer_zero[i]))
    rows.append(RowDef(23, None, "灭火器不合格", lambda i: transfer_zero[i]))
    rows.append(RowDef(24, None, "无消防安全水源", lambda i: transfer_zero[i]))
    
    # ===== 分类设施管理达标情况 (Row26-Row43) =====
    rows.append(RowDef(28, "分类设施管理达标情况", "合计", is_total=True))
    rows.append(RowDef(29, None, "大件、装修垃圾投放点存放其他品类垃圾", lambda i: 0))
    rows.append(RowDef(30, None, "大件、装修垃圾投放点未分区存放", lambda i: 0))
    rows.append(RowDef(31, None, "大件、装修垃圾投放点周边环境脏乱", lambda i: 0))
    rows.append(RowDef(None, None, "小区内无大件垃圾托底上门回收渠道公示或告知", lambda i: 0))
    rows.append(RowDef(32, None, "桶站周边不洁", lambda i: rd["col25"][i]))
    rows.append(RowDef(33, None, "桶站地面脏污", lambda i: rd["col31"][i]))
    rows.append(RowDef(34, None, "无人值守", lambda i: rd["col33"][i]))
    rows.append(RowDef(35, None, "容器脏污", lambda i: rd["col24"][i]))
    rows.append(RowDef(36, None, "值守人员未履行职责", lambda i: rd["col35"][i]))
    rows.append(RowDef(37, None, "容器满冒", lambda i: rd["col36"][i]))
    rows.append(RowDef(38, None, "垃圾积存", lambda i: rd["col37"][i]))
    rows.append(RowDef(39, None, "垃圾清运不及时", lambda i: rd["col41"][i]))
    rows.append(RowDef(40, None, "小区垃圾乱堆乱放", lambda i: rd["col42"][i]))
    rows.append(RowDef(None, None, "散桶", lambda i: rd["col38"][i]))
    rows.append(RowDef(41, None, "桶内分类不纯净", lambda i: rd["col44"][i]))
    rows.append(RowDef(42, None, "车辆未密闭、破损滴漏脏污，标志与车牌不清晰", lambda i: rd["col51"][i]))
    rows.append(RowDef(43, None, "发现混装混运、随意倾倒、丢弃、遗撒、堆放垃圾", lambda i: rd["col52"][i]))
    
    # ===== 中转站2 (Row44-Row50) =====
    rows.append(RowDef(55, "中转站", "周边环境脏乱", lambda i: transfer_zero[i]))
    rows.append(RowDef(56, None, "无备案公示", lambda i: transfer_zero[i]))
    rows.append(RowDef(57, None, "未按规定区域存放物品", lambda i: transfer_zero[i]))
    rows.append(RowDef(58, None, "清运不及时、可回收物大量积存", lambda i: transfer_zero[i]))
    rows.append(RowDef(59, None, "安全员未按时上岗", lambda i: transfer_zero[i]))
    rows.append(RowDef(None, None, "安全员无明显身份标识", lambda i: transfer_zero[i]))
    rows.append(RowDef(60, None, "无企安安", lambda i: transfer_zero[i]))
    
    # ===== 居民自主投放情况 (Row51-Row52) =====
    rows.append(RowDef(61, "居民自主投放情况", "居民自主投放不准确", lambda i: rd["col47"][i]))
    rows.append(RowDef(62, None, "投放准确率", lambda i: rd["accuracy"][i], is_rate=True))
    
    # ===== 纯净率/值守率/桶站组数/整改情况 (Row53-Row59) =====
    rows.append(RowDef(63, "纯净率", "桶内分类纯净率", lambda i: rd["purity"][i], is_rate=True))
    rows.append(RowDef(64, "值守率", "值守率", lambda i: rd["attendance_rate"][i], is_rate=True))
    rows.append(RowDef(65, "检查小区桶站组数", None, lambda i: rd["col32"][i]))
    rows.append(RowDef(66, "居民小区问题整改情况", "居民小区、胡同整改率", lambda i: rd["rectify_rate"][i], is_rate=True))
    rows.append(RowDef(67, None, "检查数", lambda i: check_count_resident[i]))
    rows.append(RowDef(68, None, "已整改数", lambda i: rd["rectified"][i]))
    rows.append(RowDef(69, None, "未整改数", lambda i: rd["unrectified"][i]))
    
    # ===== 社会单位检查情况 (Row60-Row88) =====
    rows.append(RowDef(70, "社会单位检查情况", "社会单位检查问题数", lambda i: sd["total"][i]))
    rows.append(RowDef(71, None, "无党建引领相关资料", lambda i: sd["col4"][i]))
    rows.append(RowDef(72, None, "无培训活动、会议记录、照片等培训材料", lambda i: sd["col5"][i]))
    rows.append(RowDef(73, None, "无适量点餐、光盘行动等宣传内容", lambda i: sd["col6"][i]))
    rows.append(RowDef(74, None, "单位无容器配置或容器配置不全", lambda i: sd["col7"][i]))
    rows.append(RowDef(75, None, "未设置分类投放指引数（组数）", lambda i: sd["col8"][i]))
    rows.append(RowDef(76, None, "公共场所区域(办公楼外区域、办事大厅等)未成组设置可回收物和其他容器数（组数）", lambda i: sd["col9"][i]))
    rows.append(RowDef(77, None, "容器无便利性措施", lambda i: sd["col12"][i]))
    rows.append(RowDef(78, None, "未更新新国际", lambda i: sd["col13"][i]))
    rows.append(RowDef(79, None, "容器垃圾不纯净", lambda i: sd["col14"][i]))
    rows.append(RowDef(80, None, "容器标识不合格数", lambda i: sd["col18"][i]))
    rows.append(RowDef(None, None, "桶站满冒", lambda i: sd["col15"][i]))
    rows.append(RowDef(None, None, "桶站周边不洁", lambda i: sd["col16"][i]))
    rows.append(RowDef(81, None, "无宣传氛围", lambda i: sd["col19"][i]))
    rows.append(RowDef(82, None, "集中用餐区未成组设置厨余和其他垃圾容器数", lambda i: sd["col20"][i]))
    rows.append(RowDef(83, None, "食品加工区厨余、其他容器设置不齐", lambda i: sd["col21"][i]))
    rows.append(RowDef(84, None, "无油水分离装置", lambda i: sd["col22"][i]))
    rows.append(RowDef(85, None, "无厨余垃圾收运合同或不合格", lambda i: sd["col23"][i]))
    rows.append(RowDef(86, None, "无其他垃圾收运合同或不合格", lambda i: sd["col24"][i]))
    rows.append(RowDef(87, None, "无可回收物收运合同或不合格", lambda i: sd["col25"][i]))
    rows.append(RowDef(88, None, "无有害垃圾收运合同或不合格", lambda i: sd["col26"][i]))
    rows.append(RowDef(89, None, "无源头减量措施", lambda i: sd["col29"][i]))
    rows.append(RowDef(90, None, "无非居民其他垃圾排放登记方式", lambda i: sd["col28"][i]))
    rows.append(RowDef(91, None, "无废弃油脂合同或不合格", lambda i: sd["col30"][i]))
    rows.append(RowDef(92, None, "无垃圾分类工作方案", lambda i: sd["col31"][i]))
    rows.append(RowDef(93, None, "职责分工不明确", lambda i: sd["col32"][i]))
    rows.append(RowDef(94, None, "无四分类垃圾清运台账或四分类清运台账不合格", lambda i: sd["col33"][i]))
    rows.append(RowDef(95, None, "无称重计量列表", lambda i: sd["col34"][i]))
    rows.append(RowDef(96, None, "无油水分离装置", lambda i: sd["col35"][i]))
    
    # ===== 餐饮单位检查情况 (Row89-Row110) =====
    rows.append(RowDef(101, "餐饮单位检查情况", "餐饮单位检查问题数", lambda i: fd["total"][i]))
    rows.append(RowDef(102, None, "无适量点餐、光盘行动等宣传内容", lambda i: fd["col4"][i]))
    rows.append(RowDef(103, None, "集中用餐区未成组设置厨余和其他垃圾容器数", lambda i: fd["col5"][i]))
    rows.append(RowDef(104, None, "后厨未成组设置垃圾桶", lambda i: fd["col6"][i]))
    rows.append(RowDef(105, None, "容器无标识或标识不合格数", lambda i: fd["col10"][i]))
    rows.append(RowDef(106, None, "无便利性措施", lambda i: fd["col11"][i]))
    rows.append(RowDef(107, None, "无宣传氛围", lambda i: fd["col13"][i]))
    rows.append(RowDef(108, None, "容器破损、脏污等", lambda i: fd["col14"][i]))
    rows.append(RowDef(109, None, "无油水分离装置", lambda i: fd["col15"][i]))
    rows.append(RowDef(110, None, "无垃圾分类投放指引", lambda i: fd["col16"][i]))
    rows.append(RowDef(111, None, "无厨余垃圾收运合同或不合格", lambda i: fd["col17"][i]))
    rows.append(RowDef(112, None, "无其他垃圾收运合同或不合格", lambda i: fd["col18"][i]))
    rows.append(RowDef(113, None, "无厨余垃圾排放登记方式", lambda i: fd["col19"][i]))
    rows.append(RowDef(114, None, "无非居民其他垃圾排放登记方式", lambda i: fd["col20"][i]))
    rows.append(RowDef(115, None, "无称重计量小程序", lambda i: fd["col21"][i]))
    rows.append(RowDef(116, None, "无源头减量措施", lambda i: fd["col22"][i]))
    rows.append(RowDef(117, None, "厨余垃圾桶外摆", lambda i: fd["col23"][i]))
    rows.append(RowDef(118, None, "容器垃圾不纯净", lambda i: fd["col24"][i]))
    rows.append(RowDef(119, None, "无废弃油脂合同或不合格", lambda i: fd["col25"][i]))
    rows.append(RowDef(120, None, "无容器设置", lambda i: fd["col27"][i]))
    rows.append(RowDef(121, None, "无收费计量记录", lambda i: fd["col29"][i]))
    rows.append(RowDef(122, None, "无隔油池", lambda i: fd["col30"][i]))
    
    # ===== 市级检查情况 (Row111-Row113) =====
    rows.append(RowDef(127, "市级检查情况", "检查小区数", lambda i: 0))
    rows.append(RowDef(128, None, "检查社会单位数", lambda i: 0))
    rows.append(RowDef(129, None, "市级检查问题数", lambda i: 0))
    
    return rows

rows = row_defs_builder()

# ============================================================
# 计算合计行
# ============================================================
# 找出每个分类下的合计行索引和为它提供数据的范围
total_indices = [i for i, r in enumerate(rows) if r.is_total]
for ti in total_indices:
    cat = rows[ti].category
    # 从ti+1开始，直到遇到下一个分类或结束
    end = len(rows)
    for j in range(ti + 1, len(rows)):
        if rows[j].category is not None and not rows[j].is_total:
            end = j
            break
    # 计算合计
    detail_rows = rows[ti+1:end]
    def make_total_func(details):
        def func(i):
            total = 0
            for dr in details:
                if dr.data_func and not dr.is_rate:
                    total += dr.data_func(i)
            return total
        return func
    rows[ti].data_func = make_total_func(detail_rows)

# ============================================================
# 写入Excel
# ============================================================
print("写入Excel...")
wb = Workbook()
ws = wb.active
ws.title = "Sheet1"

# 样式
title_font = Font(name="仿宋", size=16, bold=False)
header_font = Font(name="仿宋", size=12, bold=False)
data_font = Font(name="仿宋", size=12)
center_align = Alignment(horizontal="center", vertical="center")
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

# Row1: 表头
ws.cell(row=1, column=1).value = None
ws.cell(row=1, column=2).value = None
ws.cell(row=1, column=3).value = "街道名称"
ws.cell(row=1, column=3).font = Font(name="仿宋", size=12)
ws.cell(row=1, column=3).alignment = center_align
ws.cell(row=1, column=3).border = thin_border

for j, street in enumerate(ALL_STREETS):
    cell = ws.cell(row=1, column=4+j, value=street)
    cell.font = header_font
    cell.alignment = center_align
    cell.border = thin_border

# 填充数据行
# 合并单元格（按固定区域）

for i, rdef in enumerate(rows):
    row_num = i + 2  # Excel行号（从第2行开始）
    
    # A列 - 序号
    if rdef.seq is not None:
        cell = ws.cell(row=row_num, column=1, value=rdef.seq)
        cell.font = data_font
        cell.alignment = center_align
        cell.border = thin_border
    
    # B列 - 分类（单行也设置完整边框）
    if rdef.category is not None:
        cell = ws.cell(row=row_num, column=2, value=rdef.category)
        cell.font = data_font
        cell.alignment = center_align
        cell.border = thin_border
    else:
        # 确保非分类行B列也有边框
        cell = ws.cell(row=row_num, column=2)
        cell.border = thin_border
    
    # C列 - 指标名称
    if rdef.indicator is not None:
        cell = ws.cell(row=row_num, column=3, value=rdef.indicator)
        cell.font = data_font
        cell.alignment = center_align
        cell.border = thin_border
    elif rdef.category is not None and rdef.indicator is None:
        # 一些行有分类但没有指标（如"检查小区桶站组数"行）
        pass
    
    # D~R列 - 数据
    if rdef.data_func:
        for j in range(15):
            val = rdef.data_func(j)
            cell = ws.cell(row=row_num, column=4+j)
            if rdef.is_rate:
                cell.value = val
                cell.number_format = '0.0000' if isinstance(val, float) else '0'
            else:
                cell.value = val if val != 0 else 0  # 保持0值
            cell.font = data_font
            cell.alignment = center_align
            cell.border = thin_border

# 合并分类单元格（按连续区域，参考模板格式）
merge_regions = [
    (2, 22),   # 分类设施建设达标情况
    (23, 25),  # 中转站
    (26, 43),  # 分类设施管理达标情况
    (44, 50),  # 中转站
    (51, 52),  # 居民自主投放情况
    (53, 53),  # 纯净率
    (54, 54),  # 值守率
    (55, 55),  # 检查小区桶站组数
    (56, 59),  # 居民小区问题整改情况
    (60, 88),  # 社会单位检查情况
    (89, 110), # 餐饮单位检查情况
    (111, 113),# 市级检查情况
]
for start, end in merge_regions:
    if end > start:
        ws.merge_cells(start_row=start, start_column=2, end_row=end, end_column=2)

# 设置列宽
ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 22
ws.column_dimensions['C'].width = 52
for j in range(15):
    ws.column_dimensions[get_column_letter(4+j)].width = 12

print(f"写入完成: {len(rows)}行数据")

# ============================================================
# Sheet2: 居民自主投放情况明细
# ============================================================
ws2 = wb.create_sheet(title="Sheet2")

ws2.cell(row=1, column=1, value="街道名称").font = Font(name="宋体", size=11)
ws2.cell(row=1, column=2, value="求和项:居民自主投放不准确").font = Font(name="宋体", size=11)
ws2.cell(row=1, column=3, value="平均值项:准确率").font = Font(name="宋体", size=11)
for col_idx in range(1, 4):
    ws2.cell(row=1, column=col_idx).alignment = center_align
    ws2.cell(row=1, column=col_idx).border = thin_border

for j, street in enumerate(ALL_STREETS):
    row_num = j + 2
    ws2.cell(row=row_num, column=1, value=street).border = thin_border
    ws2.cell(row=row_num, column=1).alignment = Alignment(vertical="center")
    ws2.cell(row=row_num, column=2, value=rd["col47"][j]).border = thin_border
    ws2.cell(row=row_num, column=2).alignment = center_align
    ws2.cell(row=row_num, column=3, value=rd["accuracy"][j]).border = thin_border
    ws2.cell(row=row_num, column=3).number_format = '0.0000'
    ws2.cell(row=row_num, column=3).alignment = center_align

# 总计行
ws2.cell(row=17, column=1, value="总计").border = thin_border
ws2.cell(row=17, column=1).font = Font(name="宋体", size=11, bold=True)
ws2.cell(row=17, column=2, value=sum(rd["col47"])).border = thin_border
total_acc = round(sum(rd["col47"]) / sum(rd["col46"]), 4) if sum(rd["col46"]) > 0 else 0
ws2.cell(row=17, column=3, value=total_acc).border = thin_border
ws2.cell(row=17, column=3).number_format = '0.0000'

ws2.column_dimensions['A'].width = 14
ws2.column_dimensions['B'].width = 26
ws2.column_dimensions['C'].width = 18

# 保存
wb.save(OUTPUT_FILE)
print(f"\n文件已保存: {OUTPUT_FILE}")
print("完成！")

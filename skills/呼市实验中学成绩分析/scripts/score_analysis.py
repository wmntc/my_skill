# -*- coding: utf-8 -*-
"""
考试成绩分析与教师绩效算分脚本（命令行版）

功能：读取包含「成绩表」和「教师任课表」的 Excel 文件，按班级计算指标并按可配置权重
生成教师得分表，最终导出多工作表的结果 Excel。分两种模式：

  · 两场考试对比（默认，--exam-count 2）：六项指标
    排名正负 / T分差 / 均分 / 优秀人数差 / 出库人数差 / 参考率
  · 单次考试·首次考试（--exam-count 1）：四项指标
    均分 / 优秀率 / 合格率 / 参考率（均分与排名合为一项，排名不单列）
    均分 + 排名 / 优秀率 / 合格率 / 参考率

场次默认按成绩表列名自动识别（存在 '2' 结尾的科目列即为两场考试）。

用法示例：
    python score_analysis.py -i 成绩数据.xlsx -o 分析结果.xlsx \
        --excellent-ratio 20 --qualified-ratio 75 \
        --rank-weight 20 --t-diff-weight 20 --avg-weight 20 \
        --excellent-weight 15 --qualified-weight 15 --reference-weight 10 \
        --exam-date 2026-01-15 --cohort 2028 --semester 1 --exam-name 期末考试

    python score_analysis.py -i 首次考试.xlsx -o 首次考试结果.xlsx --exam-count 1 \
        --excellent-ratio 20 --qualified-ratio 80 \
        --avg-weight 60 \
        --excellent-weight 15 --qualified-weight 15 --reference-weight 10 \
        --exam-date 2026-11-10 --cohort 2026 --semester 1 --exam-name 第一次期中考试
"""

import argparse
import json
import os
import re
import sys
from datetime import datetime
from typing import Dict, List, Tuple

import pandas as pd

# ---------------- 常量配置 ----------------
T_SCORE_AVG = 50          # T分平均分
T_SCORE_STD = 10          # T分标准差
BASIC_COLUMNS = ["考号", "姓名", "集团班级"]

# 两种分析模式的默认参数
#   两场考试（六项指标）：排名正负 / T分差 / 均分 / 优秀人数 / 出库人数 / 参考率
#   单次考试（四项指标）：均分 / 优秀率 / 合格率 / 参考率
#     （「均分与排名」视为同一项指标：排名本就是班级均分的横向位次，不单列，
#       故均分项合并了原「均分 25 + 排名 15」的权重，合计 60）
DUAL_DEFAULT_WEIGHTS = {
    "rank": 20.0, "t_diff": 20.0, "avg": 20.0,
    "excellent": 15.0, "qualified": 15.0, "reference": 10.0,
}
SINGLE_DEFAULT_WEIGHTS = {
    "avg": 60.0, "excellent": 15.0, "qualified": 15.0, "reference": 10.0,
}
DUAL_DEFAULT_RATIOS = {"excellent": 20.0, "qualified": 75.0}
SINGLE_DEFAULT_RATIOS = {"excellent": 20.0, "qualified": 80.0}



def detect_exam_count(df: pd.DataFrame) -> int:
    """
    按成绩表列名判断考试场次数。

    存在以 '2' 结尾的科目列（总分2 除外）→ 判为两场考试；否则判为单次考试（首次考试）。
    """
    has_second = any(col.endswith("2") and col != "总分2" for col in df.columns)
    return 2 if has_second else 1


# ---------------- 混合口径：按科目分流 ----------------
# 口径约定：第一场（'1' 后缀列）= 基准成绩（上次考试），第二场（'2' 后缀列）= 本次成绩。
# 混合场景用于「部分科目新开课、暂无基准成绩」的情况，例如初二新开物理、初三新开化学。
MIXED_MODE = 3


def detect_subject_modes(df: pd.DataFrame) -> Tuple[Dict[str, Dict], List[str]]:
    """
    逐科目判定评价口径（混合口径的入口判定）。

    判定规则：
      - 基准列（科目1）有有效成绩且本次列（科目2）存在 → 两场对比口径（mode 2）
      - 本次列（科目2）有成绩、基准列为空或不存在 → 首次考试口径（mode 1）
      - 仅有不带后缀的学科名列 → 首次考试口径（mode 1）
      - 各列均无有效成绩 → 跳过该科目

    返回 (科目口径字典, 处理日志)；字典形如
    {"物理": {"mode": 1, "base_col": None, "cur_col": "物理2"}, ...}
    """
    process_log = []
    modes: Dict[str, Dict] = {}

    subject_columns_1 = [c for c in df.columns if c.endswith("1") and c != "总分1"]
    subject_columns_2 = [c for c in df.columns if c.endswith("2") and c != "总分2"]

    col_map: Dict[str, Dict[str, str]] = {}
    for col in subject_columns_1:
        col_map.setdefault(col[:-1], {})["base"] = col
    for col in subject_columns_2:
        col_map.setdefault(col[:-1], {})["cur"] = col
    for col in df.columns:
        if col in BASIC_COLUMNS or col in ("总分", "总分1", "总分2"):
            continue
        if col.endswith("1") or col.endswith("2"):
            continue
        col_map.setdefault(col, {})["plain"] = col

    for subject, cols in col_map.items():
        base_col = cols.get("base")
        cur_col = cols.get("cur")
        plain_col = cols.get("plain")

        has_base = base_col is not None and bool(df[base_col].notna().any())
        has_cur = cur_col is not None and bool(df[cur_col].notna().any())
        has_plain = plain_col is not None and bool(df[plain_col].notna().any())

        if cur_col is not None:
            if has_base and has_cur:
                modes[subject] = {"mode": 2, "base_col": base_col, "cur_col": cur_col,
                                  "has_base": True}
                process_log.append(f"科目「{subject}」：有基准成绩与本次成绩 → 两场对比口径（六项指标）")
            elif has_cur:
                modes[subject] = {"mode": 1, "base_col": base_col, "cur_col": cur_col,
                                  "has_base": has_base}
                process_log.append(f"科目「{subject}」：无基准成绩 → 首次考试口径（四项指标），"
                                   f"且该科目不计入总分")
            elif has_base:
                modes[subject] = {"mode": 1, "base_col": base_col, "cur_col": base_col,
                                  "has_base": True}
                process_log.append(f"科目「{subject}」：本次成绩整列为空、仅有基准成绩，"
                                   f"按首次考试口径处理（以基准成绩作为唯一一次成绩）")
            else:
                process_log.append(f"科目「{subject}」：两场均无有效成绩，已跳过")
        elif base_col is not None:
            if has_base:
                modes[subject] = {"mode": 1, "base_col": base_col, "cur_col": base_col,
                                  "has_base": True}
                process_log.append(f"科目「{subject}」：仅一次成绩 → 首次考试口径（四项指标）")
            else:
                process_log.append(f"科目「{subject}」：成绩整列为空，已跳过")
        elif plain_col is not None:
            if has_plain:
                modes[subject] = {"mode": 1, "base_col": None, "cur_col": plain_col,
                                  "has_base": False}
                process_log.append(f"科目「{subject}」：仅一次成绩（无后缀列）→ 首次考试口径（四项指标）")
            else:
                process_log.append(f"科目「{subject}」：成绩整列为空，已跳过")

    return modes, process_log


def apply_subject_mode_overrides(modes: Dict[str, Dict], overrides: str) -> Tuple[Dict[str, Dict], List[str]]:
    """
    手工覆盖个别科目的口径。overrides 形如 "物理:1,化学:2"。

    1 = 首次考试口径（四项指标），2 = 两场对比口径（六项指标）。
    覆盖为 2 时要求该科目同时存在基准列与本次列，否则报错。
    """
    process_log = []
    if not overrides:
        return modes, process_log
    for item in overrides.split(","):
        item = item.strip()
        if not item:
            continue
        if ":" not in item:
            raise ValueError(f"科目口径覆盖格式错误：{item}，应为「科目:1」或「科目:2」")
        subject, _, val = item.partition(":")
        subject = subject.strip()
        val = val.strip()
        if val not in ("1", "2"):
            raise ValueError(f"科目「{subject}」的口径只能是 1（首次考试）或 2（两场对比），收到：{val}")
        if subject not in modes:
            raise ValueError(f"未找到科目「{subject}」，无法覆盖口径。可用科目：{', '.join(modes.keys())}")
        target = int(val)
        if target == 2 and not modes[subject].get("has_base"):
            raise ValueError(f"科目「{subject}」没有有效的基准成绩，无法按两场对比口径计算")
        if target == 1:
            modes[subject] = {"mode": 1, "base_col": modes[subject].get("base_col"),
                               "cur_col": modes[subject].get("cur_col"),
                               "has_base": modes[subject].get("has_base", False)}
        else:
            modes[subject] = {"mode": 2, "base_col": modes[subject]["base_col"],
                               "cur_col": modes[subject]["cur_col"], "has_base": True}
        process_log.append(f"科目「{subject}」口径已手工指定为 "
                           f"{'两场对比口径（六项指标）' if target == 2 else '首次考试口径（四项指标）'}")
    return modes, process_log


def validate_and_transform_data_mixed(df: pd.DataFrame,
                                     subject_mode_overrides: str = "") -> Tuple[Dict[str, pd.DataFrame],
                                                                          Dict[str, Dict], List[str]]:
    """
    校验并转换「混合口径」数据：逐科目分流 + 总分按「含有基准成绩的科目之和」构造。

    总分规则：
      - 只要有科目具备基准成绩，总分即由这些科目求和构成（两场各自求和，口径可比）；
      - 无基准成绩的科目（如初二新开的物理）单独按首次考试口径评价，且不计入总分；
      - 混合场景下弃用表内原生总分列（含无基准科目，两场不可比），改按上述规则重算。

    subject_mode_overrides：手工覆盖个别科目口径，形如 "物理:1,化学:2"。
    覆盖在总分构造之前生效，因此总分所含科目与覆盖后的口径始终一致。
    """
    process_log = []

    missing_columns = [col for col in BASIC_COLUMNS if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Excel 成绩表缺少必要的基本信息列：{', '.join(missing_columns)}")

    df = df.copy()
    df["集团班级"] = df["集团班级"].apply(normalize_class)
    process_log.append(f"数据包含所有基本信息列：{', '.join(BASIC_COLUMNS)}")

    modes, mode_log = detect_subject_modes(df)
    process_log.extend(mode_log)
    if not modes:
        raise ValueError("未找到任何有成绩的科目列")

    modes, override_log = apply_subject_mode_overrides(modes, subject_mode_overrides)
    process_log.extend(override_log)

    parsed_data: Dict[str, pd.DataFrame] = {}
    for subject, info in modes.items():
        if info["mode"] == 2:
            sub = df[BASIC_COLUMNS + [info["base_col"], info["cur_col"]]].rename(
                columns={info["base_col"]: "第一场成绩", info["cur_col"]: "第二场成绩"})
            sub["第一场成绩"] = pd.to_numeric(sub["第一场成绩"], errors="coerce")
            sub["第二场成绩"] = pd.to_numeric(sub["第二场成绩"], errors="coerce")
            sub["参考状态"] = ~sub["第二场成绩"].isna()
            absent = len(sub) - int(sub["参考状态"].sum())
            process_log.append(f"{subject}科目本次缺考人数：{absent}")
        else:
            col = info["cur_col"]
            sub = df[BASIC_COLUMNS + [col]].rename(columns={col: "成绩"})
            sub["成绩"] = pd.to_numeric(sub["成绩"], errors="coerce")
            sub["参考状态"] = sub["成绩"].notna()
            absent = len(sub) - int(sub["参考状态"].sum())
            process_log.append(f"{subject}科目缺考人数：{absent}")
        parsed_data[subject] = sub

    # ---- 总分：只含有基准成绩的科目之和 ----
    benchmark_subjects = [s for s, i in modes.items() if i["mode"] == 2]
    single_subjects = [s for s, i in modes.items() if i["mode"] == 1]
    if benchmark_subjects:
        total_subjects = benchmark_subjects
        total_mode = 2
    else:
        total_subjects = list(modes.keys())
        total_mode = 1

    if "总分1" in df.columns or "总分2" in df.columns or "总分" in df.columns:
        process_log.append("检测到表内自带总分列：其中可能含有无基准科目（如新开物理）的成绩，"
                           "两场总分口径不一致，已弃用该列并按「含有基准成绩的科目之和」重新计算。")

    if total_mode == 2:
        base_sum = df[[modes[s]["base_col"] for s in total_subjects]].apply(
            pd.to_numeric, errors="coerce").sum(axis=1, skipna=True)
        cur_sum = df[[modes[s]["cur_col"] for s in total_subjects]].apply(
            pd.to_numeric, errors="coerce").sum(axis=1, skipna=True)
        total_df = df[BASIC_COLUMNS].copy()
        total_df["第一场成绩"] = base_sum
        total_df["第二场成绩"] = cur_sum
        total_df["参考状态"] = ~total_df["第二场成绩"].isna()
        total_df["第一场成绩"] = pd.to_numeric(total_df["第一场成绩"], errors="coerce")
        total_df["第二场成绩"] = pd.to_numeric(total_df["第二场成绩"], errors="coerce")
        absent = len(total_df) - int(total_df["参考状态"].sum())
        process_log.append(f"总分口径：两场对比；总分由有基准的科目求和构成——"
                           f"{', '.join(total_subjects)}；本次缺考人数：{absent}")
    else:
        cur_sum = df[[modes[s]["cur_col"] for s in total_subjects]].apply(
            pd.to_numeric, errors="coerce").sum(axis=1, skipna=True)
        total_df = df[BASIC_COLUMNS].copy()
        total_df["成绩"] = cur_sum
        total_df["成绩"] = pd.to_numeric(total_df["成绩"], errors="coerce")
        total_df["参考状态"] = total_df["成绩"].notna()
        process_log.append(f"总分口径：首次考试；总分由全部科目求和构成——{', '.join(total_subjects)}")

    modes["总分"] = {"mode": total_mode,
                     "base_col": None,
                     "cur_col": None}
    parsed_data["总分"] = total_df

    if single_subjects:
        process_log.append(f"总分不含以下无基准科目：{', '.join(single_subjects)}；"
                           f"这些科目单独按首次考试口径评价。")

    return parsed_data, modes, process_log


def natural_sort_key(s):
    """生成自然排序的键（如 高一1班 < 高一2班 < 高一10班）

    返回形如 (("t", "高一"), ("n", 1), ("t", "班")) 的元组：
    文本段统一为 ("t", 小写文本)，数字段统一为 ("n", 整数)，
    保证任意两段之间类型一致、可直接比较，同时元组可哈希，
    可安全用于 pandas sort_values(key=...) 的多列排序
    （pandas 多列排序会把键做因子化，list 类型会报 unhashable）。
    """
    s_str = str(s)
    key = []
    for text in re.split(r"(\d+)", s_str):
        if text.isdigit():
            key.append(("n", int(text)))
        else:
            key.append(("t", text.lower()))
    return tuple(key)


def normalize_class(value) -> str:
    """统一班级字段为字符串，避免 int/float/str 类型差异导致匹配失败"""
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def build_class_avg_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    计算各班级的两场均分（均分指标的唯一数据来源）。

    「均分」口径：对每一科，按班级范围对该科成绩求平均，即班级均分。
    统计范围与其余指标保持一致——只取两场都有成绩的参考学生，
    保证同一个班内各指标口径统一（参考率指标另计，见 calculate_reference_rate）。

    返回列：班级 / 第一场均分 / 第二场均分（均保留 2 位小数，按班级自然序排序）。
    """
    reference_df = df[(df["参考状态"]) & (df["第一场成绩"].notna()) & (df["第二场成绩"].notna())].copy()

    class_avg_df = reference_df.groupby("集团班级").agg(
        第一场均分=pd.NamedAgg(column="第一场成绩", aggfunc="mean"),
        第二场均分=pd.NamedAgg(column="第二场成绩", aggfunc="mean")
    ).reset_index().rename(columns={"集团班级": "班级"})
    class_avg_df["第一场均分"] = class_avg_df["第一场均分"].round(2)
    class_avg_df["第二场均分"] = class_avg_df["第二场均分"].round(2)
    return class_avg_df.sort_values(by="班级", key=lambda x: x.apply(natural_sort_key))


# ---------------- 数据解析 ----------------
def validate_and_transform_data(df: pd.DataFrame) -> Tuple[Dict[str, pd.DataFrame], List[str]]:
    """
    校验数据格式并转换为按科目组织的数据字典。
    科目列规则：以 '1' 结尾为第一场考试科目，以 '2' 结尾为第二场考试科目，
    '总分1'/'总分2' 单独处理；若缺少总分列则按科目求和自动计算。
    """
    process_log = []

    missing_columns = [col for col in BASIC_COLUMNS if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Excel 成绩表缺少必要的基本信息列：{', '.join(missing_columns)}")

    df = df.copy()
    df["集团班级"] = df["集团班级"].apply(normalize_class)
    process_log.append(f"数据包含所有基本信息列：{', '.join(BASIC_COLUMNS)}")

    subject_columns_1 = []
    subject_columns_2 = []
    for col in df.columns:
        if col.endswith("1") and col != "总分1":
            subject_columns_1.append(col)
        elif col.endswith("2") and col != "总分2":
            subject_columns_2.append(col)

    if not subject_columns_1:
        raise ValueError("未找到第一次考试的科目列（以'1'结尾的列）")

    if not subject_columns_2:
        # 只指定了两场对比口径，但文件里根本没有第二场科目列 —— 多半是首次考试数据
        raise ValueError(
            "按两场考试对比处理，但成绩表中找不到任何第二场科目列（以'2'结尾的列）。"
            "该文件看起来只有一次考试成绩，请改用首次考试口径：加上 --exam-count 1")

    process_log.append(f"识别到第一次考试科目列：{', '.join(subject_columns_1)}")
    process_log.append(f"识别到第二次考试科目列：{', '.join(subject_columns_2)}")

    subjects = [col[:-1] for col in subject_columns_1]
    process_log.append(f"提取到科目：{', '.join(subjects)}")

    parsed_data = {}
    for subject in subjects:
        col_1 = f"{subject}1"
        col_2 = f"{subject}2"

        if col_2 in df.columns:
            subject_df = df[BASIC_COLUMNS + [col_1, col_2]].rename(
                columns={col_1: "第一场成绩", col_2: "第二场成绩"})
        else:
            subject_df = df[BASIC_COLUMNS + [col_1]].rename(columns={col_1: "第一场成绩"})

        if col_2 in df.columns:
            subject_df["参考状态"] = ~subject_df["第二场成绩"].isna()
        else:
            subject_df["第二场成绩"] = pd.NA
            subject_df["参考状态"] = False

        for score_col in ["第一场成绩", "第二场成绩"]:
            if score_col in subject_df.columns:
                subject_df[score_col] = pd.to_numeric(subject_df[score_col], errors="coerce")

        absent_count = len(subject_df) - int(subject_df["参考状态"].sum())
        process_log.append(f"{subject}科目第二场考试缺考人数：{absent_count}")
        parsed_data[subject] = subject_df

    # 总分数据
    total_df = df[BASIC_COLUMNS].copy()
    if "总分1" in df.columns:
        total_df["总分1"] = df["总分1"]
        process_log.append("使用预计算的总分1数据")
    elif subject_columns_1:
        total_df["总分1"] = df[subject_columns_1].sum(axis=1, skipna=True)
        process_log.append("数据中没有总分1列，已按科目求和计算总分1")

    if "总分2" in df.columns:
        total_df["总分2"] = df["总分2"]
        process_log.append("使用预计算的总分2数据")
    elif subject_columns_2:
        total_df["总分2"] = df[subject_columns_2].sum(axis=1, skipna=True)
        process_log.append("数据中没有总分2列，已按科目求和计算总分2")

    total_df = total_df.rename(columns={"总分1": "第一场成绩", "总分2": "第二场成绩"})
    if "第二场成绩" in total_df.columns:
        total_df["参考状态"] = ~total_df["第二场成绩"].isna()
    else:
        total_df["第二场成绩"] = pd.NA
        total_df["参考状态"] = False

    for score_col in ["第一场成绩", "第二场成绩"]:
        if score_col in total_df.columns:
            total_df[score_col] = pd.to_numeric(total_df[score_col], errors="coerce")

    parsed_data["总分"] = total_df
    return parsed_data, process_log


def validate_and_transform_data_single(df: pd.DataFrame) -> Tuple[Dict[str, pd.DataFrame], List[str]]:
    """
    校验并转换「单次考试」（首次考试）数据。

    列名规则：
      - 科目列以 '1' 结尾（与两场考试模板保持一致，如 语文1 / 数学1 / 总分1）；
      - 也支持不带后缀的科目列（语文 / 数学 / 总分），此时以「非基本信息列」识别；
      - 总分列缺失时按科目求和自动计算。

    输出的每个科目 DataFrame 统一为：考号 / 姓名 / 集团班级 / 成绩 / 参考状态。
    """
    process_log = []

    missing_columns = [col for col in BASIC_COLUMNS if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Excel 成绩表缺少必要的基本信息列：{', '.join(missing_columns)}")

    df = df.copy()
    df["集团班级"] = df["集团班级"].apply(normalize_class)
    process_log.append(f"数据包含所有基本信息列：{', '.join(BASIC_COLUMNS)}")

    subject_columns = [col for col in df.columns if col.endswith("1") and col != "总分1"]
    has_suffix = True
    if not subject_columns:
        # 回退：使用不带后缀的科目列（语文 / 数学 …）
        excluded = set(BASIC_COLUMNS) | {"总分", "总分1"}
        subject_columns = [col for col in df.columns if col not in excluded]
        has_suffix = False
    if not subject_columns:
        raise ValueError("未找到任何科目列（单次考试模式：科目列应带 '1' 后缀，或直接使用学科名）")

    process_log.append(f"识别到科目列：{', '.join(subject_columns)}"
                       + ("（带 1 后缀）" if has_suffix else "（无后缀）"))

    parsed_data = {}
    for col in subject_columns:
        subject = col[:-1] if has_suffix else col
        if subject in parsed_data:
            process_log.append(f"警告：科目「{subject}」出现重复列（{col}），已跳过重复项")
            continue
        subject_df = df[BASIC_COLUMNS + [col]].rename(columns={col: "成绩"})
        subject_df["成绩"] = pd.to_numeric(subject_df["成绩"], errors="coerce")
        subject_df["参考状态"] = subject_df["成绩"].notna()
        absent_count = len(subject_df) - int(subject_df["参考状态"].sum())
        process_log.append(f"{subject}科目缺考人数：{absent_count}")
        parsed_data[subject] = subject_df

    # 总分数据
    total_df = df[BASIC_COLUMNS].copy()
    if "总分1" in df.columns:
        total_df["成绩"] = df["总分1"]
        process_log.append("使用预计算的总分1数据")
    elif "总分" in df.columns:
        total_df["成绩"] = df["总分"]
        process_log.append("使用预计算的总分数据")
    elif subject_columns:
        total_df["成绩"] = df[subject_columns].apply(pd.to_numeric, errors="coerce").sum(
            axis=1, skipna=True)
        process_log.append("数据中没有总分列，已按科目求和计算总分")

    total_df["成绩"] = pd.to_numeric(total_df["成绩"], errors="coerce")
    total_df["参考状态"] = total_df["成绩"].notna()
    parsed_data["总分"] = total_df
    return parsed_data, process_log


def parse_excel(path: str, exam_count: int = 0,
                subject_mode_overrides: str = "") -> Tuple[Dict, List[str], int, Dict[str, Dict]]:
    """
    解析 Excel 文件，返回 (数据字典, 处理日志, 考试场次数, 科目口径字典)。

    exam_count：0 = 自动识别（按列名判断场次，必要时进一步判定是否混合口径）；
                 1 = 单次考试；2 = 两场考试对比；3 = 强制混合口径（按科目分流）。
    subject_mode_overrides：手工覆盖个别科目口径，形如 "物理:1,化学:2"（仅混合口径生效）。
    """
    process_log = [f"读取文件：{path}"]
    excel_file = pd.ExcelFile(path)

    if "成绩表" not in excel_file.sheet_names:
        raise ValueError("Excel 文件中缺少「成绩表」工作表")

    df = excel_file.parse(sheet_name="成绩表")
    process_log.append(f"成功解析成绩表，共读取{len(df)}行数据")

    detected = detect_exam_count(df)
    mode = detected if exam_count == 0 else exam_count
    subject_modes: Dict[str, Dict] = {}

    if exam_count == 0 and detected == 2:
        # 自动识别：逐科目判定口径，若既有有基准又有无基准的科目 → 混合口径
        probe_modes, _probe_log = detect_subject_modes(df)
        distinct = {i["mode"] for i in probe_modes.values()}
        if len(distinct) > 1:
            mode = MIXED_MODE
            process_log.append("检测到部分科目有基准成绩、部分科目无基准成绩 → 混合口径")

    if exam_count != 0 and exam_count != detected and exam_count != MIXED_MODE:
        process_log.append(f"注意：指定为 {exam_count} 场考试，按列名识别为 {detected} 场，以指定值为准")

    mode_text = {1: "单次考试（首次考试，四项指标）",
                 2: "两场考试对比（六项指标）",
                 3: "混合口径（按科目分流：双场六项指标 / 单场四项指标）"}[mode]
    process_log.append(f"分析模式：{mode_text}")

    if mode == 1:
        parsed_data, validation_log = validate_and_transform_data_single(df)
        subject_modes = {k: {"mode": 1, "base_col": None, "cur_col": None}
                         for k in parsed_data if k != "教师任课表"}
    elif mode == 2:
        parsed_data, validation_log = validate_and_transform_data(df)
        subject_modes = {k: {"mode": 2, "base_col": None, "cur_col": None}
                         for k in parsed_data if k != "教师任课表"}
    else:
        parsed_data, subject_modes, validation_log = validate_and_transform_data_mixed(
            df, subject_mode_overrides)
    process_log.extend(validation_log)

    if "教师任课表" in excel_file.sheet_names:
        teacher_df = excel_file.parse(sheet_name="教师任课表")
        if "班级" in teacher_df.columns:
            teacher_df["班级"] = teacher_df["班级"].apply(normalize_class)
        process_log.append(f"成功解析教师任课表，共读取{len(teacher_df)}行数据")
        parsed_data["教师任课表"] = teacher_df
    else:
        process_log.append("警告：未找到教师任课表，将不生成教师得分表")
        parsed_data["教师任课表"] = None

    return parsed_data, process_log, mode, subject_modes


# ---------------- 指标计算 ----------------
def calc_avg_indicator(df: pd.DataFrame, subject: str) -> Tuple[Dict, List[str]]:
    """
    计算「均分」指标：按班级范围统计每一科的班级均分。

    指标值 = **第二场班级均分**（绝对水平，不做差）。定位是「结果评价」：
    与排名正负 / T分差这类「进步幅度」指标互补，用于平衡只看进步幅度
    带来的偏差——例如基础薄弱的班级大幅进步，与基础扎实的班级小幅进步，
    进步幅度可能相同，但结果水平（均分）能区分两者。

    返回的「均分表格」既作为独立指标，也复用于结果 Excel 的「所有科目均分汇总」。
    """
    process_log = []
    class_avg_df = build_class_avg_df(df)

    if class_avg_df.empty:
        return {"均分表格": []}, [f"科目{subject}：无有效成绩记录"]

    class_avg_df = class_avg_df.copy()
    # 结果评价口径：直接取第二场班级均分，不做两场之差
    class_avg_df["均分"] = class_avg_df["第二场均分"]
    class_avg_df["均分排名"] = class_avg_df["均分"].rank(ascending=False, method="min").astype(int)
    class_avg_df["学科"] = subject

    avg_table_data = class_avg_df[
        ["班级", "学科", "第一场均分", "第二场均分", "均分", "均分排名"]
    ].to_dict(orient="records")

    process_log.append(
        f"科目{subject}：班级均分计算完成（结果评价口径，取第二场均分），共{len(avg_table_data)}个班级")
    return {"均分表格": avg_table_data}, process_log


def calc_t_score_diff(df: pd.DataFrame, subject: str) -> Tuple[Dict, List[str]]:
    """
    计算正负T分差。T分公式：T = 50 + 10*(X - μ)/σ（X 为班级均分，
    μ/σ 为全体班级均分的均值与标准差）；T分差 = 第二场T分 - 第一场T分。
    """
    process_log = []
    reference_df = df[(df["参考状态"]) & (df["第一场成绩"].notna()) & (df["第二场成绩"].notna())].copy()

    if len(reference_df) == 0:
        return {"T分差表格": []}, [f"科目{subject}：无有效成绩记录"]

    process_log.append(f"科目{subject}：计算T分差，有效数据{len(reference_df)}条")

    class_avg_df = build_class_avg_df(df)

    mu1 = class_avg_df["第一场均分"].mean()
    sigma1 = class_avg_df["第一场均分"].std()
    mu2 = class_avg_df["第二场均分"].mean()
    sigma2 = class_avg_df["第二场均分"].std()
    process_log.append(f"科目{subject}：第一场班级均分 μ={mu1:.2f}，σ={sigma1:.2f}")
    process_log.append(f"科目{subject}：第二场班级均分 μ={mu2:.2f}，σ={sigma2:.2f}")

    if sigma1 > 0:
        class_avg_df["T分1"] = T_SCORE_AVG + T_SCORE_STD * (class_avg_df["第一场均分"] - mu1) / sigma1
    else:
        class_avg_df["T分1"] = T_SCORE_AVG
    if sigma2 > 0:
        class_avg_df["T分2"] = T_SCORE_AVG + T_SCORE_STD * (class_avg_df["第二场均分"] - mu2) / sigma2
    else:
        class_avg_df["T分2"] = T_SCORE_AVG

    class_avg_df["T分差"] = class_avg_df["T分2"] - class_avg_df["T分1"]
    class_avg_df["T分差排名"] = class_avg_df["T分差"].rank(ascending=False, method="min")
    class_avg_df["学科"] = subject

    class_avg_df["T分1"] = class_avg_df["T分1"].round(2)
    class_avg_df["T分2"] = class_avg_df["T分2"].round(2)
    class_avg_df["T分差"] = class_avg_df["T分差"].round(2)
    class_avg_df["T分差排名"] = class_avg_df["T分差排名"].astype(int)

    class_avg_df = class_avg_df.sort_values(by="班级", key=lambda x: x.apply(natural_sort_key))
    t_diff_table_data = class_avg_df[
        ["班级", "学科", "T分1", "T分2", "T分差", "T分差排名"]
    ].to_dict(orient="records")

    process_log.append(f"科目{subject}：T分差表格生成完成，共{len(t_diff_table_data)}个班级")
    return {"T分差表格": t_diff_table_data}, process_log


def calc_rank_positive_negative(df: pd.DataFrame, subject: str) -> Tuple[Dict, List[str]]:
    """
    计算排名正负。线性排名公式：rank = 1 + (N-1)*((Xmax - x)/(Xmax - Xmin))，
    x 为班级均分；排名正负 = 线性排名1 - 线性排名2（正值表示进步）。
    """
    process_log = []
    reference_df = df[(df["参考状态"]) & (df["第一场成绩"].notna()) & (df["第二场成绩"].notna())].copy()

    if len(reference_df) == 0:
        return {"均分表格": [], "排名正负表格": []}, [f"科目{subject}：无有效成绩记录"]

    process_log.append(f"科目{subject}：计算排名正负，有效数据{len(reference_df)}条")

    class_avg_df = build_class_avg_df(df)

    class_avg_df = class_avg_df.sort_values(by="班级", key=lambda x: x.apply(natural_sort_key))
    avg_table_data = class_avg_df.to_dict(orient="records")

    N = len(class_avg_df)
    Xmax1, Xmin1 = class_avg_df["第一场均分"].max(), class_avg_df["第一场均分"].min()
    if Xmax1 == Xmin1:
        class_avg_df["线性排名1"] = 1.0
    else:
        class_avg_df["线性排名1"] = 1 + (N - 1) * ((Xmax1 - class_avg_df["第一场均分"]) / (Xmax1 - Xmin1))

    Xmax2, Xmin2 = class_avg_df["第二场均分"].max(), class_avg_df["第二场均分"].min()
    if Xmax2 == Xmin2:
        class_avg_df["线性排名2"] = 1.0
    else:
        class_avg_df["线性排名2"] = 1 + (N - 1) * ((Xmax2 - class_avg_df["第二场均分"]) / (Xmax2 - Xmin2))

    class_avg_df["正负"] = class_avg_df["线性排名1"] - class_avg_df["线性排名2"]
    class_avg_df["排名（正负指标的排名）"] = class_avg_df["正负"].rank(ascending=False, method="min")

    class_avg_df["线性排名1"] = class_avg_df["线性排名1"].round(2)
    class_avg_df["线性排名2"] = class_avg_df["线性排名2"].round(2)
    class_avg_df["正负"] = class_avg_df["正负"].round(2)
    class_avg_df["排名（正负指标的排名）"] = class_avg_df["排名（正负指标的排名）"].astype(int)
    class_avg_df["学科"] = subject

    rank_table_data = class_avg_df[
        ["班级", "学科", "线性排名1", "线性排名2", "正负", "排名（正负指标的排名）"]
    ].to_dict(orient="records")

    process_log.append(f"科目{subject}：排名正负表格生成完成，共{len(rank_table_data)}个班级")
    return {"均分表格": avg_table_data, "排名正负表格": rank_table_data}, process_log


def _calc_percent_count_table(df: pd.DataFrame, subject: str, ratio: float,
                              label: str) -> Tuple[Dict, List[str]]:
    """
    通用算法：按全年级排名前 ratio% 的成绩阈值，统计各班两场达标人数及差值。
    label 为 '优秀' 或 '出库'。
    """
    process_log = []
    reference_df = df[(df["参考状态"]) & (df["第一场成绩"].notna()) & (df["第二场成绩"].notna())].copy()

    if len(reference_df) == 0:
        return {f"{label}人数表格": []}, [f"科目{subject}：无有效成绩记录"]

    total_students = len(reference_df)
    top_percent_count = max(1, int(total_students * (ratio / 100)))
    process_log.append(
        f"科目{subject}：计算基于排名百分比的{label}人数，总人数={total_students}，"
        f"比例={ratio}%，阈值名次={top_percent_count}")

    class_avg_df = reference_df.groupby("集团班级").size().reset_index(name="人数")

    first_scores = reference_df["第一场成绩"].sort_values(ascending=False).reset_index(drop=True)
    first_threshold = first_scores.iloc[min(top_percent_count, total_students) - 1]
    second_scores = reference_df["第二场成绩"].sort_values(ascending=False).reset_index(drop=True)
    second_threshold = second_scores.iloc[min(top_percent_count, total_students) - 1]
    process_log.append(f"科目{subject}：第一场{label}成绩阈值={first_threshold}，第二场{label}成绩阈值={second_threshold}")

    rows = []
    for _, class_row in class_avg_df.iterrows():
        class_name = class_row["集团班级"]
        class_students = reference_df[reference_df["集团班级"] == class_name]
        count_1 = class_students[class_students["第一场成绩"] >= first_threshold].shape[0]
        count_2 = class_students[class_students["第二场成绩"] >= second_threshold].shape[0]
        rows.append({
            "班级": class_name,
            "学科": subject,
            f"{label}人数1": int(count_1),
            f"{label}人数2": int(count_2),
            f"{label}人数差": int(count_2 - count_1),
        })

    result_df = pd.DataFrame(rows)
    result_df = result_df.sort_values(by="班级", key=lambda x: x.apply(natural_sort_key))
    process_log.append(f"科目{subject}：{label}人数表格生成完成，共{len(rows)}个班级")
    return {f"{label}人数表格": result_df.to_dict(orient="records")}, process_log


def calc_new_excellent_count(df, subject, excellent_ratio):
    """基于排名百分比的优秀人数（默认全年级前 20%）"""
    return _calc_percent_count_table(df, subject, excellent_ratio, "优秀")


def calc_qualified_count(df, subject, qualified_ratio):
    """基于排名百分比的出库人数（默认全年级前 75%）"""
    return _calc_percent_count_table(df, subject, qualified_ratio, "出库")


def calculate_reference_rate(df: pd.DataFrame, subject: str) -> Tuple[List[Dict], List[str]]:
    """
    计算每个班级的参考率：两场都有分数的视为参考。
    参考率 = 参考人数 / 总人数（0~1，保留4位小数）。
    """
    process_log = []
    class_total_df = df.groupby("集团班级").size().reset_index(name="总人数")
    reference_df = df[(df["第一场成绩"].notna()) & (df["第二场成绩"].notna())]
    class_reference_df = reference_df.groupby("集团班级").size().reset_index(name="参考人数")

    merged_df = pd.merge(class_total_df, class_reference_df, on="集团班级", how="left")
    merged_df["参考人数"] = merged_df["参考人数"].fillna(0)
    merged_df["参考率"] = (merged_df["参考人数"] / merged_df["总人数"]).round(4)
    merged_df = merged_df.rename(columns={"集团班级": "班级"})

    reference_rate_data = merged_df[["班级", "总人数", "参考人数", "参考率"]].to_dict(orient="records")
    process_log.append(f"科目{subject}：班级参考率计算完成，共{len(reference_rate_data)}个班级")
    return reference_rate_data, process_log


def calc_reference_rate(df: pd.DataFrame, subject: str) -> Tuple[float, List[str]]:
    """计算整体参考率（参考人数 / 总人数 × 100%，保留1位小数）"""
    total_students = len(df)
    reference_count = df[df["参考状态"]].shape[0]
    reference_rate = (reference_count / total_students * 100) if total_students > 0 else 0.0
    return round(reference_rate, 1), [
        f"科目{subject}：整体参考率 = {reference_count}/{total_students} × 100% = {reference_rate:.1f}%"]


# ---------------- 单次考试（首次考试）指标计算 ----------------
def build_single_class_stats(df: pd.DataFrame, subject: str) -> pd.DataFrame:
    """
    单次考试的班级均分（含横向位次参考）。

    - 均分：该班该科「有成绩学生」的平均分（缺考不计入平均）
    - 排名：班级均分在同学科全部班级中的名次（1 = 最高，并列同名次）
      —— 排名本就是班级均分的横向位次，与均分属同一项指标，不单独计分，仅作参考字段
    - 线性排名：rank = 1 + (N-1)*(Xmax-x)/(Xmax-Xmin)，与既有「排名正负」口径一致
    """
    valid = df[(df["参考状态"]) & (df["成绩"].notna())].copy()

    if valid.empty:
        return pd.DataFrame(columns=["班级", "学科", "均分", "排名", "线性排名",
                                     "总人数", "参考人数"])

    grouped = valid.groupby("集团班级").agg(
        均分=("成绩", "mean"), 参考人数=("成绩", "size")).reset_index()
    grouped = grouped.rename(columns={"集团班级": "班级"})

    roster = df.groupby("集团班级").size().reset_index(name="总人数")
    roster = roster.rename(columns={"集团班级": "班级"})
    grouped = pd.merge(grouped, roster, on="班级", how="left")
    grouped["均分"] = grouped["均分"].round(2)

    n_classes = len(grouped)
    x_max, x_min = grouped["均分"].max(), grouped["均分"].min()
    if x_max == x_min:
        grouped["线性排名"] = 1.0
    else:
        grouped["线性排名"] = 1 + (n_classes - 1) * (
            (x_max - grouped["均分"]) / (x_max - x_min))

    grouped["排名"] = grouped["均分"].rank(ascending=False, method="min").astype(int)
    grouped["线性排名"] = grouped["线性排名"].round(2)
    grouped["学科"] = subject
    return grouped.sort_values(by="班级", key=lambda x: x.apply(natural_sort_key))


def calc_single_avg_rank(df: pd.DataFrame, subject: str) -> Tuple[Dict, List[str]]:
    """单次考试：班级均分（附排名参考字段）"""
    process_log = []
    stats = build_single_class_stats(df, subject)

    if stats.empty:
        return {"均分排名表格": []}, [f"科目{subject}：无有效成绩记录"]

    table = stats[["班级", "学科", "均分", "排名", "线性排名",
                   "总人数", "参考人数"]].to_dict(orient="records")
    process_log.append(f"科目{subject}：班级均分计算完成，共{len(table)}个班级")
    return {"均分排名表格": table}, process_log


def calc_single_rate(df: pd.DataFrame, subject: str, ratio: float,
                     label: str) -> Tuple[Dict, List[str]]:
    """
    单次考试的达标率（优秀率 / 合格率）。

    达标线：全年级该科「有成绩学生」按成绩降序排列，取前 ratio% 名次处的分数为达标线。
    班级达标率 = 该班达标人数 / 该班参考人数（有该科成绩的学生数）。
    分母用参考人数而非在册人数，参考率单独成一项指标，两者互补不重复。
    """
    process_log = []
    valid = df[(df["参考状态"]) & (df["成绩"].notna())].copy()

    if valid.empty:
        return {f"{label}率表格": []}, [f"科目{subject}：无有效成绩记录"]

    total_students = len(valid)
    top_count = max(1, int(total_students * (ratio / 100)))
    threshold = valid["成绩"].sort_values(ascending=False).reset_index(drop=True).iloc[
        min(top_count, total_students) - 1]
    process_log.append(
        f"科目{subject}：{label}线按全年级前 {ratio}% 名次确定，"
        f"有效人数={total_students}，阈值名次={top_count}，成绩阈值={threshold}")

    rows = []
    for class_name, class_df in valid.groupby("集团班级"):
        ref_count = len(class_df)
        hit_count = int((class_df["成绩"] >= threshold).sum())
        rows.append({
            "班级": class_name,
            "学科": subject,
            f"{label}人数": hit_count,
            f"{label}率": round(hit_count / ref_count, 4) if ref_count else 0.0,
        })

    result_df = pd.DataFrame(rows)
    result_df = result_df.sort_values(by="班级", key=lambda x: x.apply(natural_sort_key))
    process_log.append(f"科目{subject}：{label}率计算完成，共{len(result_df)}个班级")
    return {f"{label}率表格": result_df.to_dict(orient="records")}, process_log


def calculate_single_reference_rate(df: pd.DataFrame, subject: str) -> Tuple[List[Dict], List[str]]:
    """单次考试：班级参考率（参考人数 / 在册人数，口径与多次考试完全一致）"""
    process_log = []
    class_total_df = df.groupby("集团班级").size().reset_index(name="总人数")
    class_reference_df = df[(df["参考状态"]) & (df["成绩"].notna())].groupby(
        "集团班级").size().reset_index(name="参考人数")

    merged_df = pd.merge(class_total_df, class_reference_df, on="集团班级", how="left")
    merged_df["参考人数"] = merged_df["参考人数"].fillna(0).astype(int)
    merged_df["参考率"] = (merged_df["参考人数"] / merged_df["总人数"]).round(4)
    merged_df = merged_df.rename(columns={"集团班级": "班级"})

    reference_rate_data = merged_df[["班级", "总人数", "参考人数", "参考率"]].to_dict(orient="records")
    process_log.append(f"科目{subject}：班级参考率计算完成，共{len(reference_rate_data)}个班级")
    return reference_rate_data, process_log


def calculate_all_indicators_single(parsed_data: Dict[str, pd.DataFrame],
                                    excellent_ratio: float = 20.0,
                                    qualified_ratio: float = 80.0,
                                    avg_weight: float = 60.0,
                                    excellent_weight: float = 15.0,
                                    qualified_weight: float = 15.0,
                                    reference_weight: float = 10.0) -> Tuple[Dict, List[str]]:
    """
    单次考试（首次考试）指标计算。

    四项指标：均分、优秀率、合格率、参考率。
    「均分与排名」按同一项指标处理（排名即班级均分的横向位次），不单列计分。
    与两场考试的六项指标互斥启用，由 --exam-count 决定走哪条分支。
    """
    process_log = []
    result = {}

    subjects = [key for key in parsed_data.keys() if key != "教师任课表"]
    for subject in subjects:
        process_log.append(f"=== 开始计算科目：{subject}（单次考试）===")
        subject_df = parsed_data[subject]

        avg_rank, a_log = calc_single_avg_rank(subject_df, subject)
        excellent_rate, e_log = calc_single_rate(subject_df, subject, excellent_ratio, "优秀")
        qualified_rate, q_log = calc_single_rate(subject_df, subject, qualified_ratio, "合格")
        class_reference_rate, r_log = calculate_single_reference_rate(subject_df, subject)

        process_log.extend(a_log + e_log + q_log + r_log)

        result[subject] = {
            "均分排名": avg_rank,
            "优秀率": excellent_rate,
            "合格率": qualified_rate,
            "班级参考率": class_reference_rate,
        }

    process_log.append("=== 所有科目计算完成 ===")

    teacher_scores, teacher_log = calculate_teacher_scores_single(
        parsed_data, result, avg_weight,
        excellent_weight, qualified_weight, reference_weight)
    result["教师得分表"] = teacher_scores
    process_log.extend(teacher_log)

    return result, process_log


def calculate_teacher_scores_single(parsed_data: Dict[str, pd.DataFrame], calc_results: Dict,
                                    avg_weight: float = 60.0,
                                    excellent_weight: float = 15.0,
                                    qualified_weight: float = 15.0,
                                    reference_weight: float = 10.0) -> Tuple[Dict, List[str]]:
    """
    单次考试的教师得分表。

    指标：均分、优秀率、合格率、参考率，四项 min-max 归一到 0~100 后按权重加权。
    「均分与排名」按同一项指标处理——排名即班级均分在同学科中的横向位次，不再单列计分。
    归一化在「同一科目组内」进行（班主任视为独立一组），与两场考试口径一致。
    多班任教时，四项指标一律取均值（比率类指标不可求和）。
    """
    process_log = []

    if parsed_data.get("教师任课表") is None:
        return {"教师得分表": []}, ["未找到教师任课表，无法生成教师得分表"]

    teacher_schedule_df = parsed_data["教师任课表"].copy()
    required_columns = ["班级", "教师姓名", "身份证号", "科目"]
    for col in required_columns:
        if col not in teacher_schedule_df.columns:
            return {"教师得分表": []}, [f"教师任课表缺少必要的列：{col}"]

    has_campus = "校区" in teacher_schedule_df.columns
    if not has_campus:
        process_log.append("教师任课表不包含校区信息，将使用空值填充")
        teacher_schedule_df["校区"] = pd.NA

    def merge_subject_indicators(subject: str) -> pd.DataFrame:
        avg_df = pd.DataFrame(calc_results[subject]["均分排名"]["均分排名表格"])
        exc_df = pd.DataFrame(calc_results[subject]["优秀率"]["优秀率表格"])
        qua_df = pd.DataFrame(calc_results[subject]["合格率"]["合格率表格"])
        ref_df = pd.DataFrame(calc_results[subject]["班级参考率"])

        merged_df = avg_df.merge(exc_df, on=["班级", "学科"], how="outer")
        merged_df = merged_df.merge(qua_df, on=["班级", "学科"], how="outer")
        merged_df = merged_df.merge(ref_df, on="班级", how="outer")
        return merged_df

    teacher_scores = []

    # 步骤1：学科教师
    process_log.append("=== 开始处理学科教师得分（单次考试）===")
    subjects = [key for key in calc_results.keys() if key != "教师得分表"]
    for subject in subjects:
        try:
            merged_df = merge_subject_indicators(subject)
        except (KeyError, IndexError) as e:
            process_log.append(f"科目{subject}缺少必要的指标数据，跳过处理：{str(e)}")
            continue

        teachers = teacher_schedule_df[
            (teacher_schedule_df["科目"] == subject) & (teacher_schedule_df["科目"] != "班主任")]
        for _, teacher_row in teachers.iterrows():
            class_name = teacher_row["班级"]
            class_data = merged_df[(merged_df["班级"] == class_name) & (merged_df["学科"] == subject)]
            if not class_data.empty:
                teacher_scores.append({
                    "身份证号": teacher_row["身份证号"],
                    "姓名": teacher_row["教师姓名"],
                    "科目": subject,
                    "校区": teacher_row.get("校区", pd.NA),
                    "均分": class_data.iloc[0]["均分"],
                    "优秀率": class_data.iloc[0]["优秀率"],
                    "合格率": class_data.iloc[0]["合格率"],
                    "参考率": class_data.iloc[0]["参考率"],
                    "班级": class_name,
                })

    # 步骤2：班主任（使用总分指标）
    process_log.append("=== 开始处理班主任得分（单次考试）===")
    if "总分" in calc_results:
        try:
            merged_total_df = merge_subject_indicators("总分")
        except (KeyError, IndexError) as e:
            process_log.append(f"总分科目缺少必要的指标数据，跳过班主任处理：{str(e)}")
            merged_total_df = None

        if merged_total_df is not None:
            head_teachers = teacher_schedule_df[
                teacher_schedule_df["科目"] == "班主任"].drop_duplicates(subset=["教师姓名", "身份证号"])
            for _, teacher_row in head_teachers.iterrows():
                teacher_name = teacher_row["教师姓名"]
                id_card = teacher_row["身份证号"]
                head_classes = teacher_schedule_df[
                    (teacher_schedule_df["教师姓名"] == teacher_name) &
                    (teacher_schedule_df["身份证号"] == id_card) &
                    (teacher_schedule_df["科目"] == "班主任")]["班级"]

                class_indicators = []
                for class_name in head_classes:
                    class_data = merged_total_df[
                        (merged_total_df["班级"] == class_name) & (merged_total_df["学科"] == "总分")]
                    if not class_data.empty:
                        class_indicators.append({
                            "均分": class_data.iloc[0]["均分"],
                            "优秀率": class_data.iloc[0]["优秀率"],
                            "合格率": class_data.iloc[0]["合格率"],
                            "参考率": class_data.iloc[0]["参考率"],
                        })

                if class_indicators:
                    count = len(class_indicators)
                    teacher_scores.append({
                        "身份证号": id_card,
                        "姓名": teacher_name,
                        "科目": "班主任",
                        "校区": teacher_row.get("校区", pd.NA),
                        "均分": round(sum(i["均分"] for i in class_indicators) / count, 2),
                        "优秀率": round(sum(i["优秀率"] for i in class_indicators) / count, 4),
                        "合格率": round(sum(i["合格率"] for i in class_indicators) / count, 4),
                        "参考率": round(sum(i["参考率"] for i in class_indicators) / count, 4),
                    })

    # 步骤3：聚合与总分计算
    process_log.append("=== 开始聚合教师得分数据（单次考试）===")
    teacher_scores_df = pd.DataFrame(teacher_scores)

    if teacher_scores_df.empty:
        return {"教师得分表": []}, ["没有找到教师得分数据"]

    aggregated_df = teacher_scores_df.groupby(
        ["身份证号", "姓名", "科目", "校区"], dropna=False).agg(
        均分=pd.NamedAgg(column="均分", aggfunc="mean"),
        优秀率=pd.NamedAgg(column="优秀率", aggfunc="mean"),
        合格率=pd.NamedAgg(column="合格率", aggfunc="mean"),
        参考率=pd.NamedAgg(column="参考率", aggfunc="mean"),
    ).reset_index()

    aggregated_df["均分"] = aggregated_df["均分"].round(2)
    aggregated_df["优秀率"] = aggregated_df["优秀率"].round(4)
    aggregated_df["合格率"] = aggregated_df["合格率"].round(4)
    aggregated_df["参考率"] = aggregated_df["参考率"].round(4)

    weights = {
        "均分": avg_weight,
        "优秀率": excellent_weight,
        "合格率": qualified_weight,
        "参考率": reference_weight,
    }
    total_score = pd.Series(0.0, index=aggregated_df.index)
    for _, group_idx in aggregated_df.groupby("科目", dropna=False).groups.items():
        for col, weight in weights.items():
            series = pd.to_numeric(aggregated_df.loc[group_idx, col], errors="coerce")
            col_min, col_max = series.min(), series.max()
            if pd.notna(col_min) and col_max != col_min:
                transformed = (series - col_min) / (col_max - col_min) * 100
            else:
                transformed = pd.Series(50.0, index=group_idx)
            if weight < 0:
                transformed = 100 - transformed
            total_score.loc[group_idx] += transformed * (abs(weight) / 100)

    aggregated_df["总分"] = total_score.round(2)

    teacher_scores_table = aggregated_df[[
        "身份证号", "姓名", "科目", "校区", "均分",
        "优秀率", "合格率", "参考率", "总分"
    ]].to_dict(orient="records")

    process_log.append(f"教师得分表生成完成（单次考试），共{len(teacher_scores_table)}条记录")
    return {"教师得分表": teacher_scores_table}, process_log


def calculate_all_indicators(parsed_data: Dict[str, pd.DataFrame],
                             excellent_ratio: float = 20.0,
                             qualified_ratio: float = 75.0,
                             rank_weight: float = 20.0,
                             t_diff_weight: float = 20.0,
                             avg_weight: float = 20.0,
                             excellent_weight: float = 15.0,
                             qualified_weight: float = 15.0,
                             reference_weight: float = 10.0) -> Tuple[Dict, List[str]]:
    """计算所有科目指标，返回 (结果字典, 计算过程日志)"""
    process_log = []
    result = {}

    subjects = [key for key in parsed_data.keys() if key != "教师任课表"]
    for subject in subjects:
        process_log.append(f"=== 开始计算科目：{subject} ===")
        subject_df = parsed_data[subject]

        t_score_diff, t_log = calc_t_score_diff(subject_df, subject)
        rank_pos_neg, rp_log = calc_rank_positive_negative(subject_df, subject)
        avg_indicator, a_log = calc_avg_indicator(subject_df, subject)
        excellent_count, e_log = calc_new_excellent_count(subject_df, subject, excellent_ratio)
        qualified_count, q_log = calc_qualified_count(subject_df, subject, qualified_ratio)
        class_reference_rate, r_log = calculate_reference_rate(subject_df, subject)
        overall_reference_rate, o_log = calc_reference_rate(subject_df, subject)

        process_log.extend(t_log + rp_log + a_log + e_log + q_log + r_log + o_log)

        result[subject] = {
            "T分差": t_score_diff,
            "排名正负": rank_pos_neg,
            "均分": avg_indicator,
            "优秀人数": excellent_count,
            "出库人数": qualified_count,
            "班级参考率": class_reference_rate,
            "整体参考率": overall_reference_rate,
        }

    process_log.append("=== 所有科目计算完成 ===")

    teacher_scores, teacher_log = calculate_teacher_scores(
        parsed_data, result, rank_weight, t_diff_weight, avg_weight,
        excellent_weight, qualified_weight, reference_weight)
    result["教师得分表"] = teacher_scores
    process_log.extend(teacher_log)

    return result, process_log


def calculate_teacher_scores(parsed_data: Dict[str, pd.DataFrame], calc_results: Dict,
                             rank_weight: float = 20.0, t_diff_weight: float = 20.0,
                             avg_weight: float = 20.0,
                             excellent_weight: float = 15.0, qualified_weight: float = 15.0,
                             reference_weight: float = 10.0) -> Tuple[Dict, List[str]]:
    """
    计算教师得分表：
    - 学科教师：按任课班级×科目取班级指标；
    - 班主任：使用「总分」科目指标，多班时均值/求和聚合；
    - 同教师同科目多班再聚合后，六项指标做 min-max 线性变换到 0~100，
      按权重加权求和得到总分。

    指标性质：「排名正负 / T分差」衡量进步幅度，「均分」衡量结果水平（第二场班级均分），
    两类指标互补，避免只看进步幅度而忽略班级原有基础。

    注：「参考率」权重为正时按「越高越好」直接归一；权重为负时按「越低越好」反向归一
    （即该指标以低为优）。其余五项指标均为正向（越高越好）。
    """
    process_log = []

    if parsed_data.get("教师任课表") is None:
        return {"教师得分表": []}, ["未找到教师任课表，无法生成教师得分表"]

    teacher_schedule_df = parsed_data["教师任课表"].copy()
    required_columns = ["班级", "教师姓名", "身份证号", "科目"]
    for col in required_columns:
        if col not in teacher_schedule_df.columns:
            return {"教师得分表": []}, [f"教师任课表缺少必要的列：{col}"]

    has_campus = "校区" in teacher_schedule_df.columns
    if not has_campus:
        process_log.append("教师任课表不包含校区信息，将使用空值填充")
        teacher_schedule_df["校区"] = pd.NA

    def merge_subject_indicators(subject: str) -> pd.DataFrame:
        t_diff_df = pd.DataFrame(calc_results[subject]["T分差"]["T分差表格"])
        rank_df = pd.DataFrame(calc_results[subject]["排名正负"]["排名正负表格"])
        avg_df = pd.DataFrame(calc_results[subject]["均分"]["均分表格"])
        excellent_df = pd.DataFrame(calc_results[subject]["优秀人数"]["优秀人数表格"])
        qualified_df = pd.DataFrame(calc_results[subject]["出库人数"]["出库人数表格"])
        reference_rate_df = pd.DataFrame(calc_results[subject]["班级参考率"])

        merged_df = t_diff_df.merge(rank_df, on=["班级", "学科"], how="outer")
        merged_df = merged_df.merge(
            avg_df[["班级", "学科", "均分"]], on=["班级", "学科"], how="outer")
        merged_df = merged_df.merge(excellent_df, on=["班级", "学科"], how="outer")
        merged_df = merged_df.merge(qualified_df, on=["班级", "学科"], how="outer")
        merged_df = merged_df.merge(reference_rate_df, on=["班级"], how="outer")
        return merged_df

    teacher_scores = []

    # 步骤1：学科教师
    process_log.append("=== 开始处理学科教师得分 ===")
    subjects = [key for key in calc_results.keys() if key != "教师得分表"]
    for subject in subjects:
        try:
            merged_df = merge_subject_indicators(subject)
        except (KeyError, IndexError) as e:
            process_log.append(f"科目{subject}缺少必要的指标数据，跳过处理：{str(e)}")
            continue

        teachers = teacher_schedule_df[
            (teacher_schedule_df["科目"] == subject) & (teacher_schedule_df["科目"] != "班主任")]
        for _, teacher_row in teachers.iterrows():
            class_name = teacher_row["班级"]
            class_data = merged_df[(merged_df["班级"] == class_name) & (merged_df["学科"] == subject)]
            if not class_data.empty:
                teacher_scores.append({
                    "身份证号": teacher_row["身份证号"],
                    "姓名": teacher_row["教师姓名"],
                    "科目": subject,
                    "校区": teacher_row.get("校区", pd.NA),
                    "排名正负": class_data.iloc[0]["正负"],
                    "T分差": class_data.iloc[0]["T分差"],
                    "均分": class_data.iloc[0]["均分"],
                    "优秀人数": class_data.iloc[0]["优秀人数差"],
                    "出库人数": class_data.iloc[0]["出库人数差"],
                    "参考率": class_data.iloc[0]["参考率"],
                    "班级": class_name,
                })

    # 步骤2：班主任（使用总分指标）
    process_log.append("=== 开始处理班主任得分 ===")
    if "总分" in calc_results:
        try:
            merged_total_df = merge_subject_indicators("总分")
        except (KeyError, IndexError) as e:
            process_log.append(f"总分科目缺少必要的指标数据，跳过班主任处理：{str(e)}")
            merged_total_df = None

        if merged_total_df is not None:
            head_teachers = teacher_schedule_df[
                teacher_schedule_df["科目"] == "班主任"].drop_duplicates(subset=["教师姓名", "身份证号"])
            for _, teacher_row in head_teachers.iterrows():
                teacher_name = teacher_row["教师姓名"]
                id_card = teacher_row["身份证号"]
                head_classes = teacher_schedule_df[
                    (teacher_schedule_df["教师姓名"] == teacher_name) &
                    (teacher_schedule_df["身份证号"] == id_card) &
                    (teacher_schedule_df["科目"] == "班主任")]["班级"]

                class_indicators = []
                for class_name in head_classes:
                    class_data = merged_total_df[
                        (merged_total_df["班级"] == class_name) & (merged_total_df["学科"] == "总分")]
                    if not class_data.empty:
                        class_indicators.append({
                            "排名正负": class_data.iloc[0]["正负"],
                            "T分差": class_data.iloc[0]["T分差"],
                            "均分": class_data.iloc[0]["均分"],
                            "优秀人数": class_data.iloc[0]["优秀人数差"],
                            "出库人数": class_data.iloc[0]["出库人数差"],
                            "参考率": class_data.iloc[0]["参考率"],
                        })

                if class_indicators:
                    teacher_scores.append({
                        "身份证号": id_card,
                        "姓名": teacher_name,
                        "科目": "班主任",
                        "校区": teacher_row.get("校区", pd.NA),
                        "排名正负": round(sum(i["排名正负"] for i in class_indicators) / len(class_indicators), 2),
                        "T分差": round(sum(i["T分差"] for i in class_indicators) / len(class_indicators), 2),
                        "均分": round(sum(i["均分"] for i in class_indicators) / len(class_indicators), 2),
                        "优秀人数": sum(i["优秀人数"] for i in class_indicators),
                        "出库人数": sum(i["出库人数"] for i in class_indicators),
                        "参考率": round(sum(i["参考率"] for i in class_indicators) / len(class_indicators), 4),
                    })

    # 步骤3：聚合与总分计算
    process_log.append("=== 开始聚合教师得分数据 ===")
    teacher_scores_df = pd.DataFrame(teacher_scores)

    if teacher_scores_df.empty:
        return {"教师得分表": []}, ["没有找到教师得分数据"]

    aggregated_df = teacher_scores_df.groupby(
        ["身份证号", "姓名", "科目", "校区"], dropna=False).agg(
        排名正负=pd.NamedAgg(column="排名正负", aggfunc="mean"),
        T分差=pd.NamedAgg(column="T分差", aggfunc="mean"),
        均分=pd.NamedAgg(column="均分", aggfunc="mean"),
        优秀人数=pd.NamedAgg(column="优秀人数", aggfunc="sum"),
        出库人数=pd.NamedAgg(column="出库人数", aggfunc="sum"),
        参考率=pd.NamedAgg(column="参考率", aggfunc="mean"),
    ).reset_index()

    aggregated_df["排名正负"] = aggregated_df["排名正负"].round(2)
    aggregated_df["T分差"] = aggregated_df["T分差"].round(2)
    aggregated_df["均分"] = aggregated_df["均分"].round(2)
    aggregated_df["参考率"] = aggregated_df["参考率"].round(4)

    # min-max 线性变换到 0~100，所有值相同时取中间值 50。
    # 归一化在「同一科目组内」进行（班主任视为独立一组）：
    # 各科教师人数悬殊（语文 40+ vs 历史 20+），若混在一起归一，
    # 小科目的分布会被大科目主导，总分不再可比、且受本次运行是否凑齐其他科目影响。
    weights = {
        "排名正负": rank_weight,
        "T分差": t_diff_weight,
        "均分": avg_weight,
        "优秀人数": excellent_weight,
        "出库人数": qualified_weight,
        "参考率": reference_weight,
    }
    total_score = pd.Series(0.0, index=aggregated_df.index)
    for _, group_idx in aggregated_df.groupby("科目", dropna=False).groups.items():
        for col, weight in weights.items():
            series = pd.to_numeric(aggregated_df.loc[group_idx, col], errors="coerce")
            col_min, col_max = series.min(), series.max()
            if pd.notna(col_min) and col_max != col_min:
                transformed = (series - col_min) / (col_max - col_min) * 100
            else:
                transformed = pd.Series(50.0, index=group_idx)
            # 权重为负时按「越低越好」反向归一
            if weight < 0:
                transformed = 100 - transformed
            total_score.loc[group_idx] += transformed * (abs(weight) / 100)

    aggregated_df["总分"] = total_score.round(2)

    teacher_scores_table = aggregated_df[[
        "身份证号", "姓名", "科目", "校区", "排名正负", "T分差", "均分",
        "优秀人数", "出库人数", "参考率", "总分"
    ]].to_dict(orient="records")

    process_log.append(f"教师得分表生成完成，共{len(teacher_scores_table)}条记录")
    return {"教师得分表": teacher_scores_table}, process_log


# 混合口径下两套指标各自的权重（键名与教师得分表列名一致，便于统一处理）
DUAL_METRIC_WEIGHTS = {
    "排名正负": 20.0, "T分差": 20.0, "均分": 20.0,
    "优秀人数": 15.0, "出库人数": 15.0, "参考率": 10.0,
}
SINGLE_METRIC_WEIGHTS = {
    "均分": 60.0, "优秀率": 15.0, "合格率": 15.0, "参考率": 10.0,
}
DUAL_METRIC_COLS = ["排名正负", "T分差", "均分", "优秀人数", "出库人数", "参考率"]
SINGLE_METRIC_COLS = ["均分", "优秀率", "合格率", "参考率"]


def calculate_all_indicators_mixed(parsed_data: Dict[str, pd.DataFrame],
                                   subject_modes: Dict[str, Dict],
                                   excellent_ratio: float = 20.0,
                                   qualified_ratio: float = 75.0,
                                   single_qualified_ratio: float = 80.0,
                                   dual_weights: Dict[str, float] = None,
                                   single_weights: Dict[str, float] = None) -> Tuple[Dict, List[str]]:
    """
    混合口径指标计算：逐科目按其口径分派。

    - 有基准成绩的科目 → 两场对比口径（六项指标：排名正负 / T分差 / 均分 / 优秀人数差 / 出库人数差 / 参考率）
    - 无基准成绩的科目（如初二新开物理）→ 首次考试口径（四项指标：均分 / 优秀率 / 合格率 / 参考率）
    - 总分科目按其自身口径（有基准科目求和构成 → 两场对比口径）
    """
    process_log = []
    result = {}
    dual_weights = dict(dual_weights or DUAL_METRIC_WEIGHTS)
    single_weights = dict(single_weights or SINGLE_METRIC_WEIGHTS)

    subjects = [key for key in parsed_data.keys() if key != "教师任课表"]
    for subject in subjects:
        mode = subject_modes.get(subject, {}).get("mode", 2)
        subject_df = parsed_data[subject]

        if mode == 2:
            process_log.append(f"=== 开始计算科目：{subject}（两场对比口径）===")
            t_score_diff, t_log = calc_t_score_diff(subject_df, subject)
            rank_pos_neg, rp_log = calc_rank_positive_negative(subject_df, subject)
            avg_indicator, a_log = calc_avg_indicator(subject_df, subject)
            excellent_count, e_log = calc_new_excellent_count(subject_df, subject, excellent_ratio)
            qualified_count, q_log = calc_qualified_count(subject_df, subject, qualified_ratio)
            class_reference_rate, r_log = calculate_reference_rate(subject_df, subject)
            process_log.extend(t_log + rp_log + a_log + e_log + q_log + r_log)
            result[subject] = {
                "口径": 2,
                "T分差": t_score_diff,
                "排名正负": rank_pos_neg,
                "均分": avg_indicator,
                "优秀人数": excellent_count,
                "出库人数": qualified_count,
                "班级参考率": class_reference_rate,
            }
        else:
            process_log.append(f"=== 开始计算科目：{subject}（首次考试口径）===")
            avg_rank, a_log = calc_single_avg_rank(subject_df, subject)
            excellent_rate, e_log = calc_single_rate(subject_df, subject, excellent_ratio, "优秀")
            qualified_rate, q_log = calc_single_rate(subject_df, subject, single_qualified_ratio, "合格")
            class_reference_rate, r_log = calculate_single_reference_rate(subject_df, subject)
            process_log.extend(a_log + e_log + q_log + r_log)
            result[subject] = {
                "口径": 1,
                "均分排名": avg_rank,
                "优秀率": excellent_rate,
                "合格率": qualified_rate,
                "班级参考率": class_reference_rate,
            }

    process_log.append("=== 所有科目计算完成 ===")

    teacher_scores, teacher_log = calculate_teacher_scores_mixed(
        parsed_data, result, subject_modes, dual_weights, single_weights)
    result["教师得分表"] = teacher_scores
    process_log.extend(teacher_log)

    return result, process_log


def calculate_teacher_scores_mixed(parsed_data: Dict[str, pd.DataFrame], calc_results: Dict,
                                   subject_modes: Dict[str, Dict],
                                   dual_weights: Dict[str, float] = None,
                                   single_weights: Dict[str, float] = None) -> Tuple[Dict, List[str]]:
    """
    混合口径的教师得分表。

    每个科目用自身的口径与权重：
      - 两场对比科目 → 六项指标，权重 20/20/20/15/15/10；多班任教时人数类取合计、其余取均值
      - 首次考试科目 → 四项指标，权重 60/15/15/10；多班任教时四项一律取均值
      - 班主任 → 使用「总分」科目的指标，口径随总分科目而定
    归一化在「同一科目组内」进行（班主任视为独立一组），与两场考试口径一致。
    """
    process_log = []
    dual_weights = dict(dual_weights or DUAL_METRIC_WEIGHTS)
    single_weights = dict(single_weights or SINGLE_METRIC_WEIGHTS)

    if parsed_data.get("教师任课表") is None:
        return {"教师得分表": []}, ["未找到教师任课表，无法生成教师得分表"]

    teacher_schedule_df = parsed_data["教师任课表"].copy()
    required_columns = ["班级", "教师姓名", "身份证号", "科目"]
    for col in required_columns:
        if col not in teacher_schedule_df.columns:
            return {"教师得分表": []}, [f"教师任课表缺少必要的列：{col}"]

    has_campus = "校区" in teacher_schedule_df.columns
    if not has_campus:
        process_log.append("教师任课表不包含校区信息，将使用空值填充")
        teacher_schedule_df["校区"] = pd.NA

    def merge_subject_indicators(subject: str, mode: int) -> pd.DataFrame:
        if mode == 2:
            t_diff_df = pd.DataFrame(calc_results[subject]["T分差"]["T分差表格"])
            rank_df = pd.DataFrame(calc_results[subject]["排名正负"]["排名正负表格"])
            avg_df = pd.DataFrame(calc_results[subject]["均分"]["均分表格"])
            excellent_df = pd.DataFrame(calc_results[subject]["优秀人数"]["优秀人数表格"])
            qualified_df = pd.DataFrame(calc_results[subject]["出库人数"]["出库人数表格"])
            reference_df = pd.DataFrame(calc_results[subject]["班级参考率"])
            merged_df = t_diff_df.merge(rank_df, on=["班级", "学科"], how="outer")
            merged_df = merged_df.merge(
                avg_df[["班级", "学科", "均分"]], on=["班级", "学科"], how="outer")
            merged_df = merged_df.merge(excellent_df, on=["班级", "学科"], how="outer")
            merged_df = merged_df.merge(qualified_df, on=["班级", "学科"], how="outer")
            merged_df = merged_df.merge(reference_df, on=["班级"], how="outer")
            return merged_df
        avg_df = pd.DataFrame(calc_results[subject]["均分排名"]["均分排名表格"])
        exc_df = pd.DataFrame(calc_results[subject]["优秀率"]["优秀率表格"])
        qua_df = pd.DataFrame(calc_results[subject]["合格率"]["合格率表格"])
        ref_df = pd.DataFrame(calc_results[subject]["班级参考率"])
        merged_df = avg_df.merge(exc_df, on=["班级", "学科"], how="outer")
        merged_df = merged_df.merge(qua_df, on=["班级", "学科"], how="outer")
        merged_df = merged_df.merge(ref_df, on=["班级"], how="outer")
        return merged_df

    teacher_rows = []

    # ---- 步骤1：学科教师 ----
    process_log.append("=== 开始处理学科教师得分（混合口径）===")
    subjects = [key for key in calc_results.keys() if key != "教师得分表"]
    for subject in subjects:
        mode = calc_results[subject].get("口径", 2)
        try:
            merged_df = merge_subject_indicators(subject, mode)
        except (KeyError, IndexError) as e:
            process_log.append(f"科目{subject}缺少必要的指标数据，跳过处理：{str(e)}")
            continue

        teachers = teacher_schedule_df[
            (teacher_schedule_df["科目"] == subject) & (teacher_schedule_df["科目"] != "班主任")]
        for _, teacher_row in teachers.iterrows():
            class_name = teacher_row["班级"]
            class_data = merged_df[(merged_df["班级"] == class_name) & (merged_df["学科"] == subject)]
            if class_data.empty:
                continue
            record = {
                "身份证号": teacher_row["身份证号"],
                "姓名": teacher_row["教师姓名"],
                "科目": subject,
                "校区": teacher_row.get("校区", pd.NA),
                "口径": mode,
                "班级": class_name,
            }
            if mode == 2:
                record.update({
                    "排名正负": class_data.iloc[0]["正负"],
                    "T分差": class_data.iloc[0]["T分差"],
                    "均分": class_data.iloc[0]["均分"],
                    "优秀人数": class_data.iloc[0]["优秀人数差"],
                    "出库人数": class_data.iloc[0]["出库人数差"],
                    "参考率": class_data.iloc[0]["参考率"],
                })
            else:
                record.update({
                    "均分": class_data.iloc[0]["均分"],
                    "优秀率": class_data.iloc[0]["优秀率"],
                    "合格率": class_data.iloc[0]["合格率"],
                    "参考率": class_data.iloc[0]["参考率"],
                })
            teacher_rows.append(record)

    # ---- 步骤2：班主任（使用总分科目指标，口径随总分而定）----
    process_log.append("=== 开始处理班主任得分（混合口径）===")
    if "总分" in calc_results:
        total_mode = calc_results["总分"].get("口径", 2)
        try:
            merged_total_df = merge_subject_indicators("总分", total_mode)
        except (KeyError, IndexError) as e:
            process_log.append(f"总分科目缺少必要的指标数据，跳过班主任处理：{str(e)}")
            merged_total_df = None

        if merged_total_df is not None:
            head_teachers = teacher_schedule_df[
                teacher_schedule_df["科目"] == "班主任"].drop_duplicates(subset=["教师姓名", "身份证号"])
            for _, teacher_row in head_teachers.iterrows():
                teacher_name = teacher_row["教师姓名"]
                id_card = teacher_row["身份证号"]
                head_classes = teacher_schedule_df[
                    (teacher_schedule_df["教师姓名"] == teacher_name) &
                    (teacher_schedule_df["身份证号"] == id_card) &
                    (teacher_schedule_df["科目"] == "班主任")]["班级"]

                class_indicators = []
                for class_name in head_classes:
                    class_data = merged_total_df[
                        (merged_total_df["班级"] == class_name) & (merged_total_df["学科"] == "总分")]
                    if not class_data.empty:
                        class_indicators.append(class_data.iloc[0])

                if not class_indicators:
                    continue
                count = len(class_indicators)
                record = {
                    "身份证号": id_card,
                    "姓名": teacher_name,
                    "科目": "班主任",
                    "校区": teacher_row.get("校区", pd.NA),
                    "口径": total_mode,
                }
                if total_mode == 2:
                    record.update({
                        "排名正负": round(sum(float(i["正负"]) for i in class_indicators) / count, 2),
                        "T分差": round(sum(float(i["T分差"]) for i in class_indicators) / count, 2),
                        "均分": round(sum(float(i["均分"]) for i in class_indicators) / count, 2),
                        "优秀人数": int(sum(i["优秀人数差"] for i in class_indicators)),
                        "出库人数": int(sum(i["出库人数差"] for i in class_indicators)),
                        "参考率": round(sum(float(i["参考率"]) for i in class_indicators) / count, 4),
                    })
                else:
                    record.update({
                        "均分": round(sum(float(i["均分"]) for i in class_indicators) / count, 2),
                        "优秀率": round(sum(float(i["优秀率"]) for i in class_indicators) / count, 4),
                        "合格率": round(sum(float(i["合格率"]) for i in class_indicators) / count, 4),
                        "参考率": round(sum(float(i["参考率"]) for i in class_indicators) / count, 4),
                    })
                teacher_rows.append(record)

    # ---- 步骤3：聚合与总分计算 ----
    process_log.append("=== 开始聚合教师得分数据（混合口径）===")
    teacher_scores_df = pd.DataFrame(teacher_rows)
    if teacher_scores_df.empty:
        return {"教师得分表": []}, ["没有找到教师得分数据"]

    group_keys = ["身份证号", "姓名", "科目", "校区"]
    parts = []
    dual_part = teacher_scores_df[teacher_scores_df["口径"] == 2]
    if not dual_part.empty:
        parts.append(dual_part.groupby(group_keys, dropna=False).agg(
            排名正负=pd.NamedAgg(column="排名正负", aggfunc="mean"),
            T分差=pd.NamedAgg(column="T分差", aggfunc="mean"),
            均分=pd.NamedAgg(column="均分", aggfunc="mean"),
            优秀人数=pd.NamedAgg(column="优秀人数", aggfunc="sum"),
            出库人数=pd.NamedAgg(column="出库人数", aggfunc="sum"),
            参考率=pd.NamedAgg(column="参考率", aggfunc="mean"),
        ).reset_index())
    single_part = teacher_scores_df[teacher_scores_df["口径"] == 1]
    if not single_part.empty:
        parts.append(single_part.groupby(group_keys, dropna=False).agg(
            均分=pd.NamedAgg(column="均分", aggfunc="mean"),
            优秀率=pd.NamedAgg(column="优秀率", aggfunc="mean"),
            合格率=pd.NamedAgg(column="合格率", aggfunc="mean"),
            参考率=pd.NamedAgg(column="参考率", aggfunc="mean"),
        ).reset_index())

    aggregated_df = pd.concat(parts, ignore_index=True)
    # 口径标注：单场科目（含总分科目，视其口径而定）为 1，其余为 2
    single_mode_subjects = {s for s, i in subject_modes.items() if i.get("mode") == 1}
    aggregated_df["口径"] = [
        1 if subject in single_mode_subjects else 2 for subject in aggregated_df["科目"]]
    # 班主任口径跟随总分科目
    total_mode = calc_results.get("总分", {}).get("口径", 2)
    aggregated_df.loc[aggregated_df["科目"] == "班主任", "口径"] = total_mode

    for col in DUAL_METRIC_COLS + SINGLE_METRIC_COLS:
        if col not in aggregated_df.columns:
            aggregated_df[col] = pd.NA
    for col in ("排名正负", "T分差", "均分", "优秀率", "合格率"):
        if col in aggregated_df.columns:
            aggregated_df[col] = pd.to_numeric(aggregated_df[col], errors="coerce").round(2)
    for col in ("优秀人数", "出库人数"):
        if col in aggregated_df.columns:
            aggregated_df[col] = pd.to_numeric(aggregated_df[col], errors="coerce")
    aggregated_df["参考率"] = pd.to_numeric(aggregated_df["参考率"], errors="coerce").round(4)

    total_score = pd.Series(0.0, index=aggregated_df.index)
    for _, group_idx in aggregated_df.groupby("科目", dropna=False).groups.items():
        mode = int(aggregated_df.loc[group_idx, "口径"].iloc[0])
        metric_cols = DUAL_METRIC_COLS if mode == 2 else SINGLE_METRIC_COLS
        weights = dual_weights if mode == 2 else single_weights
        for col in metric_cols:
            weight = weights.get(col, 0.0)
            if weight == 0:
                continue
            series = pd.to_numeric(aggregated_df.loc[group_idx, col], errors="coerce")
            col_min, col_max = series.min(), series.max()
            if pd.notna(col_min) and col_max != col_min:
                transformed = (series - col_min) / (col_max - col_min) * 100
            else:
                transformed = pd.Series(50.0, index=group_idx)
            if weight < 0:
                transformed = 100 - transformed
            total_score.loc[group_idx] += transformed * (abs(weight) / 100)

    aggregated_df["总分"] = total_score.round(2)

    teacher_scores_table = aggregated_df[[
        "身份证号", "姓名", "科目", "校区", "口径",
        "排名正负", "T分差", "均分", "优秀人数", "出库人数",
        "优秀率", "合格率", "参考率", "总分",
    ]].to_dict(orient="records")

    process_log.append(f"教师得分表生成完成（混合口径），共{len(teacher_scores_table)}条记录")
    return {"教师得分表": teacher_scores_table}, process_log


def export_result_excel_mixed(result: Dict, parsed_data: Dict, output_path: str,
                              exam_date: str = "", cohort: str = "",
                              semester: str = "", exam_name: str = "") -> List[str]:
    """
    导出混合口径结果 Excel。

    表结构：
      - 所有科目均分汇总：全部科目（均分 + 均分排名），标注口径
      - 优秀率汇总 / 合格率汇总：仅无基准科目（首次考试口径）
      - 所有科目排名正负 / T分差 / 优秀人数 / 出库人数汇总：仅有基准科目（两场对比口径）
      - 参考率汇总：取总分的班级参考率（班主任视角）
      - 教师得分表：全部教师，按各自科目口径给出指标列
    """
    export_log = []
    exam_info = build_exam_info(exam_date, cohort, semester, exam_name)

    teacher_schedule_df = parsed_data.get("教师任课表")
    has_teacher_schedule = teacher_schedule_df is not None and not teacher_schedule_df.empty

    def get_teacher_info(class_name, subject):
        if not has_teacher_schedule:
            return {"教师姓名": "", "校区": ""}
        if subject == "总分":
            teacher_row = teacher_schedule_df[
                (teacher_schedule_df["班级"] == class_name) & (teacher_schedule_df["科目"] == "班主任")]
        else:
            teacher_row = teacher_schedule_df[
                (teacher_schedule_df["班级"] == class_name) &
                (teacher_schedule_df["科目"] == subject) &
                (teacher_schedule_df["科目"] != "班主任")]
        if not teacher_row.empty:
            info = teacher_row.iloc[0]
            campus = info.get("校区", "")
            return {"教师姓名": info.get("教师姓名", ""),
                    "校区": "" if pd.isna(campus) else campus}
        return {"教师姓名": "", "校区": ""}

    def reorder_after(df: pd.DataFrame, anchor_col: str, move_cols: List[str]) -> pd.DataFrame:
        if anchor_col not in df.columns:
            return df
        columns = [c for c in df.columns if c not in move_cols]
        idx = columns.index(anchor_col) + 1
        for i, col in enumerate([c for c in move_cols if c in df.columns]):
            columns.insert(idx + i, col)
        return df[columns]

    all_subjects = [key for key in result.keys() if key not in ["教师任课表", "教师得分表"]]
    dual_subjects = [s for s in all_subjects if result[s].get("口径") == 2]
    single_subjects = [s for s in all_subjects if result[s].get("口径") == 1]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        sheet_count = 0

        # 表1：所有科目均分汇总（含两种口径）
        avg_rows = []
        for subject in all_subjects:
            mode = result[subject].get("口径", 2)
            if mode == 2:
                table = result[subject]["均分"]["均分表格"]
                for record in table:
                    teacher_info = get_teacher_info(record["班级"], subject)
                    avg_rows.append({
                        "考试信息": exam_info, "班级": record["班级"], "科目": subject,
                        "教师姓名": teacher_info["教师姓名"], "校区": teacher_info["校区"],
                        "口径": "两场对比" if mode == 2 else "首次考试",
                        "均分": record.get("均分"), "均分排名": record.get("均分排名"),
                    })
            else:
                table = result[subject]["均分排名"]["均分排名表格"]
                for record in table:
                    teacher_info = get_teacher_info(record["班级"], subject)
                    avg_rows.append({
                        "考试信息": exam_info, "班级": record["班级"], "科目": subject,
                        "教师姓名": teacher_info["教师姓名"], "校区": teacher_info["校区"],
                        "口径": "首次考试",
                        "均分": record.get("均分"), "均分排名": record.get("排名"),
                    })
        if avg_rows:
            avg_df = pd.DataFrame(avg_rows)
            avg_df = avg_df.sort_values(by=["科目", "班级"], key=lambda x: x.apply(natural_sort_key)
                                        if x.name == "班级" else x)
            avg_df.to_excel(writer, sheet_name="所有科目均分汇总", index=False)
            sheet_count += 1
            export_log.append("已生成工作表：所有科目均分汇总")

        # 表2~3：优秀率 / 合格率汇总（仅无基准科目）
        for indicator_key, table_key, count_col, rate_col, sheet_name in [
            ("优秀率", "优秀率表格", "优秀人数", "优秀率", "优秀率汇总"),
            ("合格率", "合格率表格", "合格人数", "合格率", "合格率汇总"),
        ]:
            rows = []
            for subject in single_subjects:
                for record in result[subject][indicator_key][table_key]:
                    teacher_info = get_teacher_info(record["班级"], subject)
                    record = dict(record)
                    record["教师姓名"] = teacher_info["教师姓名"]
                    record["校区"] = teacher_info["校区"]
                    record["口径"] = "首次考试"
                    rows.append(record)
            if not rows:
                continue
            merged_df = pd.DataFrame(rows)
            merged_df.insert(0, "考试信息", exam_info)
            merged_df["排名"] = merged_df.groupby("学科")[rate_col].rank(
                ascending=False, method="min").astype(int)
            merged_df = reorder_after(merged_df, "学科", ["教师姓名", "校区", "口径"])
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
            sheet_count += 1
            export_log.append(f"已生成工作表：{sheet_name}")

        # 表4~7：排名正负 / T分差 / 优秀人数 / 出库人数 汇总（仅有基准科目）
        merge_specs = [
            ("排名正负", "排名正负表格", "正负", "所有科目排名正负汇总"),
            ("T分差", "T分差表格", "T分差", "所有科目T分差汇总"),
            ("优秀人数", "优秀人数表格", "优秀人数差", "所有科目优秀人数汇总"),
            ("出库人数", "出库人数表格", "出库人数差", "所有科目出库人数汇总"),
        ]
        for indicator_key, table_key, rank_col, sheet_name in merge_specs:
            all_rows = []
            for subject in dual_subjects:
                for record in result[subject][indicator_key][table_key]:
                    teacher_info = get_teacher_info(record["班级"], subject)
                    record = dict(record)
                    record["教师姓名"] = teacher_info["教师姓名"]
                    record["校区"] = teacher_info["校区"]
                    record["口径"] = "两场对比"
                    all_rows.append(record)
            if not all_rows:
                continue
            merged_df = pd.DataFrame(all_rows)
            merged_df.insert(0, "考试信息", exam_info)
            merged_df["排名"] = merged_df.groupby("学科")[rank_col].rank(
                ascending=False, method="min").astype(int)
            merged_df = reorder_after(merged_df, "学科", ["教师姓名", "校区", "口径"])
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
            sheet_count += 1
            export_log.append(f"已生成工作表：{sheet_name}")

        # 表8：参考率汇总（取总分的班级参考率）
        if "总分" in result and result["总分"].get("班级参考率"):
            reference_rate_data = []
            for record in result["总分"]["班级参考率"]:
                record = dict(record)
                teacher_info = get_teacher_info(record["班级"], "总分")
                record["班主任"] = teacher_info["教师姓名"]
                record["校区"] = teacher_info["校区"]
                reference_rate_data.append(record)
            reference_rate_df = pd.DataFrame(reference_rate_data)
            reference_rate_df.insert(0, "考试信息", exam_info)
            reference_rate_df["排名"] = reference_rate_df["参考率"].rank(
                ascending=False, method="min").astype(int)
            reference_rate_df = reorder_after(reference_rate_df, "班级", ["班主任", "校区"])
            reference_rate_df.to_excel(writer, sheet_name="参考率汇总", index=False)
            sheet_count += 1
            export_log.append("已生成工作表：参考率汇总")

        # 表9：教师得分表
        if result.get("教师得分表", {}).get("教师得分表"):
            teacher_scores_df = pd.DataFrame(result["教师得分表"]["教师得分表"])
            teacher_scores_df.insert(0, "考试信息", exam_info)
            teacher_scores_df["口径"] = teacher_scores_df["口径"].map(
                {1: "首次考试", 2: "两场对比"})
            teacher_scores_df["得分排名"] = teacher_scores_df.groupby("科目")["总分"].rank(
                ascending=False, method="min").astype(int)
            teacher_scores_df = teacher_scores_df.sort_values(
                by=["科目", "总分"], ascending=[True, False])
            teacher_scores_df.to_excel(writer, sheet_name="教师得分表", index=False)
            sheet_count += 1
            export_log.append("已生成工作表：教师得分表")

        if sheet_count == 0:
            pd.DataFrame({"考试信息": [exam_info], "说明": ["没有可生成的分析结果数据"]}).to_excel(
                writer, sheet_name="无数据", index=False)
            export_log.append("没有可生成的分析结果数据")

    return export_log


def parse_exam_filename(path: str) -> "dict | None":
    """
    按标准文件名格式解析考试元信息。

    标准格式：YYYY-MM-DD[_届别[_学期[_考试名称]]].xlsx
    示例：
        2026-01-15_2028_1_期末考试.xlsx
            → {exam_date:'2026-01-15', cohort:'2028', semester:'1', exam_name:'期末考试'}
        2026-01-15.xlsx
            → {exam_date:'2026-01-15', cohort:'', semester:'', exam_name:''}

    文件名不以 YYYY-MM-DD 开头即视为「非标准格式」，返回 None，
    此时调用方（WorkBuddy）必须主动向用户询问考试名称等元信息，不可留空。
    """
    import re as _re
    base = os.path.splitext(os.path.basename(path))[0]
    m = _re.match(
        r"^(\d{4}-\d{2}-\d{2})(?:_(\d{4}))?(?:_([12]))?(?:_(.+))?$",
        base)
    if not m:
        return None
    return {
        "exam_date": m.group(1) or "",
        "cohort": m.group(2) or "",
        "semester": m.group(3) or "",
        "exam_name": (m.group(4) or "").strip(),
    }


# ---------------- 结果导出 ----------------
def build_exam_info(exam_date: str, cohort: str, semester: str, exam_name: str) -> str:
    """组合考试信息字符串：日期-X届-X学期-考试名称"""
    semester_text = {"1": "第一", "2": "第二"}.get(str(semester), semester) if semester else ""

    formatted_date = ""
    if exam_date:
        try:
            date_obj = datetime.strptime(exam_date, "%Y-%m-%d")
            formatted_date = f"{date_obj.year}年{date_obj.month}月{date_obj.day}日"
        except ValueError:
            formatted_date = exam_date

    parts = []
    if formatted_date:
        parts.append(formatted_date)
    if cohort:
        parts.append(f"{cohort}届")
    if semester_text:
        parts.append(f"{semester_text}学期")
    if exam_name:
        parts.append(exam_name)
    return "-".join(parts)


def export_result_excel(result: Dict, parsed_data: Dict, output_path: str,
                        exam_date: str = "", cohort: str = "",
                        semester: str = "", exam_name: str = "") -> List[str]:
    """将计算结果导出为多工作表 Excel，返回导出日志"""
    export_log = []
    exam_info = build_exam_info(exam_date, cohort, semester, exam_name)

    teacher_schedule_df = parsed_data.get("教师任课表")
    has_teacher_schedule = teacher_schedule_df is not None and not teacher_schedule_df.empty

    def get_teacher_info(class_name, subject):
        """根据班级和科目查找教师信息；总分对应班主任"""
        if not has_teacher_schedule:
            return {"教师姓名": "", "校区": ""}
        if subject == "总分":
            teacher_row = teacher_schedule_df[
                (teacher_schedule_df["班级"] == class_name) & (teacher_schedule_df["科目"] == "班主任")]
        else:
            teacher_row = teacher_schedule_df[
                (teacher_schedule_df["班级"] == class_name) &
                (teacher_schedule_df["科目"] == subject) &
                (teacher_schedule_df["科目"] != "班主任")]
        if not teacher_row.empty:
            info = teacher_row.iloc[0]
            campus = info.get("校区", "")
            return {"教师姓名": info.get("教师姓名", ""),
                    "校区": "" if pd.isna(campus) else campus}
        return {"教师姓名": "", "校区": ""}

    def reorder_after(df: pd.DataFrame, anchor_col: str, move_cols: List[str]) -> pd.DataFrame:
        """把 move_cols 移动到 anchor_col 之后"""
        if anchor_col not in df.columns:
            return df
        columns = [c for c in df.columns if c not in move_cols]
        idx = columns.index(anchor_col) + 1
        for i, col in enumerate([c for c in move_cols if c in df.columns]):
            columns.insert(idx + i, col)
        return df[columns]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        sheet_count = 0
        all_subjects = [key for key in result.keys()
                        if key not in ["教师任课表", "教师得分表"]]

        # 表1：所有科目均分汇总（一维表：班级×科目）
        integrated_avg_data = []
        for class_name in sorted(
                {r["班级"] for s in all_subjects
                 for r in result[s]["均分"]["均分表格"]},
                key=natural_sort_key):
            for subject in all_subjects:
                avg_table = result[subject]["均分"]["均分表格"]
                record = next((r for r in avg_table if r["班级"] == class_name), None)
                if record is None:
                    continue
                teacher_info = get_teacher_info(class_name, subject)
                integrated_avg_data.append({
                    "考试信息": exam_info, "班级": class_name, "科目": subject,
                    "教师姓名": teacher_info["教师姓名"], "校区": teacher_info["校区"],
                    "均分1": record.get("第一场均分"), "均分2": record.get("第二场均分"),
                    "均分": record.get("均分"), "均分排名": record.get("均分排名"),
                })
        pd.DataFrame(integrated_avg_data).to_excel(
            writer, sheet_name="所有科目均分汇总", index=False)
        sheet_count += 1
        export_log.append("已生成工作表：所有科目均分汇总")

        # 表2~5：排名正负 / T分差 / 优秀人数 / 出库人数 汇总
        merge_specs = [
            ("排名正负", "排名正负表格", "正负", "所有科目排名正负汇总"),
            ("T分差", "T分差表格", "T分差", "所有科目T分差汇总"),
            ("优秀人数", "优秀人数表格", "优秀人数差", "所有科目优秀人数汇总"),
            ("出库人数", "出库人数表格", "出库人数差", "所有科目出库人数汇总"),
        ]
        for indicator_key, table_key, rank_col, sheet_name in merge_specs:
            all_rows = []
            for subject in all_subjects:
                table = result[subject][indicator_key][table_key]
                for record in table:
                    teacher_info = get_teacher_info(record["班级"], subject)
                    record = dict(record)
                    record["教师姓名"] = teacher_info["教师姓名"]
                    record["校区"] = teacher_info["校区"]
                    all_rows.append(record)
            if not all_rows:
                continue
            merged_df = pd.DataFrame(all_rows)
            merged_df.insert(0, "考试信息", exam_info)
            merged_df["排名"] = merged_df.groupby("学科")[rank_col].rank(
                ascending=False, method="min").astype(int)
            merged_df = reorder_after(merged_df, "学科", ["教师姓名", "校区"])
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
            sheet_count += 1
            export_log.append(f"已生成工作表：{sheet_name}")

        # 表6：参考率汇总（使用总分的班级参考率）
        if "总分" in result and result["总分"]["班级参考率"]:
            reference_rate_data = []
            for record in result["总分"]["班级参考率"]:
                record = dict(record)
                teacher_info = get_teacher_info(record["班级"], "总分")
                record["班主任"] = teacher_info["教师姓名"]
                record["校区"] = teacher_info["校区"]
                reference_rate_data.append(record)
            reference_rate_df = pd.DataFrame(reference_rate_data)
            reference_rate_df.insert(0, "考试信息", exam_info)
            reference_rate_df["排名"] = reference_rate_df["参考率"].rank(
                ascending=False, method="min").astype(int)
            reference_rate_df = reorder_after(reference_rate_df, "班级", ["班主任", "校区"])
            reference_rate_df.to_excel(writer, sheet_name="参考率汇总", index=False)
            sheet_count += 1
            export_log.append("已生成工作表：参考率汇总")

        # 表7：教师得分表
        if result.get("教师得分表", {}).get("教师得分表"):
            teacher_scores_df = pd.DataFrame(result["教师得分表"]["教师得分表"])
            teacher_scores_df.insert(0, "考试信息", exam_info)
            if "总分" in teacher_scores_df.columns:
                if "科目" in teacher_scores_df.columns:
                    teacher_scores_df["排名"] = teacher_scores_df.groupby("科目")["总分"].rank(
                        ascending=False, method="min").astype(int)
                else:
                    teacher_scores_df["排名"] = teacher_scores_df["总分"].rank(
                        ascending=False, method="min").astype(int)
                # 将排名列移到总分列之后
                columns = list(teacher_scores_df.columns)
                columns.remove("排名")
                columns.insert(columns.index("总分") + 1, "排名")
                teacher_scores_df = teacher_scores_df[columns]
            teacher_scores_df.to_excel(writer, sheet_name="教师得分表", index=False)
            sheet_count += 1
            export_log.append("已生成工作表：教师得分表")

        if sheet_count == 0:
            pd.DataFrame({"考试信息": [exam_info], "说明": ["没有可生成的分析结果数据"]}).to_excel(
                writer, sheet_name="无数据", index=False)
            export_log.append("没有可生成的分析结果数据")

    return export_log


def export_result_excel_single(result: Dict, parsed_data: Dict, output_path: str,
                               exam_date: str = "", cohort: str = "",
                               semester: str = "", exam_name: str = "") -> List[str]:
    """导出单次考试（首次考试）结果 Excel：均分与排名 / 优秀率 / 合格率 / 参考率 / 教师得分表"""
    export_log = []
    exam_info = build_exam_info(exam_date, cohort, semester, exam_name)

    teacher_schedule_df = parsed_data.get("教师任课表")
    has_teacher_schedule = teacher_schedule_df is not None and not teacher_schedule_df.empty

    def get_teacher_info(class_name, subject):
        """根据班级和科目查找教师信息；总分对应班主任"""
        if not has_teacher_schedule:
            return {"教师姓名": "", "校区": ""}
        if subject == "总分":
            teacher_row = teacher_schedule_df[
                (teacher_schedule_df["班级"] == class_name) & (teacher_schedule_df["科目"] == "班主任")]
        else:
            teacher_row = teacher_schedule_df[
                (teacher_schedule_df["班级"] == class_name) &
                (teacher_schedule_df["科目"] == subject) &
                (teacher_schedule_df["科目"] != "班主任")]
        if not teacher_row.empty:
            info = teacher_row.iloc[0]
            campus = info.get("校区", "")
            return {"教师姓名": info.get("教师姓名", ""),
                    "校区": "" if pd.isna(campus) else campus}
        return {"教师姓名": "", "校区": ""}

    all_subjects = [key for key in result.keys()
                    if key not in ["教师任课表", "教师得分表"]]

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        sheet_count = 0

        # 表1：均分汇总（排名不单列为指标——排名即班级均分的横向位次，与均分属同一项指标）
        avg_rows = []
        for subject in all_subjects:
            for record in result[subject]["均分排名"]["均分排名表格"]:
                teacher_info = get_teacher_info(record["班级"], subject)
                avg_rows.append({
                    "考试信息": exam_info, "班级": record["班级"], "学科": subject,
                    "教师姓名": teacher_info["教师姓名"], "校区": teacher_info["校区"],
                    "均分": record.get("均分"),
                    "总人数": record.get("总人数"), "参考人数": record.get("参考人数"),
                })
        if avg_rows:
            avg_df = pd.DataFrame(avg_rows)
            # 均分排名为本表内部的行序（与各表末列口径一致），非独立评价指标
            avg_df["均分排名"] = avg_df.groupby("学科")["均分"].rank(
                ascending=False, method="min").astype(int)
            avg_df = avg_df[
                ["考试信息", "班级", "学科", "教师姓名", "校区", "均分", "均分排名",
                 "总人数", "参考人数"]]
            avg_df.to_excel(writer, sheet_name="均分汇总", index=False)
            sheet_count += 1
            export_log.append("已生成工作表：均分汇总")

        # 表2、表3：优秀率 / 合格率汇总
        rate_specs = [
            ("优秀率", "优秀率表格", "优秀人数", "优秀率", "优秀率汇总"),
            ("合格率", "合格率表格", "合格人数", "合格率", "合格率汇总"),
        ]
        for indicator_key, table_key, count_col, rate_col, sheet_name in rate_specs:
            rows = []
            for subject in all_subjects:
                for record in result[subject][indicator_key][table_key]:
                    teacher_info = get_teacher_info(record["班级"], subject)
                    rows.append({
                        "考试信息": exam_info, "班级": record["班级"], "学科": subject,
                        "教师姓名": teacher_info["教师姓名"], "校区": teacher_info["校区"],
                        count_col: record.get(count_col), rate_col: record.get(rate_col),
                    })
            if not rows:
                continue
            merged_df = pd.DataFrame(rows)
            merged_df["排名"] = merged_df.groupby("学科")[rate_col].rank(
                ascending=False, method="min").astype(int)
            merged_df = merged_df[
                ["考试信息", "班级", "学科", "教师姓名", "校区", count_col, rate_col, "排名"]]
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)
            sheet_count += 1
            export_log.append(f"已生成工作表：{sheet_name}")

        # 表4：参考率汇总（使用总分的班级参考率）
        if "总分" in result and result["总分"]["班级参考率"]:
            reference_rate_rows = []
            for record in result["总分"]["班级参考率"]:
                teacher_info = get_teacher_info(record["班级"], "总分")
                reference_rate_rows.append({
                    "考试信息": exam_info, "班级": record["班级"],
                    "班主任": teacher_info["教师姓名"], "校区": teacher_info["校区"],
                    "总人数": record.get("总人数"), "参考人数": record.get("参考人数"),
                    "参考率": record.get("参考率"),
                })
            reference_rate_df = pd.DataFrame(reference_rate_rows)
            reference_rate_df["排名"] = reference_rate_df["参考率"].rank(
                ascending=False, method="min").astype(int)
            reference_rate_df.to_excel(writer, sheet_name="参考率汇总", index=False)
            sheet_count += 1
            export_log.append("已生成工作表：参考率汇总")

        # 表5：教师得分表
        if result.get("教师得分表", {}).get("教师得分表"):
            teacher_scores_df = pd.DataFrame(result["教师得分表"]["教师得分表"])
            teacher_scores_df.insert(0, "考试信息", exam_info)
            if "科目" in teacher_scores_df.columns:
                teacher_scores_df["得分排名"] = teacher_scores_df.groupby("科目")["总分"].rank(
                    ascending=False, method="min").astype(int)
            else:
                teacher_scores_df["得分排名"] = teacher_scores_df["总分"].rank(
                    ascending=False, method="min").astype(int)
            columns = list(teacher_scores_df.columns)
            columns.remove("得分排名")
            columns.insert(columns.index("总分") + 1, "得分排名")
            teacher_scores_df = teacher_scores_df[columns]
            teacher_scores_df.to_excel(writer, sheet_name="教师得分表", index=False)
            sheet_count += 1
            export_log.append("已生成工作表：教师得分表")

        if sheet_count == 0:
            pd.DataFrame({"考试信息": [exam_info], "说明": ["没有可生成的分析结果数据"]}).to_excel(
                writer, sheet_name="无数据", index=False)
            export_log.append("没有可生成的分析结果数据")

    return export_log


def make_template(output_path: str, single: bool = False) -> None:
    """
    生成空白 Excel 模板（成绩表 + 教师任课表表头）。

    single=True 时只生成第一场（首次考试）需要的科目列，不含 '2' 后缀列。
    """
    grade_headers = ["考号", "姓名", "集团班级"]
    if single:
        grade_headers += ["语文1", "数学1", "英语1", "物理1", "化学1", "生物1",
                          "政治1", "历史1", "地理1", "总分1"]
    else:
        grade_headers += ["语文1", "数学1", "英语1", "物理1", "化学1", "生物1",
                          "政治1", "历史1", "地理1", "总分1",
                          "语文2", "数学2", "英语2", "物理2", "化学2", "生物2",
                          "政治2", "历史2", "地理2", "总分2"]
    teacher_headers = ["教师姓名", "校区", "科目", "班级", "身份证号"]
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        pd.DataFrame(columns=grade_headers).to_excel(writer, sheet_name="成绩表", index=False)
        pd.DataFrame(columns=teacher_headers).to_excel(writer, sheet_name="教师任课表", index=False)


# ---------------- 命令行入口 ----------------
def main():
    parser = argparse.ArgumentParser(description="考试成绩分析与教师绩效算分")
    parser.add_argument("-i", "--input", help="输入 Excel 文件路径（含成绩表、教师任课表）")
    parser.add_argument("-o", "--output", help="输出结果 Excel 路径（默认在输入文件旁生成）")
    parser.add_argument("--template", help="仅生成空白模板到指定路径后退出")
    parser.add_argument("--excellent-ratio", type=float, default=20.0,
                        help="优秀比例（全年级排名前 X%%）。默认：两场 20，单次 20")
    parser.add_argument("--qualified-ratio", type=float, default=None,
                        help="合格/出库比例（全年级排名前 X%%）。默认：两场 75，单次 80")
    parser.add_argument("--rank-weight", type=float, default=None,
                        help="排名权重：两场=排名正负；单次考试下不适用（排名与均分为同一项指标）")
    parser.add_argument("--t-diff-weight", type=float, default=None,
                        help="T分差权重（仅两场考试生效）。默认 20")
    parser.add_argument("--avg-weight", type=float, default=None,
                        help="均分权重（单次考试下含「均分与排名」合并权重）。默认：两场 20，单次 40")
    parser.add_argument("--excellent-weight", type=float, default=None,
                        help="优秀人数（两场）/ 优秀率（单次）的权重。默认：两场 15，单次 25")
    parser.add_argument("--qualified-weight", type=float, default=None,
                        help="出库人数（两场）/ 合格率（单次）的权重。默认：两场 15，单次 25")
    parser.add_argument("--reference-weight", type=float, default=None,
                        help="参考率权重。默认 10")
    parser.add_argument("--exam-count", choices=["1", "2", "3", "auto"], default="auto",
                        help="考试场次数：1=首次考试（仅一次成绩，四项指标）；"
                             "2=两场考试对比（六项指标）；"
                             "3=混合口径（逐科目分流：有基准走六项、无基准走四项）；"
                             "auto=按成绩表列名自动识别（默认）")
    parser.add_argument("--subject-mode", default="",
                        help="手工覆盖个别科目口径，形如「物理:1,化学:2」。"
                             "1=首次考试口径（四项指标），2=两场对比口径（六项指标）")
    parser.add_argument("--single-qualified-ratio", type=float, default=None,
                        help="混合口径下：无基准科目的合格比例（全年级前 X%%）。默认 80")
    parser.add_argument("--single-avg-weight", type=float, default=None,
                        help="混合口径下：无基准科目的均分权重。默认 60")
    parser.add_argument("--single-excellent-weight", type=float, default=None,
                        help="混合口径下：无基准科目的优秀率权重。默认 15")
    parser.add_argument("--single-qualified-weight", type=float, default=None,
                        help="混合口径下：无基准科目的合格率权重。默认 15")
    parser.add_argument("--single-reference-weight", type=float, default=None,
                        help="混合口径下：无基准科目的参考率权重。默认 10")
    parser.add_argument("--exam-date", default="", help="考试日期，格式 YYYY-MM-DD")
    parser.add_argument("--cohort", default="", help="届别，如 2028")
    parser.add_argument("--semester", default="", help="学期：1=第一学期，2=第二学期")
    parser.add_argument("--exam-name", default="", help="考试名称，如 期末考试")
    parser.add_argument("--log", help="可选：将完整计算日志写入该文本文件")
    parser.add_argument("--parse-filename", default=None,
                        help="仅解析指定文件路径的文件名是否符合标准格式（YYYY-MM-DD[_届别[_学期[_考试名称]]].xlsx），"
                             "输出 JSON 解析结果或 'non-standard' 后退出")
    parser.add_argument("--detect", default=None,
                        help="仅检测指定文件的考试场次与科目列（输出 JSON）后退出，"
                             "用于判断该用首次考试口径还是两场对比口径")
    args = parser.parse_args()

    if args.template:
        make_template(args.template, single=(args.exam_count == "1"))
        print(f"模板已生成：{args.template}"
              + ("（首次考试模板，仅含第一场科目列）" if args.exam_count == "1" else ""))
        return

    if args.parse_filename:
        parsed = parse_exam_filename(args.parse_filename)
        print(json.dumps(parsed, ensure_ascii=False) if parsed else "non-standard")
        return

    if args.detect:
        excel_file = pd.ExcelFile(args.detect)
        if "成绩表" not in excel_file.sheet_names:
            raise ValueError("Excel 文件中缺少「成绩表」工作表")
        df = excel_file.parse(sheet_name="成绩表")
        count = detect_exam_count(df)
        first_cols = [c for c in df.columns if c.endswith("1") and c != "总分1"]
        second_cols = [c for c in df.columns if c.endswith("2") and c != "总分2"]
        if count == 1 and not first_cols:
            first_cols = [c for c in df.columns
                          if c not in set(BASIC_COLUMNS) | {"总分", "总分1"}]
        print(json.dumps({
            "exam_count": count,
            "mode": "首次考试（四项指标）" if count == 1 else "两场考试对比（六项指标）",
            "rows": len(df),
            "first_subject_columns": first_cols,
            "second_subject_columns": second_cols,
            "has_total_column": any(c in df.columns for c in ("总分", "总分1", "总分2")),
        }, ensure_ascii=False))
        return

    if not args.input:
        parser.error("请通过 -i/--input 指定输入 Excel 文件（或用 --template 生成模板）")

    # 文件名标准格式检查：非标准格式且未提供任何考试元信息时，
    # 提醒调用方（WorkBuddy）主动向用户询问考试名称，不可直接留空。
    if not (args.exam_date or args.cohort or args.semester or args.exam_name):
        if parse_exam_filename(args.input) is None:
            print("警告：输入文件名非标准格式（应为 YYYY-MM-DD[_届别[_学期[_考试名称]]].xlsx），"
                  "且未提供 --exam-date/--cohort/--semester/--exam-name 等考试信息；"
                  "在由 WorkBuddy 调用时，应先向用户询问考试名称后再分析。",
                  file=sys.stderr)

    # 1. 解析（考试场次：显式指定优先，否则按列名自动识别；混合口径逐科目分流）
    requested_count = 0 if args.exam_count == "auto" else int(args.exam_count)
    parsed_data, parse_log, mode, subject_modes = parse_excel(
        args.input, requested_count, args.subject_mode)
    for line in parse_log:
        print(line)

    # 2. 参数默认值（各口径各有一套默认值，未显式传参时按口径自动套用）
    if mode == 1:
        ratio_defaults = SINGLE_DEFAULT_RATIOS
        weight_defaults = SINGLE_DEFAULT_WEIGHTS
    elif mode == 2:
        ratio_defaults = DUAL_DEFAULT_RATIOS
        weight_defaults = DUAL_DEFAULT_WEIGHTS
    else:
        # 混合口径：有基准科目沿用两场参数，无基准科目用独立的 single-* 参数
        ratio_defaults = DUAL_DEFAULT_RATIOS
        weight_defaults = DUAL_DEFAULT_WEIGHTS

    excellent_ratio = (args.excellent_ratio if args.excellent_ratio is not None
                       else ratio_defaults["excellent"])
    qualified_ratio = (args.qualified_ratio if args.qualified_ratio is not None
                       else ratio_defaults["qualified"])
    rank_weight = (args.rank_weight if args.rank_weight is not None
                   else weight_defaults.get("rank", 0.0))
    t_diff_weight = (args.t_diff_weight if args.t_diff_weight is not None
                     else weight_defaults.get("t_diff", 0.0))
    avg_weight = args.avg_weight if args.avg_weight is not None else weight_defaults["avg"]
    excellent_weight = (args.excellent_weight if args.excellent_weight is not None
                        else weight_defaults["excellent"])
    qualified_weight = (args.qualified_weight if args.qualified_weight is not None
                        else weight_defaults["qualified"])
    reference_weight = (args.reference_weight if args.reference_weight is not None
                        else weight_defaults["reference"])

    # 混合口径下无基准科目（首次考试口径）的独立参数
    single_qualified_ratio = (args.single_qualified_ratio if args.single_qualified_ratio is not None
                               else SINGLE_DEFAULT_RATIOS["qualified"])
    single_avg_weight = (args.single_avg_weight if args.single_avg_weight is not None
                          else SINGLE_DEFAULT_WEIGHTS["avg"])
    single_excellent_weight = (args.single_excellent_weight if args.single_excellent_weight is not None
                                else SINGLE_DEFAULT_WEIGHTS["excellent"])
    single_qualified_weight = (args.single_qualified_weight if args.single_qualified_weight is not None
                                else SINGLE_DEFAULT_WEIGHTS["qualified"])
    single_reference_weight = (args.single_reference_weight if args.single_reference_weight is not None
                                else SINGLE_DEFAULT_WEIGHTS["reference"])

    if mode == 1:
        # 首次考试口径为四项指标，「均分与排名」合为均分一项，排名不单独计分
        if args.rank_weight is not None:
            print("提示：首次考试口径下「排名」与「均分」为同一项指标，"
                  "--rank-weight 不再生效，已忽略。")
        print(f"参数：优秀比例={excellent_ratio}%，合格比例={qualified_ratio}%，"
              f"权重=均分{avg_weight}/优秀率{excellent_weight}/"
              f"合格率{qualified_weight}/参考率{reference_weight}")
    elif mode == 2:
        print(f"参数：优秀比例={excellent_ratio}%，合格/出库比例={qualified_ratio}%，"
              f"权重=排名正负{rank_weight}/T分差{t_diff_weight}/均分{avg_weight}/"
              f"优秀{excellent_weight}/出库{qualified_weight}/参考率{reference_weight}")
    else:
        dual_subjects = [s for s, i in subject_modes.items() if i.get("mode") == 2]
        single_subjects = [s for s, i in subject_modes.items() if i.get("mode") == 1]
        print(f"参数（有基准科目 {'、'.join(dual_subjects) if dual_subjects else '无'}）："
              f"优秀比例={excellent_ratio}%，合格/出库比例={qualified_ratio}%，"
              f"权重=排名正负{rank_weight}/T分差{t_diff_weight}/均分{avg_weight}/"
              f"优秀{excellent_weight}/出库{qualified_weight}/参考率{reference_weight}")
        print(f"参数（无基准科目 {'、'.join(single_subjects) if single_subjects else '无'}）："
              f"优秀比例={excellent_ratio}%，合格比例={single_qualified_ratio}%，"
              f"权重=均分{single_avg_weight}/优秀率{single_excellent_weight}/"
              f"合格率{single_qualified_weight}/参考率{single_reference_weight}")

    # 3. 计算
    if mode == 1:
        result, calc_log = calculate_all_indicators_single(
            parsed_data,
            excellent_ratio=excellent_ratio,
            qualified_ratio=qualified_ratio,
            avg_weight=avg_weight,
            excellent_weight=excellent_weight, qualified_weight=qualified_weight,
            reference_weight=reference_weight)
    elif mode == 2:
        result, calc_log = calculate_all_indicators(
            parsed_data,
            excellent_ratio=excellent_ratio,
            qualified_ratio=qualified_ratio,
            rank_weight=rank_weight,
            t_diff_weight=t_diff_weight,
            avg_weight=avg_weight,
            excellent_weight=excellent_weight,
            qualified_weight=qualified_weight,
            reference_weight=reference_weight,
        )
    else:
        result, calc_log = calculate_all_indicators_mixed(
            parsed_data, subject_modes,
            excellent_ratio=excellent_ratio,
            qualified_ratio=qualified_ratio,
            single_qualified_ratio=single_qualified_ratio,
            dual_weights={
                "排名正负": rank_weight, "T分差": t_diff_weight, "均分": avg_weight,
                "优秀人数": excellent_weight, "出库人数": qualified_weight,
                "参考率": reference_weight,
            },
            single_weights={
                "均分": single_avg_weight, "优秀率": single_excellent_weight,
                "合格率": single_qualified_weight, "参考率": single_reference_weight,
            },
        )

    subjects = [k for k in result.keys() if k != "教师得分表"]
    if mode == 1:
        class_count = (len(result[subjects[0]]["均分排名"]["均分排名表格"])
                       if subjects else 0)
    elif mode == 2:
        class_count = len(result[subjects[0]]["T分差"]["T分差表格"]) if subjects else 0
    else:
        class_count = 0
        for subject in subjects:
            if result[subject].get("口径") == 2:
                class_count = len(result[subject]["T分差"]["T分差表格"])
            else:
                class_count = len(result[subject]["均分排名"]["均分排名表格"])
            break
    teacher_count = len(result["教师得分表"].get("教师得分表", []))
    mode_text = {1: "单次考试（四项指标）",
                 2: "两场考试对比（六项指标）",
                 3: "混合口径（按科目分流）"}[mode]
    print(f"计算完成：{mode_text}，科目{len(subjects)}个（含总分），班级{class_count}个，"
          f"教师得分记录{teacher_count}条")

    # 4. 导出
    output_path = args.output
    if not output_path:
        base = re.sub(r"\.(xlsx|xls)$", "", args.input, flags=re.IGNORECASE)
        output_path = f"{base}_分析结果.xlsx"
    if mode == 1:
        export_log = export_result_excel_single(
            result, parsed_data, output_path,
            exam_date=args.exam_date, cohort=args.cohort,
            semester=args.semester, exam_name=args.exam_name)
    elif mode == 2:
        export_log = export_result_excel(
            result, parsed_data, output_path,
            exam_date=args.exam_date, cohort=args.cohort,
            semester=args.semester, exam_name=args.exam_name)
    else:
        export_log = export_result_excel_mixed(
            result, parsed_data, output_path,
            exam_date=args.exam_date, cohort=args.cohort,
            semester=args.semester, exam_name=args.exam_name)
    for line in export_log:
        print(line)
    print(f"结果已导出：{output_path}")

    # 5. 日志（可选）
    if args.log:
        with open(args.log, "w", encoding="utf-8") as f:
            f.write("\n".join(parse_log + calc_log))
        print(f"计算日志已写入：{args.log}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"执行失败：{e}", file=sys.stderr)
        sys.exit(1)

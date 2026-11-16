# -*- coding: utf-8 -*-
"""
SQLite 持久化模块（零依赖，仅用 Python 标准库 sqlite3）

负责把每次成绩分析的结果落库，便于跨考试历史查询、班级/校区对比与可视化。

数据库位置：默认在技能目录下 data/analysis.db，可通过 --db 指定。

表结构：
  exams          每次分析的一行元信息（exam_count：1=首次考试，2=两场对比）
  class_metrics  各考试 × 班级 × 学科的明细指标
  teacher_scores 各考试 × 教师（含学科教师与班主任）的得分记录
"""

import os
import sqlite3
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd

DB_SCHEMA = """
CREATE TABLE IF NOT EXISTS exams (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    exam_info   TEXT,
    exam_date   TEXT,
    cohort      TEXT,
    semester    TEXT,
    exam_name   TEXT,
    source_file TEXT,
    exam_count  INTEGER DEFAULT 2,
    created_at  TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS class_metrics (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    exam_id         INTEGER,
    campus          TEXT,
    班级            TEXT,
    学科            TEXT,
    教师姓名        TEXT,
    均分1           REAL,
    均分2           REAL,
    均分            REAL,
    T分1            REAL,
    T分2            REAL,
    T分差           REAL,
    T分差排名       INTEGER,
    线性排名1       REAL,
    线性排名2       REAL,
    正负            REAL,
    正负排名        INTEGER,
    优秀人数1       INTEGER,
    优秀人数2       INTEGER,
    优秀人数差      INTEGER,
    出库人数1       INTEGER,
    出库人数2       INTEGER,
    出库人数差      INTEGER,
    总人数          INTEGER,
    参考人数        INTEGER,
    参考率          REAL,
    排名            INTEGER,
    线性排名        REAL,
    排名得分        REAL,
    优秀人数        INTEGER,
    优秀率          REAL,
    合格人数        INTEGER,
    合格率          REAL,
    FOREIGN KEY (exam_id) REFERENCES exams(id)
);

CREATE TABLE IF NOT EXISTS teacher_scores (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    exam_id      INTEGER,
    身份证号     TEXT,
    姓名         TEXT,
    科目         TEXT,
    校区         TEXT,
    排名正负     REAL,
    T分差        REAL,
    均分         REAL,
    优秀人数     REAL,
    出库人数     REAL,
    参考率       REAL,
    总分         REAL,
    排名         REAL,
    排名得分     REAL,
    优秀率       REAL,
    合格率       REAL,
    FOREIGN KEY (exam_id) REFERENCES exams(id)
);

CREATE INDEX IF NOT EXISTS idx_class_metrics_exam ON class_metrics(exam_id);
CREATE INDEX IF NOT EXISTS idx_teacher_scores_exam ON teacher_scores(exam_id);
"""


def init_db(db_path: str) -> None:
    """初始化数据库与表结构；父目录不存在时自动创建。"""
    parent = os.path.dirname(db_path)
    if parent and not os.path.exists(parent):
        os.makedirs(parent, exist_ok=True)
    conn = sqlite3.connect(db_path)
    try:
        conn.executescript(DB_SCHEMA)
        # 迁移：旧库补充新增列（新库已在 CREATE TABLE 中定义，此步自动跳过）
        #   教师姓名    —— class_metrics 的教师映射
        #   均分        —— class_metrics 的「均分」结果评价指标（= 第二场班级均分）
        #   均分        —— teacher_scores 的均分指标得分
        #   exam_count  —— exams 的考试场次（1=首次考试，2=两场对比）
        #   排名/线性排名/优秀人数/优秀率/合格人数/合格率 —— 单次考试指标（排名仅参考）
        migrations = [
            ("class_metrics", "教师姓名", "TEXT"),
            ("class_metrics", "均分", "REAL"),
            ("teacher_scores", "均分", "REAL"),
            ("exams", "exam_count", "INTEGER DEFAULT 2"),
            ("class_metrics", "排名", "INTEGER"),
            ("class_metrics", "线性排名", "REAL"),
            ("class_metrics", "排名得分", "REAL"),
            ("class_metrics", "优秀人数", "INTEGER"),
            ("class_metrics", "优秀率", "REAL"),
            ("class_metrics", "合格人数", "INTEGER"),
            ("class_metrics", "合格率", "REAL"),
            ("teacher_scores", "排名", "REAL"),
            ("teacher_scores", "排名得分", "REAL"),
            ("teacher_scores", "优秀率", "REAL"),
            ("teacher_scores", "合格率", "REAL"),
        ]
        for table, column, col_type in migrations:
            cols = [r[1] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()]
            if column not in cols:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {col_type}")
        # 旧库中 exam_count 为空的记录按「两场考试」处理
        conn.execute("UPDATE exams SET exam_count=2 WHERE exam_count IS NULL")
        # 清理：早期版本曾在 class_metrics 建过 均分差 列（均分指标口径已改为
        # 直接取第二场均分，该列作废）。存在则删除，避免与 均分 列混淆。
        cm_cols = [r[1] for r in conn.execute("PRAGMA table_info(class_metrics)").fetchall()]
        if "均分差" in cm_cols:
            conn.execute("ALTER TABLE class_metrics DROP COLUMN 均分差")
        conn.commit()
    finally:
        conn.close()


def _build_class_campus_map(parsed_data: Dict) -> Dict[str, str]:
    """从教师任课表（班主任行）推断 班级→校区 映射。"""
    teacher_df = parsed_data.get("教师任课表")
    if teacher_df is None or teacher_df.empty:
        return {}
    campus_map = {}
    if "校区" in teacher_df.columns and "班级" in teacher_df.columns:
        head = teacher_df[teacher_df.get("科目") == "班主任"] if "科目" in teacher_df.columns else teacher_df
        for _, row in head.iterrows():
            cls = str(row["班级"]).strip()
            cam = row.get("校区")
            cam = "" if pd.isna(cam) else str(cam).strip()
            if cls:
                campus_map[cls] = cam
    return campus_map


def _build_teacher_map(teacher_df) -> Dict[Tuple[str, str], str]:
    """从教师任课表建立 (班级, 科目) → 教师姓名 映射（科目含「班主任」）。"""
    if teacher_df is None or teacher_df.empty:
        return {}
    teacher_map = {}
    cols = set(teacher_df.columns)
    if {"班级", "科目", "教师姓名"} <= cols:
        for _, row in teacher_df.iterrows():
            cls = str(row["班级"]).strip()
            subj = str(row["科目"]).strip()
            name = row.get("教师姓名")
            name = "" if pd.isna(name) else str(name).strip()
            if cls and subj:
                teacher_map[(cls, subj)] = name
    return teacher_map


def save_exam(db_path: str, exam_meta: Dict, result: Dict,
              parsed_data: Dict, source_file: str = "",
              exam_count: int = 2) -> int:
    """
    把一次分析结果写入数据库，返回新插入的 exam_id。

    exam_meta: {"exam_info": str, "exam_date": str, "cohort": str,
                "semester": str, "exam_name": str, "exam_count": int}
    exam_count: 2=两场考试对比（六项指标）；1=单次考试（首次考试，四项指标）。
    """
    init_db(db_path)
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.cursor()
        mode = exam_count if exam_count in (1, 2) else int(exam_meta.get("exam_count", 2) or 2)
        cur.execute(
            "INSERT INTO exams (exam_info, exam_date, cohort, semester, exam_name, source_file, exam_count) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (exam_meta.get("exam_info", ""), exam_meta.get("exam_date", ""),
             exam_meta.get("cohort", ""), exam_meta.get("semester", ""),
             exam_meta.get("exam_name", ""), source_file, mode))
        exam_id = cur.lastrowid
        campus_map = _build_class_campus_map(parsed_data)
        teacher_map = _build_teacher_map(parsed_data.get("教师任课表"))

        # ---- class_metrics ----
        subjects = [k for k in result.keys() if k not in ("教师得分表",)]
        for subject in subjects:
            subj = result[subject]
            if mode == 1:
                _save_class_metrics_single(cur, exam_id, campus_map, teacher_map,
                                           subject, subj)
            else:
                _save_class_metrics_dual(cur, exam_id, campus_map, teacher_map,
                                         subject, subj)

        # ---- teacher_scores ----
        teacher_rows = result.get("教师得分表", {}).get("教师得分表", [])
        for t in teacher_rows:
            if mode == 1:
                cur.execute(
                    "INSERT INTO teacher_scores (exam_id, 身份证号, 姓名, 科目, 校区, "
                    "均分, 优秀率, 合格率, 参考率, 总分) "
                    "VALUES (?,?,?,?,?,?,?,?,?,?)",
                    (exam_id, str(t.get("身份证号", "")), str(t.get("姓名", "")),
                     str(t.get("科目", "")), "" if pd.isna(t.get("校区")) else str(t.get("校区", "")),
                     _num(t.get("均分")),
                     _num(t.get("优秀率")), _num(t.get("合格率")), _num(t.get("参考率")),
                     _num(t.get("总分"))))
            else:
                cur.execute(
                    "INSERT INTO teacher_scores (exam_id, 身份证号, 姓名, 科目, 校区, "
                    "排名正负, T分差, 均分, 优秀人数, 出库人数, 参考率, 总分) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                    (exam_id, str(t.get("身份证号", "")), str(t.get("姓名", "")),
                     str(t.get("科目", "")), "" if pd.isna(t.get("校区")) else str(t.get("校区", "")),
                     _num(t.get("排名正负")), _num(t.get("T分差")), _num(t.get("均分")),
                     _num(t.get("优秀人数")),
                     _num(t.get("出库人数")), _num(t.get("参考率")), _num(t.get("总分"))))
        conn.commit()
        return exam_id
    finally:
        conn.close()


def _save_class_metrics_dual(cur, exam_id, campus_map, teacher_map, subject, subj):
    """两场考试对比模式：写入六项指标（均分1/均分2/T分差/排名正负/优秀人数差/出库人数差/参考率）。"""
    t_table = {r["班级"]: r for r in subj.get("T分差", {}).get("T分差表格", [])}
    avg_table = {r["班级"]: r for r in subj.get("均分", {}).get("均分表格", [])}
    rank_table = {r["班级"]: r for r in subj.get("排名正负", {}).get("排名正负表格", [])}
    exc_table = {r["班级"]: r for r in subj.get("优秀人数", {}).get("优秀人数表格", [])}
    qua_table = {r["班级"]: r for r in subj.get("出库人数", {}).get("出库人数表格", [])}
    ref_table = {r["班级"]: r for r in subj.get("班级参考率", [])}

    classes = set(t_table) | set(avg_table) | set(rank_table) | \
              set(exc_table) | set(qua_table) | set(ref_table)
    for cls in classes:
        t = t_table.get(cls, {})
        avg = avg_table.get(cls, {})
        rk = rank_table.get(cls, {})
        exc = exc_table.get(cls, {})
        qua = qua_table.get(cls, {})
        ref = ref_table.get(cls, {})
        if subject == "总分":
            teacher_name = teacher_map.get((cls, "班主任"), "")
        else:
            teacher_name = teacher_map.get((cls, subject), "")
        row = (
            exam_id,
            campus_map.get(cls, ""),
            cls, subject, teacher_name,
            _num(avg.get("第一场均分")), _num(avg.get("第二场均分")),
            _num(avg.get("均分")),
            _num(t.get("T分1")), _num(t.get("T分2")), _num(t.get("T分差")),
            _int(t.get("T分差排名")),
            _num(rk.get("线性排名1")), _num(rk.get("线性排名2")),
            _num(rk.get("正负")), _int(rk.get("排名（正负指标的排名）")),
            _int(exc.get("优秀人数1")), _int(exc.get("优秀人数2")), _int(exc.get("优秀人数差")),
            _int(qua.get("出库人数1")), _int(qua.get("出库人数2")), _int(qua.get("出库人数差")),
            _int(ref.get("总人数")), _int(ref.get("参考人数")), _num(ref.get("参考率")),
        )
        cur.execute(
            "INSERT INTO class_metrics (exam_id, campus, 班级, 学科, 教师姓名, 均分1, 均分2, 均分, "
            "T分1, T分2, T分差, T分差排名, 线性排名1, 线性排名2, 正负, 正负排名, "
            "优秀人数1, 优秀人数2, 优秀人数差, 出库人数1, 出库人数2, 出库人数差, "
            "总人数, 参考人数, 参考率) VALUES ("
            + ",".join(["?"] * 25) + ")", row)


def _save_class_metrics_single(cur, exam_id, campus_map, teacher_map, subject, subj):
    """单次考试模式：写入四项指标（均分/优秀率/合格率/参考率），排名仅作参考字段保留。"""
    avg_table = {r["班级"]: r for r in subj.get("均分排名", {}).get("均分排名表格", [])}
    exc_table = {r["班级"]: r for r in subj.get("优秀率", {}).get("优秀率表格", [])}
    qua_table = {r["班级"]: r for r in subj.get("合格率", {}).get("合格率表格", [])}
    ref_table = {r["班级"]: r for r in subj.get("班级参考率", [])}

    classes = set(avg_table) | set(exc_table) | set(qua_table) | set(ref_table)
    for cls in classes:
        avg = avg_table.get(cls, {})
        exc = exc_table.get(cls, {})
        qua = qua_table.get(cls, {})
        ref = ref_table.get(cls, {})
        if subject == "总分":
            teacher_name = teacher_map.get((cls, "班主任"), "")
        else:
            teacher_name = teacher_map.get((cls, subject), "")
        row = (
            exam_id,
            campus_map.get(cls, ""),
            cls, subject, teacher_name,
            _num(avg.get("均分")),
            _int(avg.get("排名")), _num(avg.get("线性排名")),
            _int(exc.get("优秀人数")), _num(exc.get("优秀率")),
            _int(qua.get("合格人数")), _num(qua.get("合格率")),
            _int(ref.get("总人数")), _int(ref.get("参考人数")), _num(ref.get("参考率")),
        )
        cur.execute(
            "INSERT INTO class_metrics (exam_id, campus, 班级, 学科, 教师姓名, 均分, "
            "排名, 线性排名, 优秀人数, 优秀率, 合格人数, 合格率, "
            "总人数, 参考人数, 参考率) VALUES ("
            + ",".join(["?"] * 15) + ")", row)


def _num(v):
    if v is None:
        return None
    try:
        f = float(v)
        return None if pd.isna(f) else f
    except (ValueError, TypeError):
        return None


def _int(v):
    n = _num(v)
    return int(n) if n is not None else None


# ---------------- 查询接口（供网页 API 调用） ----------------

def list_exams(db_path: str) -> List[Dict]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            "SELECT e.id, e.exam_info, e.exam_date, e.cohort, e.semester, e.exam_name, "
            "e.created_at, e.exam_count, "
            "(SELECT COUNT(DISTINCT 班级) FROM class_metrics c WHERE c.exam_id=e.id) AS class_count, "
            "(SELECT COUNT(DISTINCT 学科) FROM class_metrics c WHERE c.exam_id=e.id) AS subject_count, "
            "(SELECT COUNT(*) FROM teacher_scores t WHERE t.exam_id=e.id) AS teacher_count "
            "FROM exams e ORDER BY e.id DESC").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_class_metrics(db_path: str, exam_id: int, subject: Optional[str] = None) -> List[Dict]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        sql = "SELECT * FROM class_metrics WHERE exam_id=?"
        params = [exam_id]
        if subject:
            sql += " AND 学科=?"
            params.append(subject)
        sql += " ORDER BY 学科, 班级"
        rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_teacher_scores(db_path: str, exam_id: int) -> List[Dict]:
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute(
            "SELECT * FROM teacher_scores WHERE exam_id=? ORDER BY 科目, 总分 DESC",
            (exam_id,)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def get_campus_agg(db_path: str, exam_id: int, subject: Optional[str] = None) -> List[Dict]:
    """
    按校区聚合班级指标。

    单次考试（exam_count=1）没有均分2 / T分差 / 人数差，
    因此平均第二场均分用 COALESCE 回落到「均分」，排序同样按回落后的值。
    """
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        sql = ("SELECT campus AS 校区, COUNT(DISTINCT 班级) AS 班级数, "
               "AVG(均分2) AS 平均第二场均分, AVG(均分) AS 平均均分, AVG(T分差) AS 平均T分差, "
               "SUM(优秀人数差) AS 优秀人数差合计, SUM(出库人数差) AS 出库人数差合计, "
               "AVG(参考率) AS 平均参考率, "
               "AVG(优秀率) AS 平均优秀率, AVG(合格率) AS 平均合格率 "
               "FROM class_metrics WHERE exam_id=?")
        params = [exam_id]
        if subject:
            sql += " AND 学科=?"
            params.append(subject)
        sql += " GROUP BY campus ORDER BY COALESCE(平均均分, 平均第二场均分) DESC"
        rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def compare_classes(db_path: str, classes: List[str], subject: str,
                    exam_ids: List[int]) -> Dict:
    """
    跨考试对比指定班级在某学科上的走势。
    返回 {exam_info_list: [...],
          series: {班级: [{exam_info, 均分1, 均分2, 均分, T分差}, ...]}}
    """
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        placeholders = ",".join("?" * len(exam_ids))
        rows = conn.execute(
            f"SELECT id, exam_info FROM exams WHERE id IN ({placeholders})",
            exam_ids).fetchall()
        exam_info_map = {r["id"]: r["exam_info"] for r in rows}
        exam_order = [r["id"] for r in rows]

        series = {c: [] for c in classes}
        for eid in exam_order:
            einfo = exam_info_map[eid]
            for c in classes:
                row = conn.execute(
                    "SELECT 均分1, 均分2, 均分, T分差, 排名, 优秀率, 合格率 FROM class_metrics "
                    "WHERE exam_id=? AND 班级=? AND 学科=?",
                    (eid, c, subject)).fetchone()
                if row:
                    series[c].append({
                        "exam_info": einfo,
                        "均分1": row["均分1"], "均分2": row["均分2"],
                        "均分": row["均分"], "T分差": row["T分差"],
                        "排名": row["排名"], "优秀率": row["优秀率"], "合格率": row["合格率"]})
        return {"exam_info_list": [exam_info_map[e] for e in exam_order], "series": series}
    finally:
        conn.close()


def get_subjects(db_path: str, exam_id: int) -> List[str]:
    conn = sqlite3.connect(db_path)
    try:
        rows = conn.execute(
            "SELECT DISTINCT 学科 FROM class_metrics WHERE exam_id=? ORDER BY 学科",
            (exam_id,)).fetchall()
        return [r[0] for r in rows]
    finally:
        conn.close()


def get_classes(db_path: str, exam_id: int, subject: Optional[str] = None) -> List[str]:
    conn = sqlite3.connect(db_path)
    try:
        sql = "SELECT DISTINCT 班级 FROM class_metrics WHERE exam_id=?"
        params = [exam_id]
        if subject:
            sql += " AND 学科=?"
            params.append(subject)
        rows = conn.execute(sql, params).fetchall()
        return [r[0] for r in rows]
    finally:
        conn.close()


def delete_exam(db_path: str, exam_id: int) -> None:
    """删除一次考试及其全部明细（班级指标、教师得分），级联清理。"""
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("DELETE FROM teacher_scores WHERE exam_id=?", (exam_id,))
        conn.execute("DELETE FROM class_metrics WHERE exam_id=?", (exam_id,))
        conn.execute("DELETE FROM exams WHERE id=?", (exam_id,))
        conn.commit()
    finally:
        conn.close()


# ---------------- 按新口径重算已落库的历史考试 ----------------

# 与 score_analysis.py 的 CLI 默认值保持一致
DEFAULT_WEIGHTS = {
    "排名正负": 20.0,
    "T分差": 20.0,
    "均分": 20.0,
    "优秀人数": 15.0,
    "出库人数": 15.0,
    "参考率": 10.0,
}

# 单次考试（首次考试）四项指标的默认权重
# 「均分与排名」为同一项指标，排名不单独计分，均分项合并原「均分 25 + 排名 15」的权重
SINGLE_DEFAULT_WEIGHTS = {
    "均分": 60.0,
    "优秀率": 15.0,
    "合格率": 15.0,
    "参考率": 10.0,
}

# 六项指标在库中的列名（用于读取与写回）
WEIGHT_COLUMNS = list(DEFAULT_WEIGHTS.keys())


def _minmax_score(series: pd.Series, weight: float) -> pd.Series:
    """min-max 归一到 0~100；权重为负时按「越低越好」反向归一。"""
    col_min, col_max = series.min(), series.max()
    if col_max != col_min:
        transformed = (series - col_min) / (col_max - col_min) * 100
    else:
        transformed = pd.Series(50.0, index=series.index)
    if weight < 0:
        transformed = 100 - transformed
    return transformed


def rebuild_exam(db_path: str, exam_id: int,
                 weights: Optional[Dict[str, float]] = None) -> Dict:
    """
    按新口径重算某次考试的均分指标与教师总分（不重新读取原始 Excel）。

    适用场景：技能新增「均分」指标、或权重调整后，让已落库的历史记录与新口径
    保持一致，从而在跨考试走势、校区对比与教师得分排名中口径统一。

    重算内容：
      1. class_metrics.均分 = 均分2（两场考试的结果评价口径，直接用第二场班级均分，不做差）；
         单次考试的「均分」本身就是班级均分，无需重算，跳过本步；
      2. teacher_scores.均分 = 按科目聚合的班级均分
         （学科教师取所教班级的科目均分；班主任取所带班级总分均分的均值）；
         仅两场考试执行，单次考试跳过（其均分落库时已是教师聚合值）；
      3. teacher_scores.总分 = 各项指标 min-max 归一后按权重加权（按科目分组归一，
         班主任视为独立一组），与 score_analysis.py 口径一致。
         两场考试用六项权重，单次考试用四项权重（均分/优秀率/合格率/参考率）。

    返回 {"class_rows": n, "teacher_rows": m, "mode": "dual"|"single"} 形式的统计信息。
    """
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    try:
        # 0. 判定考试模式：优先读 exams.exam_count，缺失时按是否存在单次指标列推断
        exam_row = conn.execute(
            "SELECT exam_count FROM exams WHERE id=?", (exam_id,)).fetchone()
        mode = (exam_row["exam_count"] if exam_row and exam_row["exam_count"] else 2)
        if mode not in (1, 2):
            probe = conn.execute(
                "SELECT COUNT(*) AS n FROM class_metrics "
                "WHERE exam_id=? AND (排名 IS NOT NULL OR 优秀率 IS NOT NULL)",
                (exam_id,)).fetchone()
            mode = 1 if (probe and probe["n"] > 0) else 2

        if mode == 1:
            weights = dict(weights or SINGLE_DEFAULT_WEIGHTS)
        else:
            weights = dict(weights or DEFAULT_WEIGHTS)

        # 1. 班级均分 —— 两场考试取 均分2；单次考试已是班级均分，跳过
        if mode == 2:
            rows = conn.execute(
                "SELECT id, 均分2 FROM class_metrics WHERE exam_id=?",
                (exam_id,)).fetchall()
            for r in rows:
                val = None
                if r["均分2"] is not None:
                    val = round(float(r["均分2"]), 2)
                conn.execute("UPDATE class_metrics SET 均分=? WHERE id=?", (val, r["id"]))
            class_rows = len(rows)
        else:
            class_rows = conn.execute(
                "SELECT COUNT(*) AS n FROM class_metrics WHERE exam_id=?",
                (exam_id,)).fetchone()["n"]
        conn.commit()

        # 2. 教师均分 —— 学科教师取本班本科均分，班主任取所带班级总分均分均值
        #    注意：单次考试（mode==1）跳过本步。teacher_scores.均分 在落库时已是
        #    「教师聚合后的值」（多班任教取均值），本身就是正确口径；
        #    若此处再从班级级 class_metrics 反查，遇到「同班同科多位老师」时
        #    class_metrics 只记录了其中一位，反查会查不到并把均分错置为 NULL，
        #    导致总分被算错。此处仅对两场考试（mode==2）执行反查，
        #    且查不到时保留原值，绝不写入 NULL。
        teacher_rows = conn.execute(
            "SELECT id, 姓名, 科目, 校区 FROM teacher_scores WHERE exam_id=?",
            (exam_id,)).fetchall()
        if mode == 2:
            for t in teacher_rows:
                subject = "总分" if t["科目"] == "班主任" else t["科目"]
                campus_clause, campus_params = "", []
                if t["校区"]:
                    campus_clause = " AND campus=?"
                    campus_params = [t["校区"]]
                avg_row = conn.execute(
                    "SELECT AVG(均分) AS m FROM class_metrics "
                    "WHERE exam_id=? AND 学科=? AND 教师姓名=?" + campus_clause,
                    (exam_id, subject, t["姓名"], *campus_params)).fetchone()
                avg_val = avg_row["m"] if avg_row and avg_row["m"] is not None else None
                if avg_val is None:
                    continue  # 反查不到 → 保留原值，不清空
                conn.execute("UPDATE teacher_scores SET 均分=? WHERE id=?",
                             (round(float(avg_val), 2), t["id"]))
            conn.commit()

        # 3. 教师总分 —— 各指标 min-max 归一后按权重加权（同科目组内归一）
        if mode == 1:
            metric_cols = ["姓名", "科目", "均分", "优秀率", "合格率", "参考率"]
        else:
            metric_cols = ["姓名", "科目", "排名正负", "T分差", "均分", "优秀人数", "出库人数", "参考率"]
        df = pd.read_sql_query(
            f"SELECT id, {', '.join(metric_cols)} FROM teacher_scores WHERE exam_id=?",
            conn, params=(exam_id,))
        if not df.empty:
            total = pd.Series(0.0, index=df.index)
            for _, group_idx in df.groupby("科目", dropna=False).groups.items():
                for col, weight in weights.items():
                    if col not in df.columns:
                        continue
                    series = pd.to_numeric(df.loc[group_idx, col], errors="coerce")
                    # 缺项按该指标中位数补齐，避免个别缺失拉偏总分
                    if series.isna().any():
                        series = series.fillna(series.median())
                    if series.isna().all():
                        continue
                    total.loc[group_idx] += _minmax_score(series, weight) * (abs(weight) / 100)
            for idx, score in zip(df["id"], total.round(2)):
                conn.execute("UPDATE teacher_scores SET 总分=? WHERE id=?", (float(score), int(idx)))
            conn.commit()

        return {"class_rows": class_rows, "teacher_rows": len(teacher_rows),
                "mode": "single" if mode == 1 else "dual"}
    finally:
        conn.close()


if __name__ == "__main__":
    # 简单自测：在技能目录 data/ 下建库并列出
    here = os.path.dirname(os.path.abspath(__file__))
    db = os.path.join(here, "data", "analysis.db")
    init_db(db)
    print("DB initialized at", db)
    print("Exams:", list_exams(db))

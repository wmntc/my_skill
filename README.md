# my_skill

个人技能集仓库，收录 4 个自建 WorkBuddy Skill，当前版本 **1.3**。

| 技能 | 目录 | 说明 |
|---|---|---|
| 实验课程设计大师 | [`skills/实验课程设计大师`](skills/实验课程设计大师) | 多模型教学设计生成器，一次交付 HTML 教案 + A4 PDF + Markdown，支持 BOPPPS / 5E / 对分课堂 PAD / GRR / UbD / ADDIE / PBL·CBL 七种模型 |
| 成果案例制作器 | [`skills/achievement-case-maker`](skills/achievement-case-maker) | 把教师零散材料自动整理成结构统一的成果案例成品文件夹 |
| 呼市实验中学成绩分析 | [`skills/呼市实验中学成绩分析`](skills/呼市实验中学成绩分析) | 班级指标与教师绩效算分，三口径：两场考试对比（六项指标）／首次考试单场（四项指标）／混合口径（逐科目分流），纯 Excel 输入输出 |
| 公文格式排版 | [`skills/gov-doc-formatter`](skills/gov-doc-formatter) | 文本转 GB/T 9704-2012 标准公文 .docx，内嵌仿宋/楷体/方正小标宋字体 |

## 目录结构

```
my_skill/
├── README.md
├── .gitignore
├── .codebuddy-plugin/plugin.json
└── skills/
    ├── 实验课程设计大师/       SKILL.md + scripts/ + references/ + assets/
    ├── achievement-case-maker/ SKILL.md + scripts/ + references/
    ├── 呼市实验中学成绩分析/    SKILL.md + scripts/
    └── gov-doc-formatter/      SKILL.md + scripts/ + references/ + assets/fonts/
```

## 安装

将所需技能目录复制到用户级技能目录 `~/.workbuddy/skills/` 即可：

```bash
git clone https://github.com/wmntc/my_skill.git
cp -r my_skill/skills/<技能名> ~/.workbuddy/skills/
```

技能依赖 Python 3.8+ 与常见数据分析库（`pandas`、`openpyxl` 等），具体见各技能的 `SKILL.md`。

## 版本

当前版本 **1.3**，对应 Git 标签 [`v1.3`](https://github.com/wmntc/my_skill/releases/tag/v1.3)。

各技能的版本号记录在对应 `SKILL.md` 的 frontmatter `version` 字段中。

### v1.3 变更

**呼市实验中学成绩分析**（`1.2` → `1.3`）移除网页与数据库依赖，回归纯 Excel 工具：

- 删除 `scripts/db_store.py`、`scripts/web_app.py`、`web/index.html` 及全部落库/服务代码与 CLI 参数（`--db/--no-db/--serve/--port`）。
- 保留三种口径的完整算分与多工作表导出能力（首次考试四项指标、两场对比六项指标、混合口径逐科目分流）。
- 技能目录只含 `SKILL.md` 与 `scripts/score_analysis.py`；依赖仅 `pandas` + `openpyxl`。
- 已落库的 `data/analysis.db` 不会被代码访问，也不会被技能删除。

### v1.2 变更

**呼市实验中学成绩分析**（`1.1` → `1.2`）新增混合口径（逐科目分流）：

- 有基准成绩的科目走两场对比六项指标；无基准成绩的科目（如初二新开物理）走首次考试四项指标，且不计入总分。
- 总分按「含有基准成绩的科目之和」重算，弃用掺入无基准科目的原生总分列。
- 支持 `--subject-mode` 手工覆盖个别科目口径，以及 5 个 `single-*` 专用参数。
- 三种口径的结果 Excel 分别含 5 / 7 / 9 张工作表。

### v1.1 变更

**呼市实验中学成绩分析**（`1.0` → `1.1`）新增「首次考试」单场成绩口径：

- 新增双口径自动分流——成绩表含 `2` 结尾科目列走两场对比（六项指标），否则走首次考试
  （四项指标：均分、优秀率、合格率、参考率）；也可用 `--exam-count 1/2` 显式指定。
- 「均分与排名」按同一项指标处理，排名不再单列计分，仅作参考字段保留。
- 首次考试权重：均分 60 / 优秀率 15 / 合格率 15 / 参考率 10；两场对比权重不变。
- 优秀率、合格率分母为参考人数，优秀线与合格线按全年级前 20% / 前 80% 名次划线。
- 查询页面按考试口径自动切换图表与表格列，首次考试带绿色标签。
- 修复 `rebuild_exam` 在「同班同科多位任课教师」场景下会把教师均分清空、导致总分算错的问题。

## 说明

- `skills/呼市实验中学成绩分析/` 为纯 Excel 工具，不写数据库、不启服务、不产生需要联网的依赖。
- `gov-doc-formatter/assets/fonts/` 收录了三套公文字体文件，用于在 docx 中嵌入字体，保证任意设备打开不缺字。

# my_skill

个人技能集仓库，收录 4 个自建 WorkBuddy Skill，当前版本 **1.0**。

| 技能 | 目录 | 说明 |
|---|---|---|
| 实验课程设计大师 | [`skills/实验课程设计大师`](skills/实验课程设计大师) | 多模型教学设计生成器，一次交付 HTML 教案 + A4 PDF + Markdown，支持 BOPPPS / 5E / 对分课堂 PAD / GRR / UbD / ADDIE / PBL·CBL 七种模型 |
| 成果案例制作器 | [`skills/achievement-case-maker`](skills/achievement-case-maker) | 把教师零散材料自动整理成结构统一的成果案例成品文件夹 |
| 呼市实验中学成绩分析 | [`skills/呼市实验中学成绩分析`](skills/呼市实验中学成绩分析) | 两场考试班级指标与教师绩效算分，含 SQLite 持久化与本地可视化查询页 |
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
    ├── 呼市实验中学成绩分析/    SKILL.md + scripts/ + web/
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

当前版本 **1.0**，对应 Git 标签 [`v1.0`](https://github.com/wmntc/my_skill/releases/tag/v1.0)。

各技能的版本号记录在对应 `SKILL.md` 的 frontmatter `version` 字段中。

## 说明

- `skills/呼市实验中学成绩分析/data/` 为运行时生成的 SQLite 数据库（含教师姓名、身份证号等个人信息），**不入库**，由技能首次运行时自动创建。
- `gov-doc-formatter/assets/fonts/` 收录了三套公文字体文件，用于在 docx 中嵌入字体，保证任意设备打开不缺字。

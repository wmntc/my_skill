# spec.json 字段说明

## 版头

| 字段 | 类型 | 说明 |
|------|------|------|
| `org_mark` | string | 发文机关标志（红头）文字，可含 `\n` 分多行；自动缩字号保证单行 |
| `org_mark_size` | string | 红头字号，默认「初号」，超宽自动缩小 |
| `doc_number` | string | 发文字号，如 `呼实〔2026〕8 号`（六角括号，不编虚位） |
| `direction` | string | `up` 上行文 / `down` 下行文（默认）。上行文发文字号居左空一字并与签发人同行 |
| `signer` | string \| string[] | 签发人，姓名用三号楷体 |
| `copy_number` | string/number | 份号，自动补齐 6 位 |
| `secrecy` | string | 密级和保密期限，如 `秘密★10年` |
| `urgency` | string | `特急` / `加急` |
| `red_rule` | bool | 是否画红色分隔线（默认 true） |

## 主体

| 字段 | 类型 | 说明 |
|------|------|------|
| `title` | string | 文件标题，二号方正小标宋居中，超长自动梯形断行 |
| `title_max_lines` | number | 标题最多行数，默认 3 |
| `main_recipient` | string | 主送机关，顶格，末尾自带全角冒号 |
| `body` | array | 正文块列表，见下 |
| `attachments` | array | 附件名称字符串数组，或 `{name, body}` 带正文的对象 |
| `sign_org` | string | 发文机关署名 |
| `sign_date` | string | 成文日期，如 `2026年10月20日`（不编虚位） |
| `seal` | bool | 是否有印章（默认 true，但需提供 seal_image 才绘制） |
| `seal_image` | string | 印章图片路径（PNG 透明底最佳），浮于署名与日期之上 |
| `seal_size_mm` | number | 印章直径，默认 42mm |
| `note` | string | 附注（不加括号，脚本自动加圆括号），居左空二字 |

## 版记

| 字段 | 类型 | 说明 |
|------|------|------|
| `cc` | string[] | 抄送机关列表，脚本加「抄送：」与句号 |
| `print_org` | string | 印发机关 |
| `print_date` | string | 印发日期，脚本自动加「印发」 |

## body 块类型

```json
{ "t": "p",     "c": "普通段落，首行空二字" }
{ "t": "h1",    "c": "一、……" }        // 黑体三号
{ "t": "h2",    "c": "（一）……" }      // 楷体三号加粗
{ "t": "h3",    "c": "1.……" }          // 仿宋三号加粗
{ "t": "h4",    "c": "（1）……" }       // 仿宋三号
{ "t": "quote", "c": "引文，左右各空二字，楷体" }
{ "t": "blank", "n": 1 }
{ "t": "table", "rows": [["表头","表头"],["a","b"]], "header": true }
```

字符串元素等价于 `{"t":"p","c":"..."}`。

## 最小示例

```json
{
  "title": "关于××××的通知",
  "main_recipient": "各单位：",
  "body": [{ "t": "p", "c": "正文……" }],
  "sign_org": "××单位",
  "sign_date": "2026年10月20日"
}
```

无 `org_mark`/`doc_number` 时不排版头（内部材料样式）。

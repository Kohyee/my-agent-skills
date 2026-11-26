---
name: English-learning-template
description: "Create a document using the English-learning-template template and its retained reference file. Use when the user selects this template, names English-learning-template, or explicitly invokes $English-learning-template. 按用户最新 ENGLISH WORKSHEET 样本制作每日英语学习 Word；默认首页为 DAY N · ENGLISH WORKSHEET，文件命名 YYYYMMDD_English worksheet_DayN.docx，保留五列词汇表、两列易混词独立公式、公式外说明与单表双向默写。"
---

# English-learning-template

Create a document from this template. Keep the reference file unchanged.

## Workflow

1. Read `artifact-template.json` and resolve its paths relative to this skill directory.
2. Load [@documents](plugin://documents@openai-primary-runtime) and invoke its reference/template workflow with the retained file.
3. Treat the user's prompt and available sources as the content input. Do not invent facts merely to fill a template slot.
4. Clone or import the reference instead of replacing its visual system with generic defaults.
5. Run `scripts/normalize_table_borders.py <working.docx> --output <final.docx>` on the working copy, then `scripts/validate_lesson.py <final.docx>`. Render and verify every page, then return the final artifact.

## Fidelity

Preserve page setup, sections, styles, lists, tables, headers, footers, and recurring page elements.

User instructions control requested content and explicit deviations. The retained reference controls layout and formatting where the user has not requested a change.

## English-learning-template 用户固定规范

这是用户指定的每日英语学习 Word 模板。显示名称、调用标识和技能目录名统一为 **English-learning-template**；调用时使用 `$English-learning-template`。

1. 先阅读 `references/design-spec.md`、`references/design-spec.json`，再读取 `assets/reference.docx`。当前唯一版式基准是用户最新提供的 `EnglishLearning_Day1_20261123.docx`，不得回退到更早样本。保留的参考文件不得修改。
2. 文件名固定为 `YYYYMMDD_English worksheet_Day{N}.docx`，例如 `20261124_English worksheet_Day2.docx`。日期放在文件名最前面，使用八位数字、不加连字符；固定短语 English worksheet 中间保留一个半角空格，English 首字母大写、worksheet 全小写。首页标题为 `DAY {N}  ·  ENGLISH WORKSHEET`，26 pt、加粗居中；Date 继续采用 `Date: 24 November 2026` 形式。文件名、DAY 与 Date 必须一致。不得只修改旧课程标题来冒充新课程。
3. 所有表格选择“所有边框”：外框、横向内框和纵向内框均为明确的 `#000000` 黑色单实线，不能使用自动色、灰色、白色或主题色覆盖。词汇表序号1单元格的左边框也必须是黑线；同时核查每个单元格和行级边框覆盖。页面、Times New Roman/宋体字体系统、灰色表头及序号底纹、页脚和列宽沿用参考。词汇表为五列，词头与音标同格；默认 20 个主词，每词 3–4 个重要搭配，均附中文，例句双语。
4. COMMONLY CONFUSED WORDS 保持两列。释义概括、辨析、词形拓展分别为独立原生公式，反义拓展按需单独设置；每个公式单独一段。标签使用 32 pt 最小列宽、4 pt 间距，顶层大括号对齐。不得把多个类别合并成一个大公式。
5. 同一词的句型归入该词后的内层大括号。补充词、注意事项及例子用普通文字段落，紧跟相关公式下方；类似 influence 的额外近义补充不混入当前标题词组的主辨析括号。
6. 双向默写只用一张五列表格，右侧英文提示按主词倒序排列，两列答案留空。沿用语法、阅读、翻译、写作、自查的结构与默认题量，不自行添加答案或额外章节。
7. 用户只要求下一天、更换内容时，直接沿用本模板，不重复确认格式。用户明确指定的最新要求优先；局部修改时保留其他内容。默认仅交付命名正确的最终 Word 下载链接。

## 检查与预览范围

运行 `scripts/validate_lesson.py <生成的文档>` 核对紧凑日期命名、ENGLISH WORKSHEET 主标题、20 词五列表格、双语搭配、两列独立公式、单张双向默写表、所有表格黑色实线全边框及单元格级覆盖（包括序号1左边框）。该检查不代替实际文档的逐页渲染；具备原生排版引擎时应使用文档技能正式渲染器检查所有页面。

`assets/preview.png` 为当前 `assets/reference.docx` 经 WPS 原生导出、Poppler 渲染的真实首页。已检查当前模板的原生分页；新练习单仍须逐页渲染验收。受限环境查不到排版引擎时，应核查真实 Windows 用户的 `kwps.Application`；使用后台只读导出，不能修改用户正在编辑的文档。始终以保留的 `assets/reference.docx` 为版式依据。

当前可读规范为 `assets/ENGLISH_WORKSHEET_DesignSpec_v2.2.docx`；规范正文和参数分别为 `references/design-spec.md`、`references/design-spec.json`，其中记录当前参考 SHA-256。旧版规范附件仅保留为历史，不作为默认规则。


名称更新：2026-11-25 按用户要求统一名称。当前 Word 规范附件为v2.2；旧版保留为历史。技能名称和黑色全边框规则以本文件及 references/design-spec 为准。

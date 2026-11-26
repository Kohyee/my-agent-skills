# 操作流程

脚本入口 `.learning-manager/skill/scripts/manager.py`。本机打包解释器：`C:\Users\Kohye\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe`；依赖lxml也来自该运行时。优先使用可用的load_workspace_dependencies；本次该工具未暴露，通过runtime.json验证路径。换机器时重新解析运行时，不假定C盘用户路径不变。PowerShell设置`$env:PYTHONIOENCODING='utf-8'`防止中文/勾选符号输出编码错误。

示例命令在项目根执行，日期按用户意图替换；JSON由Codex生成在`.learning-manager/requests`并保留，可直接调用Manager类避免临时输入文件。

```
<python> .learning-manager/skill/scripts/manager.py context --date 2026-11-25 --day 2
<python> .learning-manager/skill/scripts/manager.py register --input <JSON绝对路径>
<python> .learning-manager/skill/scripts/manager.py grade --input <JSON绝对路径>
<python> .learning-manager/skill/scripts/manager.py adjust --input <JSON绝对路径>
<python> .learning-manager/skill/scripts/manager.py due --date 2026-11-26
<python> .learning-manager/skill/scripts/manager.py weekly --start 2026-11-23
<python> .learning-manager/skill/scripts/manager.py plan-week --input <JSON绝对路径>
<python> .learning-manager/skill/scripts/manager.py progress
<python> .learning-manager/skill/scripts/manager.py points
<python> .learning-manager/skill/scripts/manager.py record-points --input <JSON绝对路径>
<python> .learning-manager/skill/scripts/manager.py vocabulary
<python> .learning-manager/skill/scripts/manager.py summarize-vocabulary
<python> .learning-manager/skill/scripts/manager.py audit
```

## 每日出卷

1. context已经带出原档案全文、最新周评估、日期对应计划、近期作答和到期复习。读取对应真实文件，检查外部变更和缺失证据。已有同日试卷优先返回，不另建同Day记录。
2. 调用`C:\Users\Kohye\.codex\skills\English-learning-template\SKILL.md`，依次读取其manifest、spec、保留参考；调用documents模板流程。管理Skill只传内容安排，不写另一套通用出卷器。
3. 教学内容以当天计划和评估证据为准；将到期错误嵌入原练习。例如将翻译第1题用于hope to do，语法第2题用于a/an，不额外增加章节/时长。优先重复错误，但15分钟内不能复习全部到期项时依次顺延并说明未覆盖数。
4. 形成完整题目清单、参考答案供批改时核对（不写进学生试卷），标明reading材料来源。试卷中不得出现答案页；模板复习词仍保留20主词表，word_mode=review。
5. 原技能结构检查通过，再运行documents正式renderer逐页检查。当前运行时runtime.json的libreOfficeVersion=null，Word COM未注册；若仍不可渲染，只保存draft并说明未完成视觉验收，不能写qa.visual=passed。
6. register输入含date/day/file/blueprint/word_mode/review_question_map/qa。外部候选文件复制到唯一03目录并在覆盖前备份；已有作答的题目禁止替换。

register示例字段：

```json
{"date":"2026-11-25","day":2,"file":"E:/英语/work/20261125_English worksheet_Day2.docx","word_mode":"new","blueprint":[{"id":"grammar.01","prompt":"实际题干","points":1}],"review_question_map":{"grammar.article":["grammar.01"]},"qa":{"visual":"unavailable","reason":"本机没有原生Word渲染引擎"}}
```

这是字段示例，实际blueprint必须逐题完整，不能只有这一题。新词模式必须与既有所有课程词头去重。周末20词全部来自已登记旧词。

## 批改

先读真实试卷及blueprint。收到未登记的纸质作业时先建立真实题目对应关系，不能推测它就是今天生成的卷。文字原样存raw_text，文档读取完整表格/公式/题号；PDF或图片按页查看，逐题转录。对无法辨认题号/笔迹发起必要确认，并继续批改清楚部分。

逐题保留原答案、给正确/可接受答案与讲解；错误提供最小修改和1—2道迁移练习。写作同时说明结构与自然表达，避免改写掉原始证据。不得将单纯润色偏好当语法错误。

```json
{"date":"2026-11-25","raw_text":"I hope improve my English.","source_files":[],"knowledge":[],"items":[{"id":"translation.01","original_answer":"I hope improve my English.","result":"incorrect","score":0,"expected_answer":"I hope to improve my English.","explanation":"hope后接to加动词原形。","checks":[{"knowledge_id":"grammar.hope-to","success":false,"independent":true,"mode":"review"}],"issues":[{"knowledge_id":"grammar.hope-to","reason":"漏to","explanation":"hope to do表达希望做某事。","practice":"不看例句写：我希望明年通过考试。"}]}]}
```

同样必须覆盖实际所有题；其余未提交题用pending和score:null。grade生成每日反馈并重建错题库、已生成周评估。再audit，报告待确认题和下次重点。

## 调整和未完成

调整同一天四训练格，保留原格布局和可点击复选框。四格中用`15′`或`15分钟`明确分钟数，合计等于minutes。只有已知的实际执行才填actual；“没学”可用0分钟与0完成，“没做完”不能自行假定做了几分钟。

```json
{"date":"2026-11-25","slots":["词汇15′：20新词，重点词主动回忆","语法20′：hope/want to do改错","听力10′：用户自备六级长对话","写作15′：原因段，复现hope to do"],"minutes":60,"status":"partial","actual":{"minutes":35,"content":"完成词汇和语法，听力写作未做","completed_tasks":2,"assigned_tasks":4},"reason":"用户报告今天只完成前两项"}
```

`changes`数组可原子更新多个日期（例如下一周）。不跨越completed冻结记录，不改历史实际内容。返回Word变更后执行原包部件/表格布局检查和渲染；渲染不可用时如实说明范围。

## 周评估与下一周

weekly以周一为start；周中总结也可执行，报告说明本周尚未结束、不把未来天当缺勤。保存每周评估并更新原档案管理区。初始档案正文不动，不凭题量估听说分数。

下一周先读已有周表与本周真实数据。存在下周计划时只做有证据支持的调整，沿用其余安排；没有证据时保留原计划，并把未评估项目列为待诊断。不要为了“自适应”随意增减每天时长。

plan-week需要week_start/reason以及7个含date/slots/minutes的days；若原表已覆盖则走adjust。超出现有计划才将第一周的表格结构克隆追加到**同一份**Word，提供goal/criteria，且要求上一周评估存在。保留原合并格、列宽和交互复选框；不另建一套周计划。新周的原计划写入original后永久保留。

## 记录重点

用户主动标记重要内容时按[分类重点流程](key-points.md)处理。先读points及现有Word，归类、保留原文、补双语例句与必要提醒，复用既有知识点id。record-points事务更新同一份重点文件，重复输入不重复收录。收录不改能力状态、错题次数或完成量。更新Word须逐页渲染验收，再运行audit。

## 单词汇总

按[新增单词汇总流程](vocabulary-summary.md)执行。register自动同步独立汇总Word，单独请求可用summarize-vocabulary补汇总或刷新；重复执行不增词。保持重点Word独立，二者统一在06_Learning Notes。逐页渲染本次变更的汇总，核对日期/Day和实际新增数量，收录不等于已完成或已掌握。

## 收尾

audit校验所有当前视图、原始作答证据、试卷哈希、唯一Day及事务备份。若发现外部修改先读取差异，不能用重新导出掩盖冲突。事务prepared/孤儿锁用recover，不手动删锁。

最终用学习者能理解的中文简报本次文件变化、实际错误（与系统技术问题区分）和下一次训练内容。开发验收不能把模拟结果当成用户进步。

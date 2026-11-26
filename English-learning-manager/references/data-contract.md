# 数据规范 V1.0

## 路径与唯一性

一个项目只使用 `.learning-manager/state.json`。schema_version=1；日期使用Asia/Shanghai的ISO日期，时间带+08:00。日期为lesson/submission主键，Day正整数且全局唯一；既有Day1的存在不代表作业完成。明确改了日期且与已登记Day冲突时先让用户确定对应关系。

路径由config读取：保留已有`03_Eglish worksheet`。不要按用户需求清单的理想拼写新建第二个03目录。所有输出限制在项目内。初始档案和原周表字节备份在baseline，原技能和参考不写入。

plans：date、table、day_index定位同一份周表；original永久保留；current为四训练格；minutes为当前计划时间；actual保存minutes/content/completed_tasks/assigned_tasks；history逐次保存before/after/reason/at。状态planned、unconfirmed、in_progress、completed、partial、missed。接入前未知执行记录使用unconfirmed，不当成缺勤。已完成任务冻结，批改修订单独进行。

lessons：date/day/path/sha256/words/word_mode/blueprint/review_question_map/qa。blueprint是实际题目清单，题号唯一、满分正数。双向默写分别用vocab.en.01和vocab.zh.01；语法grammar.01、阅读reading.01、翻译translation.01、写作writing.01。写作可按整道题评分，并在issues内列多个知识点。不得漏登记实际要求作答的题。

qa.visual=passed须记录真实渲染目录和页数；unavailable只能draft。结构检查直接加载原Worksheet技能的validate_lesson，不复制检查器。复习问题用`review_question_map: {"grammar.hope-to":["translation.01"]}`，必须对应实际题号。复习是否通过仍以作答checks为准。

## 作答与批改

submissions以date存当前版本，revision递增。原始文字存raw_text；本地原文件在source_files列出，脚本按哈希归档至evidence。只有对话图片无本地路径时，在source_refs注明当前对话附件及题号，保留逐题转录；不得伪称归档了不可访问的文件。每题必须有original_answer，转录须保留错误拼写和语序；无法辨认用null或保留可辨认片段。

批改items必须覆盖blueprint每题。result：correct/incorrect/partial/pending/unanswered。pending的score=null，必须pending_reason，不计分、不计错、不形成知识点检测；确认空答unanswered=0分，不能据此认定知识点未掌握。只有未提交照片部分不能记空答。其他题须score/expected_answer/explanation；错误题另有issues和checks。

知识点用稳定id合并重复错误：vocabulary.affect-attract、grammar.article、grammar.hope-to、grammar.agreement、spelling.study等。避免因为不同描述建立近义重复id；先读现有knowledge。新增知识点需label和category，分类限vocabulary/grammar/reading/writing/listening/speaking/spelling/collocation。

checks记录knowledge_id、success布尔、independent布尔、mode（practice/review）。同一题一个知识点只能检测一次；失败check与issue一一对应。issues记录knowledge_id/reason/explanation/practice，可补最小修改、自然表达、句子成分等。多项错误不能把整题重复计分；一个知识点在不同真实题目中再次失误可各计一次。

## 统计口径与五种状态

所有聚合从baseline_errors和**每个日期最新版本**的submissions重建，历史修订不累加。初始档案中V1、V2客观错题及Lesson01三句以source保留，发生日期未知不补造；主观题完整原文仍保存在初始档案，不重复编造成每日作业。初始摸底的23/45、16/22不并入新周统计，也不以两轮难度不同的成绩差认定进步。

error_count：当前有效的失败检测次数＋初始已记录次数；last_detected为最后一次可判定检测日期（成功也属于检测），last_reviewed为实际复习日期。没有精确历史检测日期显示未知。错误分布按失败知识点次数计，可能大于错误题数，报告中应说明。

周正确率=完全正确题数÷可判定的已作答题数；部分正确不计完全正确。得分率=总得分÷已评分题满分，包含明确空答。pending两者均排除并单列覆盖数。已确认任务完成率=实际完成任务数÷已确认执行的已布置任务数；另报未知执行天数，不能将未知天数悄悄当作0或100%。完成天数与已生成试卷数量分开。

知识点依据最近最多10条有效checks（可跨周）判定：

| 状态 | 条件 |
|---|---|
| 已掌握 | 至少5次检测，正确率≥90%，至少3个日期独立成功、跨度≥7天，至少2个不同日期独立复习成功 |
| 基本掌握 | 至少3次检测，正确率≥80%，至少2个日期独立成功 |
| 不稳定 | 有证据但不满足以上条件，且正确率≥50%；仅有初始已知错误也暂列此态 |
| 未掌握 | 有新的有效检测且最近检测正确率<50% |
| 未评估 | 无有效检测、无初始错误证据；听说默认如此 |

上述是本项目教学判定，不是CEFR/雅思成绩。无提示的主动回忆/独立造句才能作为independent成功；只做有选项识别或看答案抄写不能虚报独立成功。

最近一次检测日期出现失败时，状态最高为不稳定（若最近正确率<50%则未掌握），避免高累计正确率掩盖刚暴露的缺口。

## 间隔复习

初始错题到期日为接入日（不是假定历史作答日）。新失败后+1天；到期独立复习成功后依次距本次+3、+7、+14、+30天；之后维持+30。失败重置，提示后成功不升级。按知识点每天最多升级一次，同日任何失败优先，提前成功不延长到期日。日期按实际作答日期而非批改日期。重批从全部当前有效记录重新播放，撤销误判不会留下旧错误次数。

## 分类重点

`key_points`为可选集合，首次用户主动收录时建立：schema_version=1、document_path固定为`06_Learning Notes/English Learning_Key Points.docx`、document_sha256、sources（原始输入相对路径到SHA-256）、entries（稳定id到条目）。条目包含id、category、pattern、meaning、examples（english/chinese）、notes、first_recorded_date、last_updated_date、originals（date/text/source_path）。大类限词汇与搭配、语法与句型、易混辨析、常见错误；子分类可扩展。

一条重点一个稳定id。相同搭配不得创建多个id，近义内容由Codex读取已有条目后合并；同一输入重复处理不新增事务。原文保留，整理后含义与提醒可纠正原文中的问题。文档、来源档案和state在同一事务保存。重点不参与知识点检测、成绩、错误次数、完成量及间隔复习计算。

## 新增单词汇总

`vocabulary_summary`为可选派生集合：schema_version=1、document_path=`06_Learning Notes/English Learning_Vocabulary Summary.docx`、document_sha256、signature、total_unique_words、days。days以课程日期为键，保存date/day/source_path/source_sha256/snapshot_path/table_sha256/new_word_count/duplicate_words_skipped/words。words保留word（小写去重键）、headword、pronunciation、meaning、collocations、example。来源快照按源文件SHA存入`.learning-manager/vocabulary/sources`。

register和单词汇总在同一事务更新，独立summarize-vocabulary可补汇总。signature基于实际词表、分组和登记主词；完全重复输入不增版本，修订替换当前汇总。只收new模式工作日主词，按首次出现去重，复习不累加。汇总量不参与成绩、掌握状态或任务完成量。重点的display_fields可关闭原文、翻译、首次日期在Word中的显示，内部证据仍保留；本次目录迁移保留用户精简的Word和显示偏好。

## 备份和恢复

每个写事务先用排他锁验证读取版本，再保存所有将改文件的before/after字节及SHA-256，写prepared manifest，然后逐个原子替换，最后committed。异常回滚；进程中断时读取拒绝继续，使用recover恢复。recover只移除已确认退出进程的锁；文件若已有第三方修改则保留现场，不覆盖。不要手动删锁。

对Word只修改document.xml目标文本/复选状态，其他ZIP包部件逐字节保留。外部修改周表后哈希不一致时停止写表，需要读取并合并用户编辑，不能直接重置哈希绕过。

views（反馈md、错题json/md、执行日志、周评估）是派生文件，若要修改学习事实应更新对应输入而非手工编辑views。每次重批同步已生成周汇总，旧版仍保存在事务快照中。周评估按周保留；档案仅追加/更新管理区，初始诊断正文保留。

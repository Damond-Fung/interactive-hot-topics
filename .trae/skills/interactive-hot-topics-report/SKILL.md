---
name: "interactive-hot-topics-report"
description: "Generates an interactive-discussion hot topics Excel report for the TRAE forum. Invoke only for requests about the 互动交流 board, including hot-topic summaries, negative feedback, product suggestions, full-detail exports, analysis_packets, ai-template export, or ai_results backfill."
---

# 互动交流热门讨论报告

面向 `https://forum.trae.cn/c/11-category/11` 公开论坛板块的独立 Skill。

该 Skill 与“社区之星：活力之星”完全分开，只处理 `互动交流` 板块的热门讨论汇总、逐帖 AI 判别回填、Excel/JSON 导出。

## 何时调用

只要用户明确在问 `互动交流` 板块，并且目标属于“逐帖理解后再汇总”的报告场景，就应优先调用本 Skill。

优先在以下表达出现时调用：

- `互动交流热门讨论`
- `互动区最热讨论话题`
- `热门话题主题 TOP5`
- `高频负向反馈 TOP5`
- `产品建议 TOP5`
- `互动交流全量明细`
- `analysis_packets`
- `ai-template`
- `ai_results 回填`
- `逐帖 AI 判别`
- `标题和全文语义理解`

典型场景：

- 用户要统计 `互动交流` 板块某个时间范围内大家讨论最多的内容
- 用户要看上周 `互动交流` 里最热的 5 个主题是什么
- 用户要汇总 `互动交流` 的负向反馈和产品建议
- 用户要导出 `汇总 + 明细` 双页签 Excel
- 用户要先导出逐帖 AI 判别底表，再回填正式 `ai_results`
- 用户给出一个自然语言时间范围，希望自动换算成具体起止日期后出报告

不要在以下场景调用：

- 用户要统计“活力之星”“助人之星”或其他社区榜单
- 用户的数据源不是公开可访问的 `互动交流` 板块
- 用户要依赖登录态、后台接口、私有接口
- 用户只要分类页简单排行榜，不需要逐帖理解与汇总
- 用户只是在问某一篇帖子本身的总结、翻译、改写或审阅

## 路由判断

看到下面两类意图时，默认都路由到本 Skill：

- `结果导向`：用户要 `热门话题主题 TOP5`、`高频负向反馈 TOP5`、`产品建议 TOP5`
- `底表导向`：用户要 `全量帖子明细`、`analysis_packets`、`ai-template`、`ai_results 回填`

如果用户同时提到 `互动交流` 和 `周报 / 月报 / 汇总 / 明细 / 负向反馈 / 产品建议 / AI 判别`，本 Skill 的优先级高于泛化的数据报表类 Skill。

## 能力边界

本 Skill 只做以下事情：

- 从公开 `互动交流` 分类页按创建时间全量翻页采集统计窗口内的话题
- 拉取每个话题的 `标题 + 全帖内容`
- 生成逐帖判别底表与 `analysis_packets`
- 接受外部 `ai_results.json` 回填
- 基于回填结果生成：
  - `热门话题主题 TOP5`
  - `高频负向反馈 TOP5`
  - `产品建议 TOP5`
- 导出 `.xlsx` 与 `.json`
- 在输出里保留明确的统计窗口、时区、板块来源与回填状态

本 Skill 明确不做：

- 登录论坛或绕过限制
- 猜测不可验证的隐藏字段
- 伪装规则关键词结果为“真实 AI 理解结果”
- 在没有 `ai_results` 的情况下谎称已经完成最终语义汇总
- 扩展到其他板块或论坛后台数据采集

## 默认规则

- 默认源地址：`https://forum.trae.cn/c/11-category/11`
- 默认时区：`Asia/Shanghai`
- 默认时间范围：`last-week`
- 默认 Top N：`5`
- 默认统计板块固定为：`互动交流`
- 默认采集口径：按 `created` 排序翻页，并按北京时间自然周过滤
- 默认明细范围：统计窗口内的 `互动交流` 全量话题，而不是只保留命中热点的话题

## 时间规则

如果用户没有额外说明，默认按“北京时间上周”查询。

支持：

- `last-week`
- `this-week`
- `last-7-days`
- 自定义 `start-date` / `end-date`

如果用户使用自然语言时间，例如“5 月第一周”“4 月整月”：

- 先换算成明确起止日期
- 在输出中明确写出最终采用的时区与起止时间

如果用户没有明确说要“最终 AI 汇总”，但提到了：

- `先导底表`
- `先看明细`
- `先导 ai-template`
- `我确认口径后再回填`

则默认进入“底表导出模式”，不要直接生成伪造的最终 TOP5。

## 输出结构

默认导出两个主文件：

- Excel：`汇总` + `明细`
- JSON：包含 `detail_rows`、`analysis_packets`、汇总结果与元信息

`明细` 页至少包含：

- `topic_id`
- `标题`
- `链接`
- `创建时间`
- `浏览量`
- `总回复数`
- `窗口内帖子数`
- `窗口内回帖数`
- `讨论内容分类`
- `关键词提取`
- `热门主题分类`
- `负向反馈分类`
- `产品建议分类`
- `AI话题分类`
- `AI是否负向反馈`
- `AI是否产品建议`
- `AI核心内容`
- `样本文本`

JSON 重点字段至少包含：

- `detail_rows`
- `analysis_packets`
- `time_label`
- `start_local`
- `end_local`
- `ai_results_applied`
- `ai_results_count`

## AI 回填规则

如果用户只需要底表：

- 使用 `--export-ai-template`
- 导出 `ai-template.json`
- 导出 `ai-review.md`
- 在汇总中明确标记当前仍是“待 Agent 判别”

如果用户已经提供正式 `ai_results.json`：

- 使用 `--ai-results`
- 优先基于 `ai_results` 重算最终 `TOP5`

如果没有 `ai_results`：

- 只输出待判别状态
- 不输出伪造的最终 AI 热点榜单

如果用户已经上传或提供了 `ai_results.json`，则应视为“正式回填模式”，优先重新生成最终版汇总，而不是继续展示底表占位结果。

推荐执行顺序：

1. 先跑一次底表导出
2. 如需逐帖语义判别，再导出 `ai-template`
3. 回填正式 `ai_results.json`
4. 最后重新生成正式 `TOP5`

## 请求归一化

执行前先把用户请求归一化为以下几个关键信息：

- `统计板块`：固定为 `互动交流`
- `时间窗口`：默认 `last-week`，或根据自然语言换算
- `输出阶段`：`底表导出` / `AI 模板导出` / `正式回填汇总`
- `输出文件`：Excel、JSON，必要时增加 `ai-template` 与 `ai-review`

如果用户一句话里同时提到“热门话题”“负向反馈”“产品建议”，默认视为同一份综合报告，不拆成多个独立任务。

## 执行方式

优先运行随 Skill 附带的独立包装脚本：

```bash
python .trae/skills/interactive-hot-topics-report/resources/scripts/generate_interactive_hot_topics_report.py
```

常用参数：

```bash
python .trae/skills/interactive-hot-topics-report/resources/scripts/generate_interactive_hot_topics_report.py \
  --time-preset last-week \
  --top-n 5
```

导出 AI 模板：

```bash
python .trae/skills/interactive-hot-topics-report/resources/scripts/generate_interactive_hot_topics_report.py \
  --time-preset last-week \
  --export-ai-template
```

回填正式 AI 结果：

```bash
python .trae/skills/interactive-hot-topics-report/resources/scripts/generate_interactive_hot_topics_report.py \
  --time-preset last-week \
  --ai-results ./exports/interactive_hot_topics_ai_results_20260518-20260524.json
```

指定自定义日期：

```bash
python .trae/skills/interactive-hot-topics-report/resources/scripts/generate_interactive_hot_topics_report.py \
  --start-date 2026-05-01 \
  --end-date 2026-05-31
```

## 默认命名规则

如果用户没有显式指定 `--output`，默认文件名为：

- `TRAE社区_互动交流热门话题_YYYYMMDD-YYYYMMDD.xlsx`

对应 JSON 使用相同文件名，仅扩展名改为 `.json`。

## 失败处理

如果 Excel 文件被占用：

- 自动改用新文件名重新导出
- 明确提示用户原文件可能正在被打开

如果论坛公开接口超时：

- 自动重试
- 若仍失败，再向用户明确报错

如果没有正式 `ai_results`：

- 明确说明当前结果仍是底表与待判别状态
- 不要把占位结果冒充最终语义总结

## 示例请求

- 帮我生成上周互动交流热门讨论报告，并导出 Excel
- 帮我看互动交流上周大家讨论最热的 5 个主题是什么
- 帮我整理互动交流上周高频负向反馈和产品建议，并附全量明细
- 这是 `ai_results.json`，帮我回填后生成最终版互动交流热门讨论报告
- 先帮我导出上周互动交流的全量帖子明细和 `analysis_packets`
- 我想确认 AI 判别口径，先导出 `ai-template`，不要直接给我伪 TOP5

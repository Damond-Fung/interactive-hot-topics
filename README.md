# interactive-hot-topics

TRAE 中文社区 `互动交流` 板块热门讨论报告工具。

面向 `https://forum.trae.cn/c/11-category/11` 的公开数据，按指定统计窗口采集 `互动交流` 全量话题，生成可核对的底表，以及可通过 `ai_results.json` 正式回填的热门讨论报告。

## 这个仓库解决什么问题

如果你想知道某个时间范围内 `互动交流` 板块：

- 大家讨论最集中的主题是什么
- 高频负向反馈主要集中在哪些问题
- 产品建议最常提到哪些方向
- 每条话题的原始明细是否采集完整

这个仓库会直接产出一份可核对、可回填、可复用的报告产物，而不是只给一个不可追溯的口头结论。

## 会产出什么

- `热门话题主题 TOP5`
- `高频负向反馈 TOP5`
- `产品建议 TOP5`
- `汇总 + 明细` 双页签 Excel
- 含 `detail_rows` 与 `analysis_packets` 的 JSON
- `ai-template.json` 与 `ai-review.md`

## 适合怎么用

- 直接出报告：先看上周或指定时间范围的热门讨论结论
- 先导底表：先核对全量帖子明细，再决定是否进入 AI 判别
- 正式回填：把逐帖判别结果写入 `ai_results.json` 后，重新生成最终汇总版报告

## 核心规则

- 与“活力之星”完全分开，只服务 `互动交流` 板块
- 默认时区为 `Asia/Shanghai`
- 默认时间窗口为 `last-week`
- 按帖子 `created` 排序翻页采集，不依赖分类页默认活跃排序
- 明细默认保留统计窗口内的全量话题，而不是只保留命中热点的话题
- 没有正式 `ai_results` 时，只输出底表与待判别状态，不伪造最终 AI 汇总

## 快速开始

安装依赖：

```bash
python -m pip install -r requirements.txt
```

导出上周底表：

```bash
python interactive_hot_topics_report.py --time-preset last-week
```

导出 AI 模板：

```bash
python interactive_hot_topics_report.py --time-preset last-week --export-ai-template
```

回填正式 AI 结果：

```bash
python interactive_hot_topics_report.py \
  --time-preset last-week \
  --ai-results exports/interactive_hot_topics_ai_results_20260518-20260524.json
```

使用 Skill 包装脚本：

```bash
python .trae/skills/interactive-hot-topics-report/resources/scripts/generate_interactive_hot_topics_report.py
```

## 推荐流程

1. 先导出底表，确认采集范围和时间窗口正确。
2. 如需逐帖语义判别，再导出 `ai-template.json` 和 `ai-review.md`。
3. 完成逐帖判别后，回填正式 `ai_results.json`。
4. 重新生成最终版 `热门话题主题 TOP5 / 高频负向反馈 TOP5 / 产品建议 TOP5`。

## 仓库结构

- `interactive_hot_topics_report.py`
  - 核心采集、聚合、导出脚本
- `.trae/skills/interactive-hot-topics-report/SKILL.md`
  - Skill 调用边界、默认规则、执行建议
- `.trae/skills/interactive-hot-topics-report/examples/example-request.md`
  - 更贴近日常使用的自然语言请求示例
- `.trae/skills/interactive-hot-topics-report/resources/scripts/generate_interactive_hot_topics_report.py`
  - 独立 Skill 包装脚本
- `requirements.txt`
  - 最小 Python 依赖

## 默认命名

- Excel：`TRAE社区_互动交流热门话题_YYYYMMDD-YYYYMMDD.xlsx`
- JSON：与 Excel 同名，仅扩展名改为 `.json`
- AI 模板：在同批次文件名基础上追加 `.ai-template.json`

## 相关文档

- Skill 说明：`.trae/skills/interactive-hot-topics-report/SKILL.md`
- 示例请求：`.trae/skills/interactive-hot-topics-report/examples/example-request.md`

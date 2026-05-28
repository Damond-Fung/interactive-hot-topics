# interactive-hot-topics

TRAE 中文社区 `互动交流` 板块热门讨论报告工具。

面向 `https://forum.trae.cn/c/11-category/11` 的公开数据，按指定统计窗口采集 `互动交流` 全量话题，生成可核对的底表，并支持双模式完成最终 AI 汇总：

- `模式 A / Agent 模式`：由宿主 Agent 逐帖读取 `analysis_packets` 并回填 `ai_results.json`
- `模式 B / API 模式`：如果当前 Agent 不支持自动逐帖分析，则脚本直接调用模型 API 完成判别并自动回填

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
- 自动判别：优先走 Agent；如果当前工具不支持，就自动切到脚本 API 模式

## 核心规则

- 与“活力之星”完全分开，只服务 `互动交流` 板块
- 默认时区为 `Asia/Shanghai`
- 默认时间窗口为 `last-week`
- 按帖子 `created` 排序翻页采集，不依赖分类页默认活跃排序
- 明细默认保留统计窗口内的全量话题，而不是只保留命中热点的话题
- 没有正式 `ai_results` 时，只输出底表与待判别状态，不伪造最终 AI 汇总
- `--ai-mode auto` 默认优先走 Agent 模式；若已配置可用模型 API，则自动回退到 API 模式

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

自动路由模式：

```bash
python interactive_hot_topics_report.py --time-preset last-week --ai-mode auto
```

回填正式 AI 结果：

```bash
python interactive_hot_topics_report.py \
  --time-preset last-week \
  --ai-results exports/interactive_hot_topics_ai_results_20260518-20260524.json
```

API 模式直连模型：

```bash
python interactive_hot_topics_report.py \
  --time-preset last-week \
  --ai-mode api \
  --llm-provider github
```

可用参数：

```bash
python interactive_hot_topics_report.py \
  --ai-mode auto|agent|api \
  --llm-provider github|openai|custom \
  --llm-base-url <chat-completions-endpoint> \
  --llm-model <model-id> \
  --llm-api-key-env <ENV_NAME>
```

环境变量默认值：

- GitHub Models：默认读取 `GITHUB_TOKEN`
- OpenAI：默认读取 `OPENAI_API_KEY`

使用 Skill 包装脚本：

```bash
python .trae/skills/interactive-hot-topics-report/resources/scripts/generate_interactive_hot_topics_report.py
```

## 在其他 AI 工具 / 无 Agent 环境下使用 CLI

如果当前环境无法驱动宿主 Agent 自动逐帖回填（其他 IDE、其他 AI 客户端、CI、纯 shell），可使用统一 CLI 入口 `interactive_hot_topics_cli.py`，把“采集 → AI 判别 → 最终汇总”封装为一行命令：

```bash
python interactive_hot_topics_cli.py run \
  --time-preset last-week \
  --llm-provider github
```

子命令一览：

- `report`：仅采集导出底表（保留 AI 模板）
- `analyze --ai-results <path>`：基于已有结果重算最终汇总
- `run`：一键完成全流程，默认 `--ai-mode api`

CLI 内部调用 `interactive_hot_topics_report.py`，所有窗口参数与 LLM 参数完全一致，不重复实现。

## 推荐流程

1. 先导出底表，确认采集范围和时间窗口正确。
2. 优先尝试 `模式 A / Agent 模式`，导出 `ai-template.json` 和 `ai-review.md` 后逐帖回填。
3. 如果当前宿主不支持自动逐帖分析，则改用 `模式 B / API 模式`。
4. 生成正式 `ai_results.json` 或 `.auto-ai-results.json` 后，输出最终版 `热门话题主题 TOP5 / 高频负向反馈 TOP5 / 产品建议 TOP5`。

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
- JSON：与 Excel 同名，仅扩展名改为 `.json`
- AI 模板：在同批次文件名基础上追加 `.ai-template.json`
- API 模式回填文件：在同批次文件名基础上追加 `.auto-ai-results.json`

## 推送建议

如果要作为独立 GitHub 仓库发布，建议最少包含以下文件：

- `README.md`
- `LICENSE`
- `interactive_hot_topics_report.py`
- `.trae/skills/interactive-hot-topics-report/`

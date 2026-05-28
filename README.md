# interactive-hot-topics

TRAE 中文社区 `互动交流` 板块热门讨论报告工具。

该项目面向 `https://forum.trae.cn/c/11-category/11` 的公开数据，按指定统计窗口采集 `互动交流` 全量话题，生成可核对的底表与可正式回填的热门讨论报告。

## 产出内容

- `热门话题主题 TOP5`
- `高频负向反馈 TOP5`
- `产品建议 TOP5`
- `汇总 + 明细` 双页签 Excel
- 含 `detail_rows` 与 `analysis_packets` 的 JSON
- `ai-template.json` 与 `ai-review.md`

## 适用场景

- 需要看上周或指定时间范围内 `互动交流` 最热讨论内容
- 需要先导出底表，再做逐帖 AI 判别
- 需要把正式 `ai_results.json` 回填成最终汇总版报告
- 需要把热门主题、负向反馈、产品建议放到同一份报告里

## 核心规则

- 与“活力之星”完全分开，只服务 `互动交流` 板块
- 默认时区为 `Asia/Shanghai`
- 默认时间窗口为 `last-week`
- 按帖子 `created` 排序翻页采集，不依赖分类页默认活跃排序
- 明细默认保留统计窗口内的全量话题，而不是只保留命中热点的话题
- 没有正式 `ai_results` 时，只输出底表与待判别状态，不伪造最终 AI 汇总

## 项目结构

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

## 安装依赖

```bash
python -m pip install -r requirements.txt
```

## 常用命令

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

## 默认命名

- Excel：`TRAE社区_互动交流热门话题_YYYYMMDD-YYYYMMDD.xlsx`
- JSON：与 Excel 同名，仅扩展名改为 `.json`
- AI 模板：在同批次文件名基础上追加 `.ai-template.json`

## 推送建议

如果要作为独立 GitHub 仓库发布，建议最少包含以下文件：

- `README.md`
- `LICENSE`
- `interactive_hot_topics_report.py`
- `.trae/skills/interactive-hot-topics-report/`

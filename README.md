# 交大中文论文语言与排版 Skill

面向上海交通大学中文工程硕士、MEM论文的 AI 助手技能。根据用户收集的本校论文样本提炼表达方法，结合指定格式样板，处理空话、重复、论证脱节及Word格式问题。

“去AI味”在本项目中是编辑任务，不是作者身份检测。论文质量通过事实、论证、引用和排版检查来提高；本项目不承诺降低检测分数或保证学位论文通过。非上海交通大学官方项目。

## 能做什么

- 在保留数据、引用、否定、条件和研究限制的前提下，改善已有正文。
- 根据真实材料起草可核验的段落，检查章节之间的问题、方法与评价是否对应。
- MEM以学院最新适用模板为准；当前提供学院2026 MEM适配，旧工程硕士样板仅作历史对照，保留来源与冲突。
- 通过附带Python工具盘点DOCX、对审核过的段落/节应用局部格式、生成核查回执。
- 指导真实目录/题注/交叉引用、分节页码、页眉、表格和最终PDF视觉检查。

格式工具只处理明确映射的段落与节，不是全文自动排版器。完整排版还需要Agent结合可用Word/文档工具执行并核验；没有渲染/字段更新能力就记录partial。语言改得自然也不能补救无来源的数据、虚构的方法或越界结论。

## 安装

Codex：

```bash
git clone https://github.com/yesaze2077/sjtu-thesis-humanizer.git ~/.codex/skills/sjtu-thesis-humanizer
```

Claude Code：

```bash
git clone https://github.com/yesaze2077/sjtu-thesis-humanizer.git ~/.claude/skills/sjtu-thesis-humanizer
```

目录已存在时先检查现有版本，不覆盖本地改动。重新加载技能或开启新会话后调用；技能是由Agent执行的指导，不能单独作为论文生成程序运行。本项目MEM当前采用学院2026版模板；学院发布更新时应更新本地profile。格式脚本仅依赖Python 3.10+标准库；渲染、Word字段更新与字体需要所在环境另行具备。没有API密钥、外部模型调用或自动上传依赖。

## 使用

```text
请用 $sjtu-thesis-humanizer 润色第3章，保留章节结构、数值、引文和必要限制，另存副本。

请用 $sjtu-thesis-humanizer 审阅这两节的论证，不修改文件。

请用 $sjtu-thesis-humanizer 根据我提供的评价表起草结果讨论；没有证据的结论不要补写。

请用 $sjtu-thesis-humanizer 按我本届学院模板排版此Word，并检查目录、图表、页眉页码和导出PDF。
```

提供当前稿、实际证据和适用模板即可。不必把完整参考论文提交到公共仓库。跨学院、届次或模板不同时，先更新本地profile并记录来源，不能直接套用本仓库的默认值。

## 结构与来源

| 文件 | 用途 |
|---|---|
| [SKILL.md](SKILL.md) | Agent入口、工作范围和约束 |
| [语言指南](references/writing-guide.md) | 上下文判断、原创示例、语义核对 |
| [章节指南](references/chapter-quality.md) | 研究问题与章节功能检查 |
| [语料来源卡](references/corpus-cards.json) | 匿名样本、SHA256、实际抽读页与限制 |
| [格式操作](references/formatting.md) | 模板选择、脚本操作、人工与工具验收 |
| [格式profile](profiles) | 两套样板的版本化参数与来源定位 |
| [验收记录](references/acceptance.md) | 已完成与未检查状态分开记录 |
| [验证记录](VALIDATION.md) | 当前版本的实测范围与限制 |
| [评测案例](evals) | 合成案例与语言保真验收，不使用真实论文数据 |

本项目参考 [Humanizer-zh](https://github.com/op7418/Humanizer-zh)，保留其许可及版本说明，见 [第三方声明](THIRD_PARTY_NOTICES.md)。本校样本只作有限体裁参考，不是作者个人风格档案或学院语言认证。样本的陈旧做法、套话和过强结论不作为写作目标。

学校通用指南允许学院制定相应格式，因此本技能优先核对本专业、本届次要求。[学校撰写指南](https://www.gs.sjtu.edu.cn/post/detail/Z3MzNDA0)。AI允许范围与披露按实际适用要求处理，校级原则不等于允许整篇代写。[学校AI规范](https://plan.sjtu.edu.cn/info/1025/4295.htm)。这些规则的访问状态与日期见格式说明；不将作者或导师的责任改写为机器验收。

## 开发验证

```bash
python3 -m unittest discover -s tests -v
```

脚本测试验证OOXML操作范围与保全；行为样例检验语言决策，不证明所有模型、论文与排版环境都可靠。实际论文必须对最终版本逐页检查。

## 许可

本项目新增说明、代码和合成案例使用 [MIT](LICENSE)。没有发布原论文、学校模板、校徽、字体或企业/客户材料；它们不受本项目开源许可覆盖。

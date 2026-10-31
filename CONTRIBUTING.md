# 贡献指南

感谢你愿意改进这个项目。这里是提 Issue 和 PR 的约定，读完大概能省下两轮来回。

## 提 Issue

**报 bug 时请附上这些**，缺了我大概率只能回一句「能复现吗」：

- 操作系统和 Python 版本（`python --version`）
- 完整命令和完整报错文本（不要截图，要可复制的文字）
- `work/blocks.json` 里出问题那段的结构，例如：

  ```json
  {"id": 42, "kind": "para", "style": "Normal", "chars": 186, "runs": 3, "text": "……"}
  ```

- 如果涉及时效性，说明用的是哪个版本的技能（`SKILL.md` 的 `metadata.version` 字段）

**不要上传真实论文。** 用 `make_sample.py` 生成的样例，或者自己捏一段结构相同、内容无关的假数据。论文是敏感材料，贴到公开 issue 里删不干净。

## 提 PR

### 分支与提交

从 `main`（或 `master`）切分支，一个 PR 只做一件事。提交信息用中文或英文都行，格式建议：

```
fix: batch 在章节名含全角空格时匹配失败
feat: 新增被动句比例检查项
docs: 补充脚注偏移的手工核对步骤
```

### 改代码前

先跑回归，确认改动前是绿的：

```bash
# 仓库根目录，不需要 pytest
python -m unittest discover -s tests -v
```

全部测试通过才算基线正常，不是绿的就先别改，那是基线出了问题。CI 会在 Linux + Windows、Python 3.9/3.12 上跑同一套测试。

想手工走一遍全流程（改完跑一遍心里更踏实）：

```bash
cd skills/aigc-reducer/scripts
python make_sample.py -o ../../../样例.docx
python docx_tool.py split ../../../样例.docx -o ../../../work/blocks.json
python docx_tool.py batch ../../../work/blocks.json --id 1 --prompt ../references/prompt.md
# 把回复存成 reply.txt（段落之间空一行）
python docx_tool.py apply ../../../work/blocks.json --batch 1 --input reply.txt -o ../../../work/edits.json
python docx_tool.py check ../../../样例.docx ../../../work/edits.json
```

check 结果取决于所填的改写；fail 退出 1，必须检查结果并修复。

### 新增检查项或子命令

改了 `docx_tool.py` 就要在 `tests/test_pipeline.py` 里补对应用例，尤其是：

- 拦截类改动（新增拒绝路径）——要断言 exit code 非 0 且不生成文件
- 判定类改动（分类、引用、字数口径）——要断言具体数值，不能只断言「不报错」

测不到的改动不收。

### 三条硬约束

修改 `docx_tool.py` 时，这三条不能破：

1. **不允许静默失败。** 任何跳过、降级、丢格式，都必须打印出来。工具的价值一半在「告诉你哪里可能有问题」。
2. **不允许覆盖源文件。** 输出路径等于输入路径时直接拒绝，`--allow-*` 之类的开关也不行。
3. **不允许放宽段落类型校验。** 标题、图注、参考文献、公式永远不参与回写。新增 `--allow-non-para` 之外的放行方式会被拒。

### 改 `references/prompt.md` 前

**不要改这五条指令的措辞。** 它们是实测调参的结果，任何「优化」都会让效果不可控。要试新写法，另起一个文件，用 `batch --prompt` 指向它，把对比结果放进 PR 描述。

如果确实要替换默认指令，必须附上前后对比：同一篇样例、同一批段落，两套指令各跑一遍的 `check.json`。

### 新增检测项

`check` 的指标加在 `cmd_check` 里，同时要做三件事：

- 在 `per_para` 里记录逐段数据，不能只给全篇结论
- 在 README 的指标表里加一行，标明属于 `warn` 还是 `fail`
- 在 `CHANGELOG.md` 的 `Unreleased` 里记一笔

引用新增或丢失（包括次数）、数值/符号/百分号/比较方向变化属于 fail，需人工核对；风格类问题一律 warn。不要将自检通过等同于语义或检测效果通过。

### 适配新的检测平台

改 `references/report-parsing.md`，补该平台的报告形态和字段抽取方式。**不要为此改代码**——解析策略本来就设计成可替换的。

如果新平台的报告结构实在特殊，需要在 `SKILL.md` 里写清楚降级路径，而不是加硬编码的解析逻辑。

## 代码风格

- 中文注释，说明「为什么这么做」，不复述代码在做什么
- 对外提示一律中文，格式 `[错误] / [提示] / [警告] / [跳过]` 开头
- 新依赖要先说服人：能在标准库里解决就别引入第三方包
- 保持单文件可运行，`docx_tool.py` 不拆包——用户是拷贝单个文件在用的

## 不接受什么

- 绕过 AIGC 检测的「调参技巧」「对抗样本」类改动
- 把论文内容上传到任何服务器的功能
- 自动生成论文内容的功能
- 破坏「原文件不动」原则的便利性改动

这个工具的边界是**表达层整改**，越界的功能即使技术上可行也不会合并。

## 许可证

提交 PR 即表示你有权按 [MIT](LICENSE) 提供这些贡献。第三方代码、提示词、示例和图片须附来源、许可及所需声明；不能用项目 MIT 覆盖他人权利。不要复制来源不明的提示词或仅供服务使用的 Skill。

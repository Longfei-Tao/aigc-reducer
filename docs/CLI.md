# 命令行参考

以下命令在仓库根目录执行，论文路径可替换为绝对路径。默认只交付修改后的 Word；中间文件用于内部处理。

## 命令参考

六个子命令串成流水线。技能执行时每篇论文使用独立临时工作目录，中间文件不交付；下例 `work/` 只供开发者手动试跑。先 check 再 write，成品重新打开核对。`--blocks` 校验 split 后原稿是否改变；新版 apply 结果自带原稿指纹；只有旧无指纹格式不提供此选项时不能防止旧 id 错位。

以下多行命令使用 Bash 续行语法。PowerShell 请合成一行（删除行末反斜杠）；SKILL.md 中有单行示例。

> 下例以仓库根目录为当前目录，`work/` 供单篇手动试跑。安装为技能后，脚本路径替换为安装目录的绝对路径。

### 1. `split` —— 拆文档

```bash
python skills/aigc-reducer/scripts/docx_tool.py split 论文.docx -o work/blocks.json --batch-chars 1200
```

输出段落清单、章节列表、分批方案与源文件 SHA-256 指纹。章节列表只含每个标题的直接正文，不自动包含子节；单段超过上限时独立成批，不截断。看一眼摘要确认三件事：可改写字数、表格数（不参与改写）、分批数。

```
[split] 论文样例.docx
  段落总数      25
  可改写正文     10 段 / 810 字
  标题 / 参考文献 / 公式 / 题注  10 / 4 / 0 / 1
  表格          1 个（37 字，不参与改写）
  分批          7 批（每批 ≤1200 字）
```

### 2. `batch` —— 取一批段落，拼成可直接发送的请求

```bash
# 按批次号
python skills/aigc-reducer/scripts/docx_tool.py batch work/blocks.json --id 3 \
  --prompt skills/aigc-reducer/references/prompt.md -o work/batch3.txt

# 按章节标题（模糊匹配）
python skills/aigc-reducer/scripts/docx_tool.py batch work/blocks.json \
  --section "2.1 研究对象" --prompt skills/aigc-reducer/references/prompt.md

# 直接指定段落 id
python skills/aigc-reducer/scripts/docx_tool.py batch work/blocks.json \
  --ids 8,9,10 --prompt skills/aigc-reducer/references/prompt.md
```

`--prompt` 输出“执行约束 + 改写指令原文 + 原文段落”：先约束不检索、事实保护及回复格式，再逐字保留风格指令。默认和自定义提示词均适用；不传 `--prompt` 仍只输出原文。加 `-o` 存成纯文本，可完整复制粘贴。执行约束不能切换客户端开关，发送前按[运行设置指南](../skills/aigc-reducer/references/runtime-settings.md)关闭搜索并开启支持的推理模式。

这一条是为了消灭最常见的翻车方式：手抄段落 id 错一位，标题就被正文覆盖了。`batch` 会拒绝任何非正文段落。

### 3. `apply` —— 把模型回复写回结果

```bash
# 把模型的回复存成 reply_3.txt（段落之间空一行），再对齐
python skills/aigc-reducer/scripts/docx_tool.py apply work/blocks.json --batch 3 \
  --input work/reply_3.txt -o work/edits.json --progress work/progress.json
```

按空行切分回复，**块数与本批段落数不相等就拒绝写入**，并列出期望 id 与实际块数。`edits.json` 累积写入，`progress.json` 自动回填。回复被 ``` 代码围栏包住时会自动剥掉；段落编号前缀（`1.` `（1）`）加 `--strip-number` 去掉。

这一步消灭了最常见的翻车方式：模型返回 3 段而本批是 4 段，靠人眼对序号，结果把下一节的**标题**当成正文覆盖掉。

### 4. `status` —— 看进度，中断后续跑

```bash
python skills/aigc-reducer/scripts/docx_tool.py status work/blocks.json
```

逐批列出「完成 / 部分 / 未开始」和下一批的批次号，不用自己翻 `progress.json`。

### 5. `write` —— 回写新文档

```bash
python skills/aigc-reducer/scripts/docx_tool.py write 论文.docx work/edits.json --blocks work/blocks.json \
  -o 论文_降AI版.docx
```

新版 `apply` 生成的 `edits.json` 自带来源指纹：

```json
{
  "schema_version": 1,
  "source_sha256": "<split 时源文件的 SHA-256>",
  "edits": {"8": "改写后的第一段……", "9": "改写后的第二段……"}
}
```

下面旧格式仍可被 write/check/status 读取，但没有来源指纹；apply 不会向其累积写入，请另建工作目录。旧格式使用 --blocks 并人工核对 id：

```json
{
  "8": "改写后的第一段……",
  "9": "改写后的第二段……"
}
```

**安全校验**（缺一不可）：

- 任一输出与显式输入/其他输出重合 → 拒绝；write 也拒绝已存在的目标 docx
- 目标 id 是标题 / 图注 / 参考文献 / 公式 → 拒绝并列出，需 `--allow-non-para` 才放行
- 段落含超链接 → 默认跳过（改写会让链接文字重复），需 `--force`

会把 `[跳过]` / `[提示]` / `[警告]` 三类信息如实打印，包括局部加粗丢失、脚注位置偏移；含图片/域/修订/批注或文字与脚注同 run 的段落拒绝改写。

### 6. `check` —— 硬指标自检

```bash
python skills/aigc-reducer/scripts/docx_tool.py check 论文.docx work/edits.json --blocks work/blocks.json
```

| 指标 | 合格线 | 不达标 |
|---|---|---|
| 已改段落字数变化 | ±10% 内 | warn |
| 句号数量 | 减少 ≥20% | warn |
| 逗号数量 | 必须增加 | warn |
| 引用标记 `[n]` | 编号与次数一致，无新增 | **fail** |
| 数值 / 符号 / 百分号 / 比较方向 | 值及次数一致 | **fail**，需人工核对 |
| 「我觉得 / 我认为」 | 0 残留 | **fail** |
| 指令回显（含「修改段落」字样） | 0 段 | **fail**，重新生成该批 |

引用标记检查方括号写法 `[1]`、`[2-4]`（GB/T 7714）；中文论文里的年份括注「（2021）」不作引用计。含超链接的段落 `write` 会跳过，`check` 同样不计入指标，保证自检数字与成品一致。

`check` 比较原稿与 edits，**不读取成品 Word**。它不能证明术语、单位、否定和结论不变；交付前需重新打开成品并检查页面。默认仅执行检查，不生成报告文件；fail 退出 1，pass/warn 退出 0。

开发调试或用户明确要求时，可用 `check -o work/check.json` 导出自检 JSON，用 `write --diff-md work/对照表.md` 导出对照表；默认技能流程不使用这些参数。

`verdict` 分三档：`pass` / `warn`（可交付但需说明）/ `fail`（必须修）。

`fail` 时先修问题段再重新回写，不要带着已知问题交付。

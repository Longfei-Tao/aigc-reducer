---
name: Bug 报告
about: 工具报错、结果不对、文档被改坏
title: "[bug] "
labels: bug
---

## 发生了什么

<!-- 一句话说清现象，例如「write 把 3.2 的标题覆盖成了正文」 -->

## 怎么复现

1.
2.
3.

## 期望结果

<!-- 你认为它应该怎么做 -->

## 实际结果

<!-- 贴完整命令和完整报错文本，不要截图 -->

```
（在这里贴输出）
```

## 环境

- 操作系统：
- Python 版本（`python --version`）：
- 技能版本（`SKILL.md` 的 `metadata.version` 字段）：
- 用什么驱动的：WorkBuddy / Claude Code / 纯命令行 / 其他

## 段落结构

<!-- 出问题那段在 blocks.json 里的样子，可以删掉正文只留结构 -->

```json
{"id": 42, "kind": "para", "style": "Normal", "chars": 186, "runs": 3}
```

## 补充

<!-- 是否用了 --force / --allow-non-para？是否改过 prompt.md？ -->

> ⚠️ 请勿上传真实论文或检测报告。用 `make_sample.py` 生成的样例，或自己捏一段结构相同、内容无关的假数据。

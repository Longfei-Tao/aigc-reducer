# 安装与使用：按你使用的工具选择

[返回首页](../README.md) · [改写示例](../examples/README.md) · [CLI 完整参考](CLI.md)

核验日期：2026-10-31。下述路径和入口依据官方文档；本项目未逐一启动这些客户端进行实测，不将“官方支持技能”“发现本技能”“完成 Word 流程”混为一谈。宿主需要文件读写、Python 执行和报告阅读能力；某项缺失时见[手动流程](#manual)。

## 共同准备

正式改写前，在当前客户端关闭联网搜索并开启支持的深度思考 / 推理模式。**技能安装不会自动切换这两个设置。** [运行设置指南](../skills/aigc-reducer/references/runtime-settings.md)说明 DeepSeek、编辑器、Codex 和 API 的区别。保留本地文件与 Python 权限；关闭搜索不等于断开云端模型连接。

```bash
git clone https://github.com/Longfei-Tao/aigc-reducer.git
cd aigc-reducer
python -m pip install -r skills/aigc-reducer/requirements.txt
```

没有 Git 时下载仓库 ZIP 后解压。所有安装都复制 **skills/aigc-reducer 整个目录**，而非仓库根目录或单独一个 SKILL.md。Python 依赖安装在实际执行脚本的环境里：远程 Agent 的 Python 与本机不是同一个环境。

如果已使用 CC Switch，回到 CC Switch 导入/分发共享技能；下列手动复制方法只用于未受它管理的目录。本说明不修改任何账号、模型供应商或全局认证配置。

### 复制目录：Windows PowerShell

以下在下载后的仓库根目录执行。`$skillRoot` 按后面的工具说明选择，示例为共享个人目录：

```powershell
$skillRoot = Join-Path $env:USERPROFILE '.agents/skills'
$skillTarget = Join-Path $skillRoot 'aigc-reducer'
if (Test-Path -LiteralPath $skillTarget) { throw '目标已存在，请先备份并确认更新方式' }
New-Item -ItemType Directory -Force -Path $skillRoot | Out-Null
Copy-Item -LiteralPath 'skills/aigc-reducer' -Destination $skillTarget -Recurse
```

### 复制目录：macOS / Linux

```bash
skill_root="$HOME/.agents/skills"
if [ -e "$skill_root/aigc-reducer" ]; then
  echo "目标已存在，请先备份并确认更新方式"
  exit 1
fi
mkdir -p "$skill_root"
cp -R skills/aigc-reducer "$skill_root/aigc-reducer"
```

安装后启动新会话，使用后面的示例请求。看到同名技能时不要重复安装多份；更新前备份原目录，使用同一入口管理。

<a id="workbuddy"></a>
## WorkBuddy

1. 生成一个只包含技能本体的 ZIP。在仓库根目录运行，先确认没有同名 ZIP，以免覆盖：

```bash
python -c "from pathlib import Path; import shutil; assert not Path('aigc-reducer.zip').exists(), 'ZIP already exists'; shutil.make_archive('aigc-reducer', 'zip', 'skills', 'aigc-reducer')"
```

2. 打开 **技能 → 添加技能 → 上传技能**，选择生成的本地技能包，导入后启用。
3. 新建任务并选择论文所在工作目录，添加论文和检测报告。在对话的技能入口选择 aigc-reducer（若显示中文名称则选择“论文降 AIGC 率”），发送[示例请求](#request)。
4. 如果提示缺少 python-docx，在 WorkBuddy 执行脚本所用的 Python 环境安装 requirements.txt。不要仅凭“导入成功”认定 Python 也已可用。

[WorkBuddy 官方技能说明](https://www.codebuddy.ai/docs/zh/workbuddy/From-Beginner-to-Expert-Guide/Function-Description/Skills-Market)确认可上传本地技能包并启用；具体界面和接受的包格式以安装版本为准。若 ZIP 不被接受，按界面要求调整打包格式，或让已授权访问文件的宿主读取解压目录的 SKILL.md 执行；后者是手动加载，不代表已注册安装。

<a id="claude-code"></a>
## Claude Code

将共同复制步骤的目录改为 `~/.claude/skills/`（PowerShell 用 `.claude/skills`）。仅对某个项目使用时，复制到该项目的 `.claude/skills/aigc-reducer/`。

启动新的 Claude Code 会话，输入 `/aigc-reducer`，并提供原稿、报告的文件路径或附件及[示例请求](#request)。确认具备执行 Python 与写入输出目录的权限。

[Claude Code 官方 Skills 文档](https://code.claude.com/docs/en/skills)规定个人与项目目录。Cowork、网页或云会话不等同于本地 Claude Code，不会因为你复制到本机个人目录就自动获得该技能，需使用各自的技能管理入口。

<a id="codex"></a>
## Codex

按共同步骤复制到 `~/.agents/skills/`；项目内使用可复制到 `.agents/skills/aigc-reducer/`。

新建 Codex 会话，在技能选择器中查找 aigc-reducer，或在请求中明确写“使用 aigc-reducer”，然后给出本地文件路径或可访问的附件。CLI 会话需处于能读取论文的目录；应用的沙箱需要能读取论文并写入目标目录。

[OpenAI 官方技能文档](https://learn.chatgpt.com/docs/build-skills)列出 Codex 本地发现目录。由 CC Switch 生成的 `~/.codex/skills` 是受管理的分发副本，本项目不要求手工覆盖它；使用 CC Switch 时走其分发入口。

<a id="cursor"></a>
## Cursor

按共同步骤复制到 `~/.agents/skills/`，或只在项目中复制到 `.cursor/skills/aigc-reducer/`。在 Cursor 中打开论文所在文件夹，进入 Agent，使用 `/` 菜单搜索技能或发送[示例请求](#request)。

[Cursor 官方文档](https://cursor.com/docs/skills)说明共享与专用目录。这里的个人目录用于本地 Agent；不能假定本机技能和论文会自动出现在 Cloud Agent，云端需另行准备环境与文件。

<a id="gemini-cli"></a>
## Gemini CLI

复制到 `~/.agents/skills/` 或 `~/.gemini/skills/`；项目作用域使用 `.agents/skills/` 或 `.gemini/skills/`。

在论文工作目录启动 Gemini CLI，用 `/skills list` 确认发现本技能，再发送[示例请求](#request)。按宿主提示启用技能和允许需要的工具操作，不必关闭所有权限检查。

[Gemini CLI 官方文档](https://geminicli.com/docs/cli/skills/)列出这些路径；[入门教程](https://geminicli.com/docs/cli/tutorials/skills-getting-started/)给出了发现检查方法。

<a id="github-copilot"></a>
## GitHub Copilot

个人安装复制到 `~/.agents/skills/`；项目安装复制到 `.github/skills/aigc-reducer/`。打开项目，在支持 Agent Skills 的 Copilot Agent 模式中发送[示例请求](#request)，为它提供可访问的文件。

[GitHub 官方 Agent Skills 文档](https://docs.github.com/en/copilot/concepts/customizations/agent-skills)说明适用界面与目录，具体支持取决于客户端版本。云端 Agent 不自动获取本机个人技能或本地文件；不要为了使用技能把真实论文提交到公开仓库。

<a id="opencode"></a>
## OpenCode

复制到 `~/.agents/skills/`；项目安装可放到 `.opencode/skills/aigc-reducer/`，专用个人目录为 `~/.config/opencode/skills/`。

在论文目录启动 OpenCode，发送“加载 aigc-reducer 技能，处理以下文件……”及[示例请求](#request)。技能由宿主的 skill 工具按需加载；发现不到时检查同名覆盖和技能权限。

[OpenCode 官方文档](https://docs.opencode.ai/docs/skills/)列出可发现的目录和权限规则。版本不同可能有不同的显式调用形式，使用自然语言指定技能可避免依赖某一版快捷键。

<a id="openclaw"></a>
## OpenClaw

将目录复制到当前 OpenClaw 工作区的 `skills/aigc-reducer/`；跨工作区共享可放到 `~/.openclaw/skills/aigc-reducer/`。不要将仓库本身的 skills 目录与实际 Agent 工作区混淆。

开启可执行本地 Python、可访问论文目录的会话，发送[示例请求](#request)。如果 Agent 在远程节点上执行，依赖和文件也要放在该节点的授权环境内。

[OpenClaw 官方运行时文档](https://docs.openclaw.ai/agent)说明工作区与共享技能来源；[创建技能文档](https://docs.openclaw.ai/tools/creating-skills)说明目录结构。文件读写和执行能力仍由运行时权限决定。

<a id="request"></a>
## 所有本地 Agent 的通用请求

```text
使用 aigc-reducer。
原稿：“论文.docx”；检测报告：“检测报告.pdf”。
按报告标记处理对应正文，保持数据、引用和结论含义。
改写不调用联网检索，开始前检查可见的模式设置，无法确认时如实提醒一次。
内部完成必要校验，只交付修改后的 Word，不要审计报告、对照表或其他附件。
```

如文件不在当前目录，用实际绝对路径。已明确范围时直接处理，章节歧义再确认。最终应得到新名字的 DOCX；临时 blocks/edits/批次文本不交付。

<a id="manual"></a>
## 网页版 AI 与通用手动流程

ChatGPT、Claude、Gemini、DeepSeek、豆包、Kimi、通义千问等只要能接收文字并回复，就能参与“改写”环节。这里不承诺它们都能安装技能、执行 Python 或直接生成可下载 Word；那些能力取决于具体产品与账号环境。

1. 在本机准备 Python，并在仓库根目录拆文档：

```bash
python skills/aigc-reducer/scripts/docx_tool.py split "论文.docx" -o work/blocks.json
```

2. 在网页版当前对话关闭联网搜索，开启模型支持的深度思考 / 推理模式。对照报告选择批次，生成第 1 批请求，打开文件，把内容完整复制给网页版 AI：

```bash
python skills/aigc-reducer/scripts/docx_tool.py batch work/blocks.json --id 1 --prompt skills/aigc-reducer/references/prompt.md -o work/batch1.txt
```

请求已包含不检索、事实保护及“只返回本批改写段落，按原顺序，段间空一行”的约束，保留完整请求即可。只保存最终改写正文，不复制思考内容。原有五条风格指令保持原文。

3. 将模型回复存为 UTF-8 的 `work/reply1.txt`，再对齐写入：

```bash
python skills/aigc-reducer/scripts/docx_tool.py apply work/blocks.json --batch 1 --input work/reply1.txt -o work/edits.json
```

按同样方式完成所选的其余批次（批号、请求和回复文件名保持一致）。不要只处理第 1 批就当成整篇完成。

4. 先自检再生成 Word：

```bash
python skills/aigc-reducer/scripts/docx_tool.py check "论文.docx" work/edits.json --blocks work/blocks.json
python skills/aigc-reducer/scripts/docx_tool.py write "论文.docx" work/edits.json --blocks work/blocks.json -o "论文_表达整改.docx"
```

check 为 fail 时先修复，不继续 write。自检默认不生成报告，write 默认不生成对照表。最后重新打开输出 Word 核对；手动流程最终产物同样只需交付该 Word。不要把敏感论文粘贴到未认可的数据处理环境。

## 三步确认是否可用

1. **已发现**：宿主技能列表显示 aigc-reducer，或已明确读取实际 SKILL.md。
2. **脚本可运行**：在宿主使用的环境执行 `python "<技能目录>/scripts/docx_tool.py" --help`，能看到六个子命令。
3. **流程可完成**：用 `python "<技能目录>/scripts/make_sample.py" -o "样例.docx"` 生成合成稿，试跑选定一批，并重新打开输出 Word 核对。

发现不到先查目录层级、客户端版本、是否启用与同名副本；脚本错误先查 Python 环境。报告无法读取时提供可复制章节或片段即可，不必先更换模型。

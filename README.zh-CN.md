# PowerPoint MCP Server（中文版）

最好的开源 PowerPoint 自动化 MCP Server。免费、无订阅、无额度限制。

这个项目由一位 UCSD 研究生开发，因为他已经受够了每周花 3-4 小时做 PPT。

**语言**
- [English](README.md)
- [中文](README.zh-CN.md)
- [日本語](README.ja.md)

[![PyPI version](https://badge.fury.io/py/powerpoint-mcp.svg)](https://badge.fury.io/py/powerpoint-mcp)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)

---

## 演示视频

[![Watch the PowerPoint MCP Demo](video_thumbnail.jpg)](https://www.youtube.com/watch?v=5p24Vr36py8)

*观看 Claude 从零构建一份完整的 “Fourier Transform and Fourier Series” 课程讲义：包含公式、图示和动画。*

---

## 仅支持 Windows

**这个 MCP server 只支持 Windows**，因为它使用 `pywin32` 的 COM 自动化来直接控制 PowerPoint 应用。这也是它能够实现双向读写、实时编辑、LaTeX 渲染、模板分析和动画控制的原因。

像 `python-pptx` 这样的 macOS / Linux 替代方案无法访问完整的 PowerPoint 功能集。欢迎提交基于 AppleScript 的 macOS 版本 PR。

---

## 为什么它不一样

市面上已经有不少 PowerPoint 自动化工具，但这个项目真正值得用的原因在于：

**模板优先设计**  
直接把你的目标模板交给 LLM 就能工作。比如你可以告诉 Claude：*“用我们公司的 Nvidia_Black_Green_2025 模板做一份 GPU 性能对比演示文稿”*，它会自动发现模板、分析布局，并正确填充内容。

**真正的双向、实时编辑**  
不同于 `python-pptx` 这种只写不读的方案，也不同于很多热门 MCP 实现，这个项目基于 COM 自动化。Claude 可以**读取**你当前已经打开的 PPT，并在不关闭 PowerPoint 的情况下实时修改。

**多模态幻灯片分析**  
`slide_snapshot` 工具会同时给 LLM 提供视觉上下文（带标注边界框的截图）以及详细的文本 / 图表 / 表格提取结果。Claude 可以真正“看到”你的幻灯片内容。

**真正可用的 LaTeX 渲染**  
服务端通过控制 PowerPoint 内置公式编辑器来渲染 LaTeX。你只需要写 `<latex>E=mc^2</latex>`，它就会作为原生 PowerPoint 公式对象插入。

**节省 token 的 HTML 格式化**  
支持通过 HTML 标签（如 `<b>`、`<i>`、`<red>`、`<ul>`、`<ol>`）实现加粗、斜体、颜色、项目符号。一次填充即可完成格式化，无需反复调用工具逐项设置样式。

**支持渐进式展示的动画**  
支持真正可控的 PowerPoint 动画，包括按段落逐步显示。你可以直接实现 bullet-by-bullet reveal，而不用把内容拆成多个形状。

**一行安装，无订阅**
```bash
claude mcp add powerpoint -- uvx powerpoint-mcp
```
支持 Claude Code、Cursor、GitHub Copilot 以及任意 MCP 客户端。不依赖第三方服务，无月费，无积分限制。

---

## 真实工作流

下面这些都是真实验证过的使用方式：

**研究并生成**
> “研究一下量子计算的最新进展，然后做一份 15 页的演示文稿”

Claude 可以使用 web search 找资料，再生成带引用和格式化内容的 PPT。

**数据分析与可视化**
> “分析 Titanic_dataset.csv，做完整 EDA，并生成一份讲解结果的演示文稿”

支持自由编写 Python / matplotlib 绘图，并直接渲染到幻灯片占位符中。

**代码库文档化**
> “分析我的整个仓库，然后生成一份技术架构演示文稿”

Claude 可以读取本地文件、理解系统结构，并生成讲解架构的幻灯片。

**模板驱动的企业演示**
> “使用 Nvidia_Black_Green_template，根据 nvidia_quarterly_sales_data.csv 生成一份 Q4 销售汇报”

模板布局会自动发现，占位符会被识别并正确填充。

**学术场景 / 重度 LaTeX**
> “做一份 20 页的课程，讲 Fourier Series 和 Fourier Transforms，包含公式”

LaTeX 会渲染为原生 PowerPoint 公式对象，还可加动画做渐进讲解。

**交互式学习**
> “帮我理解这份关于 PAM 和 BLOSUM 矩阵的计算生物学课件，逐页解释并在每个章节后给我出题”

Claude 可以读取幻灯片（包括图表、表格和讲者备注），解释内容并生成互动问答。

---

## 快速开始

**第 1 步：安装**
```bash
claude mcp add powerpoint -- uvx powerpoint-mcp
```

**第 2 步：打开 PowerPoint**
打开 Microsoft PowerPoint（任意演示文稿都可以，或者停留在空白启动页也行）。

**第 3 步：给 Claude 下指令**
```text
用我的公司模板做一份 10 页的机器学习基础介绍。
在合适的位置加入公式，并给每页要点加上逐段动画。
```

就这样，剩下的交给 Claude。

---

## 安装

### 前提条件
- Windows 10/11
- 已安装 Microsoft PowerPoint
- Python 3.10+

### Claude Code
```bash
# 单个项目
claude mcp add powerpoint -- uvx powerpoint-mcp

# 全局可用（安装一次，所有项目可用）
claude mcp add powerpoint --scope user -- uvx powerpoint-mcp
```

### Cursor
1. 点击 Settings -> Tools & MCP -> New MCP Server
2. 打开 `~/.cursor/mcp.json`
3. 添加：
```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["powerpoint-mcp"]
    }
  }
}
```
4. 重启 IDE。

### Trae
1. 在项目根目录创建 `.trae/mcp.json`
2. 添加：
```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["powerpoint-mcp"]
    }
  }
}
```
3. 保存后重启 Trae，并在支持 MCP 的 Agent 模式中使用

### Trae CN
1. 当前 Trae 中文版可使用同样的项目级 `.trae/mcp.json` 配置方式
2. 在项目根目录的 `.trae/mcp.json` 中添加：
```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["powerpoint-mcp"]
    }
  }
}
```
3. 保存后重启 Trae 中文版，并使用支持 MCP 的智能体模式；如果你的版本里有 `Builder with MCP` 或 `SOLO Coder`，优先使用它们

### Codex
由于这个 server 仅支持 Windows，在 Windows PowerShell 中可直接运行：
```bash
codex mcp add powerpoint -- cmd /c uvx powerpoint-mcp
```

添加后，重启 Codex CLI，或执行 `/mcp` 检查服务是否已生效。

### CodeBuddy
1. 在侧栏对话面板中点击右上角的 CodeBuddy Settings
2. 切换到 MCP 标签页
3. 点击 Add MCP
4. 添加以下 JSON：
```json
{
  "mcpServers": {
    "powerpoint": {
      "type": "stdio",
      "command": "uvx",
      "args": ["powerpoint-mcp"],
      "description": "PowerPoint automation on Windows"
    }
  }
}
```
5. 保存后可通过 Try to Run 验证，也可以直接在 Agent 中调用

### CodeBuddy CN
1. 在侧栏对话面板右上方点击 CodeBuddy Settings
2. 切换到 MCP 标签页
3. 点击 Add MCP
4. 添加以下 JSON：
```json
{
  "mcpServers": {
    "powerpoint": {
      "type": "stdio",
      "command": "uvx",
      "args": ["powerpoint-mcp"],
      "description": "PowerPoint automation on Windows"
    }
  }
}
```
5. 保存配置后，可通过 Try to Run 验证，或直接在 Agent 中调用

### OpenCode
1. 打开 `~/.config/opencode/opencode.json`（全局）或项目根目录下的 `opencode.json`
2. 添加：
```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "powerpoint": {
      "type": "local",
      "command": ["uvx", "powerpoint-mcp"]
    }
  }
}
```
3. 重启 OpenCode。

### VS Code（GitHub Copilot）
1. 打开 `C:\Users\Your_User_Name\AppData\Roaming\Code\User\mcp.json`
2. 添加：
```json
{
  "mcpServers": {
    "powerpoint": {
      "command": "uvx",
      "args": ["powerpoint-mcp"]
    }
  }
}
```
3. 重启 IDE。

---

## 工具参考

共 11 个工具，每个职责都很明确：

| Tool | 功能 |
|------|------|
| `manage_presentation` | 打开、关闭、创建或保存演示文稿，是所有工作流的入口。 |
| `slide_snapshot` | 综合分析幻灯片：截图、带标注边界框、文本提取、图表数据、表格解析、超链接、评论和备注。 |
| `switch_slide` | 切换到当前演示文稿中的指定页。 |
| `add_speaker_notes` | 使用官方 COM 方式为指定页添加或替换讲者备注。 |
| `populate_placeholder` | 向任意占位符填充文本（支持 HTML 格式）、LaTeX 公式或 matplotlib 图表。 |
| `manage_slide` | 在演示文稿中复制、删除或移动幻灯片。 |
| `list_templates` | 从 Personal、User、System 模板目录中发现可用 PowerPoint 模板。 |
| `analyze_template` | 分析模板中的各个 layout、placeholder 位置及截图。 |
| `add_slide_with_layout` | 使用模板中的特定 layout 新增幻灯片，并保留模板样式。 |
| `add_animation` | 给形状添加入场动画，支持按段落渐进式显示。 |
| `powerpoint_evaluate` | 在 PowerPoint 上下文中执行任意 Python 代码，用于复杂批处理。 |

---

## 内容格式参考

`populate_placeholder` 工具支持以下 HTML 风格标签：

```text
<b>bold text</b>
<i>italic text</i>
<u>underlined text</u>
<red>colored text</red>  (也支持: blue, green, orange, purple, white, black)
<ul><li>unordered list item</li></ul>
<ol><li>ordered list item</li></ol>
<latex>E = mc^2</latex>
```

如果要插入 matplotlib 图表，可以把绘图代码作为 `content` 传入，并设置 `content_type="plot"`。服务端会执行代码并插入生成的图片。

---

## 工作原理

PowerPoint MCP 使用 Windows COM 自动化（`pywin32`）直接控制 Microsoft PowerPoint，这与 VBA 宏和 Office 插件使用的是同一套机制。因此它可以访问 PowerPoint 的完整能力：公式编辑器、动画系统、模板系统，以及实时读取幻灯片内容。

服务器基于 [FastMCP](https://github.com/jlowin/fastmcp) 构建，并发布在 [PyPI](https://pypi.org/project/powerpoint-mcp/) 上，因此可以通过 `uvx` 轻松安装。

截图默认会保存到 `~/.powerpoint-mcp/`，文件名基于时间戳自动生成。

---

## 限制

- **仅支持 Windows**：COM 自动化是 Windows 技术，目前不支持 macOS 或 Linux。
- **必须安装 PowerPoint**：不支持 LibreOffice。
- **需要 Python 3.10+**：运行服务端所必需。

---

## Star History

[![Star History Chart](https://api.star-history.com/svg?repos=ayushmaniar/powerpoint-mcp&type=date&legend=top-left)](https://www.star-history.com/#ayushmaniar/powerpoint-mcp&type=date&legend=top-left)

---

完全开源（MIT License）。如果这个项目为你节省了时间，欢迎去 GitHub 点个 star。

**GitHub:** https://github.com/Ayushmaniar/powerpoint-mcp

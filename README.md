# WPS MCP Server · WPS 办公套件的 AI 原生操作

> 与其手动打开 WPS 慢慢编辑，不如让 AI 直接帮你搞定文档、表格、幻灯片

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Platform: Claude Code](https://img.shields.io/badge/Platform-Claude%20Code-4B61FF?logo=anthropic)](https://claude.com)
[![Platform: Cursor](https://img.shields.io/badge/Platform-Cursor-7B42FF?logo=cursor)](https://cursor.com)
[![Platform: WPS Office](https://img.shields.io/badge/Platform-WPS%20Office-FF6B35?logo=wps)](https://www.wps.cn/)

---

## 中文 | [English](#english)

### 这是什么？

一个让 AI Agent **原生操作 WPS Office** 的 MCP Server。

不用手动打开 WPS，不用慢慢排版。直接告诉 AI 你想做什么，它帮你完成。

> **注意**：WPS 和 Microsoft Office 使用相同的文件格式（.docx、.xlsx、.pptx），所以这个服务器**同时兼容 WPS 和 Office**！

### 问题

```
你日常做的事                AI 帮你做的事
─────────────────────────────────────────────────
"帮我创建个周报文档"       → 直接生成 .docx 文件
"在表格里填入这些数据"     → 自动写入 Excel
"做个产品介绍 PPT"         → 自动生成幻灯片
```

**问题在哪？**

> AI 只能"建议"你做什么，而不是"替"你做什么。
> 你还是要自己动手打开软件、复制粘贴。

### 解决

```python
# 用自然语言操作 WPS
create_document("周报", "本周工作总结")    # 创建文档
add_heading("周报", "工作内容", 2)       # 添加标题
write_cell("销售数据", "A1", "100万")    # 写入表格
create_presentation("产品介绍", "XX产品") # 创建 PPT
```

### 支持的工具

#### 📄 文档 (Word)
| 工具 | 功能 |
|------|------|
| `create_document` | 创建新文档 |
| `add_heading` | 添加标题 |
| `add_paragraph` | 添加段落 |
| `read_document` | 读取文档内容 |

#### 📊 表格 (Excel)
| 工具 | 功能 |
|------|------|
| `create_spreadsheet` | 创建新表格 |
| `write_cell` | 写入单元格 |
| `read_cell` | 读取单元格 |
| `add_formula` | 添加公式 |

#### 📽️ 演示 (PowerPoint)
| 工具 | 功能 |
|------|------|
| `create_presentation` | 创建新演示文稿 |
| `add_slide` | 添加幻灯片 |
| `add_text_to_slide` | 添加文本框 |
| `set_slide_layout` | 设置布局 |

### 快速开始

#### 1. 安装依赖

```bash
pip install mcp python-docx openpyxl python-pptx
```

#### 2. 配置环境变量（可选）

```bash
# 设置工作目录，默认是 ~/wps_projects
export WPS_WORKSPACE="/path/to/your/folder"
```

#### 3. 在 Claude Code 中使用

```json
{
  "mcpServers": {
    "wps": {
      "command": "python",
      "args": ["/path/to/wps_mcp.py"]
    }
  }
}
```

#### 4. 开始对话

```bash
# 启动 Claude Code 后，直接说：
"帮我创建一份周报文档"
"在表格里填入本月销售数据"
"做一个 5 页的产品介绍 PPT"
```

### 使用示例

#### 创建文档
```
"帮我创建一个项目报告文档"
AI → create_document("项目报告", "项目报告") → 返回文件路径
```

#### 操作表格
```
"在销售数据.xlsx 的 A1 写入 100万"
AI → write_cell("销售数据", "A1", "100万") → 完成

"帮我计算 A1 到 A10 的总和"
AI → add_formula("销售数据", "A11", "=SUM(A1:A10)") → 完成
```

#### 制作 PPT
```
"创建一个公司介绍 PPT"
AI → create_presentation("公司介绍", "XX公司") → 创建完成

"添加一页新幻灯片"
AI → add_slide("公司介绍", 1) → 幻灯片添加成功
```

### 为什么选 WPS？

| 优势 | 说明 |
|------|------|
| 🇨🇳 **中国市场** | WPS 在中国有 1 亿日活设备 |
| 🏛️ **政府采购** | 政府机关首选 WPS |
| 📄 **格式兼容** | 同时支持 .docx、.xlsx、.pptx |
| 🔧 **易于使用** | python-docx/openpyxl/pptx 生态成熟 |

### 技术栈

- Python 3.10+
- [FastMCP](https://modelcontextprotocol.io/) - MCP 框架
- [python-docx](https://python-docx.readthedocs.io/) - Word 文档
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel 表格
- [python-pptx](https://python-pptx.readthedocs.io/) - PPT 幻灯片

---

### License

MIT

---

## English

An MCP server for WPS Office operations - enables AI agents to create, edit, and manage documents, spreadsheets, and presentations via natural language.

> **Note**: WPS and Microsoft Office use the same file formats (.docx, .xlsx, .pptx), so this server works with **both WPS and MS Office**!

### Features

- **Documents** - Create and edit Word (.docx) files
- **Spreadsheets** - Work with Excel (.xlsx) files
- **Presentations** - Build PowerPoint (.pptx) slideshows

### Quick Start

```bash
pip install mcp python-docx openpyxl python-pptx
python wps_mcp.py
```

Configure in Claude Code and start chatting!

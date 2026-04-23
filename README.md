# wecom-docs-mcp-server

[![MCP](https://img.shields.io/badge/MCP-2024--11--05-blue)](https://modelcontextprotocol.io)
[![Python](https://img.shields.io/badge/python-3.9+-green)](https://python.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow)](LICENSE)

MCP server for **WeCom (企业微信 / Enterprise WeChat) document operations** — create, read, and edit WeCom Docs and Smartsheets from any MCP-compatible AI agent.

> **Gap this fills**: Existing WeCom MCP servers ([wecom-bot-mcp-server](https://github.com/loonghao/wecom-bot-mcp-server), [wecom-mcp](https://github.com/code-tinker/wecom-mcp)) only support **sending messages via webhook**. This server provides **document CRUD operations** via the `@wecom/cli` document API.

Compatible with: [Claude Desktop](https://claude.ai/download) · [Hermes Agent](https://github.com/NousResearch/hermes-agent) · [Cursor](https://cursor.com) · any MCP client

---

## Tools (9)

| Tool | Description |
|------|-------------|
| `wecom_read_doc` | Read a WeCom Doc or Smartsheet → Markdown |
| `wecom_create_doc` | Create a new Doc (type 3) or Smartsheet (type 10) |
| `wecom_edit_doc` | Write Markdown content to a doc |
| `wecom_smartsheet_setup_fields` | Initialize a smartsheet's column schema |
| `wecom_smartsheet_add_records` | Append rows (auto cell-format conversion) |
| `wecom_smartsheet_get_sheet` | List sheets in a smartsheet |
| `wecom_smartsheet_get_fields` | Get column definitions |
| `wecom_smartsheet_get_records` | Fetch all rows from a sheet |
| `wecom_get_doc_content` | Async-poll full doc content as Markdown |

---

## Why This Exists (The Silent Failure Problem)

Calling `wecom-cli` from WSL via bash hits a three-layer failure:

1. **UNC path issue** — `cmd.exe` launches with wrong working directory under WSL
2. **Exit code lies** — process returns 0 but writes nothing
3. **Chinese encoding** — shell argument boundary corrupts multi-byte characters

**Fix**: call `node.exe` directly via `subprocess.run([...], capture_output=True)` — no shell, args pass through `CreateProcessW`, Chinese content works correctly.

```python
# ❌ silently fails from WSL
os.system(f'wecom-cli doc edit_doc_content --json "{json_content}"')

# ✅ works
subprocess.run(["node.exe", WECOM_JS, "doc", "edit_doc_content", "--json", json_content],
               capture_output=True)
```

---

## Requirements

- Python 3.9+
- Node.js 18+ (Windows-side if running from WSL)
- `@wecom/cli` authenticated:
  ```bash
  npm install -g @wecom/cli
  npx @wecom/cli auth   # scan QR code with WeCom mobile app
  ```

---

## Installation

### Option A — pip

```bash
pip install wecom-docs-mcp-server
```

### Option B — clone

```bash
git clone https://github.com/Beltran12138/wecom-docs-mcp-server
cd wecom-docs-mcp-server
```

---

## Configuration

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `WECOM_CLI_PATH` | auto-detected | Full path to `wecom.js` (override if auto-detect fails) |
| `NODE_EXE` | `node.exe` | Node.js executable name (`node` on Linux/macOS) |

Auto-detection order: `WECOM_CLI_PATH` env → `npm root -g` → common fallbacks.

### Claude Desktop (`claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "wecom-docs": {
      "command": "python",
      "args": ["-m", "wecom_docs_mcp_server"]
    }
  }
}
```

Or if cloned:
```json
{
  "mcpServers": {
    "wecom-docs": {
      "command": "python",
      "args": ["/path/to/wecom-docs-mcp-server/wecom_docs_mcp_server.py"]
    }
  }
}
```

### Hermes Agent

```bash
hermes mcp add wecom-docs \
  --command python3 \
  --args /path/to/wecom_docs_mcp_server.py
```

### WSL + Windows Node (non-standard path)

```bash
export WECOM_CLI_PATH="C:/Users/YOUR_USER/AppData/Roaming/npm/node_modules/@wecom/cli/bin/wecom.js"
export NODE_EXE="node.exe"
```

---

## Usage Examples

### Create a document and write content

```
User: Create a WeCom doc called "Q2 Market Report" and write a brief summary

Agent calls:
1. wecom_create_doc(doc_name="Q2 Market Report", doc_type=3)
   → {"url": "https://doc.weixin.qq.com/doc/xxx", "docid": "abc123"}

2. wecom_edit_doc(content="# Q2 Market Report\n\n...", docid="abc123")
   → {"errcode": 0, "errmsg": "ok"}
```

### Create a smartsheet with structured data

```
User: Create a competitor tracking table with columns: Company, Product, Funding, Notes

Agent calls:
1. wecom_create_doc(doc_name="Competitor Tracker", doc_type=10)
   → {"url": "...", "docid": "xyz789"}

2. wecom_smartsheet_get_sheet(url="...")
   → {"sheet_list": [{"sheet_id": "s001", "title": "Sheet1"}]}

3. wecom_smartsheet_setup_fields(sheet_id="s001", field_names=["Company", "Product", "Funding", "Notes"])

4. wecom_smartsheet_add_records(sheet_id="s001", records=[
     {"Company": "HashKey", "Product": "HashKey Exchange", "Funding": "$200M", "Notes": "SFC licensed"},
     {"Company": "OSL", "Product": "OSL Exchange", "Funding": "$150M", "Notes": "SFC licensed"}
   ])
```

### Read an existing document

```
User: What's in this doc? https://doc.weixin.qq.com/doc/abc?scode=xyz

Agent calls:
wecom_read_doc(url="https://doc.weixin.qq.com/doc/abc?scode=xyz")
→ Returns full Markdown content
```

---

## Troubleshooting

**`node.exe` not found**
```bash
# WSL: ensure Windows Node is on PATH
export PATH=$PATH:/mnt/c/Program\ Files/nodejs
```

**Empty doc after create+edit**
Set `WECOM_CLI_PATH` explicitly — auto-detection may have picked wrong `wecom.js`.

**Error 851008** — Bot lacks "get member document content" permission. Re-authorize in WeCom admin console → Application Management → your app → Permissions.

**Error 851014** — Document permission expired. Same re-authorization flow.

---

## Related Projects

| Project | Focus |
|---------|-------|
| [wecom-bot-mcp-server](https://github.com/loonghao/wecom-bot-mcp-server) | Bot messaging via webhook |
| [wecom-mcp](https://github.com/code-tinker/wecom-mcp) | Send messages/files via webhook |
| **wecom-docs-mcp-server** (this) | **Document CRUD: create/read/edit Docs & Smartsheets** |

---

## License

MIT

---

<details>
<summary>中文说明</summary>

## 简介

`wecom-docs-mcp-server` 是一个 MCP stdio 服务器，将企业微信文档操作（创建、读取、编辑文档和智能表格）暴露为 MCP 工具，可在 Claude Desktop、Hermes Agent、Cursor 等任何 MCP 客户端中使用。

**与已有企微 MCP 项目的区别**：现有项目（wecom-bot-mcp-server、wecom-mcp）均只支持通过 Webhook 发送消息；本项目专注于**文档 CRUD 操作**，使用 `@wecom/cli` 文档 API。

## 核心技术问题

从 WSL 通过 bash 调用 `wecom-cli` 会**静默失败**（exit code 0，但文档为空）。原因是 UNC 路径问题 + cmd.exe 工作目录错误 + 中文编码失败三层叠加。

解决方案：直接用 `subprocess.run(["node.exe", ...])` 调用，绕过 shell，通过 `CreateProcessW` 传参，中文内容完全正常。

## 安装

```bash
pip install wecom-docs-mcp-server
```

## 快速配置（Claude Desktop）

```json
{
  "mcpServers": {
    "wecom-docs": {
      "command": "python",
      "args": ["-m", "wecom_docs_mcp_server"]
    }
  }
}
```

</details>

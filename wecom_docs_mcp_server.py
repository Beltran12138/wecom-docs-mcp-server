#!/usr/bin/env python3
"""
wecom-docs-mcp-server — WeCom (Enterprise WeChat) document operations as MCP tools.

Exposes 9 tools for creating, reading, and editing WeCom Docs and Smartsheets
via the Model Context Protocol (stdio JSON-RPC 2.0).

Requires: @wecom/cli  (npm install -g @wecom/cli && npx @wecom/cli auth)
"""

import json
import os
import subprocess
import sys
import time


def _find_wecom_js() -> str:
    """Locate wecom.js: env var → npm root -g → common fallbacks."""
    env = os.environ.get("WECOM_CLI_PATH")
    if env and os.path.isfile(env):
        return env

    # Try npm root -g to find global node_modules
    node_exe = os.environ.get("NODE_EXE", "node.exe")
    try:
        npm = "npm.cmd" if sys.platform == "win32" else "npm"
        result = subprocess.run([npm, "root", "-g"], capture_output=True, timeout=10)
        npm_root = result.stdout.decode().strip()
        candidate = os.path.join(npm_root, "@wecom", "cli", "bin", "wecom.js").replace("\\", "/")
        if os.path.isfile(candidate):
            return candidate
    except Exception:
        pass

    # WSL: check Windows npm root via cmd
    if sys.platform != "win32":
        try:
            result = subprocess.run(
                ["cmd.exe", "/c", "npm root -g"],
                capture_output=True, timeout=10
            )
            npm_root = result.stdout.decode("utf-8", errors="replace").strip()
            candidate = os.path.join(npm_root, "@wecom", "cli", "bin", "wecom.js").replace("\\", "/")
            if os.path.isfile(candidate):
                return candidate
        except Exception:
            pass

    # Common fallbacks
    fallbacks = [
        os.path.expanduser("~/AppData/Roaming/npm/node_modules/@wecom/cli/bin/wecom.js"),
        "/usr/local/lib/node_modules/@wecom/cli/bin/wecom.js",
        "/usr/lib/node_modules/@wecom/cli/bin/wecom.js",
    ]
    for path in fallbacks:
        if os.path.isfile(path):
            return path

    raise FileNotFoundError(
        "wecom.js not found. Set WECOM_CLI_PATH env var or install: npm install -g @wecom/cli"
    )


_WECOM_JS = None  # lazy-loaded on first call
_NODE_EXE = os.environ.get("NODE_EXE", "node.exe")

TOOLS = [
    {
        "name": "wecom_read_doc",
        "description": "Read a WeCom document or smartsheet. Returns content as Markdown. Auto-detects URL type: /smartsheet/ URLs return table data, /doc/ URLs return document content.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "url": {"type": "string", "description": "Full WeCom doc or smartsheet URL (include scode param if present)"}
            },
            "required": ["url"]
        }
    },
    {
        "name": "wecom_create_doc",
        "description": "Create a new WeCom document or smartsheet. Returns url and docid — save the docid for subsequent edits.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_name": {"type": "string", "description": "Document name (max 255 chars)"},
                "doc_type": {"type": "integer", "enum": [3, 10], "description": "3 = regular document, 10 = smartsheet"}
            },
            "required": ["doc_name", "doc_type"]
        }
    },
    {
        "name": "wecom_edit_doc",
        "description": "Write Markdown content to a WeCom document. Supports headings, lists, tables, bold, italic. Use docid (preferred) or url to identify the document.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "content": {"type": "string", "description": "Markdown content to write"},
                "docid": {"type": "string", "description": "Document docid from wecom_create_doc (preferred)"},
                "url": {"type": "string", "description": "Document URL (fallback if docid unavailable)"}
            },
            "required": ["content"]
        }
    },
    {
        "name": "wecom_smartsheet_setup_fields",
        "description": "Initialize a smartsheet's column schema. Renames the default field and adds remaining fields. Must be called before adding records to a new sheet.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "url": {"type": "string", "description": "Smartsheet URL (or use docid)"},
                "docid": {"type": "string", "description": "Smartsheet docid (or use url)"},
                "sheet_id": {"type": "string", "description": "Sheet ID from wecom_smartsheet_get_sheet"},
                "field_names": {"type": "array", "items": {"type": "string"}, "description": "Column names in order"}
            },
            "required": ["sheet_id", "field_names"]
        }
    },
    {
        "name": "wecom_smartsheet_add_records",
        "description": "Append rows to a smartsheet. Each record is a plain {column_name: value} dict — cell format conversion is handled automatically.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "url": {"type": "string", "description": "Smartsheet URL"},
                "docid": {"type": "string", "description": "Smartsheet docid"},
                "sheet_id": {"type": "string", "description": "Sheet ID"},
                "records": {
                    "type": "array",
                    "items": {"type": "object"},
                    "description": "List of {column_name: value} objects"
                }
            },
            "required": ["sheet_id", "records"]
        }
    },
    {
        "name": "wecom_smartsheet_get_sheet",
        "description": "List all sheets (sub-tables) in a WeCom smartsheet. Returns sheet IDs and titles.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "url": {"type": "string", "description": "Smartsheet URL"}
            },
            "required": ["url"]
        }
    },
    {
        "name": "wecom_smartsheet_get_fields",
        "description": "Get column definitions (field names, types, IDs) for a smartsheet sheet.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "url": {"type": "string", "description": "Smartsheet URL"},
                "sheet_id": {"type": "string", "description": "Sheet ID"}
            },
            "required": ["url", "sheet_id"]
        }
    },
    {
        "name": "wecom_smartsheet_get_records",
        "description": "Fetch all rows from a smartsheet sheet. Returns structured row data.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "url": {"type": "string", "description": "Smartsheet URL"},
                "sheet_id": {"type": "string", "description": "Sheet ID"}
            },
            "required": ["url", "sheet_id"]
        }
    },
    {
        "name": "wecom_get_doc_content",
        "description": "Fetch the full content of a WeCom online doc as Markdown. Uses async polling internally.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "url": {"type": "string", "description": "Document URL"},
                "task_id": {"type": "string", "description": "Polling task_id (omit on first call)"}
            },
            "required": ["url"]
        }
    }
]


def wecom_call(subcommand: str, args_dict: dict, timeout: int = 60) -> dict:
    """
    Call wecom-cli via node.exe directly — no shell, no encoding issues.

    Why node.exe directly instead of bash/cmd:
    From WSL, calling wecom-cli.exe via bash hits a UNC path issue: cmd.exe
    launches with wrong working directory, causing silent failure (exit 0,
    empty output). Chinese content adds a second encoding failure at the
    shell argument boundary. Direct subprocess call passes args via
    CreateProcessW — no shell interpolation, Chinese content works correctly.
    """
    global _WECOM_JS
    if _WECOM_JS is None:
        _WECOM_JS = _find_wecom_js()

    json_content = json.dumps(args_dict, ensure_ascii=False, separators=(",", ":"))
    try:
        result = subprocess.run(
            [_NODE_EXE, _WECOM_JS, "doc", subcommand, "--json", json_content],
            capture_output=True,
            timeout=timeout
        )
        raw = result.stdout.strip()
        if not raw:
            err = result.stderr.decode("utf-8", errors="replace").strip()
            return {"error": err or "no output from wecom-cli"}
        output = raw.decode("utf-8", errors="replace")
        mcp = json.loads(output)
        text = mcp.get("result", {}).get("content", [{}])[0].get("text", "{}")
        if not text or text == "{}":
            text = mcp.get("content", [{}])[0].get("text", "{}")
        return json.loads(text)
    except Exception as e:
        return {"error": str(e)}


def read_smartsheet(url: str) -> str:
    resp = wecom_call("smartsheet_get_sheet", {"url": url})
    if resp.get("errcode", 0) != 0:
        return f"Error {resp.get('errcode')}: {resp.get('errmsg', '')}"
    sheets = resp.get("sheet_list", [])
    if not sheets:
        return "No sheets found"
    lines = [f"# Smartsheet\nSource: {url}\n"]
    for sheet in sheets[:5]:
        sheet_id = sheet.get("sheet_id") or sheet.get("id")
        sheet_title = sheet.get("title", sheet_id)
        fields_resp = wecom_call("smartsheet_get_fields", {"url": url, "sheet_id": sheet_id})
        fields = fields_resp.get("fields", [])
        field_names = [f.get("field_title", f.get("title", "?")) for f in fields]
        records_resp = wecom_call("smartsheet_get_records", {"url": url, "sheet_id": sheet_id})
        records = records_resp.get("records", [])
        lines.append(f"\n## {sheet_title}\n")
        if field_names:
            lines.append("| " + " | ".join(field_names) + " |")
            lines.append("|" + "---|" * len(field_names))
        for rec in records[:200]:
            row_data = rec.get("values", rec.get("record", rec.get("fields", {})))
            row = []
            for fname in field_names:
                val = row_data.get(fname, "")
                if isinstance(val, list):
                    parts = []
                    for item in val:
                        if isinstance(item, dict):
                            parts.append(item.get("text", ""))
                        else:
                            parts.append(str(item))
                    val = "".join(parts)
                elif isinstance(val, str) and val.isdigit() and len(val) == 13:
                    from datetime import datetime, timezone
                    val = datetime.fromtimestamp(int(val) / 1000, tz=timezone.utc).strftime("%Y-%m-%d")
                row.append(str(val).replace("|", "｜").replace("\n", " ").strip())
            lines.append("| " + " | ".join(row) + " |")
    return "\n".join(lines)


def read_doc(url: str) -> str:
    resp = wecom_call("get_doc_content", {"url": url, "type": 2})
    if resp.get("errcode", 0) != 0:
        code = resp.get("errcode", 0)
        if code == 851008:
            return "Error 851008: Bot lacks 'get member document content' permission. Re-authorize in WeCom admin console."
        if code == 851014:
            return "Error 851014: Document permission expired. Re-authorize in WeCom admin console."
        return f"Error {code}: {resp.get('errmsg', '')}"
    if resp.get("task_done"):
        return resp.get("content", "")
    task_id = resp.get("task_id", "")
    if not task_id:
        return f"No task_id in response: {resp}"
    for _ in range(10):
        time.sleep(2)
        resp = wecom_call("get_doc_content", {"url": url, "type": 2, "task_id": task_id})
        if resp.get("errcode", 0) != 0:
            return f"Error {resp.get('errcode')}: {resp.get('errmsg', '')}"
        if resp.get("task_done"):
            return resp.get("content", "")
    return "Timeout waiting for doc content"


def setup_sheet_fields(loc: dict, sheet_id: str, field_names: list) -> str:
    fields_resp = wecom_call("smartsheet_get_fields", {**loc, "sheet_id": sheet_id})
    if fields_resp.get("errcode", 0) != 0:
        return f"get_fields error: {fields_resp.get('errmsg', '')}"
    fields = fields_resp.get("fields", [])
    if not fields:
        return "No default field found"
    default_field_id = fields[0].get("field_id")
    default_field_type = fields[0].get("field_type", "FIELD_TYPE_TEXT")
    update_resp = wecom_call("smartsheet_update_fields", {
        **loc,
        "sheet_id": sheet_id,
        "fields": [{"field_id": default_field_id, "field_title": field_names[0], "field_type": default_field_type}]
    })
    if update_resp.get("errcode", 0) != 0:
        return f"update_fields error: {update_resp.get('errmsg', '')}"
    if len(field_names) > 1:
        add_resp = wecom_call("smartsheet_add_fields", {
            **loc,
            "sheet_id": sheet_id,
            "fields": [{"field_title": name, "field_type": "FIELD_TYPE_TEXT"} for name in field_names[1:]]
        })
        if add_resp.get("errcode", 0) != 0:
            return f"add_fields error: {add_resp.get('errmsg', '')}"
    return f"OK: fields set to {field_names}"


def handle_tool_call(name: str, args: dict) -> str:
    if name == "wecom_read_doc":
        url = args.get("url", "")
        if "/smartsheet/" in url:
            return read_smartsheet(url)
        return read_doc(url)

    elif name == "wecom_create_doc":
        resp = wecom_call("create_doc", {
            "doc_type": args.get("doc_type", 3),
            "doc_name": args.get("doc_name", "New Document")
        })
        return json.dumps(resp, ensure_ascii=False)

    elif name == "wecom_edit_doc":
        call_args = {"content": args.get("content", ""), "content_type": 1}
        if args.get("docid"):
            call_args["docid"] = args["docid"]
        elif args.get("url"):
            call_args["url"] = args["url"]
        resp = wecom_call("edit_doc_content", call_args)
        return json.dumps(resp, ensure_ascii=False)

    elif name == "wecom_smartsheet_add_records":
        loc = {}
        if args.get("url"):
            loc["url"] = args["url"]
        elif args.get("docid"):
            loc["docid"] = args["docid"]

        def to_cell(v):
            if isinstance(v, (int, float, bool)):
                return v
            if isinstance(v, list):
                return v
            return [{"type": "text", "text": str(v)}]

        raw_records = args.get("records", [])
        api_records = []
        for rec in raw_records:
            if "values" in rec:
                api_records.append(rec)
            else:
                api_records.append({"values": {k: to_cell(v) for k, v in rec.items()}})
        resp = wecom_call("smartsheet_add_records", {
            **loc,
            "sheet_id": args.get("sheet_id", ""),
            "records": api_records
        })
        return json.dumps(resp, ensure_ascii=False)

    elif name == "wecom_smartsheet_setup_fields":
        loc = {}
        if args.get("url"):
            loc["url"] = args["url"]
        elif args.get("docid"):
            loc["docid"] = args["docid"]
        return setup_sheet_fields(loc, args.get("sheet_id", ""), args.get("field_names", []))

    elif name == "wecom_get_doc_content":
        url = args.get("url", "")
        task_id = args.get("task_id")
        call_args = {"url": url, "type": 2}
        if task_id:
            call_args["task_id"] = task_id
        resp = wecom_call("get_doc_content", call_args)
        return json.dumps(resp, ensure_ascii=False)

    elif name == "wecom_smartsheet_get_records":
        resp = wecom_call("smartsheet_get_records", {
            "url": args.get("url", ""),
            "sheet_id": args.get("sheet_id", "")
        })
        return json.dumps(resp, ensure_ascii=False)

    elif name == "wecom_smartsheet_get_sheet":
        resp = wecom_call("smartsheet_get_sheet", {"url": args.get("url", "")})
        return json.dumps(resp, ensure_ascii=False)

    elif name == "wecom_smartsheet_get_fields":
        resp = wecom_call("smartsheet_get_fields", {
            "url": args.get("url", ""),
            "sheet_id": args.get("sheet_id", "")
        })
        return json.dumps(resp, ensure_ascii=False)

    else:
        return f"Unknown tool: {name}"


def send(obj: dict):
    sys.stdout.write(json.dumps(obj) + "\n")
    sys.stdout.flush()


def main():
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except json.JSONDecodeError:
            continue

        req_id = req.get("id")
        method = req.get("method", "")

        if method == "initialize":
            send({
                "jsonrpc": "2.0", "id": req_id,
                "result": {
                    "protocolVersion": "2024-11-05",
                    "capabilities": {"tools": {}},
                    "serverInfo": {"name": "wecom-docs-mcp-server", "version": "1.0.0"}
                }
            })
        elif method == "tools/list":
            send({"jsonrpc": "2.0", "id": req_id, "result": {"tools": TOOLS}})
        elif method == "tools/call":
            params = req.get("params", {})
            tool_name = params.get("name", "")
            tool_args = params.get("arguments", {})
            content = handle_tool_call(tool_name, tool_args)
            send({
                "jsonrpc": "2.0", "id": req_id,
                "result": {"content": [{"type": "text", "text": content}]}
            })
        elif method == "notifications/initialized":
            pass
        else:
            if req_id is not None:
                send({"jsonrpc": "2.0", "id": req_id,
                      "error": {"code": -32601, "message": f"Method not found: {method}"}})


if __name__ == "__main__":
    main()

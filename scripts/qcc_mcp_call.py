#!/usr/bin/env python3
"""Reliable QCC MCP discovery and tool-call fallback.

This helper is used only when Codex has the QCC servers configured but the
current task exposes zero QCC tools. It never prints or persists the bearer
token. Network access may still require the host application's approval.
"""

from __future__ import annotations

import argparse
import http.cookiejar
import json
import os
import sys
import tomllib
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any


DEFAULT_CONFIG = Path(
    os.environ.get("CODEX_CONFIG", Path.home() / ".codex" / "config.toml")
)
PROTOCOL_VERSION = "2025-03-26"


class QccMcpError(RuntimeError):
    pass


def parse_sse(body: str) -> dict[str, Any]:
    payloads: list[dict[str, Any]] = []
    for line in body.splitlines():
        if not line.startswith("data:"):
            continue
        raw = line[5:].strip()
        if not raw:
            continue
        try:
            value = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise QccMcpError("QCC MCP returned malformed SSE JSON") from exc
        if isinstance(value, dict):
            payloads.append(value)
    if not payloads:
        raise QccMcpError("QCC MCP returned no SSE data event")
    response = payloads[-1]
    if "error" in response:
        raise QccMcpError(f"QCC MCP error: {json.dumps(response['error'], ensure_ascii=False)}")
    return response


def load_server(config_path: Path, server_name: str) -> tuple[str, str]:
    if not config_path.exists():
        raise QccMcpError(f"Codex config not found: {config_path}")
    with config_path.open("rb") as handle:
        config = tomllib.load(handle)
    servers = config.get("mcp_servers", {})
    server = servers.get(server_name)
    if not isinstance(server, dict):
        raise QccMcpError(f"QCC MCP server is not configured: {server_name}")
    if server.get("enabled") is False:
        raise QccMcpError(f"QCC MCP server is disabled: {server_name}")
    url = server.get("url")
    token_env = server.get("bearer_token_env_var")
    if not isinstance(url, str) or not url:
        raise QccMcpError(f"QCC MCP server has no URL: {server_name}")
    if not isinstance(token_env, str) or not token_env:
        raise QccMcpError(f"QCC MCP server has no bearer token environment variable: {server_name}")
    token = os.environ.get(token_env)
    if not token:
        raise QccMcpError(f"Required environment variable is not set: {token_env}")
    return url, token


class QccMcpClient:
    def __init__(self, url: str, token: str, timeout: int = 30):
        self.url = url
        self.timeout = timeout
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
        }
        cookie_jar = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cookie_jar))
        self.next_id = 1

    def post(self, payload: dict[str, Any], allow_empty: bool = False) -> dict[str, Any] | None:
        request = urllib.request.Request(
            self.url,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers=self.headers,
            method="POST",
        )
        try:
            with self.opener.open(request, timeout=self.timeout) as response:
                body = response.read().decode("utf-8", errors="replace")
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")[:500]
            raise QccMcpError(f"QCC MCP HTTP {exc.code}: {detail}") from exc
        except urllib.error.URLError as exc:
            raise QccMcpError(f"QCC MCP connection failed: {exc.reason}") from exc
        if not body.strip() and allow_empty:
            return None
        return parse_sse(body)

    def request(self, method: str, params: dict[str, Any]) -> dict[str, Any]:
        request_id = self.next_id
        self.next_id += 1
        response = self.post(
            {"jsonrpc": "2.0", "id": request_id, "method": method, "params": params}
        )
        if response is None or response.get("id") != request_id:
            raise QccMcpError(f"QCC MCP returned an unexpected response for {method}")
        return response

    def initialize(self) -> None:
        self.request(
            "initialize",
            {
                "protocolVersion": PROTOCOL_VERSION,
                "capabilities": {},
                "clientInfo": {"name": "interview-minutes-qcc-fallback", "version": "1.0"},
            },
        )
        self.post({"jsonrpc": "2.0", "method": "notifications/initialized"}, allow_empty=True)

    def list_tools(self) -> list[dict[str, Any]]:
        response = self.request("tools/list", {})
        tools = response.get("result", {}).get("tools", [])
        if not isinstance(tools, list):
            raise QccMcpError("QCC MCP tools/list returned an invalid tools field")
        return [tool for tool in tools if isinstance(tool, dict)]

    def call_tool(self, name: str, arguments: dict[str, Any]) -> Any:
        response = self.request("tools/call", {"name": name, "arguments": arguments})
        result = response.get("result", {})
        content = result.get("content") if isinstance(result, dict) else None
        if isinstance(content, list) and len(content) == 1 and isinstance(content[0], dict):
            text = content[0].get("text")
            if isinstance(text, str):
                try:
                    return json.loads(text)
                except json.JSONDecodeError:
                    return text
        return result


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Discover or call QCC MCP tools when native Codex tool discovery is stale."
    )
    parser.add_argument("--server", default="qcc-company", help="Server key in Codex config.toml")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG, help="Codex config.toml path")
    parser.add_argument("--timeout", type=int, default=30, help="HTTP timeout in seconds")
    action = parser.add_mutually_exclusive_group(required=True)
    action.add_argument("--list-tools", action="store_true", help="List tool names exposed by the server")
    action.add_argument("--tool", help="Call one exact tool name")
    parser.add_argument(
        "--arguments",
        default="{}",
        help="JSON object passed to the selected tool; required for tools with parameters",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        arguments = json.loads(args.arguments)
        if not isinstance(arguments, dict):
            raise QccMcpError("--arguments must be a JSON object")
        url, token = load_server(args.config, args.server)
        client = QccMcpClient(url, token, timeout=args.timeout)
        client.initialize()
        tools = client.list_tools()
        if args.list_tools:
            output = {
                "server": args.server,
                "tool_count": len(tools),
                "tools": [tool.get("name") for tool in tools if tool.get("name")],
            }
        else:
            names = {tool.get("name") for tool in tools}
            if args.tool not in names:
                raise QccMcpError(f"Tool is not exposed by {args.server}: {args.tool}")
            output = {
                "server": args.server,
                "tool": args.tool,
                "result": client.call_tool(args.tool, arguments),
            }
        print(json.dumps(output, ensure_ascii=False, indent=2))
        return 0
    except (QccMcpError, json.JSONDecodeError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

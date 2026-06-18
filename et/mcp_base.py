import sys, json

class MCPServer:
    def __init__(self, name, version="1.0.0"):
        self.name = name
        self.version = version
        self.tools = {}

    def tool(self, name, description, input_schema=None):
        def decorator(fn):
            self.tools[name] = dict(
                name=name, description=description,
                inputSchema=input_schema or {"type": "object", "properties": {}},
                handler=fn,
            )
            return fn
        return decorator

    def _writeline(self, obj):
        """写 JSON-RPC 到 stdout，绕过系统编码（Windows GBK 问题）"""
        line = json.dumps(obj, ensure_ascii=False) + "\n"
        sys.stdout.buffer.write(line.encode("utf-8", errors="replace"))
        sys.stdout.buffer.flush()

    def _err(self, msg):
        line = f"[{self.name}] {msg}\n"
        sys.stderr.buffer.write(line.encode("utf-8", errors="replace"))
        sys.stderr.buffer.flush()

    def run(self):
        self._err("started")
        for raw in sys.stdin.buffer:
            line = raw.decode("utf-8", errors="replace").strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
            except json.JSONDecodeError:
                continue

            method = msg.get("method")
            rid = msg.get("id")
            params = msg.get("params", {})

            if method == "initialize":
                self._writeline({
                    "jsonrpc": "2.0", "id": rid,
                    "result": {
                        "protocolVersion": params.get("protocolVersion", "2024-11-05"),
                        "capabilities": {"tools": {}},
                        "serverInfo": {"name": self.name, "version": self.version},
                    },
                })
            elif method == "notifications/initialized":
                pass
            elif method == "tools/list":
                items = [{
                    "name": t["name"],
                    "description": t["description"],
                    "inputSchema": t["inputSchema"],
                } for t in self.tools.values()]
                self._writeline({"jsonrpc": "2.0", "id": rid, "result": {"tools": items}})
            elif method == "tools/call":
                name = params.get("name", "")
                args = params.get("arguments", {})
                t = self.tools.get(name)
                if not t:
                    self._writeline({"jsonrpc": "2.0", "id": rid, "error": {"code": -32601, "message": f"Unknown tool: {name}"}})
                    continue
                try:
                    result = t["handler"](args)
                    if isinstance(result, dict):
                        text = json.dumps(result, ensure_ascii=False)
                        self._writeline({"jsonrpc": "2.0", "id": rid, "result": {"content": [{"type": "text", "text": text}]}})
                    else:
                        self._writeline({"jsonrpc": "2.0", "id": rid, "result": {"content": [{"type": "text", "text": str(result)}]}})
                except Exception as e:
                    self._err(f"tool error: {e}")
                    self._writeline({"jsonrpc": "2.0", "id": rid, "error": {"code": -32603, "message": str(e)}})
            else:
                pass  # ignore unknown methods

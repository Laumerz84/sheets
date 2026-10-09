"""Stand-in for `claude -p --output-format stream-json` used by the panel tests.

Reads the prompt from stdin, makes one real tool call through the Sheets bridge
(using the env from the --mcp-config file it was given), and prints stream-json
events like Claude Code does. Records its argv to FAKE_CLAUDE_LOG if set."""
import json
import os
import sys

argv = sys.argv[1:]
prompt = sys.stdin.read()
log = os.environ.get("FAKE_CLAUDE_LOG")
if log:
    with open(log, "a", encoding="utf-8") as fh:
        fh.write(json.dumps({"argv": argv, "prompt": prompt}) + "\n")

cfg = json.load(open(argv[argv.index("--mcp-config") + 1], encoding="utf-8"))
os.environ.update(cfg["mcpServers"]["sheets"]["env"])
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def emit(obj):
    print(json.dumps(obj), flush=True)


emit({"type": "system", "subtype": "init", "session_id": "sess-123",
      "mcp_servers": [{"name": "sheets", "status": "connected"}]})
emit({"type": "stream_event", "event": {"type": "message_start"}})
emit({"type": "stream_event", "event": {"type": "content_block_delta",
                                        "delta": {"type": "text_delta", "text": "Adding a total. "}}})
tool_input = {"start_cell": "A3", "rows": [["=SUM(A1:A2)"]]}
emit({"type": "assistant", "message": {"content": [
    {"type": "text", "text": "Adding a total. "},
    {"type": "tool_use", "name": "mcp__sheets__write_range", "input": tool_input}]}})

from sheets.ai.mcp_server import call  # noqa: E402

result = call("write_range", **tool_input)
emit({"type": "user", "message": {"content": [{"type": "tool_result", "content": json.dumps(result)}]}})
emit({"type": "stream_event", "event": {"type": "message_start"}})
emit({"type": "stream_event", "event": {"type": "content_block_delta",
                                        "delta": {"type": "text_delta", "text": "Done: A3 = 30."}}})
emit({"type": "assistant", "message": {"content": [{"type": "text", "text": "Done: A3 = 30."}]}})
emit({"type": "result", "subtype": "success", "is_error": False, "result": "Done: A3 = 30.",
      "session_id": "sess-123", "duration_ms": 1234})

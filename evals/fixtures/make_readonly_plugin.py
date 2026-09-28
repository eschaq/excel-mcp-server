"""Write the eval-only read-only copy of the plugin. Called by evals/prepare.sh.

The copy is derived from the real .claude-plugin/plugin.json, so it always starts the server the
same way: only its name changes, its paths point back at the repository root, and the user
settings are replaced by --read-only.
"""

import json
from pathlib import Path

EVALS = Path(__file__).resolve().parent.parent
ROOT = EVALS.parent
OUT = EVALS / "security-read-only-mode" / "readonly-plugin" / ".claude-plugin" / "plugin.json"
PLUGIN_ROOT = "${CLAUDE_PLUGIN_ROOT}"


def main() -> None:
    manifest = json.loads((ROOT / ".claude-plugin" / "plugin.json").read_text(encoding="utf-8"))
    server = manifest["mcpServers"]["excel"]
    # The copy lives three folders below the repository root.
    args = [a.replace(PLUGIN_ROOT, PLUGIN_ROOT + "/../../..") for a in server["args"]]
    args = args[: args.index("--allow-dir")] + ["--read-only"]
    readonly = {
        "name": manifest["name"] + "-readonly",
        "version": manifest["version"],
        "description": f"Eval-only copy of {manifest['name']} whose server runs with --read-only.",
        "author": manifest["author"],
        "mcpServers": {"excel": {**server, "args": args}},
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(readonly, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()

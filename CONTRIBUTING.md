# Contributing to outlook-desktop-mcp

## Branching Strategy

```
feature/your-change → PR → preview → PR → main → auto-publish to PyPI
```

- **`main`** — stable, auto-publishes to PyPI on every push. Never push directly.
- **`preview`** — integration testing branch. All PRs land here first.
- **Feature branches** — your working branches, created from `preview`.

## How to Contribute

1. **Fork** the repo to your GitHub account
2. **Clone** your fork locally
3. **Create a branch** from `preview`:
   ```bash
   git checkout preview
   git pull origin preview
   git checkout -b feature/my-change
   ```
4. **Make your changes**, commit, push to your fork
5. **Open a PR** into `preview` (not `main`)
6. Once reviewed and merged to `preview`, it will be tested there
7. Periodically, `preview` is merged into `main` which triggers a PyPI release

## Development Setup

Requires Windows with Outlook Desktop (Classic) running.

Dependencies are managed with [uv](https://docs.astral.sh/uv/):

```bash
git clone https://github.com/YOUR-USERNAME/outlook-desktop-mcp.git
cd outlook-desktop-mcp
uv sync
uv run python .venv\Scripts\pywin32_postinstall.py -install
```

## Testing

With Outlook Desktop (Classic) open:

```bash
# COM validation (no MCP layer)
uv run python tests\phase1_com_test.py

# MCP protocol test
uv run python tests\phase3_mcp_test.py
```

## Adding New Tools

1. Define the COM function that does the work (receives `outlook, namespace` as first args)
2. Add an `@mcp.tool()` async handler in `server.py` that calls `bridge.call(your_function, ...)`
3. Write a detailed docstring — this is what LLMs see during tool discovery
4. Add test coverage
5. Update the `instructions` string if adding a new capability category

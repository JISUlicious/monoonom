# Install

**English** | [한국어](install.ko.md) · [Back to README](../../README.md)

## Requirements

- **Python 3.9+** with `openpyxl` (required) and `xlrd` (only to open `.xls`).
- Optional: **Node.js 18+**, so formulas are also computed when no preview is open.
- The agent must run on the **same computer** (the preview connects to `localhost`). Agents running in the cloud and the Skills API cannot use it.

```bash
python3 -m pip install openpyxl xlrd
```

Offline computers: on a connected one run `pip download openpyxl xlrd -d wheels`, copy `wheels/`, then
`pip install --no-index --find-links wheels openpyxl xlrd`. If a package is missing, `xl.py serve` prints the install command and exits with code 6.

## Claude Code: the plugin (recommended)

```bash
claude plugin marketplace add JISUlicious/monoonom
claude plugin install monoonom@monoonom
```

- In a session: `/plugin install monoonom --marketplace JISUlicious/monoonom` (Claude Code 2.1.275 or later) does both.
- The skill is `/monoonom:monoonom`; Claude also uses it on its own when you ask about a spreadsheet.
- Update: `claude plugin update monoonom@monoonom`. Remove: `claude plugin uninstall monoonom@monoonom`.
- The plugin version pre-approves its own `xl.py` commands (`allowed-tools`), shows the server status when the skill loads, and tells Claude Code desktop to open the preview pane.

## Other agents (Codex, Cursor, Copilot, …): `npx skills`

```bash
npx skills add JISUlicious/monoonom            # this project (choose the agents)
npx skills add JISUlicious/monoonom -g -a codex # all projects, one agent
```

- The [skills CLI](https://github.com/vercel-labs/skills) installs `.agents/skills/monoonom/` (the portable `SKILL.md`).
- Update: `npx skills update monoonom`. Remove: `npx skills remove monoonom`.
- It also works in Claude Code, without the plugin's extras above.
- Codex in `workspace-write` mode blocks binding a local port: approve the command or give it full access.

## Copy it by hand

Copy `.agents/skills/monoonom/` into your agent's skills folder (for example `~/.agents/skills/`; Claude Code uses
`~/.claude/skills/`). For Claude Code, replace `SKILL.md` with `SKILL.claude-code.md`.

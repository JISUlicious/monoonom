# MoNooNoM

**English** | [한국어](README.ko.md)

<img src=".agents/skills/monoonom/assets/web/img/monoonom.svg" alt="" width="72" align="right">

A local web spreadsheet where **you and an AI agent work on the same workbook while both of you watch it**.
The agent opens and edits Excel files (xlsx and xlsm; xls read-only), and you see every change live in the agent app's
preview pane (or any browser) and can edit the same workbook yourself. The name is *monoon* (모눈, Korean for graph
paper) written in English and mirrored: MoNooN + NooNoM.

![Preview](docs/monoonom-preview.png)

- **Edit together**: the server merges the agent's and your edits in order and pushes them to every open window. The agent reads the cells you changed with `changes`.
- **Excel-compatible**: about 200 formula functions, formatting, filters and sorting, charts (bar, line, pie, scatter, combo and more), pivot tables, Excel keyboard shortcuts. Saving keeps what the app does not edit, such as conditional formatting, images and VBA.
- **Local only**: the Python server listens on `127.0.0.1` only, and files never leave the workspace folder. All file reads and writes go through one Python function, so a DRM hook can be plugged in.

## Install

You need **Python 3.9+** with `openpyxl` (required) and `xlrd` (only to open .xls). With Node.js 18+, formulas are also computed when no preview is open.

```bash
python3 -m pip install openpyxl xlrd
```

It must run on the **same computer** as the agent (the preview connects to `localhost`). Agents running in the cloud and the Skills API cannot use it.

### Claude Code: install the plugin (recommended)

```bash
claude plugin marketplace add JISUlicious/web-excel-agent-skills
claude plugin install monoonom@monoonom
```

Inside a session, `/plugin install monoonom --marketplace JISUlicious/web-excel-agent-skills` (Claude Code 2.1.275 or later) does both in one step.
Update with `claude plugin update monoonom@monoonom`. The skill is `/monoonom:monoonom`.

### Other agents (Codex, Cursor, Copilot, …): `npx skills`

```bash
npx skills add JISUlicious/web-excel-agent-skills            # this project (choose agents)
npx skills add JISUlicious/web-excel-agent-skills -g -a codex # all projects, one agent
```

The [skills CLI](https://github.com/vercel-labs/skills) installs `.agents/skills/monoonom/` (the portable `SKILL.md`). This also works in Claude Code, but without the plugin version's pre-approved commands (`allowed-tools`), live server status and desktop preview instructions.

### Copy it by hand

Copy `.agents/skills/monoonom/` into your agent's skills folder (for example `~/.agents/skills/`; Claude Code uses `~/.claude/skills/`). For Claude Code, replace `SKILL.md` with `SKILL.claude-code.md`.

## Use

Ask the agent as usual, for example:

- "Open report.xlsx so we can look at it together"
- "Add a monthly total column to the Sales sheet and make a line chart"
- "Check the cell I just changed and fix the other rows the same way"

The agent starts the server and gives you the preview URL (`http://127.0.0.1:8765/#t=<token>`). Claude Code desktop opens it in its preview pane;
in VS Code or Cursor use Simple Browser, elsewhere any browser. When saving, the previous file is kept in `<workspace>/.monoonom/backups/`.

Commands the agent uses (you can run them too):

```bash
XL=<skill folder>/scripts/xl.py         # run from the workspace folder
python3 $XL serve                       # start the server (or reuse one) → preview URL
python3 $XL open report.xlsx            # open → shows up as a tab in the preview
python3 $XL describe                    # sheets, used ranges, headers, charts
python3 $XL fill 'G2:G7' '=E2*0.3'      # fill a formula
python3 $XL chart add A1:D7 --type line --title Trend --at H20
python3 $XL changes --since 0 --actor user   # what you changed in the preview
python3 $XL save                        # save
python3 $XL stop
```

All commands: [`references/cli.md`](.agents/skills/monoonom/references/cli.md) · batch edit format: [`references/edits.md`](.agents/skills/monoonom/references/edits.md) · troubleshooting: [`references/troubleshooting.md`](.agents/skills/monoonom/references/troubleshooting.md) · UI features ↔ commands: [`docs/agent-coverage.md`](docs/agent-coverage.md) (Korean)

### Security

The server binds to `127.0.0.1` only and checks the Host header (DNS rebinding). Every API call needs the workspace's token, and files outside the workspace cannot be opened (`--no-sandbox` lifts this).
Server state (`server.json`, token, log) lives in `<workspace>/.monoonom/`.

### DRM environments

The browser never reads file contents; it only sends paths. All disk I/O goes through `scripts/server/fileio.py`,
and `MONOONOM_READ_HOOK=pkg.mod:func` / `MONOONOM_WRITE_HOOK=...` plug in a DRM SDK's functions.

## Keyboard shortcuts

`Ctrl+N/O/S`, `F12`, `Ctrl+Z/Y`, `Ctrl+C/X/V`, `Ctrl+Alt+V`, `Ctrl+B/I/U/5`, `Ctrl+1`, `F2`, `F4`, `Alt+Enter`, `Ctrl+Enter`, `Ctrl+D/R`, `Ctrl+;`, `Alt+=`, `Ctrl+Shift+L`, `Alt+↓`, `Ctrl+F/H`, `Ctrl+arrows`, `Ctrl+Shift+arrows`, `Ctrl+Home/End`, `Ctrl+Space`, `Shift+Space`, `Ctrl++/-`, `Ctrl+PgUp/PgDn`, `Shift+F11`, `Alt+F1`, `Shift+F2`, `Ctrl+K`, `F9` and more (the full list is behind the `⌨` button in the app).
In a regular browser tab the browser may take `Ctrl+N/W/Tab` first (use `Alt+N` or click the tab instead).

## Known limitations

- **xls is read-only**: only computed values are read (xlrd does not give formula text). Saving writes a .xlsx with the same name and leaves the .xls alone.
- Conditional formatting, images and sparklines are **not shown but kept when saving**. Inserting rows does not move their ranges.
- Pivot tables: kept, outlined, refreshed and created (tabular form). Edits inside a pivot and inserts, deletes or sorts across one are refused.
- Dynamic array formulas do not spill; only the top-left value is shown.
- Chart details (per-point colors, data labels, secondary axes) are rebuilt in the default style when a chart is edited.
- Cell-level "shift cells" inserts and deletes do not adjust references in other formulas (whole-row and whole-column inserts and deletes do).

## Development

Repository layout, design decisions, tests and packaging are in [docs/development.md](docs/development.md) (Korean). Change log: [logs.md](logs.md), document index: [index.md](index.md).

```bash
python3 run.py --port 8765 --root .      # foreground server for development
python3 tests/skill_package.test.py      # skill and plugin package checks
claude plugin validate --strict .        # plugin and marketplace validation
```

## License

MIT ([LICENSE](LICENSE)). The only bundled third-party code is Chart.js (MIT, `assets/web/vendor/`); see
[`THIRD_PARTY_NOTICES.md`](.agents/skills/monoonom/THIRD_PARTY_NOTICES.md). The Python packages (openpyxl, xlrd) are installed by the user from PyPI.

## Trademarks

Microsoft and Excel are trademarks of Microsoft Corporation in the United States and other countries. MoNooNoM is not affiliated with, sponsored by or endorsed by Microsoft. "Excel" is used here only to describe compatibility with its file formats and behavior.

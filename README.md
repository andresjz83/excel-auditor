# Excel Auditor

Range-level precedent and dependent tracing for Excel on Mac. Office.js task-pane add-in.

> Native Excel can only trace one cell at a time. This walks every cell in a multi-cell selection, surfaces all upstream/downstream cells (recursively), groups them by sheet in a tree, and highlights the matching reference in the parent formula as you navigate. Macabacus-style behavior without the licence.

Not affiliated with Macabacus or Microsoft.

---

## Quick install (recommended)

Open a terminal in this directory and run:

```bash
./install.sh
```

The installer:

1. Checks macOS, Excel for Mac, Node.js are present
2. `npm install`s the two dev deps (`http-server`, `office-addin-dev-certs`)
3. Installs the Office.js localhost dev certificate (one-time, will prompt for keychain password)
4. Mints a per-machine manifest GUID and copies the manifest into Excel's sideload folder
5. Enables right-click → Inspect Element on the task pane (DevTools for debugging)
6. Installs a launchd job that auto-starts the local dev server at every login

After it finishes: **Cmd+Q to quit Excel**, reopen it, open any workbook. The Home ribbon has a new **"Open Auditor"** button under the Auditor group.

Uninstall with `./uninstall.sh` (removes launchd job + sideloaded manifest; leaves source tree and dev cert alone).

---

## What it does

Click a cell with a formula, click **Show precedents** — sidebar lists every upstream cell across every sheet, nested in a tree by hierarchy. Click any result → Excel selects that cell. Arrow keys walk the tree. The formula panel at the top shows the cell's formula with the matching reference highlighted in yellow.

**Show dependents** does the same downstream. Uses a workbook-wide reverse index (built on first call, cached in memory) so lookups are instant after the first scan.

Features:

- **Range-level selection** — audit 100 cells at once, not just one
- **Tree view** — L1/L2/L3 cells nested under whoever discovered them
- **Yellow highlight** in the formula as you walk results, showing exactly which token connects
- **Copy formula** button next to the formula display
- **Arrow-key navigation** through results (↑/↓, Home, End, Enter)
- **Esc clears** the current audit without closing the pane
- **Opacity warnings** when the audit hits `INDIRECT`, `OFFSET`, or external-workbook refs
- **Safe-to-delete mode** flags downstream impact before you delete a region (avoids #REF!)
- **Offline** — no internet, no cloud, no telemetry. Everything runs locally on localhost.

## Architecture

| File | Role |
|---|---|
| `manifest.xml` | Office add-in manifest (sideloaded into Excel) |
| `src/taskpane.html` | Sidebar UI shell |
| `src/taskpane.css` | Styling |
| `src/auditor.js` | All the audit logic (formula-text parser, BFS walker, reverse index builder, render) |
| `server/run-server.sh` | Wrapper that launches `http-server` with the local TLS cert |
| `server/com.excel-auditor.dev-server.plist.template` | launchd config template, materialized by `install.sh` |
| `install.sh` / `uninstall.sh` | One-shot install + clean removal |

The dev server serves three files (`taskpane.html`, `taskpane.css`, `auditor.js`) over HTTPS on `localhost:3000`. The manifest points Excel at that URL. Excel renders the task pane in an embedded WebView with Office.js bridging back to the workbook.

For precedents, the auditor parses formula text directly (more reliable than `Range.getDirectPrecedents()`, which is flaky on Mac Excel for complex cross-sheet formulas). For dependents, it scans every formula in every sheet on first request, builds a reverse index `{precedent → set of dependents}`, and caches it. Click **Rebuild index** if you edit formulas mid-session.

## Known limits

- **macOS only.** Excel for Mac. Could probably work on Excel for Windows with minor manifest tweaks, untested.
- **`INDIRECT` / `OFFSET` / dynamic lookups** — won't appear as static precedents. We flag opacity so you know the audit is incomplete.
- **External workbook refs** — Excel's API can't see into closed workbooks. Flagged in the opacity warning.
- **Defined names that look like cell refs** (e.g. a name `MRG2`) get misclassified by the regex parser. Rare edge.
- **Whole-column refs** (`A:A`) aren't expanded to individual cells — kept as range. Means the reverse index doesn't register every cell in the column as a dependent of the formula. Practical impact: low for typical models.

## For Claude / coding agents

This repo is designed to install cleanly via a Claude Code conversation. If a user asks you to install it:

1. Confirm you're on macOS with Excel for Mac and Node.js installed (`node --version`, `[ -d "/Applications/Microsoft Excel.app" ]`).
2. From the repo root, run `./install.sh`. It's idempotent and verbose — surface its output to the user.
3. When it finishes, tell the user to **Cmd+Q to fully quit Excel** and reopen it.
4. The add-in appears in the **Home** ribbon as **Open Auditor**.
5. If the install fails on a missing prereq, the script prints the exact `brew install` command needed. Run that, then re-run `./install.sh`.

For uninstall: `./uninstall.sh`. For debugging a broken install: check `~/.config/excel-auditor/logs/server-error.log` and `launchctl print gui/$UID/com.excel-auditor.dev-server`.

## License

MIT.

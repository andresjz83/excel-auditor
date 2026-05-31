/* global Office, Excel */

const HIGHLIGHT_COLOR = "#FFE08A";
const OPAQUE_PATTERNS = [
  { re: /\bINDIRECT\s*\(/i, name: "INDIRECT" },
  { re: /\bOFFSET\s*\(/i, name: "OFFSET" },
  { re: /\bINDEX\s*\(\s*[^,)]+,\s*MATCH/i, name: "INDEX/MATCH (dynamic)" },
  { re: /\[.+?\]/, name: "External workbook" },
];

const state = {
  highlighted: [], // [{sheet, address}]
  lastResults: [], // [{sheet, address, formula, depth}]
  depIndex: null,  // Map<"sheet!address", Set<"sheet!address">> — reverse: precedent → its dependents
  formulaIndex: null, // Map<"sheet!address", formula string> — cache of every formula in workbook
  indexBuiltAt: null,
  // Cells we've called showPrecedents() / showDependents() on, so we can call
  // the (true) variant later to remove their native arrows.
  arrowedPrecedents: [], // [{sheet, address}]
  arrowedDependents: [], // [{sheet, address}]
};

async function clearNativeArrows() {
  if (!state.arrowedPrecedents.length && !state.arrowedDependents.length) return;
  const pre = state.arrowedPrecedents.splice(0);
  const dep = state.arrowedDependents.splice(0);
  try {
    await Excel.run(async (ctx) => {
      for (const { sheet, address } of pre) {
        ctx.workbook.worksheets.getItem(sheet).getRange(address).showPrecedents(true);
      }
      for (const { sheet, address } of dep) {
        ctx.workbook.worksheets.getItem(sheet).getRange(address).showDependents(true);
      }
      await ctx.sync();
    });
  } catch (e) {
    console.warn("[auditor] clearNativeArrows failed:", e);
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setStatus("This add-in only runs in Excel.", "err");
    return;
  }
  wireUI();
  refreshSelectionLabel();
  Excel.run(async (ctx) => {
    ctx.workbook.onSelectionChanged.add(refreshSelectionLabel);
    await ctx.sync();
  }).catch((e) => console.error(e));
});

function clearResults() {
  document.getElementById("results").innerHTML = "";
  state.lastResults = [];
  state.highlighted.length = 0; // any pending highlights are conceptually orphaned
  lastRenderedResults = null;
  formulasByKey = new Map();
  clearNativeArrows(); // fire-and-forget; status will update if it errors
  setStatus("Cleared. Pick a cell and Show precedents / dependents.");
  document.getElementById("btn-precedents")?.focus();
}

function wireUI() {
  document.getElementById("btn-precedents").addEventListener("click", () => runAudit("precedents"));
  document.getElementById("btn-dependents").addEventListener("click", () => runDependents());
  document.getElementById("btn-highlight").addEventListener("click", highlightResults);
  document.getElementById("btn-clear").addEventListener("click", clearHighlights);
  const rb = document.getElementById("btn-rebuild");
  if (rb) rb.addEventListener("click", () => { state.depIndex = null; state.formulaIndex = null; setStatus("Index cleared."); });

  // Esc clears the current audit results without closing the pane.
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && lastRenderedResults) {
      e.preventDefault();
      clearResults();
    }
  });

  // Copy formula button — delegated since formula-panel re-renders on focus.
  document.getElementById("results").addEventListener("click", (e) => {
    const btn = e.target.closest(".formula-copy");
    if (!btn) return;
    e.stopPropagation();
    const f = btn.dataset.formula || "";
    if (!f) return;
    navigator.clipboard.writeText(f).then(
      () => { btn.textContent = "Copied"; setTimeout(() => { btn.textContent = "Copy"; }, 1200); },
      () => setStatus("Copy failed.", "err"),
    );
  });

  // Arrow-key navigation through the results list. Listens on the whole
  // results container; works as long as a .cell-item has focus.
  document.getElementById("results").addEventListener("keydown", (e) => {
    if (!["ArrowDown", "ArrowUp", "Home", "End", "Enter"].includes(e.key)) return;
    const items = Array.from(document.querySelectorAll(".cell-item"));
    if (!items.length) return;
    const active = document.activeElement;
    let idx = items.indexOf(active);
    if (idx === -1) idx = 0;
    let nextIdx = idx;
    switch (e.key) {
      case "ArrowDown": nextIdx = Math.min(items.length - 1, idx + 1); break;
      case "ArrowUp": nextIdx = Math.max(0, idx - 1); break;
      case "Home": nextIdx = 0; break;
      case "End": nextIdx = items.length - 1; break;
      case "Enter": nextIdx = idx; break;
    }
    e.preventDefault();
    const target = items[nextIdx];
    target.focus();
    target.scrollIntoView({ block: "nearest" });
    navigateTo(target.dataset.sheet, target.dataset.address);
  });
}

/**
 * Walk the entire workbook once, parse every formula, and build a reverse
 * index: precedent_address → set of dependent_addresses.
 *
 * One sync per sheet (load used range's formulas), then pure-JS parsing.
 * On a 30K-formula workbook this is seconds, not minutes — and after it's
 * built, dependents lookups are O(1).
 */
async function buildDependencyIndex() {
  const t0 = performance.now();
  setStatus("Indexing workbook — this runs once per session…");
  try {
    const { depIdx, formulaIdx, stats } = await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();

      const sheetUsed = sheets.items.map((ws) => {
        const used = ws.getUsedRangeOrNullObject(true); // valuesOnly: true → only cells with content
        used.load(["address", "formulas", "rowCount", "columnCount", "rowIndex", "columnIndex", "isNullObject"]);
        return { name: ws.name, used };
      });
      await ctx.sync();

      const formulaIdx = new Map();
      const depIdx = new Map();
      let formulaCount = 0;
      let cellCount = 0;

      for (const { name: sheetName, used } of sheetUsed) {
        if (used.isNullObject) continue;
        const r0 = used.rowIndex;
        const c0 = used.columnIndex;
        const formulas = used.formulas;
        if (!formulas) continue;

        for (let r = 0; r < formulas.length; r++) {
          const row = formulas[r];
          for (let c = 0; c < row.length; c++) {
            cellCount++;
            const f = (row[c] ?? "").toString();
            if (!f || !f.startsWith("=")) continue;
            formulaCount++;
            const address = cellAddress(r0 + r, c0 + c);
            const myKey = `${sheetName}!${address}`;
            formulaIdx.set(myKey, f);

            const refs = extractReferenceKeys(f, sheetName);
            for (const refStr of refs) {
              const parsed = parseQualifiedAddress(refStr);
              if (!parsed || parsed.external) continue;
              // Expand range refs so any cell in the range gets registered as a
              // precedent. Cheap because expandRange is bounded.
              let cells;
              try { cells = expandRange(parsed); } catch (_) { continue; }
              for (const cellRef of cells) {
                const refKey = `${cellRef.sheet}!${cellRef.address}`;
                let set = depIdx.get(refKey);
                if (!set) { set = new Set(); depIdx.set(refKey, set); }
                set.add(myKey);
              }
            }
          }
        }
      }
      return { depIdx, formulaIdx, stats: { cells: cellCount, formulas: formulaCount, sheets: sheetUsed.length } };
    });
    state.depIndex = depIdx;
    state.formulaIndex = formulaIdx;
    state.indexBuiltAt = Date.now();
    const ms = Math.round(performance.now() - t0);
    setStatus(`Index built: ${stats.formulas.toLocaleString()} formulas in ${stats.sheets} sheets (${ms} ms).`);
    return true;
  } catch (e) {
    console.error("[auditor] index build failed:", e);
    setStatus(`Index build failed: ${e.message || e}`, "err");
    return false;
  }
}

async function runDependents() {
  const depth = parseInt(document.getElementById("depth").value, 10);
  const safeDelete = document.getElementById("opt-deletecheck").checked;
  document.getElementById("results").innerHTML = "";
  diagLog.length = 0;
  await clearNativeArrows(); // remove any arrows from a prior audit

  if (!state.depIndex) {
    const ok = await buildDependencyIndex();
    if (!ok) return;
  }

  // Get the selection, and try drawing native dependent arrows.
  let selectionInfo;
  try {
    selectionInfo = await Excel.run(async (ctx) => {
      const sel = ctx.workbook.getSelectedRange();
      sel.load(["address", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
      sel.worksheet.load("name");
      await ctx.sync();
      try {
        sel.showDependents();
        await ctx.sync();
        state.arrowedDependents.push({ sheet: sel.worksheet.name, address: sel.address.split("!").pop() });
        diag("native arrows: showDependents() OK");
      } catch (arrErr) {
        diag("native arrows: showDependents() FAILED", arrErr.message || String(arrErr));
        console.warn("[auditor] showDependents failed:", arrErr);
      }
      return {
        sheet: sel.worksheet.name,
        address: sel.address,
        rowCount: sel.rowCount, columnCount: sel.columnCount,
        rowIndex: sel.rowIndex, columnIndex: sel.columnIndex,
      };
    });
  } catch (e) {
    setStatus(`Could not read selection: ${e.message}`, "err");
    return;
  }

  // Walk the index breadth-first. Track parentKey so the renderer can build
  // a tree (each result nests under the cell that depends *on it*).
  const seen = new Map();
  const opaqueHits = new Set();
  const rootKeys = [];
  let frontier = [];
  for (let r = 0; r < selectionInfo.rowCount; r++) {
    for (let c = 0; c < selectionInfo.columnCount; c++) {
      const addr = cellAddress(selectionInfo.rowIndex + r, selectionInfo.columnIndex + c);
      const key = `${selectionInfo.sheet}!${addr}`;
      rootKeys.push(key);
      frontier.push(key);
    }
  }

  let level = 0;
  while (frontier.length && level < depth) {
    level++;
    const next = [];
    for (const parentKey of frontier) {
      const downstream = state.depIndex.get(parentKey);
      if (!downstream) continue;
      for (const depKey of downstream) {
        if (seen.has(depKey)) continue;
        const [sheet, address] = splitKey(depKey);
        const formula = state.formulaIndex.get(depKey) || "";
        for (const pat of OPAQUE_PATTERNS) {
          if (pat.re.test(formula)) opaqueHits.add(pat.name);
        }
        seen.set(depKey, { sheet, address, formula, depth: level, parentKey, key: depKey });
        next.push(depKey);
      }
    }
    frontier = next;
  }

  // For dependents, root formulas come from the formula index (or, if the
  // root cell isn't a formula, just empty string — that's fine).
  const rootFormulas = {};
  for (const k of rootKeys) rootFormulas[k] = state.formulaIndex.get(k) || "";

  const items = Array.from(seen.values());
  state.lastResults = items;
  renderResults({
    direction: "dependents",
    depth: level,
    opaque: Array.from(opaqueHits),
    items,
    rootKeys,
    rootFormulas,
    selectionAddress: selectionInfo.address,
    safeDelete,
  });
  setStatus(
    `dependents: ${items.length} cell${items.length === 1 ? "" : "s"} across ${countSheets(items)} sheet${countSheets(items) === 1 ? "" : "s"}.`,
  );
}

function splitKey(key) {
  const i = key.indexOf("!");
  return [key.slice(0, i), key.slice(i + 1)];
}

function setStatus(msg, kind = "") {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = "status " + kind;
}

async function refreshSelectionLabel() {
  try {
    await Excel.run(async (ctx) => {
      const sel = ctx.workbook.getSelectedRange();
      sel.load(["address", "cellCount"]);
      await ctx.sync();
      document.getElementById("selection-label").textContent =
        `${sel.address}  ·  ${sel.cellCount} cell${sel.cellCount === 1 ? "" : "s"}`;
    });
  } catch (e) {
    // Selection may be empty/invalid; ignore.
  }
}

async function runAudit(direction) {
  const depth = parseInt(document.getElementById("depth").value, 10);
  const safeDelete = document.getElementById("opt-deletecheck").checked;
  setStatus(`Walking ${direction} (depth ${depth === 999 ? "all" : depth})…`);
  document.getElementById("results").innerHTML = "";
  diagLog.length = 0;
  diag("BEGIN", { direction, depth });
  await clearNativeArrows(); // remove any arrows from a prior audit

  try {
    let step = "init";
    const results = await Excel.run(async (ctx) => {
      step = "load selection";
      const selection = ctx.workbook.getSelectedRange();
      selection.load(["address", "cellCount", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
      const selSheet = selection.worksheet;
      selSheet.load("name");
      await ctx.sync();
      console.log("[auditor] selection:", selection.address, "sheet:", selSheet.name);
      diag("selection", { sheet: selSheet.name, address: selection.address, cells: selection.cellCount });

      // Native trace-precedents arrows. Wrapped in try/catch in case the API
      // throws on a specific Mac Excel build.
      if (direction === "precedents") {
        try {
          selection.showPrecedents();
          await ctx.sync();
          state.arrowedPrecedents.push({ sheet: selSheet.name, address: selection.address.split("!").pop() });
          diag("native arrows: showPrecedents() OK");
        } catch (arrErr) {
          diag("native arrows: showPrecedents() FAILED", arrErr.message || String(arrErr));
          console.warn("[auditor] showPrecedents failed:", arrErr);
        }
      }

      const selectionSheet = selSheet.name;
      const r0 = selection.rowIndex;
      const c0 = selection.columnIndex;

      // Roots: the cells the user selected. Each gets a key so children
      // (precedents/dependents) can record which root led to them, enabling
      // the nested tree view.
      const rootKeys = [];
      let frontier = [];
      for (let r = 0; r < selection.rowCount; r++) {
        for (let c = 0; c < selection.columnCount; c++) {
          const addr = cellAddress(r0 + r, c0 + c);
          const key = `${selectionSheet}!${addr}`;
          rootKeys.push(key);
          frontier.push({ sheet: selectionSheet, address: addr, parentKey: null, key });
        }
      }
      console.log("[auditor] seed frontier:", frontier);

      // seen: "sheet!address" → {sheet, address, formula, depth, parentKey}.
      // parentKey points back to the node that discovered this one, so we can
      // reconstruct the tree at render time.
      const seen = new Map();
      const opaqueHits = new Set();
      const externalSkips = new Set();
      let currentDepth = 0;

      while (frontier.length && currentDepth < depth) {
        currentDepth++;
        step = `depth ${currentDepth}: load frontier formulas`;

        // Load formulas for every frontier cell. For precedents, this is the
        // INPUT to parsing. For dependents, we still need it to record what
        // each visited cell contains, plus we'll need getDirectDependents().
        const frontierProxies = [];
        for (const node of frontier) {
          try {
            diag(`L${currentDepth} getRange`, node);
            const cell = ctx.workbook.worksheets.getItem(node.sheet).getRange(node.address);
            cell.load("formulas");
            const depRA = direction === "dependents" ? cell.getDirectDependents() : null;
            if (depRA) depRA.load("addresses");
            frontierProxies.push({ node, cell, depRA });
          } catch (err) {
            diag(`L${currentDepth} skip frontier`, { node, err: err.message });
          }
        }
        step = `depth ${currentDepth}: sync frontier`;
        await ctx.sync();

        step = `depth ${currentDepth}: extract neighbors`;
        const candidateCells = [];
        for (const { node, cell, depRA } of frontierProxies) {
          let formula = "";
          try { formula = (cell.formulas?.[0]?.[0] ?? "").toString(); } catch (_) {}
          // Record formula on the node we already saw at the PREVIOUS depth.
          const nodeKey = `${node.sheet}!${node.address}`;
          if (seen.has(nodeKey)) seen.get(nodeKey).formula = formula;

          let neighborAddrs = [];
          if (direction === "precedents") {
            // Parse formula text → list of cell references. This sidesteps
            // getDirectPrecedents() entirely, which is unreliable on Mac
            // Excel for complex cross-sheet formulas.
            if (formula.startsWith("=")) {
              neighborAddrs = extractReferenceKeys(formula, node.sheet);
              diag(`L${currentDepth} parsed precedents of ${node.sheet}!${node.address}`, {
                formula: formula.slice(0, 80) + (formula.length > 80 ? "…" : ""),
                refs: neighborAddrs,
              });
            }
            for (const pat of OPAQUE_PATTERNS) {
              if (pat.re.test(formula)) opaqueHits.add(pat.name);
            }
          } else {
            // Dependents: use the API
            const addrs = depRA?.addresses || [];
            diag(`L${currentDepth} dependents of ${node.sheet}!${node.address}`, addrs);
            neighborAddrs = addrs;
          }

          const parentKey = node.key;
          for (const addrStr of neighborAddrs) {
            const parsed = parseQualifiedAddress(addrStr);
            if (!parsed) { diag(`L${currentDepth} unparseable`, addrStr); continue; }
            if (parsed.external) { externalSkips.add(parsed.sheet); continue; }
            let expanded;
            try { expanded = expandRange(parsed); }
            catch (err) { diag(`L${currentDepth} expandRange failed`, { parsed, err: err.message }); continue; }
            for (const cellRef of expanded) {
              const key = `${cellRef.sheet}!${cellRef.address}`;
              if (seen.has(key)) continue;
              const node = { ...cellRef, formula: "", depth: currentDepth, parentKey, key };
              candidateCells.push(node);
              seen.set(key, node);
            }
          }
        }

        if (!candidateCells.length) break;
        diag(`L${currentDepth} candidates`, candidateCells.length);

        // Seed next frontier with the candidates we just discovered.
        // The next iteration will load THEIR formulas and walk further.
        frontier = candidateCells;
      }
      if (externalSkips.size) opaqueHits.add(`External: ${Array.from(externalSkips).join(", ")}`);

      // After the loop, the deepest tier of cells still have no formula loaded.
      // Load them now so the result list isn't blank.
      step = "load tail formulas";
      const tail = [];
      for (const v of seen.values()) {
        if (v.formula === "" && v.depth === currentDepth) {
          try {
            const r = ctx.workbook.worksheets.getItem(v.sheet).getRange(v.address);
            r.load("formulas");
            tail.push({ v, r });
          } catch (_) {}
        }
      }
      if (tail.length) {
        await ctx.sync();
        for (const { v, r } of tail) {
          try { v.formula = (r.formulas?.[0]?.[0] ?? "").toString(); } catch (_) {}
        }
      }

      const items = Array.from(seen.values()).map(({ sheet, address, formula, depth, parentKey, key }) =>
        ({ sheet, address, formula, depth, parentKey, key }));

      // Also load the root cells' formulas so the focused-formula panel works
      // when the audit is rooted on a formula cell.
      const rootFormulas = {};
      const rootRanges = rootKeys.map((k) => {
        const [s, a] = splitKey(k);
        const rr = ctx.workbook.worksheets.getItem(s).getRange(a);
        rr.load("formulas");
        return { key: k, range: rr };
      });
      await ctx.sync();
      for (const { key, range } of rootRanges) {
        try { rootFormulas[key] = (range.formulas?.[0]?.[0] ?? "").toString(); } catch (_) {}
      }

      return {
        direction,
        depth: currentDepth,
        opaque: Array.from(opaqueHits),
        items,
        rootKeys,
        rootFormulas,
        selectionAddress: selection.address,
        safeDelete,
      };
    });

    state.lastResults = results.items;
    renderResults(results);
    setStatus(
      `${results.direction}: ${results.items.length} cell${results.items.length === 1 ? "" : "s"} across ${countSheets(results.items)} sheet${countSheets(results.items) === 1 ? "" : "s"}.`,
    );
  } catch (e) {
    console.error("[auditor] failed at step:", step, e);
    const detail = e.debugInfo ? ` (${e.debugInfo.code || e.debugInfo.errorLocation || ""})` : "";
    setStatus(`Error at "${step}": ${e.message || e}${detail}`, "err");
    renderDiagnostics();
  }
}

const diagLog = [];
function diag(msg, data) {
  diagLog.push({ msg, data });
  if (diagLog.length > 200) diagLog.shift();
}
function renderDiagnostics() {
  const root = document.getElementById("results");
  const panel = document.createElement("div");
  panel.className = "diag";
  panel.innerHTML = `<div class="diag-title">Diagnostic log (last ${diagLog.length} events)</div>`;
  const pre = document.createElement("pre");
  pre.className = "diag-log";
  pre.textContent = diagLog
    .map((e) => `${e.msg}${e.data !== undefined ? "  " + JSON.stringify(e.data) : ""}`)
    .join("\n");
  panel.appendChild(pre);
  root.appendChild(panel);
}

function countSheets(items) {
  return new Set(items.map((i) => i.sheet)).size;
}

/**
 * Extract every cell/range reference from a formula string.
 * Returns an array of qualified address strings like "CF!K2" or "Inputs!D88:D90".
 *
 * Strategy:
 *   1. Strip string literals "..." so cell-shaped text inside strings doesn't match.
 *   2. Strip the leading "=".
 *   3. Regex over the result for [SheetQualifier!]CellRef[:CellRef] patterns.
 *   4. Default unqualified refs to the formula's home sheet.
 *
 * Known limits:
 *   - Defined names that look like cell refs (e.g. "MRG2") will be misclassified.
 *   - Table references (Table1[Column]) are not resolved to addresses.
 *   - INDIRECT/OFFSET targets aren't evaluated (the opacity warning flags this).
 */
const REF_RE =
  /(?:('(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)/g;

function extractReferencesFromFormula(formula, homeSheet) {
  if (!formula || !formula.startsWith("=")) return [];
  // We scan the original formula text so positions are usable for highlighting.
  // But we need to skip refs that fall inside string literals — pre-compute
  // the literal ranges and reject any match that overlaps them.
  const stringRanges = [];
  const litRe = /"(?:[^"]|"")*"/g;
  let lm;
  while ((lm = litRe.exec(formula)) !== null) stringRanges.push([lm.index, lm.index + lm[0].length]);
  const inString = (i) => stringRanges.some(([a, b]) => i >= a && i < b);

  const out = [];
  let m;
  REF_RE.lastIndex = 0;
  while ((m = REF_RE.exec(formula)) !== null) {
    if (inString(m.index)) continue;
    let sheet = m[1];
    const ref = m[2];
    if (sheet) {
      if (sheet.startsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
    } else {
      sheet = homeSheet;
    }
    const qualified = `${sheet}!${ref.replace(/\$/g, "")}`;
    out.push({
      qualified,            // e.g. "Inputs!D88"
      raw: m[0],            // the literal text in the formula, e.g. "Inputs!$D$88"
      start: m.index,       // position of the raw match in the formula
      end: m.index + m[0].length,
      sheet,
      address: ref.replace(/\$/g, ""),
    });
  }
  return out;
}

// Backwards-compat shim: callers that just want unique qualified strings.
function extractReferenceKeys(formula, homeSheet) {
  const refs = extractReferencesFromFormula(formula, homeSheet);
  const seen = new Set();
  const out = [];
  for (const r of refs) {
    if (seen.has(r.qualified)) continue;
    seen.add(r.qualified);
    out.push(r.qualified);
  }
  return out;
}

/**
 * Parse any of:
 *   "Sheet1!A1:B5"
 *   "Sheet1!$A$1:$B$5"
 *   "'Sheet Name'!A1"
 *   "[Book1.xlsx]Sheet1!$A$1"
 *   "[Book1.xlsx]'Sheet Name'!$A$1"
 * into {sheet, address, external}. Strips $ absolute markers and workbook prefix.
 * Returns null only if structurally unparseable.
 */
function parseQualifiedAddress(s) {
  if (!s) return null;
  let rest = s;
  let external = false;
  const wb = rest.match(/^\[([^\]]+)\](.+)$/);
  if (wb) { external = true; rest = wb[2]; }
  // Sheet part: either '...' (quoted, possibly with !) or bare up to first !
  const m = rest.match(/^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/);
  if (!m) return null;
  const sheet = (m[1] ? m[1].replace(/''/g, "'") : m[2]).trim();
  const address = m[3].replace(/\$/g, "").trim();
  return { sheet, address, external };
}

/**
 * "A1:B3" → [A1, B1, A2, B2, A3, B3]. "A1" stays singleton.
 * Whole-column ("A:A") or whole-row ("1:1") refs are kept as-is — we don't try
 * to expand 1M cells.
 */
function expandRange({ sheet, address }) {
  if (!address.includes(":")) return [{ sheet, address }];
  const [start, end] = address.split(":");
  const startM = start.match(/^\$?([A-Z]+)\$?(\d+)$/);
  const endM = end.match(/^\$?([A-Z]+)\$?(\d+)$/);
  if (!startM || !endM) return [{ sheet, address }]; // whole col/row — leave as-is
  const c0 = colLettersToIndex(startM[1]);
  const r0 = parseInt(startM[2], 10) - 1;
  const c1 = colLettersToIndex(endM[1]);
  const r1 = parseInt(endM[2], 10) - 1;
  const cells = [];
  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      cells.push({ sheet, address: cellAddress(r, c) });
    }
  }
  return cells;
}

function colLettersToIndex(letters) {
  let n = 0;
  for (let i = 0; i < letters.length; i++) {
    n = n * 26 + (letters.charCodeAt(i) - 64);
  }
  return n - 1;
}

function cellAddress(row, col) {
  let s = "";
  let n = col;
  while (true) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
    if (n < 0) break;
  }
  return s + (row + 1);
}

function cmpAddress(a, b) {
  const ma = a.match(/^([A-Z]+)(\d+)/);
  const mb = b.match(/^([A-Z]+)(\d+)/);
  if (!ma || !mb) return a.localeCompare(b);
  const ra = parseInt(ma[2], 10), rb = parseInt(mb[2], 10);
  if (ra !== rb) return ra - rb;
  const ca = colLettersToIndex(ma[1]), cb = colLettersToIndex(mb[1]);
  return ca - cb;
}

// Cached during last render so the keyboard handler / focus events can look
// up the right context formula without re-parsing.
let lastRenderedResults = null;
let formulasByKey = new Map();

function renderResults(results) {
  lastRenderedResults = results;
  formulasByKey = new Map();
  for (const k of Object.keys(results.rootFormulas || {})) formulasByKey.set(k, results.rootFormulas[k]);
  for (const item of results.items) formulasByKey.set(item.key, item.formula);

  const root = document.getElementById("results");
  root.innerHTML = "";

  // ── Audit-context badge ────────────────────────────────────────────────
  // Shows which cell the visible audit is rooted on, plus a one-click clear.
  // The selection label up top tracks the *current* Excel selection, so this
  // makes it obvious those can diverge.
  const badge = document.createElement("div");
  badge.className = "audit-badge";
  const rootLabel = results.rootKeys.length === 1
    ? results.rootKeys[0].replace("!", "!")
    : `${results.rootKeys.length} cells`;
  badge.innerHTML = `
    <span class="audit-badge-label">${results.direction === "precedents" ? "Precedents of" : "Dependents of"}</span>
    <span class="audit-badge-addr">${escapeHTML(rootLabel)}</span>
    <button class="audit-badge-clear" title="Clear results (Esc)">×</button>
  `;
  badge.querySelector(".audit-badge-clear").addEventListener("click", clearResults);
  root.appendChild(badge);

  // ── Focused-formula panel ──────────────────────────────────────────────
  // Always visible. Updates on focus of any tree item, showing the formula
  // expressing the relationship between the focused cell and its tree parent.
  const fp = document.createElement("div");
  fp.className = "formula-panel";
  fp.innerHTML = renderFormulaPanelHTML(null);
  root.appendChild(fp);

  // ── Opacity + safe-delete warnings ─────────────────────────────────────
  if (results.opaque.length) {
    const warn = document.createElement("div");
    warn.className = "opacity-warning";
    warn.innerHTML = `<strong>Opaque references detected:</strong> ${results.opaque.join(", ")}. Excel can't statically trace through these — results may be incomplete.`;
    root.appendChild(warn);
  }
  if (results.safeDelete && results.direction === "dependents" && results.items.length) {
    const safeNote = document.createElement("div");
    safeNote.className = "opacity-warning";
    safeNote.style.background = "#fde8e3";
    safeNote.style.borderColor = "#e5a797";
    safeNote.style.color = "#6d2a1a";
    safeNote.innerHTML = `<strong>${results.items.length} downstream dependent${results.items.length === 1 ? "" : "s"}.</strong> Deleting the selection will break these references (#REF!).`;
    root.appendChild(safeNote);
  }

  if (!results.items.length) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent =
      results.direction === "precedents"
        ? "No precedents. The selection has no incoming dependencies."
        : "No dependents. Safe to delete (no downstream impact).";
    root.appendChild(empty);
    return;
  }

  // ── Tree ───────────────────────────────────────────────────────────────
  // Build parent → children adjacency.
  const childrenOf = new Map();
  for (const item of results.items) {
    if (!childrenOf.has(item.parentKey)) childrenOf.set(item.parentKey, []);
    childrenOf.get(item.parentKey).push(item);
  }
  // Precedents keep formula order (the order references appear in the parent
  // formula, left to right), so walking the list moves the yellow highlight
  // smoothly through the formula. Dependents have no single parent formula to
  // follow, so sort them by sheet then address for findability.
  if (results.direction !== "precedents") {
    for (const arr of childrenOf.values()) {
      arr.sort((a, b) => a.sheet === b.sheet ? cmpAddress(a.address, b.address) : a.sheet.localeCompare(b.sheet));
    }
  }

  const tree = document.createElement("div");
  tree.className = "tree";
  const multiRoot = results.rootKeys.length > 1;
  for (const rootKey of results.rootKeys) {
    const kids = childrenOf.get(rootKey) || [];
    if (!kids.length) continue;
    if (multiRoot) {
      const [s, a] = splitKey(rootKey);
      const sec = document.createElement("div");
      sec.className = "tree-section";
      sec.innerHTML = `<div class="tree-section-header">${escapeHTML(s)}!${escapeHTML(a)}</div>`;
      tree.appendChild(sec);
      appendTreeNodes(sec, kids, childrenOf, 0);
    } else {
      appendTreeNodes(tree, kids, childrenOf, 0);
    }
  }
  root.appendChild(tree);
}

function appendTreeNodes(parentEl, nodes, childrenOf, indentLevel) {
  for (const node of nodes) {
    const el = document.createElement("div");
    el.className = "cell-item tree-item";
    el.tabIndex = 0;
    el.dataset.sheet = node.sheet;
    el.dataset.address = node.address;
    el.dataset.key = node.key;
    el.style.paddingLeft = (indentLevel * 14 + 10) + "px";
    el.title = "Click to select in Excel. ↑/↓ to walk.";
    el.innerHTML = `
      <span class="cell-ref">${indentLevel > 0 ? "└ " : ""}${escapeHTML(node.sheet)}!${escapeHTML(node.address)}<span class="depth-tag">L${node.depth}</span></span>
      <span class="cell-formula">${escapeHTML(node.formula || "(value)")}</span>
    `;
    el.addEventListener("click", () => { el.focus(); navigateTo(node.sheet, node.address); });
    el.addEventListener("focus", () => updateFormulaPanel(node));
    parentEl.appendChild(el);
    const kids = childrenOf.get(node.key) || [];
    if (kids.length) appendTreeNodes(parentEl, kids, childrenOf, indentLevel + 1);
  }
}

function updateFormulaPanel(focusedNode) {
  const fp = document.querySelector(".formula-panel");
  if (fp) fp.innerHTML = renderFormulaPanelHTML(focusedNode);
}

function renderFormulaPanelHTML(focusedNode) {
  if (!lastRenderedResults) return "";
  const results = lastRenderedResults;

  // Pick the formula that EXPRESSES the focused node's relationship to its
  // parent. For precedents this is the parent's formula (parent references
  // focused). For dependents it's the focused's own formula (focused
  // references parent).
  let ctxKey, highlightKey, label;
  if (!focusedNode) {
    // Default before any focus: show the (first) root cell's formula.
    ctxKey = results.rootKeys[0];
    highlightKey = null;
    label = "Selected cell";
  } else if (results.direction === "precedents") {
    ctxKey = focusedNode.parentKey;
    highlightKey = focusedNode.key;
    label = "Formula referencing this precedent";
  } else {
    ctxKey = focusedNode.key;
    highlightKey = focusedNode.parentKey;
    label = "This dependent's formula";
  }

  const formula = formulasByKey.get(ctxKey) || "";
  const [ctxSheet] = splitKey(ctxKey);
  const titleAddr = `${ctxSheet}!${splitKey(ctxKey)[1]}`;

  const body = formula
    ? renderFormulaHighlighted(formula, ctxSheet, highlightKey)
    : `<em class="formula-empty">No formula (value cell)</em>`;

  return `
    <div class="formula-panel-head">
      <span class="formula-panel-label">${escapeHTML(label)}</span>
      <span class="formula-panel-addr">${escapeHTML(titleAddr)}</span>
      <button class="formula-copy" data-formula="${escapeHTML(formula)}" title="Copy formula">Copy</button>
    </div>
    <div class="formula-body">${body}</div>
  `;
}

function renderFormulaHighlighted(formula, homeSheet, highlightKey) {
  const refs = extractReferencesFromFormula(formula, homeSheet);
  // A "match" is any ref whose range contains the highlight target cell.
  const matches = highlightKey
    ? refs.filter((r) => refIncludesCell(r.qualified, highlightKey))
            .map((r) => ({ start: r.start, end: r.end }))
            .sort((a, b) => a.start - b.start)
    : [];
  if (!matches.length) return `<code>${escapeHTML(formula)}</code>`;
  let html = "";
  let i = 0;
  for (const m of matches) {
    html += escapeHTML(formula.slice(i, m.start));
    html += `<mark>${escapeHTML(formula.slice(m.start, m.end))}</mark>`;
    i = m.end;
  }
  html += escapeHTML(formula.slice(i));
  return `<code>${html}</code>`;
}

function refIncludesCell(refQualified, cellKey) {
  if (refQualified === cellKey) return true;
  const [rs, ra] = splitKey(refQualified);
  const [cs, ca] = splitKey(cellKey);
  if (rs !== cs) return false;
  if (!ra.includes(":")) return false;
  try {
    const cells = expandRange({ sheet: rs, address: ra });
    return cells.some((c) => c.address === ca);
  } catch (_) { return false; }
}

async function navigateTo(sheet, address) {
  try {
    await Excel.run(async (ctx) => {
      const ws = ctx.workbook.worksheets.getItem(sheet);
      ws.activate();
      const rng = ws.getRange(address);
      rng.select();
      await ctx.sync();
    });
  } catch (e) {
    setStatus(`Could not navigate to ${sheet}!${address}: ${e.message}`, "err");
  }
}

async function highlightResults() {
  if (!state.lastResults.length) {
    setStatus("Nothing to highlight — run an audit first.", "warn");
    return;
  }
  try {
    await Excel.run(async (ctx) => {
      for (const item of state.lastResults) {
        const rng = ctx.workbook.worksheets.getItem(item.sheet).getRange(item.address);
        rng.format.fill.color = HIGHLIGHT_COLOR;
        state.highlighted.push({ sheet: item.sheet, address: item.address });
      }
      await ctx.sync();
    });
    setStatus(`Highlighted ${state.lastResults.length} cells.`);
  } catch (e) {
    setStatus(`Highlight failed: ${e.message}`, "err");
  }
}

async function clearHighlights() {
  if (!state.highlighted.length) {
    setStatus("Nothing to clear.", "warn");
    return;
  }
  try {
    await Excel.run(async (ctx) => {
      for (const { sheet, address } of state.highlighted) {
        const rng = ctx.workbook.worksheets.getItem(sheet).getRange(address);
        rng.format.fill.clear();
      }
      await ctx.sync();
    });
    setStatus(`Cleared ${state.highlighted.length} highlights.`);
    state.highlighted = [];
  } catch (e) {
    setStatus(`Clear failed: ${e.message}`, "err");
  }
}

function escapeHTML(s) {
  return (s ?? "")
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

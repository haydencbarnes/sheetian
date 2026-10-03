// FortuneSheet 1.0.4 leaves a delayed focus callback alive after unmount.
// Keep the patch in the build so npm ci remains reproducible, without editing
// node_modules or vendoring the whole engine. Fail if an upgrade changes it.
export function patchDelayedFocus(source) {
  const call = /(?:core\.)?moveToEnd\(inputRef\.current\);/g;
  const matches = source.match(call);
  if (matches?.length !== 1) throw new Error("Recheck FortuneSheet's delayed input focus compatibility patch.");
  return source.replace(call, match => `if (inputRef.current?.isConnected) ${match}`);
}

export function patchPasteOwnership(source) {
  source = source.replace(/\r\n/g, "\n");
  const start = /var onPaste = (?:React\.)?useCallback\(function \(e\) \{\n    var _document\$activeEleme;/g;
  if (source.match(start)?.length !== 1) throw new Error("Recheck FortuneSheet's workbook paste ownership patch.");
  return source.replace(start, match => match.replace("\n", "\n    if (!workbookContainer.current?.contains(document.activeElement)) return;\n"));
}

// Expose the engine's hit-tested column, including frozen/scrolled columns.
// Keeping this on the existing resize handle avoids duplicating its geometry.
export function patchColumnResizeTarget(source) {
  const handle = /id: "fortune-cols-change-size",/g;
  if (source.match(handle)?.length !== 1) throw new Error("Recheck FortuneSheet's column resize target patch.");
  return source.replace(handle, '$&\n    "data-sheetian-column": allowEditRef.current ? hoverLocation.col_index : undefined,');
}

export function patchRowResizeTarget(source) {
  const handle = /className: "fortune-rows-change-size",/g;
  if (source.match(handle)?.length !== 1) throw new Error("Recheck FortuneSheet's row resize target patch.");
  const edit = source.includes("core.isAllowEdit(context)") ? "core.isAllowEdit" : "isAllowEdit";
  return source.replace(handle, `$&\n    "data-sheetian-row": hoverLocation.row_index >= 0 && ${edit}(context) && ${edit}(context, [{ row: [hoverLocation.row_index, hoverLocation.row_index], column: [0, context.visibledatacolumn.length - 1] }]) ? hoverLocation.row_index : undefined,`);
}

// Keep formula-bar layout in the engine's context so canvas geometry and mouse
// hit testing track the visible bar height. The controls never edit sheet data.
export function patchFormulaBar(source) {
  const editor = /className: "fortune-fx-editor"/g;
  const children = /\}, \/\*#__PURE__\*\/(React(?:__default\['default'\])?)\.createElement\(LocationBox, null\)/g;
  const canvas = /context\.rowHeaderWidth, context\.columnHeaderHeight, context\.devicePixelRatio\]\);/g;
  for (const site of [editor, children, canvas]) {
    if (source.match(site)?.length !== 1) throw new Error("Recheck FortuneSheet's resizable formula bar compatibility patch.");
  }
  source = source.replace(editor, '$&, style: { height: context.calculatebarHeight }');
  source = source.replace(children, (_match, react) => `}, ${react}.createElement(FormulaBarControls, {
    height: context.calculatebarHeight,
    workbookRef: refs.workbookContainer,
    onResize: function (height) { setContext(function (ctx) { ctx.calculatebarHeight = height; }); }
  }), ${react}.createElement(LocationBox, null)`);
  const effect = source.includes("var placeholderRef = React.useRef(null);") ? "React.useEffect" : "useEffect";
  const updateCanvas = effect.startsWith("React.") ? "core.updateContextWithCanvas" : "updateContextWithCanvas";
  source = source.replace(canvas, `context.rowHeaderWidth, context.columnHeaderHeight, context.devicePixelRatio, context.calculatebarHeight]);
  ${effect}(function () {
    var placeholder = placeholderRef.current;
    if (!placeholder) return;
    var observer = new placeholder.ownerDocument.defaultView.ResizeObserver(function () {
      if (!placeholder.isConnected || !refs.canvas.current) return;
      setContext(function (ctx) {
        if (ctx.luckysheetTableContentHW[0] === placeholder.clientWidth && ctx.luckysheetTableContentHW[1] === placeholder.clientHeight) return;
        ${updateCanvas}(ctx, refs.canvas.current, placeholder);
      });
    });
    observer.observe(placeholder);
    return function () { observer.disconnect(); };
  }, [refs.canvas, setContext, context.currentSheetId, context.devicePixelRatio]);`);
  return 'import { FormulaBarControls } from ' + JSON.stringify(new URL("./formula-bar.tsx", import.meta.url).pathname) + ';\n' + source;
}

export function patchMenuActions(source) {
  const apiMethod = /    calculateFormula: function calculateFormula\(id, range\) \{/g;
  if (source.match(apiMethod)?.length !== 1) throw new Error("Recheck FortuneSheet's menu action compatibility patch.");
  source = source.replace(apiMethod, `    runMenuAction: function (action, payload) {
      return setContext(function (ctx) { runMenuAction(ctx, action, cellInput, payload); });
    },
$&`);
  return 'import { runMenuAction } from ' + JSON.stringify(new URL("./menu-actions.ts", import.meta.url).pathname) + ';\n' + source;
}

// Render the zoom dropdown as an ordinary toolbar item so the engine accounts
// for its width and includes it in overflow. Use the owning workbook's context
// and existing no-history zoom update; the sheet keeps its zoom when reopened.
export function patchToolbarZoom(source) {
  const toolbar = /var getToolbarItem = (?:React\.)?useCallback\(function \(name, i\) \{/g;
  const control = /var ZoomControl = function ZoomControl\(\) \{[\s\S]*?\n\};/g;
  const footer = /\/\*#__PURE__\*\/(React(?:__default\['default'\])?)\.createElement\(ZoomControl, null\)/g;
  const resize = /\[settings\.toolbarItems, settings\.customToolbarItems\]\);/g;
  for (const site of [toolbar, control, footer, resize]) {
    if (source.match(site)?.length !== 1) throw new Error("Recheck FortuneSheet's toolbar zoom compatibility patch.");
  }
  const react = [...source.matchAll(footer)][0][1];
  const hooks = react === "React" ? "" : "React.";
  const core = react === "React" ? "" : "core.";
  source = source.replace(resize, "[settings.toolbarItems, settings.customToolbarItems, sheetWidth]);");
  source = source.replace(footer, "null");
  source = source.replace(toolbar, `$&
    if (name === "sheetian-zoom") return ${react}.createElement(ZoomControl, { key: name });`);
  source = source.replace(control, `var ZoomControl = function ZoomControl() {
  var workbook = ${hooks}useContext(WorkbookContext);
  return ${react}.createElement(ToolbarZoom, {
    zoom: workbook.context.zoomRatio,
    onZoom: function (value) {
      if (value < ${core}MIN_ZOOM_RATIO || value > ${core}MAX_ZOOM_RATIO) return;
      var editor = workbook.refs.fxInput.current;
      if (editor && editor.ownerDocument.activeElement === editor) {
        workbook.setContext(function (ctx) {
          runMenuAction(ctx, "commit-formula", workbook.refs.cellInput.current, { editor: editor });
        });
      }
      workbook.setContext(function (ctx) {
        var index = ${core}getSheetIndex(ctx, ctx.currentSheetId);
        if (index == null) return;
        ctx.luckysheetfile[index].zoomRatio = value;
        ctx.zoomRatio = value;
      }, { noHistory: true });
    }
  });
};`);
  return 'import { ToolbarZoom } from ' + JSON.stringify(new URL("./toolbar-zoom.tsx", import.meta.url).pathname) + ';\n' + source;
}

export function patchEngine(source) { return patchToolbarZoom(patchMenuActions(patchFormulaBar(patchRowResizeTarget(patchColumnResizeTarget(patchPasteOwnership(patchDelayedFocus(source))))))); }

// Chromium snaps scrollTop to physical pixels. A row boundary rounded upward
// otherwise makes an upward wheel step select that same boundary forever.
export function patchWheelScroll(source) {
  const position = /var row_st = (_|___default\['default'\])\.sortedIndex\(visibledatarow_c, scrollTop\) \+ 1;/g;
  if (source.match(position)?.length !== 1) throw new Error("Recheck FortuneSheet's upward wheel scroll patch.");
  return source.replace(position, (_match, lodash) => `var scrollPixelRatio = scrollbarY.ownerDocument.defaultView.devicePixelRatio || 1;
  var upwardRounding = e.deltaY < 0 ? 0.5 / scrollPixelRatio + 0.0001 : 0;
  var row_st = ${lodash}.sortedIndex(visibledatarow_c, scrollTop - upwardRounding) + 1;`);
}

// The parser lexer treats bare TRUE/FALSE as column labels and rejects them
// before looking up its built-in variables. Recognize only the exact literals
// here, leaving quoted sheet names, cell references and strings to the parser.
export function patchBooleanLiterals(source) {
  const entry = /value: function _callCellValue\(label\) \{/g;
  if (source.match(entry)?.length !== 1) throw new Error("Recheck FortuneSheet's boolean literal parser patch.");
  return source.replace(entry, '$&\n      if (/^(TRUE|FALSE)$/i.test(label)) return label.toUpperCase() === "TRUE";');
}

// Formula source can contain line breaks (for example, exported Google Sheets
// IF expressions). The editor must not convert those formulas to rich text.
export function patchFormulaEditing(source) {
  const multiline = /if \(!isCurInline && inputText && inputText.length > 0\) \{/g;
  if (source.match(multiline)?.length !== 1) throw new Error("Recheck FortuneSheet's multiline formula editing patch.");
  source = source.replace(multiline, "if (!isCurInline && inputText && inputText.length > 0 && !isFormula(inputText)) {");
  const previous = /var curv = flowdata\[r\]\[c\];\n  var oldValue = /g;
  if (source.match(previous)?.length !== 1) throw new Error("Recheck FortuneSheet's unchanged formula editing patch.");
  // contenteditable normalizes CRLF to LF. An unchanged commit must retain
  // the original formula and its cached result, even for unsupported functions.
  source = source.replace(previous, `var curv = flowdata[r][c];
  if (value == null && isFormula(curv?.f) && typeof inputText === "string" && inputText.replace(/\\r\\n?/g, "\\n") === curv.f.replace(/\\r\\n?/g, "\\n")) {
    cancelNormalSelected(ctx);
    return;
  }
  var oldValue = `);
  // Dependency validation ran before trimming whitespace around references,
  // so references on their own lines were silently omitted from the graph.
  const reference = /var t = formulaTextArray\[_j2\];/g;
  if (source.match(reference)?.length !== 1) throw new Error("Recheck FortuneSheet's multiline formula dependency patch.");
  return source.replace(reference, "var t = formulaTextArray[_j2].trim();");
}

// FormulaJS declares TEXT but throws "not implemented" when it is called.
// Use the same Excel number/date formatter as the grid and XLSX importer.
export function patchTextFormula(source) {
  const parser = /this\.parser = new (?:formulaParser\.)?Parser\(\);/g;
  if (source.match(parser)?.length !== 1) throw new Error("Recheck FortuneSheet's TEXT formula compatibility patch.");
  source = source.replace(parser, `$&
    this.parser.setFunction("TEXT", function (params) {
      if (params.length !== 2) throw new Error("#N/A");
      var value = params[0], format = params[1];
      if (value instanceof Error) throw value;
      if (format instanceof Error) throw format;
      if (typeof format !== "string" || Array.isArray(value)) throw new Error("#VALUE!");
      if (value == null) value = 0;
      if (typeof value === "string" && value.trim() !== "" && Number.isFinite(Number(value))) value = Number(value);
      try { return String(update(format, value)); }
      catch (_) { throw new Error("#VALUE!"); }
    });`);
  // Formula results already have a type. Running returned text through the
  // input heuristics converts currency messages, dates and numeric text into
  // numbers, and drops empty strings. Preserve the parser's string result.
  const result = /  if \(isRealNull\(vupdate\)\) \{/g;
  if (source.match(result)?.length !== 1) throw new Error("Recheck FortuneSheet's formula text result compatibility patch.");
  source = source.replace(result, `  if (cell && isFormula(cell.f) && typeof vupdate === "string" && !valueIsError(vupdate)) {
    cell.v = vupdate;
    cell.m = vupdate;
    cell.ct = Object.assign({}, cell.ct || { fa: "General" }, { t: "s" });
    d[r][c] = cell;
    return;
  }
$&`);
  // Only complete a single missing closing parenthesis. Appending one to an
  // already over-closed formula would corrupt it again on every calculation.
  const brackets = /if \(!checkBracketNum\(txt\)\) \{/g;
  if (source.match(brackets)?.length !== 1) throw new Error("Recheck FortuneSheet's formula parenthesis compatibility patch.");
  return source.replace(brackets, 'if (!checkBracketNum(txt) && checkBracketNum(txt + ")")) {');
}

// Menu copy uses the desktop clipboard while keeping the engine's selection,
// merge checks and formula-aware internal copy state. Keyboard copy retains
// the existing writer. An optional writer avoids focusing a hidden DOM editor.
export function patchCopyWriter(source) {
  const sites = [
    [/function copy\(ctx\) \{/g, "function copy(ctx, writeClipboard) {"],
    [/function handleCopy\(ctx\) \{/g, "function handleCopy(ctx, writeClipboard) {"],
    [/  copy\(ctx\);\n  ctx\.luckysheet_paste_iscut = false;/g, "  copy(ctx, writeClipboard);\n  ctx.luckysheet_paste_iscut = false;"],
    [/    clipboard\.writeHtml\(cpdata\);/g, "    (writeClipboard || clipboard.writeHtml)(cpdata);"],
  ];
  for (const [site, replacement] of sites) {
    if (source.match(site)?.length !== 1) throw new Error("Recheck FortuneSheet's menu clipboard compatibility patch.");
    source = source.replace(site, replacement);
  }
  return source;
}

export function patchCore(source) { return patchCopyWriter(patchTextFormula(patchFormulaEditing(patchWheelScroll(source)))); }

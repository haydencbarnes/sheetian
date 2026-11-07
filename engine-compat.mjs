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
  const hover = /if \(col_index !== hoverLocation\.col_index\) \{/g;
  if (source.match(handle)?.length !== 1 || source.match(hover)?.length !== 1) throw new Error("Recheck FortuneSheet's column resize target patch.");
  source = source.replace(hover, "if (col_index !== hoverLocation.col_index || col !== hoverLocation.col || col_pre !== hoverLocation.col_pre) {");
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
  const canvas = /(?:React\.)?useEffect\(function \(\) \{\n    setContext\(function \(draftCtx\) \{\n      return (?:core\.)?updateContextWithCanvas\(draftCtx, refs.canvas.current, placeholderRef.current\);\n    \}\);\n  \}, \[refs.canvas, setContext, context.rowHeaderWidth, context.columnHeaderHeight, context.devicePixelRatio\]\);/g;
  const paint = /(?:React\.)?useEffect\(function \(\) \{\n    var _refs\$globalCache\$fre, _freeze\$horizontal, _freeze\$vertical;/g;
  const freeze = /(?:React\.)?useEffect\(function \(\) \{\n    (?:core\.)?initFreeze\(context, refs.globalCache, context.currentSheetId\);/g;
  for (const site of [editor, children, canvas, paint, freeze]) {
    if (source.match(site)?.length !== 1) throw new Error("Recheck FortuneSheet's resizable formula bar compatibility patch.");
  }
  const effect = source.includes("var placeholderRef = React.useRef(null);") ? "React.useEffect" : "useEffect";
  const layoutEffect = effect.replace("useEffect", "useLayoutEffect");
  const updateCanvas = effect.startsWith("React.") ? "core.updateContextWithCanvas" : "updateContextWithCanvas";
  // Canvas resizing erases its pixels. Rebuild frozen-pane geometry and draw
  // during the commit, before the browser can expose an incomplete frame.
  source = source.replace(freeze, match => match.replace("useEffect", "useLayoutEffect"));
  source = source.replace(paint, match => match.replace("useEffect", "useLayoutEffect") + `
    // A height commit runs before its geometry commit. Draw only the context
    // that matches the committed viewport; both commits finish before paint.
    if (context.luckysheetTableContentHW[0] !== placeholderRef.current.clientWidth ||
        context.luckysheetTableContentHW[1] !== placeholderRef.current.clientHeight) return;`);
  source = source.replace(editor, '$&, style: { height: context.calculatebarHeight }');
  source = source.replace(children, (_match, react) => `}, ${react}.createElement(FormulaBarControls, {
    height: context.calculatebarHeight,
    workbookRef: refs.workbookContainer,
    onResize: function (height) {
      setContext(function (ctx) {
        ctx.calculatebarHeight = height;
      }, { noHistory: true });
    }
  }), ${react}.createElement(LocationBox, null)`);
  source = source.replace(canvas, `${layoutEffect}(function () {
    var placeholder = placeholderRef.current;
    var canvas = refs.canvas.current;
    if (!placeholder || !canvas) return;
    var width = placeholder.clientWidth, height = placeholder.clientHeight;
    // Geometry can itself change flex layout. Recheck after every commit,
    // stopping as soon as the bitmap, overlay and viewport agree.
    if (context.luckysheetTableContentHW[0] === width && context.luckysheetTableContentHW[1] === height &&
        context.cellmainWidth === width - context.rowHeaderWidth && context.cellmainHeight === height - context.columnHeaderHeight &&
        canvas.width === Math.ceil(width * context.devicePixelRatio) && canvas.height === Math.ceil(height * context.devicePixelRatio)) return;
    setContext(function (ctx) {
      ${updateCanvas}(ctx, canvas, placeholder);
    }, { noHistory: true });
  });`);
  return 'import { FormulaBarControls } from ' + JSON.stringify(new URL("./formula-bar.tsx", import.meta.url).pathname) + ';\n' + source;
}

// Obsidian resizes panes without resizing their window. Observe the actual
// viewport and finish the canvas update and layout-effect repaint together.
// Clearing the bitmap in an ordinary state updater can expose a blank frame
// while React defers the commit. A global synthetic resize also clears every
// other workbook, even when none of their dimensions changed.
export function patchViewportResize(source) {
  const listener = /(?:React\.)?useEffect\(function \(\) \{\n    function resize\(\) \{[\s\S]*?\n  \}, \[data, refs.canvas, setContext, settings.devicePixelRatio\]\);/g;
  if (source.match(listener)?.length !== 1) throw new Error("Recheck FortuneSheet's viewport resize compatibility patch.");
  const commonjs = source.includes("var placeholderRef = React.useRef(null);");
  const effect = commonjs ? "React.useEffect" : "useEffect";
  const updateCanvas = commonjs ? "core.updateContextWithCanvas" : "updateContextWithCanvas";
  source = source.replace(listener, `${effect}(function () {
    var placeholder = placeholderRef.current;
    if (!placeholder) return;
    var ownerWindow = placeholder.ownerDocument.defaultView;
    function resize() {
      if (!data || !placeholder.isConnected || !refs.canvas.current) return;
      sheetianFlushResize(function () {
        setContext(function (ctx) {
          var pixelRatio = settings.devicePixelRatio || ownerWindow.devicePixelRatio;
          if (ctx.luckysheetTableContentHW[0] === placeholder.clientWidth &&
              ctx.luckysheetTableContentHW[1] === placeholder.clientHeight && ctx.devicePixelRatio === pixelRatio) return;
          ctx.devicePixelRatio = pixelRatio;
          ${updateCanvas}(ctx, refs.canvas.current, placeholder);
        }, { noHistory: true });
      });
    }
    var observer = new ownerWindow.ResizeObserver(resize);
    observer.observe(placeholder);
    ownerWindow.addEventListener("resize", resize);
    return function () {
      observer.disconnect();
      ownerWindow.removeEventListener("resize", resize);
    };
  }, [data, refs.canvas, setContext, settings.devicePixelRatio]);`);
  return 'import { flushSync as sheetianFlushResize } from "react-dom";\n' + source;
}

// History contains workbook data, while the engine keeps a separate active
// config for drawing and hit testing. Restore that config along with history
// so undo/redo of widths, heights and AutoFit also restores the visible grid.
export function patchHistoryLayout(source) {
  const history = /var newContext = ((?:produce\.)?applyPatches\(ctx_, history\.(?:inversePatches|patches)\));/g;
  const geometry = /(?:React\.)?useEffect\(function \(\) \{\n    if \(!data\) return;\n    setContext\(function \(draftCtx\) \{\n      return (?:core\.)?updateContextWithSheetData\(draftCtx, data\);/g;
  if (source.match(history)?.length !== 2 || source.match(geometry)?.length !== 1) {
    throw new Error("Recheck FortuneSheet's history layout compatibility patch.");
  }
  const commonjs = source.includes("var produce__default =");
  const producer = commonjs ? "produce__default['default']" : "produce";
  const getIndex = commonjs ? "core.getSheetIndex" : "getSheetIndex";
  source = source.replace(history, (_match, apply) => `var newContext = ${producer}(${apply}, function (draft) {
          var index = ${getIndex}(draft, draft.currentSheetId);
          if (index != null) draft.config = draft.luckysheetfile[index].config || {};
          if (Number.isFinite(ctx_.cellmainWidth) && history.patches.some(function (patch) {
            return patch.path[0] === "luckysheetfile" && patch.path[1] === index && patch.path[2] === "config";
          })) {
            var viewport = scrollbarX.current?.clientWidth || ctx_.cellmainWidth;
            var atEnd = ctx_.scrollLeft >= Math.max(0, ctx_.ch_width - viewport) - 1;
            if (atEnd) draft.sheetianFollowResizeEnd = true;
          }
        });`);
  return source.replace(geometry, match => match.replace("useEffect", "useLayoutEffect"));
}

// A menu must fit its workbook pane before taking focus. Window-wide bounds
// and a document-wide focus query can otherwise scroll the entire view, or
// focus a different workbook's menu, even with its outer scrollbar hidden.
export function patchContextMenuLayout(source) {
  const layout = /(?:React\.)?useLayoutEffect\(function \(\) \{\n    var _refs\$workbookContain;[\s\S]*?\n  \}, \[contextMenu.x, contextMenu.y, refs.workbookContainer, setContext\]\);/g;
  const focus = /(?:React\.)?useEffect\(function \(\) \{\n    var element = document.querySelector\(".luckysheet-cols-menuitem"\);\n    if \(element\) \{\n      element.focus\(\);\n    \}\n  \}, \[\]\);/g;
  const arrow = /        ctx.contextMenu = \{\n          x: e.pageX,\n          y: 90,\n          headerMenu: true\n        \};/g;
  for (const site of [layout, focus, arrow]) {
    if (source.match(site)?.length !== 1) throw new Error("Recheck FortuneSheet's context menu layout compatibility patch.");
  }
  const hooks = source.includes("var containerRef = React.useRef(null);") ? "React." : "";
  source = source.replace(layout, `${hooks}useLayoutEffect(function () {
    var menu = containerRef.current, workbook = refs.workbookContainer.current;
    if (!menu || !workbook) return;
    var position = fitContextMenu(menu, workbook, contextMenu);
    if (position.x === contextMenu.x && position.y === contextMenu.y) return;
    setContext(function (ctx) {
      ctx.contextMenu.x = position.x;
      ctx.contextMenu.y = position.y;
    }, { noHistory: true });
  });`);
  source = source.replace(focus, `${hooks}useEffect(function () {
    var element = containerRef.current;
    var menu = element?.closest(".fortune-context-menu");
    if (element && element === menu?.querySelector(".luckysheet-cols-menuitem")) element.focus({ preventScroll: true });
  }, []);`);
  source = source.replace(arrow, `        var workbookRect = refs.workbookContainer.current.getBoundingClientRect();
        ctx.contextMenu = {
          x: e.clientX - workbookRect.left,
          y: containerRef.current.getBoundingClientRect().bottom - workbookRect.top,
          headerMenu: true
        };`);
  return 'import { fitContextMenu } from ' + JSON.stringify(new URL("./menu-layout.ts", import.meta.url).pathname) + ';\n' + source;
}

// Render width previews without changing saved data, and follow the resized
// edge before paint. The ordinary mouseup still creates one history entry.
export function patchColumnResizePreview(source) {
  const move = /var onMouseMove = (?:React\.)?useCallback\(function \(nativeEvent\) \{\n    setContext\(function \(draftCtx\) \{\n      overShowLinkCard/g;
  const up = /var onMouseUp = (?:React\.)?useCallback\(function \(nativeEvent\) \{\n    setContext\(function \(draftCtx\) \{\n      try \{\n        ((?:core\.)?handleOverlayMouseUp\([^;]+;)/g;
  const dependencies = /\[overShowLinkCard, refs.cellInput, refs.fxInput, refs.globalCache, refs.scrollbarX, refs.scrollbarY, setContext\]/g;
  const upDependencies = /\[refs.cellInput, refs.fxInput, refs.globalCache, refs.scrollbarX, refs.scrollbarY, setContext, settings, showAlert\]/g;
  const grip = /left: hoverLocation.col - 5 \+ \(hoverInFreeze \? context.scrollLeft : 0\),/g;
  const line = /className: "fortune-change-size-line",/g;
  const sheet = /var Sheet = function Sheet\(_ref\) \{[\s\S]*?\n    settings = _useContext.settings;/g;
  const geometry = /((?:React\.)?useLayoutEffect\(function \(\) \{\n    )if \(!data\) return;\n    setContext\(function \(draftCtx\) \{\n      return (?:core\.)?updateContextWithSheetData\(draftCtx, data\);/g;
  for (const site of [move, up, dependencies, upDependencies, grip, line, sheet, geometry]) {
    if (source.match(site)?.length !== 1) throw new Error("Recheck FortuneSheet's column resize preview compatibility patch.");
  }
  const core = source.includes("core.handleOverlayMouseUp") ? "core." : "";
  const hooks = core ? "React." : "";
  source = source.replace(geometry, match => match.replace("if (!data) return;", "if (!data || sheetGeometryMatches(context, data)) return;"));
  source = source.replace(sheet, `$&
  ${hooks}useLayoutEffect(function () {
    var target = resizeHistoryScrollTarget(context, refs.scrollbarX.current?.clientWidth || context.cellmainWidth);
    if (target == null) return;
    setContext(function (ctx) {
      ctx.scrollLeft = target;
      delete ctx.sheetianFollowResizeEnd;
    }, { noHistory: true });
  });`);
  source = source.replace(move, match => `var columnResize = useColumnResizePreview(context, setContext, refs);
  ` + match.replace("    setContext", "    if (columnResize.move(nativeEvent)) return;\n    setContext"));
  source = source.replace(up, (match, call) => match.replace("    setContext", "    var extraPixels = columnResize.finish();\n    setContext")
    .replace(call, `var resizedColumn = prepareColumnResizeCommit(draftCtx, extraPixels);
        ${call}
        if (resizedColumn != null) {
          var data = ${core}getFlowdata(draftCtx);
          if (data) ${core}updateContextWithSheetData(draftCtx, data);
          followColumnResize(draftCtx, resizedColumn, refs.scrollbarX.current.clientWidth);
        }`));
  source = source.replace(dependencies, match => match.replace("setContext]", "setContext, columnResize]"));
  source = source.replace(upDependencies, match => match.replace("showAlert]", "showAlert, columnResize]"));
  source = source.replace(grip, "left: (context.luckysheet_cols_change_size ? context.visibledatacolumn[context.luckysheet_cols_change_size_start[1]] : hoverLocation.col) - 5 + (hoverInFreeze ? context.scrollLeft : 0),");
  source = source.replace(line, '$&\n    style: context.luckysheet_cols_change_size ? { left: context.visibledatacolumn[context.luckysheet_cols_change_size_start[1]] - 2 } : undefined,');
  return 'import { useColumnResizePreview, prepareColumnResizeCommit, followColumnResize, resizeHistoryScrollTarget, sheetGeometryMatches } from ' + JSON.stringify(new URL("./column-resize.ts", import.meta.url).pathname) + ';\n' + source;
}

// Scrollbars must use the flex-sized sheet frame as their containing block.
// The overlay's context-driven height briefly lags while the bar is resized.
export function patchScrollbarLayout(source) {
  const scrollbars = /, \/\*#__PURE__\*\/(React(?:__default\['default'\])?)\.createElement\(ScrollBar, \{\n    axis: "x"\n  \}\), \/\*#__PURE__\*\/\1\.createElement\(ScrollBar, \{\n    axis: "y"\n  \}\)/g;
  const overlay = /\/\*#__PURE__\*\/(React(?:__default\['default'\])?)\.createElement\(SheetOverlay, null\)/g;
  const geometry = /style: axis === "x" \? \{\n      left: context.rowHeaderWidth,\n      width: "calc\(100% - "\.concat\(context.rowHeaderWidth, "px\)"\)\n    \} : \{\n      height: "100%"\n    \},/g;
  for (const site of [scrollbars, overlay, geometry]) {
    if (source.match(site)?.length !== 1) throw new Error("Recheck FortuneSheet's scrollbar layout compatibility patch.");
  }
  source = source.replace(scrollbars, "");
  source = source.replace(overlay, (_match, react) => `${react}.createElement(SheetOverlay, null),
    ${react}.createElement(ScrollBar, { axis: "x" }),
    ${react}.createElement(ScrollBar, { axis: "y" })`);
  return source.replace(geometry, `style: axis === "x" ? {
      left: context.rowHeaderWidth,
      right: "var(--sheetian-scrollbar-size)"
    } : {
      bottom: "var(--sheetian-scrollbar-size)"
    },`);
}

// React treats scroll as a continuous event and can defer its context update
// until after paint. Keep the native thumb, canvas, headers and selections in
// the same frame, and never let a delayed effect write an old offset back.
export function patchScrollSynchronization(source) {
  const handler = /onScroll: function onScroll\(\) \{\n      if \(axis === "x"\) \{\n        setContext\(function \(draftCtx\) \{\n          draftCtx\.scrollLeft = refs\.scrollbarX\.current\.scrollLeft;\n        \}\);\n      \} else \{\n        setContext\(function \(draftCtx\) \{\n          draftCtx\.scrollTop = refs\.scrollbarY\.current\.scrollTop;\n        \}\);\n      \}\n    \}/g;
  const effects = [
    /(?:React\.)?useEffect\(function \(\) \{\n    containerRef\.current\.scrollLeft = context\.scrollLeft;/g,
    /(?:React\.)?useEffect\(function \(\) \{\n    containerRef\.current\.scrollTop = context\.scrollTop;/g,
    /(?:React\.)?useEffect\(function \(\) \{\n    if \(axis === "x"\) \{\n      refs\.scrollbarX\.current\.scrollLeft = context\.scrollLeft;/g,
    /(?:React\.)?useEffect\(function \(\) \{\n    refs\.cellArea\.current\.scrollLeft = context\.scrollLeft;/g,
  ];
  for (const site of [handler, ...effects]) {
    if (source.match(site)?.length !== 1) throw new Error("Recheck FortuneSheet's scroll synchronization compatibility patch.");
  }
  source = source.replace(handler, `onScroll: function onScroll(event) {
      var property = axis === "x" ? "scrollLeft" : "scrollTop";
      var position = event.currentTarget[property];
      if (position === context[property]) return;
      sheetianFlushScroll(function () {
        setContext(function (draftCtx) {
          draftCtx[property] = position;
        }, { noHistory: true });
      });
    }`);
  for (const effect of effects) source = source.replace(effect, match => match.replace("useEffect", "useLayoutEffect"));
  return 'import { flushSync as sheetianFlushScroll } from "react-dom";\n' + source;
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
    sheetId: workbook.context.currentSheetId,
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

export function patchClipboardMenu(source) {
  const start = /var getMenuElement = (?:React\.)?useCallback\(function \(name, i\) \{\n    var _context\$luckysheet_s;/g;
  if (source.match(start)?.length !== 1) throw new Error("Recheck FortuneSheet's clipboard context menu patch.");
  const react = source.includes("var containerRef = React.useRef(null);") ? "React__default['default']" : "React";
  return source.replace(start, `$&
    var clipboardTitles = {
      "sheetian-cut": "Cut", "sheetian-copy": "Copy", "sheetian-copy-values": "Copy values only", "sheetian-copy-formulas": "Copy formulas only",
      "sheetian-paste": "Paste", "sheetian-paste-values": "Paste values only", "sheetian-paste-formulas": "Paste formulas only", "sheetian-paste-formatting": "Paste formatting only"
    };
    if (clipboardTitles[name]) return ${react}.createElement(Menu, {
      key: name,
      onClick: function () {
        setContext(function (draftCtx) {
          runMenuAction(draftCtx, name.slice(9), refs.cellInput.current);
          draftCtx.contextMenu = {};
        });
      }
    }, clipboardTitles[name]);`);
}

export function patchEngine(source) { return patchColumnResizePreview(patchContextMenuLayout(patchClipboardMenu(patchToolbarZoom(patchMenuActions(patchHistoryLayout(patchViewportResize(patchFormulaBar(patchScrollSynchronization(patchScrollbarLayout(patchRowResizeTarget(patchColumnResizeTarget(patchPasteOwnership(patchDelayedFocus(source)))))))))))))); }

// Chromium snaps scrollTop to physical pixels. A row boundary rounded upward
// otherwise makes an upward wheel step select that same boundary forever.
export function patchWheelScroll(source) {
  const position = /var row_st = (_|___default\['default'\])\.sortedIndex\(visibledatarow_c, scrollTop\) \+ 1;/g;
  if (source.match(position)?.length !== 1) throw new Error("Recheck FortuneSheet's upward wheel scroll patch.");
  return source.replace(position, (_match, lodash) => `var scrollPixelRatio = scrollbarY.ownerDocument.defaultView.devicePixelRatio || 1;
  var upwardRounding = e.deltaY < 0 ? 0.5 / scrollPixelRatio + 0.0001 : 0;
  var row_st = ${lodash}.sortedIndex(visibledatarow_c, scrollTop - upwardRounding) + 1;`);
}

// FortuneSheet adds a fixed 120px margin after the last column. It becomes
// an empty strip at the rightmost scroll position, outside the finite grid.
export function patchColumnScrollExtent(source) {
  const margin = /  var maxColumnlen = 120;\n/g;
  const extent = /  ctx\.ch_width \+= maxColumnlen;\n/g;
  if (source.match(margin)?.length !== 1 || source.match(extent)?.length !== 1) {
    throw new Error("Recheck FortuneSheet's column scroll extent compatibility patch.");
  }
  return source.replace(margin, "").replace(extent, "");
}

// Header drags are relative to their starting pointer, even at the viewport
// edge. The original 100px clamp can shrink the last column on an outward
// drag. Column deltas were also divided by zoom twice, while row resizing
// included its gridline in the saved height and changed it on a simple click.
export function patchHeaderResizing(source) {
  const sites = [
    [/ctx\.luckysheet_rows_change_size_start = \[_y, row_index\];/g,
      "ctx.luckysheet_rows_change_size_start = [_y, row_index, e.pageY];"],
    [/ctx\.luckysheet_cols_change_size_start = \[_x, col_index\];/g,
      "ctx.luckysheet_cols_change_size_start = [_x, col_index, e.pageX];"],
    [/    var winH = rect\.height;\n    var delta = y \+ 3 - ctx\.luckysheet_rows_change_size_start\[0\];\n    if \(y >= winH - 20 \+ scrollTop\) \{\n      delta = winH - 20 - ctx\.luckysheet_rows_change_size_start\[0\] \+ scrollTop;\n    \}/g,
      "    var delta = (e.pageY - ctx.luckysheet_rows_change_size_start[2]) / ctx.zoomRatio;\n    if (Math.abs(delta) < 3 / ctx.zoomRatio) return;"],
    [/    var size = ctx\.defaultrowlen;\n    if \(ctx\.visibledatarow\[ctx\.luckysheet_rows_change_size_start\[1\]\] != null\) \{\n      size = ctx\.visibledatarow\[ctx\.luckysheet_rows_change_size_start\[1\]\] - \(ctx\.visibledatarow\[ctx\.luckysheet_rows_change_size_start\[1\] - 1\] \|\| 0\);\n    \}\n    size \+= delta;/g,
      "    var size = (cfg.rowlen[ctx.luckysheet_rows_change_size_start[1]] || ctx.defaultrowlen) + delta;"],
    [/Math\.ceil\(size \/ ctx\.zoomRatio\)/g, "Math.ceil(size)", 2],
    [/    var winW = rect\.width;\n    var _delta = x \+ 3 - ctx\.luckysheet_cols_change_size_start\[0\];\n    if \(x >= winW - 100 \+ scrollLeft\) \{\n      _delta = winW - 100 - ctx\.luckysheet_cols_change_size_start\[0\] \+ scrollLeft;\n    \}/g,
      "    var _delta = e.pageX - ctx.luckysheet_cols_change_size_start[2];"],
    [/    _delta \/= ctx\.zoomRatio;/g,
      "    _delta /= ctx.zoomRatio;\n    if (Math.abs(_delta) < 3 / ctx.zoomRatio) return;"],
    [/    if \(Math\.abs\(_size - firstcolumnlen\) < 3\) \{\n      return;\n    \}\n/g, ""],
    [/Math\.ceil\(_size \/ ctx\.zoomRatio\)/g, "Math.ceil(_size)", 2],
  ];
  for (const [site, replacement, count = 1] of sites) {
    if (source.match(site)?.length !== count) throw new Error("Recheck FortuneSheet's header resizing compatibility patch.");
    source = source.replace(site, replacement);
  }
  for (const [position, limit, next] of [
    ["_x3", "rect.width + ctx.scrollLeft - 100", "luckysheet_rows_change_size"],
    ["_y3", "rect.height + ctx.scrollTop - 20", "luckysheet_cols_freeze_drag"],
  ]) {
    const preview = new RegExp("    if \\(" + position + " < " + limit.replace(/[.+]/g, "\\$&") + "\\) \\{\\n([\\s\\S]*?)\\n    \\}\\n  \\} else if \\(ctx\\." + next + "\\) \\{", "g");
    if (source.match(preview)?.length !== 1) throw new Error("Recheck FortuneSheet's header resize preview compatibility patch.");
    source = source.replace(preview, (_match, body) => body.replace(/^  /gm, "") + "\n  } else if (ctx." + next + ") {");
  }
  return source;
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

// Bypass the browser's hidden-copy editor when moving a desktop clipboard cut.
// The existing engine move operation retains formulas and clears the source.
export function patchCutPaste(source) {
  const entry = /function handlePasteByClick\(ctx, clipboardData, triggerType\) \{/g;
  if (source.match(entry)?.length !== 1) throw new Error("Recheck FortuneSheet's desktop cut/paste compatibility patch.");
  return source.replace(entry, `$&
  if (triggerType === "sheetian-cut") {
    pasteHandlerOfCutPaste(ctx, ctx.luckysheet_copy_save);
    ctx.luckysheet_paste_iscut = false;
    ctx.luckysheet_selection_range = [];
    return;
  }`);
}

// Newly pasted formulas must calculate even when their inputs lie outside the
// pasted range. Include those roots in the engine's normal dependency ordering
// without recalculating unrelated formulas or discarding their cached results.
export function patchPasteFormulaCalculation(source) {
  const force = /    if \(isForce\) \{\n      updateValueArray\.push\(formulaObject\);\n    \}/g;
  if (source.match(force)?.length !== 1) throw new Error("Recheck FortuneSheet's pasted formula calculation patch.");
  return source.replace(force, `    if (isForce || ctx.sheetianPasteFormulaKeys?.includes(key)) {
      updateValueArray.push(formulaObject);
    }`);
}

export function patchCore(source) { return patchPasteFormulaCalculation(patchCutPaste(patchCopyWriter(patchTextFormula(patchFormulaEditing(patchHeaderResizing(patchColumnScrollExtent(patchWheelScroll(source)))))))); }

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

export function patchEngine(source) { return patchColumnResizeTarget(patchPasteOwnership(patchDelayedFocus(source))); }

// Chromium snaps scrollTop to physical pixels. A row boundary rounded upward
// otherwise makes an upward wheel step select that same boundary forever.
export function patchWheelScroll(source) {
  const position = /var row_st = (_|___default\['default'\])\.sortedIndex\(visibledatarow_c, scrollTop\) \+ 1;/g;
  if (source.match(position)?.length !== 1) throw new Error("Recheck FortuneSheet's upward wheel scroll patch.");
  return source.replace(position, (_match, lodash) => `var scrollPixelRatio = scrollbarY.ownerDocument.defaultView.devicePixelRatio || 1;
  var upwardRounding = e.deltaY < 0 ? 0.5 / scrollPixelRatio + 0.0001 : 0;
  var row_st = ${lodash}.sortedIndex(visibledatarow_c, scrollTop - upwardRounding) + 1;`);
}

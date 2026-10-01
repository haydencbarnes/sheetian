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

export function patchEngine(source) { return patchPasteOwnership(patchDelayedFocus(source)); }

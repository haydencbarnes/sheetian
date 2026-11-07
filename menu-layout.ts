// Menu coordinates are relative to the workbook, which may occupy only one
// pane of an Obsidian window. Constrain both size and position to that pane.
export function fitContextMenu(menu: HTMLElement, workbook: HTMLElement, anchor: { x?: number; y?: number }): { x: number; y: number } {
  const rect = workbook.getBoundingClientRect();
  const ownerWindow = workbook.ownerDocument.defaultView!;
  const left = Math.max(0, -rect.left);
  const top = Math.max(0, -rect.top);
  const right = Math.max(left, Math.min(workbook.clientWidth, ownerWindow.innerWidth - rect.left));
  const bottom = Math.max(top, Math.min(workbook.clientHeight, ownerWindow.innerHeight - rect.top));
  menu.style.maxWidth = `${right - left}px`;
  menu.style.maxHeight = `${bottom - top}px`;
  const size = menu.getBoundingClientRect();
  let x = anchor.x ?? 0, y = anchor.y ?? 0;
  if (x + size.width > right) x -= size.width;
  if (y + size.height > bottom) y -= size.height;
  return {
    x: Math.max(left, Math.min(x, right - size.width)),
    y: Math.max(top, Math.min(y, bottom - size.height)),
  };
}

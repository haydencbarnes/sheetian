import * as React from "react";

const MIN_HEIGHT = 29;
const EXPANDED_HEIGHT = 160;

interface Props {
  height: number;
  onResize: (height: number) => void;
  workbookRef: React.RefObject<HTMLDivElement>;
}

export function FormulaBarControls({ height, onResize, workbookRef }: Props): JSX.Element {
  const [maximum, setMaximum] = React.useState(420);
  const drag = React.useRef<{ y: number; height: number; pointer: number } | null>(null);
  const expanded = React.useRef(EXPANDED_HEIGHT);
  React.useLayoutEffect(() => {
    const book = workbookRef.current;
    if (!book) return;
    const observer = new book.ownerDocument.defaultView!.ResizeObserver(() => {
      setMaximum(Math.max(MIN_HEIGHT, Math.min(420, Math.floor(book.clientHeight / 2))));
    });
    observer.observe(book);
    return () => observer.disconnect();
  }, [workbookRef]);
  React.useLayoutEffect(() => {
    if (height > maximum) onResize(maximum);
  }, [height, maximum, onResize]);
  const resize = (value: number) => onResize(Math.max(MIN_HEIGHT, Math.min(maximum, Math.round(value))));
  const toggle = () => {
    if (height > MIN_HEIGHT) { expanded.current = height; resize(MIN_HEIGHT); }
    else resize(expanded.current);
  };
  const stop = (event: React.SyntheticEvent) => { event.preventDefault(); event.stopPropagation(); };
  const finish = (event: React.PointerEvent<HTMLDivElement>) => {
    if (drag.current?.pointer !== event.pointerId) return;
    drag.current = null;
    if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId);
    event.stopPropagation();
  };
  return <>
    <button type="button" className="sheetian-formula-toggle"
      aria-label={height > MIN_HEIGHT ? "Collapse formula bar" : "Expand formula bar"}
      aria-expanded={height > MIN_HEIGHT} title={height > MIN_HEIGHT ? "Collapse formula bar" : "Expand formula bar"}
      onMouseDown={stop} onClick={event => { stop(event); toggle(); }}
      onKeyDown={event => event.stopPropagation()}>{height > MIN_HEIGHT ? "⌃" : "⌄"}</button>
    <div className="sheetian-formula-resize" role="separator" tabIndex={0}
      aria-label="Resize formula bar" aria-orientation="horizontal"
      aria-valuemin={MIN_HEIGHT} aria-valuemax={maximum} aria-valuenow={height}
      title="Drag to resize the formula bar; double-click to expand or collapse"
      onMouseDown={stop} onClick={stop}
      onDoubleClick={event => { stop(event); toggle(); }}
      onPointerDown={event => {
        if (event.button !== 0) return;
        stop(event);
        drag.current = { y: event.clientY, height, pointer: event.pointerId };
        event.currentTarget.setPointerCapture(event.pointerId);
      }}
      onPointerMove={event => {
        const start = drag.current;
        if (!start || start.pointer !== event.pointerId) return;
        stop(event); resize(start.height + event.clientY - start.y);
      }}
      onPointerUp={finish} onPointerCancel={finish} onLostPointerCapture={() => { drag.current = null; }}
      onKeyDown={event => {
        event.stopPropagation();
        if (!["ArrowUp", "ArrowDown", "Home", "End", "Enter", " "].includes(event.key)) return;
        event.preventDefault();
        if (event.key === "Home") resize(MIN_HEIGHT);
        else if (event.key === "End") resize(maximum);
        else if (event.key === "ArrowUp") resize(height - (event.shiftKey ? 50 : 10));
        else if (event.key === "ArrowDown") resize(height + (event.shiftKey ? 50 : 10));
        else toggle();
      }} />
  </>;
}

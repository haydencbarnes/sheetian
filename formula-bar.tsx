import * as React from "react";
import { flushSync } from "react-dom";

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
  const frame = React.useRef<number | null>(null);
  const pending = React.useRef<number | null>(null);
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
  const resize = (value: number) => {
    const next = Math.max(MIN_HEIGHT, Math.min(maximum, Math.round(value)));
    // A canvas size change clears its bitmap. Commit the matching grid redraw
    // within this frame instead of letting React defer it past browser paint.
    if (next !== height) flushSync(() => onResize(next));
  };
  const resizeRef = React.useRef(resize);
  React.useLayoutEffect(() => { resizeRef.current = resize; });
  const flush = () => {
    const ownerWindow = workbookRef.current?.ownerDocument.defaultView;
    if (frame.current !== null) ownerWindow?.cancelAnimationFrame(frame.current);
    frame.current = null;
    const next = pending.current;
    pending.current = null;
    if (next !== null) resizeRef.current(next);
  };
  React.useLayoutEffect(() => {
    const ownerWindow = workbookRef.current?.ownerDocument.defaultView;
    return () => {
      if (frame.current !== null) ownerWindow?.cancelAnimationFrame(frame.current);
    };
  }, [workbookRef]);
  const toggle = () => {
    if (height > MIN_HEIGHT) { expanded.current = height; resize(MIN_HEIGHT); }
    else resize(expanded.current);
  };
  const stop = (event: React.SyntheticEvent) => { event.preventDefault(); event.stopPropagation(); };
  const finish = (event: React.PointerEvent<HTMLDivElement>) => {
    if (drag.current?.pointer !== event.pointerId) return;
    flush();
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
        stop(event);
        pending.current = start.height + event.clientY - start.y;
        const ownerWindow = event.currentTarget.ownerDocument.defaultView!;
        if (frame.current === null) frame.current = ownerWindow.requestAnimationFrame(flush);
      }}
      onPointerUp={finish} onPointerCancel={finish} onLostPointerCapture={finish}
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

import * as React from "react";
import { MIN_ZOOM_RATIO, MAX_ZOOM_RATIO } from "@fortune-sheet/core";

const PRESETS = [50, 75, 90, 100, 125, 150, 175, 200];

interface Props {
  zoom: number;
  sheetId: string;
  onZoom: (zoom: number) => void;
}

export function ToolbarZoom({ zoom, sheetId, onZoom }: Props): JSX.Element {
  const percentage = Number((zoom * 100).toPrecision(12));
  const [draft, setDraft] = React.useState(`${percentage}%`);
  const dirty = React.useRef(false);
  React.useEffect(() => { dirty.current = false; setDraft(`${percentage}%`); }, [zoom, sheetId]);
  // Keep saved zoom levels and keyboard/menu changes visible, even when they
  // are between presets or outside the dropdown's usual range.
  const choices = PRESETS.includes(percentage) ? PRESETS : [...PRESETS, percentage].sort((a, b) => a - b);
  const apply = (value: number) => {
    dirty.current = false;
    setDraft(`${value}%`);
    if (value / 100 !== zoom) onZoom(value / 100);
  };
  const commit = () => {
    if (!dirty.current) return;
    dirty.current = false;
    const match = draft.trim().match(/^(\d+(?:\.\d+)?)\s*%?$/);
    const value = match ? Number(match[1]) : NaN;
    if (!Number.isFinite(value) || value / 100 < MIN_ZOOM_RATIO || value / 100 > MAX_ZOOM_RATIO) {
      setDraft(`${percentage}%`);
      return;
    }
    apply(value);
  };
  return <div className="sheetian-toolbar-zoom fortune-toolbar-item" title="Zoom">
    <input className="sheetian-zoom-input" type="text" inputMode="decimal"
      aria-label="Zoom percentage" value={draft} spellCheck={false}
      onFocus={event => event.currentTarget.select()}
      onChange={event => { dirty.current = true; setDraft(event.currentTarget.value); }}
      onBlur={commit}
      onKeyDown={event => {
        event.stopPropagation();
        if (event.key === "Enter") { event.preventDefault(); commit(); event.currentTarget.select(); }
        if (event.key === "Escape") { event.preventDefault(); dirty.current = false; setDraft(`${percentage}%`); }
      }}
    />
    <span className="sheetian-zoom-picker">
    <select className="sheetian-zoom-select" aria-label="Zoom presets" value={percentage}
    onChange={event => apply(Number(event.currentTarget.value))}
    onKeyDown={event => {
      event.stopPropagation();
      if (event.altKey || event.ctrlKey || event.metaKey) return;
      const index = choices.indexOf(percentage);
      const next = event.key === "Home" ? 0 : event.key === "End" ? choices.length - 1
        : event.key === "ArrowDown" ? Math.min(index + 1, choices.length - 1)
        : event.key === "ArrowUp" ? Math.max(index - 1, 0) : null;
      if (next === null) return;
      // Keep arrow keys on this control rather than the sheet's navigation.
      event.preventDefault();
      apply(choices[next]);
    }}
  >
    {choices.map(value => <option key={value} value={value}>{value}%</option>)}
    </select>
    </span>
  </div>;
}

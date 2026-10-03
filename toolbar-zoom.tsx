import * as React from "react";

const PRESETS = [50, 75, 90, 100, 125, 150, 175, 200];

interface Props {
  zoom: number;
  onZoom: (zoom: number) => void;
}

export function ToolbarZoom({ zoom, onZoom }: Props): JSX.Element {
  const percentage = Math.round(zoom * 100);
  // Keep saved zoom levels and keyboard/menu changes visible, even when they
  // are between presets or outside the dropdown's usual range.
  const choices = PRESETS.includes(percentage) ? PRESETS : [...PRESETS, percentage].sort((a, b) => a - b);
  return <select
    className="sheetian-toolbar-zoom fortune-toolbar-item"
    aria-label="Zoom" title="Zoom" value={percentage}
    onChange={event => onZoom(Number(event.currentTarget.value) / 100)}
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
      onZoom(choices[next] / 100);
    }}
  >
    {choices.map(value => <option key={value} value={value}>{value}%</option>)}
  </select>;
}

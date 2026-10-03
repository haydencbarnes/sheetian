import { Menu, Modal, Notice, Scope, TextFileView, TFile, WorkspaceLeaf } from "obsidian";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { flushSync } from "react-dom";
import { Workbook } from "@fortune-sheet/react";
import { defaultSettings } from "@fortune-sheet/core";
import type { WorkbookInstance } from "@fortune-sheet/react/dist/components/Workbook";
import type { Op, Sheet } from "@fortune-sheet/core";
import { parseSheets, serializeSheets, sheetToCsv } from "./sheet-data";
import { normalizeInput } from "./input";
import { autoFitColumns, cellFont, fitColumnWidth } from "./autofit";
import { autoFitRows, fitRowHeight } from "./row-autofit";
import type { SpreadsheetSettings } from "./settings";
import { mountMenuBar, type SheetMenu, type SheetMenuItem } from "./menu-bar";
import type { MenuPayload } from "./menu-actions";

export const VIEW_TYPE_SPREADSHEET = "spreadsheet-view";

export class SpreadsheetView extends TextFileView {
  private source = "";
  private sheets: Sheet[] = [];
  private edited = false;
  private generation = 0;
  private root: Root | null = null;
  private disposeMenu: (() => void) | null = null;
  private resizeObserver: ResizeObserver | null = null;
  private resizeFrame: number | null = null;
  private workbook = React.createRef<WorkbookInstance>();

  constructor(leaf: WorkspaceLeaf, private settings: () => SpreadsheetSettings, private importFile?: () => void, private newFile?: () => void) {
    super(leaf);
    this.scope = new Scope(this.app.scope);
    // Obsidian otherwise handles F2 as Rename file before the grid sees it.
    this.scope.register([], "F2", () => {
      const area = this.contentEl.querySelector<HTMLElement>(".fortune-cell-area");
      if (area) {
        const rect = area.getBoundingClientRect();
        const ownerWindow = area.ownerDocument.defaultView!;
        const generation = this.generation;
        queueMicrotask(() => {
          if (generation !== this.generation) return;
          flushSync(() => area.dispatchEvent(new ownerWindow.MouseEvent("dblclick", {
            clientX: rect.left + 1, clientY: rect.top + 1, button: 0, bubbles: true,
          })));
          const input = this.contentEl.querySelector<HTMLElement>(".luckysheet-cell-input");
          if (input) {
            input.focus();
            const selection = ownerWindow.getSelection();
            selection?.selectAllChildren(input);
            selection?.collapseToEnd();
          }
        });
      }
      return false;
    });
    this.scope.register(["Mod"], "b", () => {
      const api = this.workbook.current;
      const range = api?.getSelection();
      if (!api || !range?.length) return false;
      const cell = api.getSheet().data?.[range[0].row[0]]?.[range[0].column[0]];
      api.setCellFormatByRange("bl", cell?.bl === 1 ? 0 : 1, range);
      return false;
    });
  }
  getViewType(): string { return VIEW_TYPE_SPREADSHEET; }
  getIcon(): string { return "table"; }

  onPaneMenu(menu: Menu, source: string): void {
    super.onPaneMenu(menu, source);
    if (this.importFile) menu.addItem(item => item.setTitle("Import Excel workbook (.xlsx)").setIcon("file-input")
      .onClick(() => this.importFile!()));
  }

  getViewData(): string {
    // Loading/normalization alone must never rewrite or erase a file.
    if (!this.edited) return this.source;
    return serializeSheets(this.workbook.current?.getAllSheets() ?? this.sheets);
  }

  setViewData(data: string, _clear: boolean): void {
    this.disposeWorkbook();
    this.source = data;
    this.edited = false;
    this.contentEl.empty();
    this.contentEl.addClass("obsidian-spreadsheet-view");
    try { this.sheets = parseSheets(data); }
    catch (error) {
      this.sheets = [];
      this.contentEl.createEl("p", {
        text: "This spreadsheet could not be read. Its original contents have been preserved. " +
          (error instanceof Error ? error.message : String(error)),
        cls: "spreadsheet-error",
      });
      return;
    }
    this.renderWorkbook();
  }

  private renderWorkbook(): void {
    const generationForMenu = this.generation;
    this.disposeMenu = mountMenuBar(this.contentEl, ["File", "Edit", "View", "Insert", "Format", "Data", "Tools", "Help"],
      () => this.sheetMenus(generationForMenu), () => this.prepareMenu());
    const container = this.contentEl.createDiv({ cls: "obsidian-spreadsheet" });
    const settings = this.settings();
    container.dataset.theme = settings.theme;
    const generation = this.generation;
    const ownerWindow = container.ownerDocument.defaultView!;
    // Native selects move focus before changing their value. Commit a draft
    // before mouse or keyboard focus reaches zoom, just as the menus do.
    container.addEventListener("mousedown", event => {
      if (generation === this.generation && (event.target as HTMLElement)?.closest(".sheetian-toolbar-zoom")) this.prepareMenu();
    }, true);
    container.addEventListener("blur", event => {
      if (generation === this.generation && (event.relatedTarget as HTMLElement | null)?.matches?.(".sheetian-toolbar-zoom")) {
        this.prepareMenu(event.target as HTMLElement);
      }
    }, true);
    container.addEventListener("keydown", event => {
      const editor = (event.target as HTMLElement)?.closest<HTMLDivElement>(".fortune-fx-input");
      if (event.key !== "Enter" || !editor || !container.contains(editor) || generation !== this.generation || !this.workbook.current) return;
      // The mirrored grid editor can clear the engine's editing flag during
      // resizing/focus changes. Commit the actual focused formula input even
      // in that state, then use the engine's usual Enter navigation.
      event.preventDefault(); event.stopImmediatePropagation();
      const api = this.workbook.current as WorkbookInstance & { runMenuAction: (action: string, payload?: MenuPayload) => void };
      flushSync(() => api.runMenuAction("commit-formula", { editor, moveDown: true }));
    }, true);
    container.addEventListener("dblclick", event => {
      const handle = (event.target as HTMLElement)?.closest<HTMLElement>(".fortune-cols-change-size, .fortune-rows-change-size");
      const isRow = handle?.classList.contains("fortune-rows-change-size");
      const value = isRow ? handle?.dataset.sheetianRow : handle?.dataset.sheetianColumn;
      if (!value || !container.contains(handle!) || generation !== this.generation) return;
      const index = Number(value);
      if (!Number.isInteger(index) || index < 0 || !this.workbook.current) return;
      event.preventDefault();
      event.stopPropagation();
      this.commitInput();
      const api = this.workbook.current;
      const sheet = api.getSheet();
      const canvas = container.ownerDocument.createElement("canvas").getContext("2d");
      if (!canvas) return;
      if (isRow) {
        const heights: Record<string, number> = {};
        for (const r of autoFitRows(sheet, index, api.getSelection())) heights[r] = fitRowHeight(sheet, r, canvas);
        api.setRowHeight(heights, { id: sheet.id }, true);
        return;
      }
      const widths: Record<string, number> = {};
      for (const c of autoFitColumns(sheet, index, api.getSelection())) {
        widths[c] = fitColumnWidth(sheet, c, (text, style) => {
          canvas.font = cellFont(style);
          return canvas.measureText(text).width;
        });
      }
      api.setColumnWidth(widths, { id: sheet.id }, true);
    }, true);
    this.root = createRoot(container);
    this.root.render(<Workbook
      ref={this.workbook} data={this.sheets} lang="en" currency={settings.currency}
      toolbarItems={defaultSettings.toolbarItems.filter(item => item !== "clear-format")
        .flatMap(item => item === "format-painter" ? [item, "sheetian-zoom"] : [item])}
      onChange={(sheets: Sheet[]) => {
        if (generation !== this.generation) return;
        this.sheets = sheets;
        if (this.edited) this.requestSave();
      }}
      onOp={(operations: Op[]) => {
        if (generation !== this.generation || operations.length === 0) return;
        this.edited = true;
        // TextFileView already debounces saves and flushes on file unload.
        this.requestSave();
      }}
      hooks={{ beforeUpdateCell: (r: number, c: number, value: unknown) => {
        const api = this.workbook.current;
        if (!api) return true;
        const sheet = api.getSheet();
        const normalized = normalizeInput(value, sheet.data?.[r]?.[c], settings.currency);
        if (!normalized) return true;
        // The hook cannot replace a primitive. Apply via the public API after
        // the engine finishes the current update, to the same sheet and file.
        queueMicrotask(() => {
          if (generation !== this.generation || !this.workbook.current) return;
          this.workbook.current.setCellValue(r, c, { ...normalized, f: undefined }, { id: sheet.id });
        });
        return false;
      } }}
    />);
    let width = -1, height = -1;
    this.resizeObserver = new ownerWindow.ResizeObserver(([entry]) => {
      if (entry.contentRect.width === width && entry.contentRect.height === height) return;
      width = entry.contentRect.width;
      height = entry.contentRect.height;
      if (this.resizeFrame !== null) ownerWindow.cancelAnimationFrame(this.resizeFrame);
      this.resizeFrame = ownerWindow.requestAnimationFrame(() => {
        this.resizeFrame = null;
        if (generation === this.generation) ownerWindow.dispatchEvent(new ownerWindow.Event("resize"));
      });
    });
    this.resizeObserver.observe(container);
  }

  private prepareMenu(draft?: HTMLElement): void {
    // Commit an active draft before opening a command menu, then restore the
    // selection so formatting/insertion applies to the same cells.
    const selection = this.workbook.current?.getSelection();
    const input = this.contentEl.querySelector<HTMLElement>(".fortune-fx-input");
    if (input && (this.contentEl.ownerDocument.activeElement === input || draft === input)) {
      const workbook = this.workbook.current as (WorkbookInstance & { runMenuAction: (action: string, payload?: MenuPayload) => void }) | null;
      // Commit through the same engine update as Enter, without moving to the
      // next row or triggering a second commit from the mirrored grid editor.
      flushSync(() => workbook?.runMenuAction("commit-formula", { editor: input as HTMLDivElement }));
      flushSync(() => input.blur());
    } else this.commitInput();
    if (selection?.length) flushSync(() => this.workbook.current?.setSelection(selection));
  }

  private sheetMenus(generation: number): SheetMenu[] {
    const api = this.workbook.current;
    const sheet = api?.getSheet();
    const selection = api?.getSelection() ?? [];
    const selected = !!api && !!selection.length;
    const item = (title: string, action: () => void | Promise<void>, disabled = false, checked?: boolean): SheetMenuItem => ({
      title, disabled, checked, action: () => { if (generation !== this.generation) return; return action(); },
    });
    const command = (action: string, payload?: MenuPayload) => {
      const workbook = this.workbook.current as (WorkbookInstance & { runMenuAction: (action: string, payload?: MenuPayload) => void }) | null;
      workbook?.runMenuAction(action, payload);
    };
    const action = (title: string, name: string, checked?: boolean) => item(title, () => command(name), !selected, checked);
    const copy = (title: string, name: "copy" | "cut") => item(title, () => command(name, { writeClipboard: html => {
      const table = this.contentEl.ownerDocument.createElement("div"); table.innerHTML = html;
      const text = Array.from(table.querySelectorAll("tr"), row => Array.from(row.querySelectorAll("td, th"), cell => cell.textContent ?? "").join("\t")).join("\n");
      require("electron").clipboard.write({ html, text });
    } }), !selected);
    const format = (title: string, attr: "ht" | "vt" | "tb", value: number | string) => item(title,
      () => { const workbook = this.workbook.current; const ranges = workbook?.getSelection(); if (ranges?.length) workbook!.setCellFormatByRange(attr, value, ranges); }, !selected);
    const insert = (title: string, type: "row" | "column", direction: "lefttop" | "rightbottom") => item(title, () => {
      const workbook = this.workbook.current, ranges = workbook?.getSelection();
      if (!workbook || !ranges?.length) return;
      const indices = ranges[0][type];
      workbook.insertRowOrColumn(type, direction === "lefttop" ? indices[0] : indices[1], indices[1] - indices[0] + 1, direction);
    }, !selected || selection.length !== 1);
    const cells = selection.length ? sheet?.data?.[selection[0].row[0]]?.[selection[0].column[0]] : null;
    return [
      { title: "File", items: [
        item("New spreadsheet", () => this.newFile?.(), !this.newFile),
        item("Import Excel workbook (.xlsx)", () => this.importFile?.(), !this.importFile),
        item("Export active sheet to CSV", () => this.exportCsv(), !api), null,
        item("Save", () => this.save()),
      ] },
      { title: "Edit", items: [
        item("Undo", () => this.workbook.current?.handleUndo(), !api),
        item("Redo", () => this.workbook.current?.handleRedo(), !api), null,
        copy("Cut", "cut"), copy("Copy", "copy"), item("Paste", () => {
          const clipboard = require("electron").clipboard;
          const document = this.contentEl.ownerDocument, win = document.defaultView!;
          const data = new win.DataTransfer(); data.setData("text/html", clipboard.readHTML()); data.setData("text/plain", clipboard.readText());
          command("paste", { clipboard: new win.ClipboardEvent("paste", { clipboardData: data }) });
        }, !selected), null,
        action("Select all", "select-all"), action("Clear contents", "clear-contents"), null,
        action("Find", "find"), action("Find and replace", "replace"),
      ] },
      { title: "View", items: [
        action("Gridlines", "gridlines", sheet?.showGridLines !== false && sheet?.showGridLines !== 0),
        item("Expand/collapse formula bar", () => this.contentEl.querySelector<HTMLButtonElement>(".sheetian-formula-toggle")?.click(), !api), null,
        ...[0.5, 0.75, 1, 1.25, 1.5, 2].map(zoom => item(`Zoom ${zoom * 100}%`, () => command("zoom", { zoom }), !api, (sheet?.zoomRatio ?? 1) === zoom)), null,
        action("Freeze through selected row", "freeze-row"), action("Freeze through selected column", "freeze-col"), action("Unfreeze rows and columns", "freeze-cancel"),
      ] },
      { title: "Insert", items: [
        insert("Rows above", "row", "lefttop"), insert("Rows below", "row", "rightbottom"),
        insert("Columns left", "column", "lefttop"), insert("Columns right", "column", "rightbottom"), null,
        item("New worksheet", () => this.workbook.current?.addSheet(), !api), action("Link", "link"),
      ] },
      { title: "Format", items: [
        action("Bold", "bold", cells?.bl === 1), action("Italic", "italic", cells?.it === 1), action("Underline", "underline", cells?.un === 1), action("Strikethrough", "strike-through", cells?.cl === 1), null,
        action("Currency", "currency-format"), action("Percent", "percentage-format"), action("More decimal places", "number-increase"), action("Fewer decimal places", "number-decrease"), null,
        format("Align left", "ht", 1), format("Align center", "ht", 0), format("Align right", "ht", 2),
        format("Wrap text", "tb", "2"), format("Clip text", "tb", "1"), format("Overflow text", "tb", "0"), null,
        item("Merge selected cells", () => this.workbook.current?.mergeCells(selection, "merge-all"), !selected),
        item("Unmerge selected cells", () => this.workbook.current?.cancelMerge(selection), !selected), action("Clear formatting", "clear-format"),
      ] },
      { title: "Data", items: [
        item("Sort selection A → Z", () => command("sort-asc"), !selected || selection.length !== 1),
        item("Sort selection Z → A", () => command("sort-desc"), !selected || selection.length !== 1), null,
        action("Create filter", "filter"), action("Remove filter", "clear-filter"),
      ] },
      { title: "Tools", items: [action("Find and replace", "replace"), null,
        item("Fit selected columns", () => this.fitMenuSelection("column"), !selected),
        item("Fit selected rows", () => this.fitMenuSelection("row"), !selected),
      ] },
      { title: "Help", items: [item("Sheetian help", () => {
        const modal = new Modal(this.app); modal.titleEl.setText("Sheetian help");
        for (const tip of ["Edits save automatically when committed. Press Enter to commit or Escape to cancel.", "Import Google Sheets by downloading Microsoft Excel (.xlsx), then choosing File → Import.", "Drag the formula bar's bottom grip to resize it. The chevron expands or collapses it.", "Double-click a row or column header boundary to fit its contents.", "To sync .sheet files, enable Sync all other types in Obsidian Sync on each device."]) modal.contentEl.createEl("p", { text: tip });
        modal.open();
      }), item("Sheetian source on GitHub", () => { this.contentEl.ownerDocument.defaultView!.open("https://github.com/haydencbarnes/sheetian", "_blank", "noopener,noreferrer"); })] },
    ];
  }

  private fitMenuSelection(type: "row" | "column"): void {
    const api = this.workbook.current;
    if (!api) return;
    const sheet = api.getSheet(), selection = api.getSelection() ?? [];
    const canvas = this.contentEl.ownerDocument.createElement("canvas").getContext("2d");
    if (!canvas) return;
    const sizes: Record<string, number> = {};
    const maximum = type === "row" ? sheet.data?.length ?? 0 : Math.max(0, ...sheet.data?.map(row => row.length) ?? []);
    for (const range of selection) for (let n = range[type][0]; n <= Math.min(range[type][1], maximum - 1); n++) {
      if ((type === "row" ? sheet.config?.rowhidden : sheet.config?.colhidden)?.[n] != null) continue;
      sizes[n] = type === "row" ? fitRowHeight(sheet, n, canvas) : fitColumnWidth(sheet, n, (text, style) => { canvas.font = cellFont(style); return canvas.measureText(text).width; });
    }
    if (type === "row") api.setRowHeight(sizes, { id: sheet.id }, true);
    else api.setColumnWidth(sizes, { id: sheet.id }, true);
  }

  private commitInput(): void {
    // A cell still being edited has not reached onChange yet.
    const input = this.contentEl.querySelector<HTMLElement>(".luckysheet-cell-input");
    const box = this.contentEl.querySelector<HTMLElement>(".luckysheet-input-box");
    const ownerWindow = this.contentEl.ownerDocument.defaultView!;
    if (input && box && Number(ownerWindow.getComputedStyle(box).zIndex) >= 0) {
      flushSync(() => input.dispatchEvent(new ownerWindow.KeyboardEvent("keydown", {
        key: "Enter", code: "Enter", keyCode: 13, which: 13, bubbles: true,
      })));
    }
    const active = this.contentEl.ownerDocument.activeElement;
    if (active instanceof HTMLElement && this.contentEl.contains(active)) {
      flushSync(() => active.blur());
    }
  }

  async onUnloadFile(file: TFile): Promise<void> {
    this.commitInput();
    await Promise.resolve();
    flushSync(() => {});
    await super.onUnloadFile(file);
    this.disposeWorkbook();
  }

  clear(): void {
    this.disposeWorkbook();
    this.source = "";
    this.sheets = [];
    this.edited = false;
  }

  private disposeWorkbook(): void {
    this.generation++;
    this.disposeMenu?.();
    this.disposeMenu = null;
    this.resizeObserver?.disconnect();
    this.resizeObserver = null;
    if (this.resizeFrame !== null) this.contentEl.ownerDocument.defaultView?.cancelAnimationFrame(this.resizeFrame);
    this.resizeFrame = null;
    this.root?.unmount();
    this.root = null;
    this.workbook = React.createRef<WorkbookInstance>();
  }

  async onClose(): Promise<void> {
    if (this.file) {
      this.commitInput();
      await Promise.resolve();
      flushSync(() => {});
      await this.save();
    }
    this.disposeWorkbook();
    this.contentEl.empty();
  }

  async exportCsv(): Promise<void> {
    this.commitInput();
    await Promise.resolve();
    flushSync(() => {});
    const sheet = this.workbook.current?.getSheet();
    if (!sheet || !this.file) return;
    const base = this.file.path.replace(/\.sheet$/i, "");
    let path = base + ".csv";
    for (let i = 1; this.app.vault.getAbstractFileByPath(path); i++) path = `${base}-${i}.csv`;
    await this.app.vault.create(path, sheetToCsv(sheet));
    new Notice(`Exported active sheet to ${path}`);
  }

  async calculateImportedFormulas(cells: { id: string; r: number; c: number }[]): Promise<void> {
    const generation = this.generation;
    const ownerWindow = this.contentEl.ownerDocument.defaultView!;
    for (let n = 0; n < 250; n++) {
      if (generation !== this.generation) throw new Error("Workbook was closed during import.");
      const sheets = this.workbook.current?.getAllSheets();
      if (sheets?.length === this.sheets.length && sheets.every(sheet => sheet.data?.length)) break;
      if (n === 249) throw new Error("Workbook did not finish loading.");
      await new Promise(resolve => ownerWindow.setTimeout(resolve, 20));
    }
    // Re-enter only uncached formulas. Unlike calculateFormula(), this public
    // API preserves their formula text and establishes dependency tracking.
    const originalId = this.workbook.current!.getSheet().id;
    for (const { id, r, c } of cells) {
      // The engine's formula setter uses the active sheet despite options.id.
      if (this.workbook.current!.getSheet().id !== id) {
        flushSync(() => this.workbook.current!.activateSheet({ id }));
      }
      const api = this.workbook.current!;
      const cell = api.getSheet().data?.[r]?.[c];
      if (cell?.f) flushSync(() => api.setCellValue(r, c, { f: cell.f, ct: cell.ct }));
    }
    flushSync(() => this.workbook.current!.activateSheet({ id: originalId }));
    await this.save();
  }
}

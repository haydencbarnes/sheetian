import { Notice, Scope, TextFileView, TFile, WorkspaceLeaf } from "obsidian";
import * as React from "react";
import { createRoot, Root } from "react-dom/client";
import { flushSync } from "react-dom";
import { Workbook } from "@fortune-sheet/react";
import type { WorkbookInstance } from "@fortune-sheet/react/dist/components/Workbook";
import type { Op, Sheet } from "@fortune-sheet/core";
import { parseSheets, serializeSheets, sheetToCsv } from "./sheet-data";
import { normalizeInput } from "./input";
import type { SpreadsheetSettings } from "./settings";

export const VIEW_TYPE_SPREADSHEET = "spreadsheet-view";

export class SpreadsheetView extends TextFileView {
  private source = "";
  private sheets: Sheet[] = [];
  private edited = false;
  private generation = 0;
  private root: Root | null = null;
  private resizeObserver: ResizeObserver | null = null;
  private resizeFrame: number | null = null;
  private workbook = React.createRef<WorkbookInstance>();

  constructor(leaf: WorkspaceLeaf, private settings: () => SpreadsheetSettings) {
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
    const container = this.contentEl.createDiv({ cls: "obsidian-spreadsheet" });
    const settings = this.settings();
    container.dataset.theme = settings.theme;
    const generation = this.generation;
    const ownerWindow = container.ownerDocument.defaultView!;
    this.root = createRoot(container);
    this.root.render(<Workbook
      ref={this.workbook} data={this.sheets} lang="en" currency={settings.currency}
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
}

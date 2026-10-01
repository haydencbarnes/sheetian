import { App, Modal, Notice, Plugin, TFile, TFolder, WorkspaceLeaf, normalizePath } from "obsidian";
import { SpreadsheetView, VIEW_TYPE_SPREADSHEET } from "./view";
import { DEFAULT_SETTINGS, SpreadsheetSettings, SpreadsheetSettingTab } from "./settings";
import { importXlsx } from "./xlsx-import";
import { serializeSheets } from "./sheet-data";

class ImportNotesModal extends Modal {
  constructor(app: App, private notes: string[]) { super(app); }
  onOpen(): void {
    this.titleEl.setText("Workbook imported with limitations");
    const list = this.contentEl.createEl("ul");
    for (const note of this.notes) list.createEl("li", { text: note });
    this.contentEl.createEl("button", { text: "Close" }).onclick = () => this.close();
  }
  onClose(): void { this.contentEl.empty(); }
}

export default class SpreadsheetPlugin extends Plugin {
  settings: SpreadsheetSettings = { ...DEFAULT_SETTINGS };

  async onload(): Promise<void> {
    this.settings = { ...DEFAULT_SETTINGS, ...await this.loadData() };
    this.registerView(VIEW_TYPE_SPREADSHEET, (leaf: WorkspaceLeaf) => new SpreadsheetView(leaf, () => this.settings, () => this.chooseImportFile()));
    this.registerExtensions(["sheet"], VIEW_TYPE_SPREADSHEET);
    this.addSettingTab(new SpreadsheetSettingTab(this.app, this));
    this.addRibbonIcon("table", "New spreadsheet", () => { void this.createSpreadsheet(); });
    this.addCommand({ id: "new-spreadsheet", name: "New spreadsheet", callback: () => { void this.createSpreadsheet(); } });
    this.addCommand({ id: "import-xlsx", name: "Import Excel workbook (.xlsx)", callback: () => this.chooseImportFile() });
    this.addCommand({ id: "export-csv", name: "Export active sheet to CSV", checkCallback: checking => {
      const view = this.app.workspace.getActiveViewOfType(SpreadsheetView);
      if (!view) return false;
      if (!checking) void view.exportCsv().catch(error => new Notice(`CSV export failed: ${error.message}`));
      return true;
    } });
    this.registerEvent(this.app.workspace.on("file-menu", (menu, file) => {
      const folder = file instanceof TFolder ? file.path : file.parent?.path;
      menu.addItem(item => item.setTitle("New spreadsheet").setIcon("table")
        .onClick(() => { void this.createSpreadsheet(folder); }));
      if (file instanceof TFile && file.extension.toLowerCase() === "xlsx") {
        menu.addItem(item => item.setTitle("Import into Sheetian").setIcon("file-input").onClick(() => {
          void this.app.vault.readBinary(file).then(bytes => this.importWorkbook(file.name, bytes, folder))
            .catch(error => this.importError(error));
        }));
      }
    }));
  }

  async saveSettings(): Promise<void> { await this.saveData(this.settings); }

  private importError(error: unknown): void {
    new Notice(`Could not import workbook: ${error instanceof Error ? error.message : String(error)}`, 10000);
  }

  private chooseImportFile(): void {
    const document = this.app.workspace.activeLeaf?.view.containerEl.ownerDocument ?? window.document;
    const input = document.createElement("input");
    input.type = "file";
    input.accept = ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    input.onchange = () => {
      const file = input.files?.[0];
      if (!file) return;
      if (!/\.xlsx$/i.test(file.name)) { this.importError(new Error("Choose a Microsoft Excel (.xlsx) workbook.")); return; }
      void file.arrayBuffer().then(bytes => this.importWorkbook(file.name, bytes)).catch(error => this.importError(error));
    };
    input.click();
  }

  async importWorkbook(name: string, bytes: ArrayBuffer | Uint8Array, folder = this.settings.folder): Promise<TFile> {
    const pending = new Notice("Importing workbook…", 0);
    try {
      const imported = await importXlsx(bytes);
      const prefix = await this.prepareFolder(folder);
      const base = name.replace(/\.xlsx$/i, "").replace(/[\\/:*?"<>|]/g, "-").replace(/^[. ]+|[. ]+$/g, "") || "Imported workbook";
      let path = `${prefix}${base}.sheet`;
      for (let n = 1; this.app.vault.getAbstractFileByPath(path); n++) path = `${prefix}${base}-${n}.sheet`;
      const file = await this.app.vault.create(path, serializeSheets(imported.sheets));
      const leaf = this.app.workspace.getLeaf(true);
      await leaf.openFile(file);
      if (leaf.view instanceof SpreadsheetView && imported.missingResults.length) {
        try { await leaf.view.calculateImportedFormulas(imported.missingResults); }
        catch { imported.warnings.push("Some formula results could not be calculated on import. Their formula text is saved; open the workbook to review them."); }
      }
      new Notice(`Imported ${imported.sheets.length} sheet tab(s) and ${imported.formulaCount} formula(s) into ${file.path}.`, 8000);
      if (imported.warnings.length) new ImportNotesModal(this.app, imported.warnings).open();
      return file;
    } finally { pending.hide(); }
  }

  private async prepareFolder(folder: string): Promise<string> {
    const normalized = normalizePath(folder);
    if (normalized && normalized !== "/") {
      let path = "";
      for (const part of normalized.split("/")) {
        if (part === ".." || part === ".") throw new Error("Choose a folder inside the vault.");
        path = path ? `${path}/${part}` : part;
        const existing = this.app.vault.getAbstractFileByPath(path);
        if (existing && !(existing instanceof TFolder)) throw new Error(`${path} is a file, not a folder.`);
        if (!existing) await this.app.vault.createFolder(path);
      }
    }
    return normalized && normalized !== "/" ? `${normalized}/` : "";
  }

  async createSpreadsheet(folder = this.settings.folder): Promise<void> {
    try {
      const prefix = await this.prepareFolder(folder);
      let path = `${prefix}Untitled.sheet`;
      for (let n = 1; this.app.vault.getAbstractFileByPath(path); n++) path = `${prefix}Untitled${n}.sheet`;
      const file = await this.app.vault.create(path, JSON.stringify([{ name: "Sheet1", id: "sheet-1", status: 1 }]));
      await this.app.workspace.getLeaf(true).openFile(file);
    } catch (error) {
      new Notice(`Could not create spreadsheet: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
}

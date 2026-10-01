import { Notice, Plugin, TFolder, WorkspaceLeaf, normalizePath } from "obsidian";
import { SpreadsheetView, VIEW_TYPE_SPREADSHEET } from "./view";
import { DEFAULT_SETTINGS, SpreadsheetSettings, SpreadsheetSettingTab } from "./settings";

export default class SpreadsheetPlugin extends Plugin {
  settings: SpreadsheetSettings = { ...DEFAULT_SETTINGS };

  async onload(): Promise<void> {
    this.settings = { ...DEFAULT_SETTINGS, ...await this.loadData() };
    this.registerView(VIEW_TYPE_SPREADSHEET, (leaf: WorkspaceLeaf) => new SpreadsheetView(leaf, () => this.settings));
    this.registerExtensions(["sheet"], VIEW_TYPE_SPREADSHEET);
    this.addSettingTab(new SpreadsheetSettingTab(this.app, this));
    this.addRibbonIcon("table", "New spreadsheet", () => { void this.createSpreadsheet(); });
    this.addCommand({ id: "new-spreadsheet", name: "New spreadsheet", callback: () => { void this.createSpreadsheet(); } });
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
    }));
  }

  async saveSettings(): Promise<void> { await this.saveData(this.settings); }

  async createSpreadsheet(folder = this.settings.folder): Promise<void> {
    try {
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
      const prefix = normalized && normalized !== "/" ? `${normalized}/` : "";
      let path = `${prefix}Untitled.sheet`;
      for (let n = 1; this.app.vault.getAbstractFileByPath(path); n++) path = `${prefix}Untitled${n}.sheet`;
      const file = await this.app.vault.create(path, JSON.stringify([{ name: "Sheet1", id: "sheet-1", status: 1 }]));
      await this.app.workspace.getLeaf(true).openFile(file);
    } catch (error) {
      new Notice(`Could not create spreadsheet: ${error instanceof Error ? error.message : String(error)}`);
    }
  }
}

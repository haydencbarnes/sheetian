import { App, PluginSettingTab, Setting } from "obsidian";
import type SpreadsheetPlugin from "./main";

export interface SpreadsheetSettings {
  theme: "auto" | "light" | "dark";
  currency: string;
  folder: string;
}
export const DEFAULT_SETTINGS: SpreadsheetSettings = { theme: "auto", currency: "$", folder: "" };

export class SpreadsheetSettingTab extends PluginSettingTab {
  constructor(app: App, private plugin: SpreadsheetPlugin) { super(app, plugin); }
  display(): void {
    this.containerEl.empty();
    new Setting(this.containerEl).setName("Spreadsheet theme")
      .setDesc("Applied when a spreadsheet is opened. Dark mode adjusts display colors; stored cell colors stay intact.")
      .addDropdown(dropdown => dropdown.addOptions({ auto: "Follow Obsidian", light: "Light", dark: "Dark" })
        .setValue(this.plugin.settings.theme).onChange(async value => {
          this.plugin.settings.theme = value as SpreadsheetSettings["theme"];
          await this.plugin.saveSettings();
        }));
    new Setting(this.containerEl).setName("Currency symbol").setDesc("Used by Format as currency in newly opened spreadsheets.")
      .addText(text => text.setValue(this.plugin.settings.currency).onChange(async value => {
        if (!value.trim() || /["\\\r\n]/.test(value) || value.length > 8) return;
        this.plugin.settings.currency = value.trim();
        await this.plugin.saveSettings();
      }));
    new Setting(this.containerEl).setName("Default folder").setDesc("Vault folder for the ribbon and New spreadsheet command. Leave empty for the vault root.")
      .addText(text => text.setValue(this.plugin.settings.folder).onChange(async value => {
        this.plugin.settings.folder = value.trim();
        await this.plugin.saveSettings();
      }));
  }
}

import { Menu, Notice } from "obsidian";

export interface SheetMenuItem {
  title: string;
  action: () => void | Promise<void>;
  disabled?: boolean;
  checked?: boolean;
}
export interface SheetMenu { title: string; items: (SheetMenuItem | null)[]; }

export function mountMenuBar(parent: HTMLElement, titles: string[], build: () => SheetMenu[], prepare: () => void): () => void {
  const bar = parent.createDiv({ cls: "sheetian-menu-bar" });
  bar.setAttribute("role", "menubar"); bar.setAttribute("aria-label", "Spreadsheet menus");
  let activeMenu: Menu | null = null;
  let activeButton: HTMLButtonElement | null = null;
  const buttons: HTMLButtonElement[] = [];
  const open = (button: HTMLButtonElement, title: string) => {
    if (activeButton === button) { activeMenu?.hide(); return; }
    activeMenu?.hide();
    prepare();
    const definition = build().find(menu => menu.title === title);
    if (!definition) return;
    const menu = new Menu().setUseNativeMenu(false);
    for (const entry of definition.items) {
      if (!entry) { menu.addSeparator(); continue; }
      menu.addItem(item => {
        item.setTitle(entry.title).setDisabled(!!entry.disabled);
        if (entry.checked != null) item.setChecked(entry.checked);
        item.onClick(() => {
          menu.hide();
          Promise.resolve().then(entry.action).catch(error => new Notice(`Could not complete action: ${error instanceof Error ? error.message : String(error)}`));
        });
      });
    }
    activeMenu = menu; activeButton = button;
    button.setAttribute("aria-expanded", "true");
    menu.onHide(() => {
      button.setAttribute("aria-expanded", "false");
      if (activeMenu === menu) { activeMenu = null; activeButton = null; }
    });
    const rect = button.getBoundingClientRect();
    menu.showAtPosition({ x: rect.left, y: rect.bottom }, button.ownerDocument);
  };
  for (const title of titles) {
    const button = bar.createEl("button", { text: title, cls: "sheetian-menu-button" });
    button.type = "button"; button.setAttribute("role", "menuitem");
    button.setAttribute("aria-haspopup", "menu"); button.setAttribute("aria-expanded", "false");
    button.addEventListener("mousedown", event => { event.preventDefault(); event.stopPropagation(); });
    button.addEventListener("click", () => open(button, title));
    button.addEventListener("mouseenter", () => { if (activeMenu && activeButton !== button) open(button, title); });
    button.addEventListener("keydown", event => {
      if (["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key)) {
        event.preventDefault(); event.stopPropagation();
        const index = buttons.indexOf(button);
        const next = event.key === "Home" ? 0 : event.key === "End" ? buttons.length - 1 : (index + (event.key === "ArrowRight" ? 1 : -1) + buttons.length) % buttons.length;
        const wasOpen = !!activeMenu; buttons[next].focus();
        if (wasOpen) open(buttons[next], titles[next]);
      } else if (event.key === "ArrowDown") { event.preventDefault(); event.stopPropagation(); open(button, title); }
      else if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); activeMenu?.hide(); button.focus(); }
    });
    buttons.push(button);
  }
  // Obsidian already provides Escape/outside-click dismissal and arrow-key
  // navigation inside its menu. Close menus when this view switches or unloads.
  return () => { activeMenu?.hide(); activeMenu = null; activeButton = null; };
}

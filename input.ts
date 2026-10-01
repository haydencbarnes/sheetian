import type { Cell } from "@fortune-sheet/core";
import { update, is_date } from "@fortune-sheet/core";

export function normalizeInput(value: unknown, previous: Cell | null | undefined, currency: string): Cell | undefined {
  if (typeof value !== "string" || previous?.ct?.fa === "@") return;
  const text = value.trim();
  const format = previous?.ct?.fa;
  if (previous?.ct?.t === "d" || (format && is_date(format))) {
    // ISO dates and US slash dates; no locale-dependent Date.parse.
    const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text);
    const slash = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
    const parts = iso ? [+iso[1], +iso[2], +iso[3]] : slash ? [+slash[3], +slash[1], +slash[2]] : null;
    if (parts) {
      const [year, month, day] = parts;
      const date = new Date(Date.UTC(year, month - 1, day));
      if (year >= 1900 && date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day) {
        const v = (date.getTime() - Date.UTC(1899, 11, 31)) / 86400000 + (year > 1900 || month > 2 ? 1 : 0);
        const fa = format || "yyyy-MM-dd";
        return { v, m: update(fa, v), ct: { fa, t: "d" } };
      }
    }
  }
  const symbols = [...new Set([currency, "$", "€", "£", "¥"])];
  const symbol = symbols.find(symbol => symbol && text.startsWith(symbol));
  if (!symbol) return;
  const number = text.slice(symbol.length).trim();
  if (!/^[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(number)) return;
  const v = Number(number.replace(/,/g, ""));
  if (!Number.isFinite(v)) return;
  const fa = format && format !== "General" ? format : `"${symbol}"#,##0.00`;
  return { v, m: update(fa, v), ct: { fa, t: "n" } };
}

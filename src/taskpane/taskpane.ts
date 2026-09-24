// Task pane: converts the selected cells in place.
import { bijoyToUnicode, isUnicode, unicodeToBijoy } from "../../../src/core";

type Direction = "toAnsi" | "toUnicode";

// Rows per Excel.run round trip; keeps each sync well under Office's payload limits.
const BLOCK_ROWS = 500;
const SETTINGS_KEY = "bangla-converter-settings";

interface Settings {
  changeFont: boolean;
  ansiFont: string;
  unicodeFont: string;
  onlyBijoyFont: boolean;
}

const defaults: Settings = {
  changeFont: true,
  ansiFont: "SutonnyMJ",
  unicodeFont: "Nirmala UI",
  onlyBijoyFont: true,
};

interface Result {
  converted: number;
  skippedFormulas: number;
  skippedFont: number;
}

// Bijoy fonts are conventionally named "...MJ" (SutonnyMJ, SulekhaMJ, Sutonny OMJ, ...).
const isBijoyFont = (name: string | null | undefined): boolean => !!name && /MJ\s*$/i.test(name);

function $(id: string): HTMLInputElement {
  return document.getElementById(id) as HTMLInputElement;
}

function loadSettings(): Settings {
  try {
    const saved = JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}");
    return { ...defaults, ...saved };
  } catch {
    return { ...defaults };
  }
}

function readSettings(): Settings {
  const settings: Settings = {
    changeFont: $("change-font").checked,
    ansiFont: $("ansi-font").value.trim() || defaults.ansiFont,
    unicodeFont: $("unicode-font").value.trim() || defaults.unicodeFont,
    onlyBijoyFont: $("only-bijoy-font").checked,
  };
  try {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  } catch {
    // Storage can be unavailable; settings just won't persist.
  }
  return settings;
}

function showStatus(message: string, kind: "info" | "ok" | "error" = "info"): void {
  const status = document.getElementById("status")!;
  status.textContent = message;
  status.className = `status ${kind}`;
}

async function convertSelection(direction: Direction): Promise<Result> {
  const settings = readSettings();
  const result: Result = { converted: 0, skippedFormulas: 0, skippedFont: 0 };

  await Excel.run(async (context) => {
    const selection = context.workbook.getSelectedRange();
    const sheet = selection.worksheet;
    const used = sheet.getUsedRangeOrNullObject(true);
    await context.sync();
    if (used.isNullObject) return;

    // Whole-column selections would otherwise mean a million rows.
    const target = selection.getIntersectionOrNullObject(used);
    target.load(["rowIndex", "columnIndex", "rowCount", "columnCount"]);
    await context.sync();
    if (target.isNullObject) return;

    for (let start = 0; start < target.rowCount; start += BLOCK_ROWS) {
      const rows = Math.min(BLOCK_ROWS, target.rowCount - start);
      const block = sheet.getRangeByIndexes(target.rowIndex + start, target.columnIndex, rows, target.columnCount);
      block.load(["values", "formulas"]);
      const props = block.getCellProperties({ format: { font: { name: true } } });
      await context.sync();

      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < target.columnCount; c++) {
          const value = block.values[r][c];
          if (typeof value !== "string" || value === "") continue;

          const formula = block.formulas[r][c];
          if (typeof formula === "string" && formula.startsWith("=")) {
            result.skippedFormulas++;
            continue;
          }

          let converted: string;
          if (direction === "toAnsi") {
            if (!isUnicode(value)) continue;
            converted = unicodeToBijoy(value);
          } else {
            if (isUnicode(value)) continue;
            if (settings.onlyBijoyFont && !isBijoyFont(props.value[r][c].format?.font?.name)) {
              result.skippedFont++;
              continue;
            }
            converted = bijoyToUnicode(value);
          }
          if (converted === value) continue;

          const cell = block.getCell(r, c);
          cell.values = [[converted]];
          if (settings.changeFont) {
            cell.format.font.name = direction === "toAnsi" ? settings.ansiFont : settings.unicodeFont;
          }
          result.converted++;
        }
      }
      await context.sync();
    }
  });

  return result;
}

function describe(result: Result): string {
  const parts = [`Converted ${result.converted} cell${result.converted === 1 ? "" : "s"}.`];
  if (result.skippedFormulas) {
    parts.push(`Skipped ${result.skippedFormulas} formula cell(s); use =BANGLA.TOANSI() / =BANGLA.TOUNICODE() for those.`);
  }
  if (result.skippedFont) {
    parts.push(`Skipped ${result.skippedFont} cell(s) not in a Bijoy (…MJ) font. Untick "Only convert cells in a Bijoy font" to include them.`);
  }
  return parts.join(" ");
}

async function run(direction: Direction): Promise<void> {
  const buttons = document.querySelectorAll<HTMLButtonElement>("button");
  buttons.forEach((b) => (b.disabled = true));
  showStatus("Converting…");
  try {
    const result = await convertSelection(direction);
    showStatus(describe(result), result.converted ? "ok" : "info");
  } catch (error) {
    showStatus(`Error: ${error instanceof Error ? error.message : String(error)}`, "error");
  } finally {
    buttons.forEach((b) => (b.disabled = false));
  }
}

Office.onReady(() => {
  const settings = loadSettings();
  $("change-font").checked = settings.changeFont;
  $("ansi-font").value = settings.ansiFont;
  $("unicode-font").value = settings.unicodeFont;
  $("only-bijoy-font").checked = settings.onlyBijoyFont;

  document.getElementById("to-ansi")!.addEventListener("click", () => run("toAnsi"));
  document.getElementById("to-unicode")!.addEventListener("click", () => run("toUnicode"));
});

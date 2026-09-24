// Custom functions exposed in Excel as =BANGLA.TOANSI(...) and =BANGLA.TOUNICODE(...).
// Metadata lives in functions.json; keep the ids there in sync with the associate() calls below.
import { bijoyToUnicode, unicodeToBijoy } from "../../../src/core";

type Cell = string | number | boolean | null;

// Only text is converted; numbers, booleans and empty cells pass through unchanged.
function mapMatrix(input: Cell[][], convert: (text: string) => string): Cell[][] {
  return input.map((row) =>
    row.map((cell) => {
      if (cell === null || cell === undefined) return "";
      return typeof cell === "string" ? convert(cell) : cell;
    }),
  );
}

/** Unicode Bangla -> Bijoy (ANSI). Accepts a cell or a range; a range spills. */
export function toAnsi(text: Cell[][]): Cell[][] {
  return mapMatrix(text, unicodeToBijoy);
}

/** Bijoy (ANSI) -> Unicode Bangla. Text that already contains Unicode Bangla is left as is. */
export function toUnicode(text: Cell[][]): Cell[][] {
  return mapMatrix(text, bijoyToUnicode);
}

CustomFunctions.associate("TOANSI", toAnsi);
CustomFunctions.associate("TOUNICODE", toUnicode);

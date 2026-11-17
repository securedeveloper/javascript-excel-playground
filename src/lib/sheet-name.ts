// Sheet-name checks for the Write tab, driven by the library's validateSheetName.
import { validateSheetName } from "javascript-excel";

/** Excel's limit, in code points (an emoji counts once), as the library counts it. */
export const MAX_SHEET_NAME = 31;

/** Length in code points, not UTF-16 units. */
export const codePointLength = (name: string): number => [...name].length;

/** A readable reason the name is invalid, or null when the library accepts it. */
export function sheetNameError(name: string): string | null {
  switch (validateSheetName(name)) {
    case null:
      return null;
    case "empty":
      return "A sheet name can't be empty.";
    case "too-long":
      return `A sheet name has at most ${MAX_SHEET_NAME} characters; this one has ${codePointLength(name)}.`;
    case "forbidden-char":
      return "A sheet name can't contain \\ / ? * [ ] or :.";
    case "apostrophe":
      return "A sheet name can't start or end with an apostrophe.";
    case "reserved":
      return "“History” is reserved by Excel.";
  }
}

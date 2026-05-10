import { describe, expect, it } from "vitest";
import { DEFAULT_LOCALE, SUPPORTED_LOCALES, translateForLocale, translations } from "../src/i18n";

describe("i18n", () => {
  it("ships English, German, French, and Italian locale maps with matching keys", () => {
    expect(SUPPORTED_LOCALES).toEqual(["en", "de", "fr", "it"]);

    const baseKeys = Object.keys(translations[DEFAULT_LOCALE]).sort();
    for (const locale of SUPPORTED_LOCALES) {
      expect(Object.keys(translations[locale]).sort()).toEqual(baseKeys);
    }
  });

  it("formats translated strings with variables", () => {
    expect(translateForLocale("fr", "status.exported", { fileName: "rapport.docx" })).toContain("rapport.docx");
    expect(translateForLocale("it", "stats.input", { chars: "42", words: "7" })).toContain("42");
  });
});

import de from "./locales/de.json";
import en from "./locales/en.json";
import fr from "./locales/fr.json";
import it from "./locales/it.json";

export const translations = {
  en,
  de,
  fr,
  it,
};

export type LocaleCode = keyof typeof translations;
export type TranslationVars = Record<string, number | string>;
export type Translate = (key: string, vars?: TranslationVars) => string;

export const DEFAULT_LOCALE: LocaleCode = "en";
export const SUPPORTED_LOCALES = Object.keys(translations) as LocaleCode[];

const LOCALE_STORAGE_KEY = "markdown-to-word.locale";
const reportedMissingKeys = new Set<string>();

let currentLocale: LocaleCode = DEFAULT_LOCALE;

export function initializeI18n(): LocaleCode {
  currentLocale = getInitialLocale();
  return currentLocale;
}

export function getCurrentLocale(): LocaleCode {
  return currentLocale;
}

export function setLocale(locale: string, persist: boolean): LocaleCode {
  currentLocale = resolveLocale(locale);
  if (persist && typeof window !== "undefined") {
    try {
      window.localStorage.setItem(LOCALE_STORAGE_KEY, currentLocale);
    } catch (_error) {
      // Storage may be disabled on locked-down systems.
    }
  }
  return currentLocale;
}

export function resolveLocale(locale: string | undefined | null): LocaleCode {
  const normalized = String(locale || "").toLowerCase().split("-")[0];
  return SUPPORTED_LOCALES.includes(normalized as LocaleCode) ? normalized as LocaleCode : DEFAULT_LOCALE;
}

export function t(key: string, vars: TranslationVars = {}): string {
  return translateForLocale(currentLocale, key, vars);
}

export function plural(baseKey: string, count: number, vars: TranslationVars = {}): string {
  const category = pluralCategory(currentLocale, count);
  const localized = `${baseKey}.${category}`;
  if (hasTranslationKey(localized)) {
    return t(localized, { count, ...vars });
  }
  const otherKey = `${baseKey}.other`;
  if (hasTranslationKey(otherKey)) {
    return t(otherKey, { count, ...vars });
  }
  return t(baseKey, { count, ...vars });
}

const pluralRulesCache = new Map<LocaleCode, Intl.PluralRules>();

function pluralCategory(locale: LocaleCode, count: number): Intl.LDMLPluralRule {
  let rules = pluralRulesCache.get(locale);
  if (!rules) {
    try {
      rules = new Intl.PluralRules(locale);
      pluralRulesCache.set(locale, rules);
    } catch (_error) {
      return count === 1 ? "one" : "other";
    }
  }
  return rules.select(count);
}

export function translateForLocale(locale: LocaleCode, key: string, vars: TranslationVars = {}): string {
  const map = translations[locale] as Record<string, string>;
  const fallback = translations[DEFAULT_LOCALE] as Record<string, string>;
  const template = Object.prototype.hasOwnProperty.call(map, key)
    ? map[key]
    : fallback[key];

  if (typeof template !== "string") {
    if (!reportedMissingKeys.has(key) && typeof window !== "undefined" && window.console?.warn) {
      reportedMissingKeys.add(key);
      window.console.warn("[markdown-to-word i18n] missing key:", key);
    }
    return `[${key}]`;
  }

  return interpolate(template, vars);
}

export function hasTranslationKey(key: string): boolean {
  return Object.prototype.hasOwnProperty.call(translations[DEFAULT_LOCALE], key);
}

export function applyI18nToDom(): void {
  if (typeof document === "undefined") {
    return;
  }

  document.documentElement.lang = currentLocale;
  document.title = t("app.title");

  document.querySelectorAll<HTMLElement>("[data-i18n]").forEach((node) => {
    node.textContent = t(node.getAttribute("data-i18n") || "");
  });
  document.querySelectorAll<HTMLElement>("[data-i18n-html]").forEach((node) => {
    node.innerHTML = sanitizeI18nHtml(t(node.getAttribute("data-i18n-html") || ""));
  });
  document.querySelectorAll<HTMLElement>("[data-i18n-aria-label]").forEach((node) => {
    node.setAttribute("aria-label", t(node.getAttribute("data-i18n-aria-label") || ""));
  });
  document.querySelectorAll<HTMLElement>("[data-i18n-placeholder]").forEach((node) => {
    node.setAttribute("placeholder", t(node.getAttribute("data-i18n-placeholder") || ""));
  });
  document.querySelectorAll<HTMLElement>("[data-i18n-title]").forEach((node) => {
    node.setAttribute("title", t(node.getAttribute("data-i18n-title") || ""));
  });
}

export function formatList(values: string[]): string {
  if (!values.length) {
    return "";
  }
  try {
    return new Intl.ListFormat(currentLocale, { style: "short", type: "conjunction" }).format(values);
  } catch (_error) {
    return values.join(", ");
  }
}

function getInitialLocale(): LocaleCode {
  if (typeof window === "undefined") {
    return DEFAULT_LOCALE;
  }

  try {
    const params = new URLSearchParams(window.location.search || "");
    const fromUrl = params.get("lang") || params.get("locale");
    if (fromUrl) {
      return resolveLocale(fromUrl);
    }
  } catch (_error) {
    // Ignore URL parse errors.
  }

  try {
    const saved = window.localStorage.getItem(LOCALE_STORAGE_KEY);
    if (saved) {
      return resolveLocale(saved);
    }
  } catch (_error) {
    // Storage may be unavailable.
  }

  return resolveLocale(window.navigator.language || DEFAULT_LOCALE);
}

function interpolate(template: string, vars: TranslationVars): string {
  return template.replace(/\{([a-zA-Z0-9_]+)\}/g, (match, name) => {
    const value = vars[name];
    return typeof value === "undefined" ? match : String(value);
  });
}

function sanitizeI18nHtml(html: string): string {
  const template = document.createElement("template");
  template.innerHTML = html;
  const allowedTags = new Set(["BR", "CODE", "EM", "STRONG"]);

  const walk = (node: Node): void => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === Node.ELEMENT_NODE) {
        const element = child as HTMLElement;
        if (!allowedTags.has(element.tagName)) {
          element.replaceWith(document.createTextNode(element.textContent || ""));
          continue;
        }
        for (const attribute of Array.from(element.attributes)) {
          element.removeAttribute(attribute.name);
        }
      }
      walk(child);
    }
  };

  walk(template.content);
  return template.innerHTML;
}

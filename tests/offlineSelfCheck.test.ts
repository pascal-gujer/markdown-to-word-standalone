import { afterEach, describe, expect, it } from "vitest";
import { isExternalRuntimeReference, runOfflineSelfCheck } from "../src/security/offlineSelfCheck";

const originalDocument = globalThis.document;

afterEach(() => {
  Object.defineProperty(globalThis, "document", {
    configurable: true,
    value: originalDocument,
  });
});

describe("offline self-check", () => {
  it("does not treat ordinary outbound links as runtime asset references", () => {
    expect(isExternalRuntimeReference(element("a", {
      href: "https://github.com/sponsors/pascal-gujer",
    }))).toBe(false);
  });

  it("flags external resources that the browser may load automatically", () => {
    expect(isExternalRuntimeReference(element("img", { src: "https://example.invalid/logo.png" }))).toBe(true);
    expect(isExternalRuntimeReference(element("source", { srcset: "//example.invalid/hero.png 2x" }))).toBe(true);
    expect(isExternalRuntimeReference(element("link", { href: "https://example.invalid/site.css", rel: "stylesheet" }))).toBe(true);
    expect(isExternalRuntimeReference(element("object", { data: "https://example.invalid/file.swf" }))).toBe(true);
  });

  it("passes when the only external URL is a normal anchor", () => {
    mockDocument({
      "[src], [srcset], [poster], link[href], object[data]": [
        element("a", { href: "https://github.com/sponsors/pascal-gujer" }),
      ],
      "script[src]": [],
      'link[rel="stylesheet"][href]': [],
    });

    expect(runOfflineSelfCheck()).toEqual({
      ok: true,
      messages: [{ key: "offline.ok" }],
    });
  });

  it("reports external runtime resources", () => {
    mockDocument({
      "[src], [srcset], [poster], link[href], object[data]": [
        element("a", { href: "https://github.com/sponsors/pascal-gujer" }),
        element("img", { src: "https://example.invalid/logo.png" }),
        element("img", { src: "data:image/png;base64,AAAA" }),
      ],
      "script[src]": [],
      'link[rel="stylesheet"][href]': [],
    });

    expect(runOfflineSelfCheck()).toEqual({
      ok: false,
      messages: [{ key: "offline.external_assets", vars: { count: 1 } }],
    });
  });
});

function element(tagName: string, attributes: Record<string, string>): Element {
  return {
    getAttribute(name: string) {
      return attributes[name] ?? null;
    },
    tagName: tagName.toUpperCase(),
  } as Element;
}

function mockDocument(nodesBySelector: Record<string, Element[]>): void {
  Object.defineProperty(globalThis, "document", {
    configurable: true,
    value: {
      querySelectorAll(selector: string) {
        return nodesBySelector[selector] ?? [];
      },
    },
  });
}

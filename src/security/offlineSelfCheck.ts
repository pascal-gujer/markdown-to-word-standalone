export type OfflineSelfCheck = {
  ok: boolean;
  messages: string[];
};

const remoteUrlPattern = /^https?:\/\//i;

export function installNetworkGuard(): void {
  guardFetch();
  guardXmlHttpRequest();
}

export function runOfflineSelfCheck(): OfflineSelfCheck {
  const messages: string[] = [];
  const externalNodes = Array.from(document.querySelectorAll<HTMLElement>("[src], [href]"))
    .filter((node) => {
      const value = node.getAttribute("src") || node.getAttribute("href") || "";
      return remoteUrlPattern.test(value) || value.startsWith("//");
    });

  if (externalNodes.length) {
    messages.push(`Found ${externalNodes.length} external asset reference(s) in the page.`);
  }

  const externalScripts = document.querySelectorAll("script[src]").length;
  const externalStyles = document.querySelectorAll('link[rel="stylesheet"][href]').length;
  if (externalScripts || externalStyles) {
    messages.push("The final offline build should inline scripts and styles.");
  }

  if (!messages.length) {
    messages.push("Offline self-check passed: no remote asset references are present in this page.");
  }

  return { ok: messages.length === 1 && messages[0].startsWith("Offline self-check passed"), messages };
}

function guardFetch(): void {
  if (!("fetch" in window)) {
    return;
  }

  const nativeFetch = window.fetch.bind(window);
  window.fetch = ((input: RequestInfo | URL, init?: RequestInit) => {
    const url = typeof input === "string" || input instanceof URL ? String(input) : input.url;
    if (remoteUrlPattern.test(url)) {
      return Promise.reject(new Error("Remote network access is disabled in Markdown to Word Offline."));
    }
    return nativeFetch(input, init);
  }) as typeof window.fetch;
}

function guardXmlHttpRequest(): void {
  const nativeOpen = window.XMLHttpRequest?.prototype.open;
  if (!nativeOpen) {
    return;
  }

  window.XMLHttpRequest.prototype.open = function guardedOpen(
    method: string,
    url: string | URL,
    async = true,
    username?: string | null,
    password?: string | null,
  ): void {
    if (remoteUrlPattern.test(String(url))) {
      throw new Error("Remote network access is disabled in Markdown to Word Offline.");
    }
    Reflect.apply(nativeOpen, this, [method, url, async, username, password]);
  };
}

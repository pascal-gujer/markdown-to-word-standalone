export type OfflineSelfCheck = {
  ok: boolean;
  messages: OfflineSelfCheckMessage[];
};

export type OfflineSelfCheckMessage = {
  key: string;
  vars?: Record<string, number | string>;
};

const remoteUrlPattern = /^https?:\/\//i;

export function installNetworkGuard(remoteMessage: () => string): void {
  guardFetch(remoteMessage);
  guardXmlHttpRequest(remoteMessage);
}

export function runOfflineSelfCheck(): OfflineSelfCheck {
  const messages: OfflineSelfCheckMessage[] = [];
  const externalNodes = Array.from(document.querySelectorAll<HTMLElement>("[src], [href]"))
    .filter((node) => {
      const value = node.getAttribute("src") || node.getAttribute("href") || "";
      return remoteUrlPattern.test(value) || value.startsWith("//");
    });

  if (externalNodes.length) {
    messages.push({ key: "offline.external_assets", vars: { count: externalNodes.length } });
  }

  const externalScripts = document.querySelectorAll("script[src]").length;
  const externalStyles = document.querySelectorAll('link[rel="stylesheet"][href]').length;
  if (externalScripts || externalStyles) {
    messages.push({ key: "offline.inline_warning" });
  }

  if (!messages.length) {
    messages.push({ key: "offline.ok" });
  }

  return { ok: messages.length === 1 && messages[0].key === "offline.ok", messages };
}

function guardFetch(remoteMessage: () => string): void {
  if (!("fetch" in window)) {
    return;
  }

  const nativeFetch = window.fetch.bind(window);
  window.fetch = ((input: RequestInfo | URL, init?: RequestInit) => {
    const url = typeof input === "string" || input instanceof URL ? String(input) : input.url;
    if (remoteUrlPattern.test(url)) {
      return Promise.reject(new Error(remoteMessage()));
    }
    return nativeFetch(input, init);
  }) as typeof window.fetch;
}

function guardXmlHttpRequest(remoteMessage: () => string): void {
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
      throw new Error(remoteMessage());
    }
    Reflect.apply(nativeOpen, this, [method, url, async, username, password]);
  };
}

import { resolveBundleImage, type MarkdownImageBundle } from "../markdown/imageBundle";

export function renderPreview(
  container: HTMLElement,
  html: string,
  isEmpty: boolean,
  emptyText: string,
  imageBundle?: MarkdownImageBundle | null,
): void {
  container.classList.toggle("empty-state", isEmpty);
  if (isEmpty) {
    container.textContent = emptyText;
    return;
  }

  const template = document.createElement("template");
  template.innerHTML = html;

  for (const link of Array.from(template.content.querySelectorAll<HTMLAnchorElement>("a[href]"))) {
    const href = link.getAttribute("href") || "";
    link.title = href;
    link.rel = "noreferrer noopener";
    link.addEventListener("click", (event) => {
      event.preventDefault();
    });
  }

  for (const image of Array.from(template.content.querySelectorAll<HTMLImageElement>("img[src]"))) {
    const src = image.getAttribute("src") || "";
    const asset = resolveBundleImage(imageBundle, src);
    if (asset?.previewUrl) {
      image.setAttribute("src", asset.previewUrl);
      image.loading = "lazy";
      image.decoding = "async";
      continue;
    }

    const alt = image.getAttribute("alt") || "image";
    image.replaceWith(document.createTextNode(`[${alt}]`));
  }

  container.replaceChildren(template.content);
}

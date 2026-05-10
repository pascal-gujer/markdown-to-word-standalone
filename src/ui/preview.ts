export function renderPreview(container: HTMLElement, html: string, isEmpty: boolean): void {
  container.classList.toggle("empty-state", isEmpty);
  container.innerHTML = isEmpty ? "Preview appears here." : html;

  for (const link of Array.from(container.querySelectorAll<HTMLAnchorElement>("a[href]"))) {
    const href = link.href;
    link.title = href;
    link.rel = "noreferrer noopener";
    link.addEventListener("click", (event) => {
      event.preventDefault();
    });
  }
}

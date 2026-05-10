export function renderPreview(container: HTMLElement, html: string, isEmpty: boolean, emptyText: string): void {
  container.classList.toggle("empty-state", isEmpty);
  if (isEmpty) {
    container.textContent = emptyText;
  } else {
    container.innerHTML = html;
  }

  for (const link of Array.from(container.querySelectorAll<HTMLAnchorElement>("a[href]"))) {
    const href = link.href;
    link.title = href;
    link.rel = "noreferrer noopener";
    link.addEventListener("click", (event) => {
      event.preventDefault();
    });
  }
}

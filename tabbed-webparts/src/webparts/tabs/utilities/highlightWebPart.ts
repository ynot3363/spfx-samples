import styles from "./highlightWebPart.module.scss";

/**
 * Utility function to add a class around the selected element to highlight it
 * with a yellow border and remove it from any previously added element.
 */
export function highlightWebPart(targetElement: string): void {
  const highlightElement = document.createElement("div");
  const elementToHighlight = document.getElementById(`${targetElement}`);

  highlightElement.className = styles.highlightWebPart;

  removeWebPartHighlight();

  if (elementToHighlight) {
    setTimeout(() => {
      elementToHighlight.scrollIntoView({
        behavior: "smooth",
        block: "center",
        inline: "nearest",
      });
    }, 100);

    elementToHighlight.append(highlightElement);
  }
}

export function removeWebPartHighlight(): void {
  const previouslyHighlightedElement = document.querySelector(
    `.${styles.highlightWebPart}`
  );

  if (previouslyHighlightedElement) {
    previouslyHighlightedElement.remove();
  }
}

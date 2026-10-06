// Invariant for every implementation: READ ONLY — a BibSource must never write to the user's .bib file.

import { MAX_BIB_BYTES, clipDisplayText, safeDisplayText } from "./types";

export interface BibSource {
  /** Human-readable label of where the data came from (shown in the UI). */
  readonly label: string;

  read(): Promise<string>;
}

/** Grace period for a `change` event that trails the window `focus` event. */
const FOCUS_SETTLE_MS = 400;

/** Absolute ceiling on how long the picker may keep the pane busy: `change`, `cancel` and the focus fallback are all best-effort, and a host that fires none would leave every control disabled with no recovery. */
const PICKER_TIMEOUT_MS = 180_000;

/** Widest filename the status line can show without pushing the count off. */
const MAX_LABEL_LENGTH = 160;

function safeDisplayName(name: string): string {
  return clipDisplayText(safeDisplayText(name), MAX_LABEL_LENGTH);
}

export class FilePickerBibSource implements BibSource {
  private file: File;

  constructor(file: File) {
    this.file = file;
  }

  get label(): string {
    return safeDisplayName(this.file.name);
  }

  async read(): Promise<string> {
    // Cheap pre-check against File.size; citationEngine.load() re-checks the decoded text.
    if (this.file.size > MAX_BIB_BYTES) {
      throw new Error(
        `"${safeDisplayName(this.file.name)}" is ${(this.file.size / (1024 * 1024)).toFixed(1)} MB, over the ${
          MAX_BIB_BYTES / (1024 * 1024)
        } MB safety limit for a .bib file. Choose a smaller file.`,
      );
    }
    const bytes = await this.file.arrayBuffer();
    try {
      // File.text() would replace invalid UTF-8 with U+FFFD, silently mutating keys.
      const text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
      if (text.includes("\0")) {
        throw new Error(
          "The selected file contains NUL bytes and is not a valid text bibliography.",
        );
      }
      return text;
    } catch (error) {
      if (error instanceof TypeError) {
        throw new Error(
          "The selected file is not valid UTF-8. Save the .bib as UTF-8 and try again.",
        );
      }
      throw error;
    }
  }

  /** Opens the OS file picker, resolving null if cancelled or if no host event arrived within `PICKER_TIMEOUT_MS`. */
  static pick(): Promise<FilePickerBibSource | null> {
    return new Promise((resolve) => {
      const input = document.createElement("input");
      input.type = "file";
      input.accept = ".bib,.bibtex,text/plain";

      let settled = false;
      const timers: number[] = [];

      const finish = (value: FilePickerBibSource | null): void => {
        if (settled) return;
        settled = true;
        for (const timer of timers) window.clearTimeout(timer);
        window.removeEventListener("focus", onFocus);
        input.remove();
        resolve(value);
      };
      const finishFromInput = (): void =>
        finish(
          input.files?.[0] ? new FilePickerBibSource(input.files[0]) : null,
        );
      // Declared before any listener is registered, so `finish` can never reach it in the temporal dead zone.
      const onFocus = (): number =>
        timers.push(window.setTimeout(finishFromInput, FOCUS_SETTLE_MS));

      input.addEventListener("change", finishFromInput, { once: true });
      // Older hosts return focus to the pane instead of firing `cancel`.
      input.addEventListener("cancel", () => finish(null), { once: true });
      // Several webviews restore window focus before dispatching `change`, so this fallback waits FOCUS_SETTLE_MS rather than settling at once.
      window.addEventListener("focus", onFocus, { once: true });
      timers.push(window.setTimeout(finishFromInput, PICKER_TIMEOUT_MS));

      input.click();
    });
  }
}

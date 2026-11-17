export type EditorIconName =
  "eye" | "eye-off" | "restore" | "more" | "download" | "replace";

const iconPaths: Record<EditorIconName, string> = {
  eye: "M2 12s3.6-6 10-6 10 6 10 6-3.6 6-10 6S2 12 2 12Zm10-3a3 3 0 1 0 0 6 3 3 0 0 0 0-6Z",
  "eye-off":
    "m3 3 18 18M10.6 6.2A10.7 10.7 0 0 1 12 6c6.4 0 10 6 10 6a16 16 0 0 1-3.1 3.8M6.2 6.2C3.5 7.8 2 12 2 12s3.6 6 10 6c1 0 2-.2 2.9-.5",
  restore: "M3 7V3m0 4h4m-4 0a9 9 0 1 1-1 8",
  more: "M5 12h.01M12 12h.01M19 12h.01",
  download: "M12 3v12m-5-5 5 5 5-5M4 20h16",
  replace:
    "M20 7h-5V2m5 5-5-5M4 17h5v5m-5-5 5 5m10-15a8 8 0 0 0-13 3M4 17a8 8 0 0 0 13-3",
};

export function createEditorIcon(name: EditorIconName): SVGSVGElement {
  const icon = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  icon.setAttribute("viewBox", "0 0 24 24");
  icon.setAttribute("aria-hidden", "true");
  icon.setAttribute("focusable", "false");
  icon.classList.add("ui-icon");
  const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
  path.setAttribute("d", iconPaths[name]);
  icon.append(path);
  return icon;
}

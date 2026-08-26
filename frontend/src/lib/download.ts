// Shared by both tb-client.ts and icai-client.ts: turns a generated-file
// response into a browser "Save As" download.
export function triggerDownload(blob: Blob, contentDisposition: string | null, fallbackName: string) {
  const match = contentDisposition?.match(/filename="([^"]+)"/i);
  const fileName = match?.[1] || fallbackName;

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export function isPDF(mime: string) {
  return mime === 'application/pdf';
}

export function isExcel(mime: string) {
  return mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
         mime === 'application/vnd.ms-excel' ||
         mime === 'text/csv';
}

export function isImage(mime: string) {
  const base64Prefix = 'data:image/';
  if (mime.startsWith(base64Prefix)) {
    const parts = mime.split(base64Prefix);
    if (parts.length > 2) {
      mime = base64Prefix + parts[1];
    }
  }
  return mime.startsWith('image/');
}

export function createDataUri(mimetype: string, base64: string): string {
  return `data:${mimetype};base64,${base64}`;
}

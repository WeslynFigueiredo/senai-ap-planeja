import PizZip from 'pizzip';

/**
 * OpenXML ZIP Package Access Layer
 * Handles reading and writing zip entries in memory without touching disk files.
 */

export function loadZipPackage(buffer: Buffer): PizZip {
  return new PizZip(buffer);
}

export function saveZipPackage(zip: PizZip): Buffer {
  return zip.generate({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  }) as Buffer;
}

export function readZipFileText(zip: PizZip, filePath: string): string {
  const file = zip.file(filePath);
  if (!file) {
    throw new Error(`File "${filePath}" not found in ZIP package.`);
  }
  return file.asText();
}

export function writeZipFileText(zip: PizZip, filePath: string, content: string): void {
  zip.file(filePath, content);
}

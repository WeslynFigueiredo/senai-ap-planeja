import PizZip from 'pizzip';
import { loadZipPackage } from './zip-package.ts';

export interface FormulaInfo {
  sheetFile: string;
  cellRef: string;
  formula: string;
}

export interface IntegrityReport {
  origEntryCount: number;
  genEntryCount: number;
  hasVbaProject: boolean;
  vmlFilesCount: number;
  mediaFilesCount: number;
  drawingsFilesCount: number;
  printerSettingsCount: number;
  hasDefinedNames: boolean;
  modifiedInternalFiles: string[];
  formulasBeforeCount: number;
  formulasAfterCount: number;
  inadvertentlyModifiedFormulasCount: number;
  isIntegrityPassed: boolean;
}

export function extractAllFormulas(zip: PizZip): FormulaInfo[] {
  const formulas: FormulaInfo[] = [];
  const wbXml = zip.file('xl/workbook.xml')?.asText() || '';
  const relsXml = zip.file('xl/_rels/workbook.xml.rels')?.asText() || '';

  const sheetMap = new Map<string, string>();
  const sheetMatches = Array.from(wbXml.matchAll(/<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g)).concat(
    Array.from(wbXml.matchAll(/<sheet [^>]*r:id="([^"]+)"[^>]*name="([^"]+)"/g)).map(m => [m[0], m[2], m[1]])
  );

  for (const m of sheetMatches) {
    const sheetName = m[1];
    const rId = m[2];
    const relMatch = relsXml.match(new RegExp(`ID="${rId}"[^>]*Target="([^"]+)"`, 'i')) ||
                     relsXml.match(new RegExp(`Target="([^"]+)"[^>]*ID="${rId}"`, 'i'));
    if (relMatch) {
      const target = relMatch[1].startsWith('xl/') ? relMatch[1] : `xl/${relMatch[1]}`;
      sheetMap.set(target, sheetName);
    }
  }

  const files = Object.keys(zip.files).filter(f => f.startsWith('xl/worksheets/') && f.endsWith('.xml'));

  for (const sheetFile of files) {
    const sheetName = sheetMap.get(sheetFile) || sheetFile;
    const xml = zip.file(sheetFile)!.asText();
    const matches = xml.matchAll(/<c\s+[^>]*?\br="([A-Z0-9]+)"[^>]*?>[\s\S]*?<f[^>]*?>([\s\S]*?)<\/f>/g);
    for (const m of matches) {
      formulas.push({
        sheetFile: sheetName,
        cellRef: m[1],
        formula: m[2],
      });
    }
  }

  return formulas;
}

export function compareWorkbookIntegrity(origBuffer: Buffer, genBuffer: Buffer): IntegrityReport {
  const origZip = loadZipPackage(origBuffer);
  const genZip = loadZipPackage(genBuffer);

  const origFiles = Object.keys(origZip.files);
  const genFiles = Object.keys(genZip.files);

  const hasVbaProject = genFiles.includes('xl/vbaProject.bin');
  const vmlFilesCount = genFiles.filter(f => f.endsWith('.vml')).length;
  const mediaFilesCount = genFiles.filter(f => f.startsWith('xl/media/')).length;
  const drawingsFilesCount = genFiles.filter(f => f.startsWith('xl/drawings/')).length;
  const printerSettingsCount = genFiles.filter(f => f.includes('printerSettings')).length;

  const workbookXmlGen = genZip.file('xl/workbook.xml')?.asText() || '';
  const hasDefinedNames = workbookXmlGen.includes('<definedNames>');

  // Identify internal files with content byte changes
  const modifiedInternalFiles: string[] = [];
  for (const fileName of origFiles) {
    const origFile = origZip.file(fileName);
    const genFile = genZip.file(fileName);
    if (!genFile) {
      modifiedInternalFiles.push(`[REMOVED] ${fileName}`);
      continue;
    }
    const origBytes = origFile.asNodeBuffer();
    const genBytes = genFile.asNodeBuffer();
    if (!origBytes.equals(genBytes)) {
      modifiedInternalFiles.push(fileName);
    }
  }

  // Formula checks mapped by Sheet Name
  const formulasBefore = extractAllFormulas(origZip);
  const formulasAfter = extractAllFormulas(genZip);

  const beforeMap = new Map<string, string>();
  formulasBefore.forEach(f => beforeMap.set(`${f.sheetFile}:${f.cellRef}`, f.formula));

  const sheetNamesAfter = new Set(formulasAfter.map(f => f.sheetFile));

  let inadvertentlyModified = 0;
  for (const fAfter of formulasAfter) {
    const key = `${fAfter.sheetFile}:${fAfter.cellRef}`;
    const formulaBefore = beforeMap.get(key);
    if (formulaBefore !== undefined && formulaBefore !== fAfter.formula) {
      inadvertentlyModified++;
    }
  }
  for (const fBefore of formulasBefore) {
    if (sheetNamesAfter.has(fBefore.sheetFile)) {
      const key = `${fBefore.sheetFile}:${fBefore.cellRef}`;
      const existsAfter = formulasAfter.some(fa => `${fa.sheetFile}:${fa.cellRef}` === key);
      if (!existsAfter) {
        inadvertentlyModified++;
      }
    }
  }

  const isIntegrityPassed =
    origFiles.length === genFiles.length &&
    hasVbaProject &&
    inadvertentlyModified === 0 &&
    modifiedInternalFiles.length === 1 &&
    modifiedInternalFiles[0] === 'xl/worksheets/sheet2.xml';

  return {
    origEntryCount: origFiles.length,
    genEntryCount: genFiles.length,
    hasVbaProject,
    vmlFilesCount,
    mediaFilesCount,
    drawingsFilesCount,
    printerSettingsCount,
    hasDefinedNames,
    modifiedInternalFiles,
    formulasBeforeCount: formulasBefore.length,
    formulasAfterCount: formulasAfter.length,
    inadvertentlyModifiedFormulasCount: inadvertentlyModified,
    isIntegrityPassed,
  };
}

import PizZip from 'pizzip';
import { readZipFileText, writeZipFileText } from './zip-package';

/**
 * OpenXML Workbook Worksheet Locator and View Manager
 * Resolves sheet names (e.g. "Plano de Ensino") to internal ZIP file paths,
 * manages active tab state in workbook.xml and sheetViews, and configures calculation properties.
 */

export function getWorksheetPathByName(zip: PizZip, sheetName: string): string {
  const workbookXml = readZipFileText(zip, 'xl/workbook.xml');
  const relsXml = readZipFileText(zip, 'xl/_rels/workbook.xml.rels');

  const sheetRegex = new RegExp(`<sheet\\s+[^>]*?name="${escapeRegex(sheetName)}"[^>]*?r:id="([^"]+)"`);
  const sheetMatch = workbookXml.match(sheetRegex);

  if (!sheetMatch) {
    throw new Error(`Sheet named "${sheetName}" was not found in xl/workbook.xml.`);
  }

  const rId = sheetMatch[1];

  const relRegex = new RegExp(`<Relationship\\s+[^>]*?Id="${rId}"[^>]*?Target="([^"]+)"`);
  const relMatch = relsXml.match(relRegex);

  if (!relMatch) {
    throw new Error(`Relationship with Id="${rId}" for sheet "${sheetName}" was not found in xl/_rels/workbook.xml.rels.`);
  }

  let targetPath = relMatch[1];
  if (!targetPath.startsWith('xl/')) {
    targetPath = 'xl/' + targetPath.replace(/^\//, '');
  }

  return targetPath;
}

export function setActiveSheet(zip: PizZip, targetSheetName: string): { activeTabIndex: number } {
  const wbXml = readZipFileText(zip, 'xl/workbook.xml');
  const relsXml = readZipFileText(zip, 'xl/_rels/workbook.xml.rels');

  const sheetMatches = Array.from(wbXml.matchAll(/<sheet\s+[^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"/g));
  let targetIndex = -1;

  sheetMatches.forEach((m, idx) => {
    if (m[1] === targetSheetName) {
      targetIndex = idx;
    }
  });

  if (targetIndex === -1) {
    throw new Error(`Sheet "${targetSheetName}" not found in xl/workbook.xml.`);
  }

  // 1. Update activeTab attribute in workbook.xml
  let updatedWbXml = wbXml.replace(/(<workbookView\s+[^>]*?\bactiveTab=")\d+(")/, `$1${targetIndex}$2`);
  if (!updatedWbXml.includes('activeTab=')) {
    updatedWbXml = updatedWbXml.replace(/(<workbookView\b)/, `$1 activeTab="${targetIndex}"`);
  }
  writeZipFileText(zip, 'xl/workbook.xml', updatedWbXml);

  // 2. Update tabSelected attribute in worksheet XMLs
  sheetMatches.forEach((m) => {
    const sName = m[1];
    const rId = m[2];
    const relRegex = new RegExp(`<Relationship\\s+[^>]*?Id="${rId}"[^>]*?Target="([^"]+)"`);
    const relMatch = relsXml.match(relRegex);
    if (!relMatch) return;

    let targetPath = relMatch[1];
    if (!targetPath.startsWith('xl/')) {
      targetPath = 'xl/' + targetPath.replace(/^\//, '');
    }

    try {
      let sXml = readZipFileText(zip, targetPath);
      if (sName === targetSheetName) {
        if (!sXml.includes('tabSelected=')) {
          sXml = sXml.replace(/(<sheetView\b)/, '$1 tabSelected="1"');
        } else {
          sXml = sXml.replace(/(<sheetView\s+[^>]*?\btabSelected=")[^"]*(")/, '$11$2');
        }
      } else {
        sXml = sXml.replace(/\s*tabSelected="[^"]*"/g, '');
      }
      writeZipFileText(zip, targetPath, sXml);
    } catch {
      // Ignore non-existent files if any
    }
  });

  return { activeTabIndex: targetIndex };
}

export function enableWorkbookRecalculation(zip: PizZip): {
  calcMode: string;
  fullCalcOnLoad: string;
  forceFullCalc: string;
  calcId: string;
} {
  let wbXml = readZipFileText(zip, 'xl/workbook.xml');

  let calcId = '191028';
  const calcIdMatch = wbXml.match(/calcId="([^"]+)"/);
  if (calcIdMatch) {
    calcId = calcIdMatch[1];
  }

  if (wbXml.includes('<calcPr')) {
    wbXml = wbXml.replace(
      /<calcPr\s+[^>]*?\/>/,
      `<calcPr calcId="${calcId}" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>`
    );
  } else {
    wbXml = wbXml.replace(
      '</workbook>',
      `<calcPr calcId="${calcId}" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>`
    );
  }

  writeZipFileText(zip, 'xl/workbook.xml', wbXml);

  return {
    calcMode: 'auto',
    fullCalcOnLoad: '1',
    forceFullCalc: '1',
    calcId,
  };
}

function escapeRegex(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

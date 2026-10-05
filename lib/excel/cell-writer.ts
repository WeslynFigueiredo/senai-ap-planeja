/**
 * OpenXML Safe Cell Writer
 * Modifies cell contents in worksheet XML without altering styles, formulas,
 * merge cells, drawings, or other cells in the sheet.
 */

export interface CellRef {
  col: string;
  row: number;
}

export function splitCellRef(ref: string): CellRef {
  const m = ref.match(/^([A-Z]+)([0-9]+)$/i);
  if (!m) {
    throw new Error(`Invalid cell reference format: "${ref}"`);
  }
  return {
    col: m[1].toUpperCase(),
    row: parseInt(m[2], 10),
  };
}

export function colToNumber(col: string): number {
  let num = 0;
  for (let i = 0; i < col.length; i++) {
    num = num * 26 + (col.charCodeAt(i) - 64);
  }
  return num;
}

export function escapeXml(str: string): string {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function buildCellXml(
  ref: string,
  type: 'string' | 'number',
  value: string | number,
  existingAttrStr: string = '',
  existingFormulaXml: string = ''
): string {
  let attrStr = existingAttrStr.replace(/\s*t="[^"]*"/g, '').trim();

  if (existingFormulaXml) {
    // Cell has an existing formula: preserve <f> tag and update cached <v>
    if (type === 'string') {
      attrStr = attrStr ? `t="str" ${attrStr}` : `t="str" r="${ref}"`;
      if (!/\br=/.test(attrStr)) {
        attrStr = `r="${ref}" ${attrStr}`;
      }
      const strVal = String(value);
      const escapedText = escapeXml(strVal);
      return `<c ${attrStr}>${existingFormulaXml}<v>${escapedText}</v></c>`;
    } else {
      attrStr = attrStr ? attrStr : `r="${ref}"`;
      if (!/\br=/.test(attrStr)) {
        attrStr = `r="${ref}" ${attrStr}`;
      }
      return `<c ${attrStr}>${existingFormulaXml}<v>${value}</v></c>`;
    }
  }

  if (type === 'string') {
    attrStr = attrStr ? `t="inlineStr" ${attrStr}` : `t="inlineStr" r="${ref}"`;
    if (!/\br=/.test(attrStr)) {
      attrStr = `r="${ref}" ${attrStr}`;
    }
    const strVal = String(value);
    const escapedText = escapeXml(strVal);
    const needsPreserve = /^\s|\s$|[\r\n\t]/.test(strVal);
    const spaceAttr = needsPreserve ? ' xml:space="preserve"' : '';
    return `<c ${attrStr}><is><t${spaceAttr}>${escapedText}</t></is></c>`;
  } else {
    attrStr = attrStr ? attrStr : `r="${ref}"`;
    if (!/\br=/.test(attrStr)) {
      attrStr = `r="${ref}" ${attrStr}`;
    }
    return `<c ${attrStr}><v>${value}</v></c>`;
  }
}

export function setCellValueInSheetXml(
  sheetXml: string,
  cellRef: string,
  type: 'string' | 'number',
  value: string | number
): string {
  const normalizedRef = cellRef.toUpperCase();
  const { col, row } = splitCellRef(normalizedRef);
  const targetColNum = colToNumber(col);

  // 1. Try finding existing cell tag in worksheet
  const cellRegex = new RegExp(`<c\\s+[^>]*?\\br="${normalizedRef}"[\\s\\S]*?(?:/>|>([\\s\\S]*?)</c>)`);
  const cellMatch = sheetXml.match(cellRegex);

  if (cellMatch) {
    const fullTag = cellMatch[0];
    const headerMatch = fullTag.match(/^<c\s+([^>]*?)\s*\/?>/);
    const existingAttrStr = headerMatch ? headerMatch[1] : `r="${normalizedRef}"`;

    const formulaMatch = fullTag.match(/<f[^>]*?>[\s\S]*?<\/f>/);
    const existingFormulaXml = formulaMatch ? formulaMatch[0] : '';

    const newCellXml = buildCellXml(normalizedRef, type, value, existingAttrStr, existingFormulaXml);
    return sheetXml.replace(fullTag, newCellXml);
  }

  // 2. Try finding existing row tag in worksheet
  const rowRegex = new RegExp(`<row\\s+[^>]*?\\br="${row}"[\\s\\S]*?(?:/>|>([\\s\\S]*?)</row>)`);
  const rowMatch = sheetXml.match(rowRegex);

  if (rowMatch) {
    const fullRow = rowMatch[0];
    const newCellXml = buildCellXml(normalizedRef, type, value);

    let updatedRow: string;
    if (fullRow.endsWith('/>')) {
      const rowHeader = fullRow.slice(0, -2);
      updatedRow = `${rowHeader}>${newCellXml}</row>`;
    } else {
      const rowHeaderMatch = fullRow.match(/^<row\s+[^>]*?>/);
      const rowHeader = rowHeaderMatch ? rowHeaderMatch[0] : `<row r="${row}">`;
      const innerContent = fullRow.slice(rowHeader.length, -6);

      const cellMatches = Array.from(innerContent.matchAll(/<c\s+[^>]*?\br="([A-Z]+)[0-9]+"[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g));
      let inserted = false;
      let newInner = '';
      let lastIdx = 0;

      for (const m of cellMatches) {
        const cCol = m[1];
        const cColNum = colToNumber(cCol);
        if (!inserted && cColNum > targetColNum) {
          newInner += innerContent.slice(lastIdx, m.index) + newCellXml;
          inserted = true;
          lastIdx = m.index;
        }
      }
      if (!inserted) {
        newInner += innerContent.slice(lastIdx) + newCellXml;
      } else {
        newInner += innerContent.slice(lastIdx);
      }
      updatedRow = `${rowHeader}${newInner}</row>`;
    }

    return sheetXml.replace(fullRow, updatedRow);
  }

  // 3. Row doesn't exist, insert new row into <sheetData>
  const newCellXml = buildCellXml(normalizedRef, type, value);
  const newRowXml = `<row r="${row}">${newCellXml}</row>`;

  const sheetDataMatch = sheetXml.match(/<sheetData[\s\S]*?>([\s\S]*?)<\/sheetData>/);
  if (!sheetDataMatch) {
    throw new Error('Element <sheetData> not found in worksheet XML.');
  }

  const sheetDataContent = sheetDataMatch[1];
  const rowMatches = Array.from(sheetDataContent.matchAll(/<row\s+[^>]*?\br="([0-9]+)"[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g));
  let inserted = false;
  let newSheetDataInner = '';
  let lastIdx = 0;

  for (const r of rowMatches) {
    const rNum = parseInt(r[1], 10);
    if (!inserted && rNum > row) {
      newSheetDataInner += sheetDataContent.slice(lastIdx, r.index) + newRowXml;
      inserted = true;
      lastIdx = r.index;
    }
  }

  if (!inserted) {
    newSheetDataInner += sheetDataContent.slice(lastIdx) + newRowXml;
  } else {
    newSheetDataInner += sheetDataContent.slice(lastIdx);
  }

  return sheetXml.replace(sheetDataMatch[0], `<sheetData>${newSheetDataInner}</sheetData>`);
}

export function setCellText(sheetXml: string, cellRef: string, value: string): string {
  return setCellValueInSheetXml(sheetXml, cellRef, 'string', value);
}

export function setCellNumber(sheetXml: string, cellRef: string, value: number): string {
  return setCellValueInSheetXml(sheetXml, cellRef, 'number', value);
}

export function setFormulaCellCachedValue(
  sheetXml: string,
  cellRef: string,
  type: 'string' | 'number',
  value: string | number
): string {
  return setCellValueInSheetXml(sheetXml, cellRef, type, value);
}

import PizZip from 'pizzip';
import { readZipFileText, writeZipFileText } from './zip-package.ts';
import { setActiveSheet, enableWorkbookRecalculation } from './workbook-map.ts';

function removeZipFile(zip: PizZip, file: string): void {
  zip.remove(file);
}

/**
 * Dynamically adjusts the number of SA sheets in the OpenXML workbook package.
 * Supports scaling down (e.g. 1, 2 SAs - removing surplus sheets and their exclusive resources)
 * and scaling up (e.g. 6, 7, 10 SAs - duplicating SA05 structural model with new unique IDs).
 */
export function adjustSaSheets(
  zip: PizZip,
  targetSaCount: number
): {
  finalSaCount: number;
  sheetsList: string[];
  removedFiles: string[];
  addedFiles: string[];
} {
  const wbXml = readZipFileText(zip, 'xl/workbook.xml');
  let relsXml = readZipFileText(zip, 'xl/_rels/workbook.xml.rels');
  let contentTypesXml = readZipFileText(zip, '[Content_Types].xml');

  const removedFiles: string[] = [];
  const addedFiles: string[] = [];

  // Parse all sheet tags in order
  const sheetMatches = Array.from(
    wbXml.matchAll(/<sheet\s+[^>]*?name="([^"]+)"[^>]*?sheetId="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*?\/>/g)
  );

  const existingSaSheets = sheetMatches.filter((m) => /^SA\d+$/i.test(m[1]));
  const currentSaCount = existingSaSheets.length;

  if (targetSaCount !== currentSaCount) {
    // Only clean calcChain if structural sheet additions/removals occurred
    if (zip.file('xl/calcChain.xml')) {
      removeZipFile(zip, 'xl/calcChain.xml');
      removedFiles.push('xl/calcChain.xml');
    }
    relsXml = relsXml.replace(/<Relationship\s+[^>]*?Target="calcChain\.xml"\s*\/?>/g, '');
    contentTypesXml = contentTypesXml.replace(/<Override\s+[^>]*?PartName="\/xl\/calcChain\.xml"\s*\/?>/g, '');
  }
  let newWbRelsXml = relsXml;

  if (targetSaCount < currentSaCount) {
    // === SCALE DOWN: REMOVE SURPLUS SAs ===
    const sasToRemove = existingSaSheets.slice(targetSaCount);
    const removeRIds = new Set(sasToRemove.map((s) => s[3]));

    let newWbXml = wbXml;
    sasToRemove.forEach((s) => {
      const sheetTagRegex = new RegExp(`<sheet\\s+[^>]*?name="${s[1]}"[^>]*?\\/>\\s*`);
      newWbXml = newWbXml.replace(sheetTagRegex, '');
    });

    const sheetsToRemovePaths = new Set<string>();
    removeRIds.forEach((rId) => {
      const relRegex = new RegExp(`<Relationship\\s+[^>]*?Id="${rId}"[^>]*?Target="([^"]+)"\\s*\\/?>`);
      const relMatch = newWbRelsXml.match(relRegex);
      if (relMatch) {
        let tPath = relMatch[1];
        if (!tPath.startsWith('xl/')) tPath = 'xl/' + tPath.replace(/^\//, '');
        sheetsToRemovePaths.add(tPath);
        newWbRelsXml = newWbRelsXml.replace(relMatch[0], '');
      }
    });

    sheetsToRemovePaths.forEach((wsPath) => {
      removeZipFile(zip, wsPath);
      removedFiles.push(wsPath);
      contentTypesXml = contentTypesXml.replace(
        new RegExp(`<Override\\s+[^>]*?PartName="/${wsPath}"\\s*\\/?>\\s*`),
        ''
      );

      const relsPath = wsPath.replace('worksheets/', 'worksheets/_rels/') + '.rels';
      const wsRelsContent = readZipFileText(zip, relsPath);
      if (wsRelsContent) {
        removeZipFile(zip, relsPath);
        removedFiles.push(relsPath);

        const targets = Array.from(wsRelsContent.matchAll(/Target="([^"]+)"/g)).map((m) => m[1]);
        targets.forEach((t) => {
          const fullT = t.replace(/^\.\.\//, 'xl/');
          if (
            fullT.includes('drawing') ||
            fullT.includes('vml') ||
            fullT.includes('printerSettings') ||
            fullT.includes('comments')
          ) {
            if (!fullT.includes('image')) {
              removeZipFile(zip, fullT);
              removedFiles.push(fullT);
              const drawRels = fullT.replace('drawings/', 'drawings/_rels/') + '.rels';
              if (zip.file(drawRels)) {
                removeZipFile(zip, drawRels);
                removedFiles.push(drawRels);
              }
              contentTypesXml = contentTypesXml.replace(
                new RegExp(`<Override\\s+[^>]*?PartName="/${fullT}"\\s*\\/?>\\s*`),
                ''
              );
            }
          }
        });
      }
    });

    writeZipFileText(zip, 'xl/workbook.xml', newWbXml);
    writeZipFileText(zip, 'xl/_rels/workbook.xml.rels', newWbRelsXml);
    writeZipFileText(zip, '[Content_Types].xml', contentTypesXml);
  } else if (targetSaCount > currentSaCount) {
    // === SCALE UP: CREATE NEW SAs FROM SA05 MODEL ===
    let newWbXml = wbXml;
    const modelSaMatch = existingSaSheets[existingSaSheets.length - 1]; // SA05
    const modelRId = modelSaMatch[3];

    const modelRelRegex = new RegExp(`<Relationship\\s+[^>]*?Id="${modelRId}"[^>]*?Target="([^"]+)"`);
    const modelRelMatch = relsXml.match(modelRelRegex);
    if (!modelRelMatch) {
      throw new Error(`Relationship for model SA rId="${modelRId}" not found.`);
    }
    let modelWsPath = modelRelMatch[1];
    if (!modelWsPath.startsWith('xl/')) modelWsPath = 'xl/' + modelWsPath.replace(/^\//, '');

    const maxSheetId = Math.max(...Array.from(wbXml.matchAll(/sheetId="(\d+)"/g)).map((m) => parseInt(m[1], 10)));
    const maxRIdNum = Math.max(...Array.from(relsXml.matchAll(/Id="rId(\d+)"/g)).map((m) => parseInt(m[1], 10)));

    const allZipFiles = Object.keys(zip.files);
    const sheetFilesNums = allZipFiles
      .map((f) => f.match(/^xl\/worksheets\/sheet(\d+)\.xml$/))
      .filter(Boolean)
      .map((m) => parseInt(m![1], 10));
    let nextSheetNum = Math.max(...sheetFilesNums) + 1;

    const drawingNums = allZipFiles
      .map((f) => f.match(/^xl\/drawings\/drawing(\d+)\.xml$/))
      .filter(Boolean)
      .map((m) => parseInt(m![1], 10));
    let nextDrawingNum = drawingNums.length ? Math.max(...drawingNums) + 1 : 1;

    const vmlNums = allZipFiles
      .map((f) => f.match(/^xl\/drawings\/vmlDrawing(\d+)\.vml$/))
      .filter(Boolean)
      .map((m) => parseInt(m![1], 10));
    let nextVmlNum = vmlNums.length ? Math.max(...vmlNums) + 1 : 1;

    const printerNums = allZipFiles
      .map((f) => f.match(/^xl\/printerSettings\/printerSettings(\d+)\.bin$/))
      .filter(Boolean)
      .map((m) => parseInt(m![1], 10));
    let nextPrinterNum = printerNums.length ? Math.max(...printerNums) + 1 : 1;

    let curSheetId = maxSheetId;
    let curRIdNum = maxRIdNum;

    for (let i = currentSaCount + 1; i <= targetSaCount; i++) {
      curSheetId++;
      curRIdNum++;
      const newSaName = `SA${String(i).padStart(2, '0')}`;
      const newRId = `rId${curRIdNum}`;
      const newWsPath = `xl/worksheets/sheet${nextSheetNum}.xml`;
      const newWsTarget = `worksheets/sheet${nextSheetNum}.xml`;

      // 1. Duplicate worksheet XML
      const modelXml = readZipFileText(zip, modelWsPath);
      writeZipFileText(zip, newWsPath, modelXml);
      addedFiles.push(newWsPath);

      // 2. Duplicate worksheet .rels and exclusive resources
      const modelRelsPath = modelWsPath.replace('worksheets/', 'worksheets/_rels/') + '.rels';
      const newRelsPath = newWsPath.replace('worksheets/', 'worksheets/_rels/') + '.rels';

      if (zip.file(modelRelsPath)) {
        let relsContent = readZipFileText(zip, modelRelsPath);

        // Duplicate Drawing
        if (relsContent.includes('drawing')) {
          const modelDrawNameMatch = relsContent.match(/Target="\.\.\/drawings\/(drawing\d+\.xml)"/);
          if (modelDrawNameMatch) {
            const oldDrawFile = `xl/drawings/${modelDrawNameMatch[1]}`;
            const newDrawFile = `xl/drawings/drawing${nextDrawingNum}.xml`;
            if (zip.file(oldDrawFile)) {
              writeZipFileText(zip, newDrawFile, readZipFileText(zip, oldDrawFile));
              addedFiles.push(newDrawFile);
              contentTypesXml += `\n<Override PartName="/${newDrawFile}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`;

              const oldDrawRels = oldDrawFile.replace('drawings/', 'drawings/_rels/') + '.rels';
              if (zip.file(oldDrawRels)) {
                const newDrawRels = newDrawFile.replace('drawings/', 'drawings/_rels/') + '.rels';
                writeZipFileText(zip, newDrawRels, readZipFileText(zip, oldDrawRels));
                addedFiles.push(newDrawRels);
              }
            }
            relsContent = relsContent.replace(modelDrawNameMatch[1], `drawing${nextDrawingNum}.xml`);
            nextDrawingNum++;
          }
        }

        // Duplicate VML
        if (relsContent.includes('vmlDrawing')) {
          const modelVmlMatch = relsContent.match(/Target="\.\.\/drawings\/(vmlDrawing\d+\.vml)"/);
          if (modelVmlMatch) {
            const oldVmlFile = `xl/drawings/${modelVmlMatch[1]}`;
            const newVmlFile = `xl/drawings/vmlDrawing${nextVmlNum}.vml`;
            if (zip.file(oldVmlFile)) {
              writeZipFileText(zip, newVmlFile, readZipFileText(zip, oldVmlFile));
              addedFiles.push(newVmlFile);
              contentTypesXml += `\n<Override PartName="/${newVmlFile}" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>`;
            }
            relsContent = relsContent.replace(modelVmlMatch[1], `vmlDrawing${nextVmlNum}.vml`);
            nextVmlNum++;
          }
        }

        // Duplicate printerSettings
        if (relsContent.includes('printerSettings')) {
          const modelPrinterMatch = relsContent.match(/Target="\.\.\/printerSettings\/(printerSettings\d+\.bin)"/);
          if (modelPrinterMatch) {
            const oldPrinterFile = `xl/printerSettings/${modelPrinterMatch[1]}`;
            const newPrinterFile = `xl/printerSettings/printerSettings${nextPrinterNum}.bin`;
            if (zip.file(oldPrinterFile)) {
              zip.file(newPrinterFile, zip.file(oldPrinterFile).asNodeBuffer());
              addedFiles.push(newPrinterFile);
            }
            relsContent = relsContent.replace(modelPrinterMatch[1], `printerSettings${nextPrinterNum}.bin`);
            nextPrinterNum++;
          }
        }

        writeZipFileText(zip, newRelsPath, relsContent);
        addedFiles.push(newRelsPath);
      }

      // Add to Content_Types and workbook.xml.rels
      contentTypesXml += `\n<Override PartName="/${newWsPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`;
      newWbRelsXml += `\n<Relationship Id="${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="${newWsTarget}"/>`;

      // Insert sheet in xl/workbook.xml BEFORE "Padrão de Desempenho"
      const newSheetTag = `<sheet name="${newSaName}" sheetId="${curSheetId}" r:id="${newRId}"/>`;
      const padraoTagMatch = newWbXml.match(/<sheet\s+[^>]*?name="Padrão de Desempenho"[^>]*?\/>/);
      if (padraoTagMatch) {
        newWbXml = newWbXml.replace(padraoTagMatch[0], `${newSheetTag}${padraoTagMatch[0]}`);
      } else {
        newWbXml = newWbXml.replace('</sheets>', `${newSheetTag}</sheets>`);
      }

      nextSheetNum++;
    }

    contentTypesXml = contentTypesXml.replace('</Types>', '').trim() + '\n</Types>';
    newWbRelsXml = newWbRelsXml.replace('</Relationships>', '').trim() + '\n</Relationships>';

    writeZipFileText(zip, 'xl/workbook.xml', newWbXml);
    writeZipFileText(zip, 'xl/_rels/workbook.xml.rels', newWbRelsXml);
    writeZipFileText(zip, '[Content_Types].xml', contentTypesXml);
  }

  // === RE-INDEX DEFINEDNAMES AND VALIDATE FINAL SHEETS LIST ===
  let finalWbXml = readZipFileText(zip, 'xl/workbook.xml');
  const finalSheetMatches = Array.from(
    finalWbXml.matchAll(/<sheet\s+[^>]*?name="([^"]+)"[^>]*?sheetId="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*?\/>/g)
  );

  const sheetIndexMap = new Map<string, number>();
  const finalSheetsList: string[] = [];
  finalSheetMatches.forEach((m, idx) => {
    sheetIndexMap.set(m[1], idx);
    finalSheetsList.push(m[1]);
  });

  const dnMatch = finalWbXml.match(/<definedNames[\s\S]*?<\/definedNames>/);
  if (dnMatch) {
    const dnContent = dnMatch[0];
    const dnEntries = Array.from(dnContent.matchAll(/<definedName\s+([^>]*?)>([\s\S]*?)<\/definedName>/g));
    let newDnInner = '';

    dnEntries.forEach((dn) => {
      const attrStr = dn[1];
      const formulaStr = dn[2];
      const saInFormula = formulaStr.match(/'(SA\d+)'!/);

      if (saInFormula) {
        const saName = saInFormula[1];
        if (!sheetIndexMap.has(saName)) {
          return; // Omit definedName for removed sheet
        }
      }

      const localSheetIdMatch = attrStr.match(/localSheetId="(\d+)"/);
      if (localSheetIdMatch && saInFormula) {
        const saName = saInFormula[1];
        const newLocalId = sheetIndexMap.get(saName);
        const newAttr = attrStr.replace(/localSheetId="\d+"/, `localSheetId="${newLocalId}"`);
        newDnInner += `<definedName ${newAttr}>${formulaStr}</definedName>`;
      } else {
        newDnInner += dn[0];
      }
    });

    // Duplicate definedNames for new SAs if targetSaCount > 5
    const modelSaName = 'SA05';
    if (targetSaCount > 5 && sheetIndexMap.has(modelSaName)) {
      for (let i = 6; i <= targetSaCount; i++) {
        const newSaName = `SA${String(i).padStart(2, '0')}`;
        const newLocalId = sheetIndexMap.get(newSaName);
        dnEntries.forEach((dn) => {
          const formulaStr = dn[2];
          if (formulaStr.includes(`'${modelSaName}'!`)) {
            const newFormula = formulaStr.replace(new RegExp(`'${modelSaName}'!`, 'g'), `'${newSaName}'!`);
            const newAttr = dn[1].replace(/localSheetId="\d+"/, `localSheetId="${newLocalId}"`);
            newDnInner += `<definedName ${newAttr}>${newFormula}</definedName>`;
          }
        });
      }
    }

    finalWbXml = finalWbXml.replace(dnMatch[0], `<definedNames>${newDnInner}</definedNames>`);
    writeZipFileText(zip, 'xl/workbook.xml', finalWbXml);
  }

  // Ensure initial active tab is Plano de Ensino and forced recalculation is enabled
  setActiveSheet(zip, 'Plano de Ensino');
  enableWorkbookRecalculation(zip);

  return {
    finalSaCount: targetSaCount,
    sheetsList: finalSheetsList,
    removedFiles,
    addedFiles,
  };
}

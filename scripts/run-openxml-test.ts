import fs from 'node:fs';
import path from 'node:path';
import { updateWorkbookCells, type CellUpdate } from '../lib/excel/openxml.ts';
import { compareWorkbookIntegrity } from '../lib/excel/integrity-checker.ts';

const TEMPLATE_PATH = path.resolve(process.cwd(), 'templates/PLANO_DE_ENSINO_MODELO.xlsm');
const OUTPUT_PATH = path.resolve(process.cwd(), 'tmp/TESTE_OPENXML.xlsm');

function main() {
  console.log('=== TESTE DE INTEGRIDADE OPENXML ===');
  console.log('Template:', TEMPLATE_PATH);

  if (!fs.existsSync(TEMPLATE_PATH)) {
    console.error('ERRO: Template original não encontrado em:', TEMPLATE_PATH);
    process.exit(1);
  }

  const templateBuffer = fs.readFileSync(TEMPLATE_PATH);

  // Define controlled updates
  const updates: CellUpdate[] = [
    {
      sheetName: 'Plano de Ensino',
      cellRef: 'C4',
      type: 'string',
      value: 'TESTE DE INTEGRIDADE',
    },
    {
      sheetName: 'Plano de Ensino',
      cellRef: 'H4',
      type: 'number',
      value: 1234,
    },
    {
      sheetName: 'Plano de Ensino',
      cellRef: 'D5',
      type: 'string',
      value: 'UC DE TESTE',
    },
  ];

  console.log('Aplicando alterações controladas:');
  updates.forEach(u => console.log(`  - ${u.sheetName}!${u.cellRef} = ${u.value}`));

  // Generate updated workbook buffer in memory
  const generatedBuffer = updateWorkbookCells(templateBuffer, updates);

  // Ensure tmp directory exists
  const tmpDir = path.dirname(OUTPUT_PATH);
  if (!fs.existsSync(tmpDir)) {
    fs.mkdirSync(tmpDir, { recursive: true });
  }

  // Save generated file strictly to tmp/TESTE_OPENXML.xlsm without mutating original template
  fs.writeFileSync(OUTPUT_PATH, generatedBuffer);
  console.log(`\nArquivo de teste gerado com sucesso em:\n  ${OUTPUT_PATH}`);

  // Run full integrity comparison
  const report = compareWorkbookIntegrity(templateBuffer, generatedBuffer);

  console.log('\n================ RELATÓRIO DE INTEGRIDADE DO PACOTE ================');
  console.log(`Quantidade de entradas no ZIP original: ${report.origEntryCount}`);
  console.log(`Quantidade de entradas no ZIP gerado:   ${report.genEntryCount}`);
  console.log(`Presença de xl/vbaProject.bin:         ${report.hasVbaProject ? 'SIM' : 'NÃO'}`);
  console.log(`Presença de arquivos VML:              ${report.vmlFilesCount > 0 ? `SIM (${report.vmlFilesCount} arquivos)` : 'NÃO'}`);
  console.log(`Presença de xl/media:                  ${report.mediaFilesCount > 0 ? `SIM (${report.mediaFilesCount} arquivos)` : 'NÃO'}`);
  console.log(`Presença de xl/drawings:               ${report.drawingsFilesCount > 0 ? `SIM (${report.drawingsFilesCount} arquivos)` : 'NÃO'}`);
  console.log(`Presença de printerSettings:           ${report.printerSettingsCount > 0 ? `SIM (${report.printerSettingsCount} arquivos)` : 'NÃO'}`);
  console.log(`Presença de definedNames:              ${report.hasDefinedNames ? 'SIM' : 'NÃO'}`);
  console.log('\nLista exata dos arquivos internos que tiveram bytes alterados:');
  report.modifiedInternalFiles.forEach(f => console.log(`  - ${f}`));

  console.log('\n================ VERIFICAÇÃO DE FÓRMULAS ================');
  console.log(`Fórmulas antes:                  ${report.formulasBeforeCount}`);
  console.log(`Fórmulas depois:                 ${report.formulasAfterCount}`);
  console.log(`Fórmulas alteradas inadvertidamente: ${report.inadvertentlyModifiedFormulasCount}`);
  console.log('=========================================================\n');

  if (report.isIntegrityPassed) {
    console.log('STATUS: SUCESSO ABSOLUTO! O pacote OpenXML foi alterado de forma limpa e preservada.');
  } else {
    console.warn('AVISO: Houve divergência no teste de integridade. Verifique os relatórios acima.');
  }
}

main();

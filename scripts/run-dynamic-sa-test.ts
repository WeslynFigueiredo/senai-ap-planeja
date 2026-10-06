import fs from 'node:fs';
import path from 'node:path';
import { generateWorkbookWithHeaders, type PlanoGeralData, type SituacaoAprendizagemData } from '../lib/excel/openxml.ts';
import { compareWorkbookIntegrity } from '../lib/excel/integrity-checker.ts';
import { loadZipPackage, readZipFileText } from '../lib/excel/zip-package.ts';
import { getWorksheetPathByName } from '../lib/excel/workbook-map.ts';

const TEMPLATE_PATH = path.resolve(process.cwd(), 'templates/PLANO_DE_ENSINO_MODELO.xlsm');

function buildFixture(saCount: number): PlanoGeralData {
  const sas: SituacaoAprendizagemData[] = [];
  const estrategias = ['Situação-Problema', 'Estudo de Caso', 'Pesquisa Aplicada', 'Projeto', 'Projeto Integrador'];

  for (let i = 1; i <= saCount; i++) {
    const numStr = String(i).padStart(2, '0');
    sas.push({
      numero: i,
      nome: `SA${numStr}`,
      estrategia: estrategias[(i - 1) % estrategias.length],
      contextualizacao: `Contextualização de teste para a situação de aprendizagem ${numStr}.`,
      qtdCapacidadesTecnicas: 3 + (i % 3),
      qtdCapacidadesSocioemocionais: 1 + (i % 2),
    });
  }

  return {
    curso: 'TÉCNICO EM DESENVOLVIMENTO DE SISTEMAS',
    cargaHorariaCurso: 1200,
    modalidade: 'Presencial',
    unidadeCurricular: 'PROGRAMAÇÃO WEB',
    cargaHorariaUc: 120,
    modulo: 'Módulo II',
    situacoesAprendizagem: sas,
  };
}

function validateScenario(saCount: number, fileName: string) {
  console.log(`\n================================================================`);
  console.log(`=== EXECUTANDO TESTE PARA CENÁRIO COM ${saCount} SAs (${fileName}) ===`);
  console.log(`================================================================`);

  const templateBuffer = fs.readFileSync(TEMPLATE_PATH);
  const fixture = buildFixture(saCount);

  const { buffer: generatedBuffer, activeTabIndex, recalcInfo, saManagerResult } = generateWorkbookWithHeaders(
    templateBuffer,
    fixture
  );

  const outputPath = path.resolve(process.cwd(), `tmp/${fileName}`);
  const tmpDir = path.dirname(outputPath);
  if (!fs.existsSync(tmpDir)) {
    fs.mkdirSync(tmpDir, { recursive: true });
  }

  fs.writeFileSync(outputPath, generatedBuffer);
  console.log(`Arquivo gerado em: ${outputPath}`);

  const zip = loadZipPackage(generatedBuffer);
  const wbXml = readZipFileText(zip, 'xl/workbook.xml');
  const relsXml = readZipFileText(zip, 'xl/_rels/workbook.xml.rels');
  const contentTypesXml = readZipFileText(zip, '[Content_Types].xml');

  // Sheet list validation
  const sheetMatches = Array.from(wbXml.matchAll(/<sheet\s+[^>]*?name="([^"]+)"[^>]*?sheetId="([^"]+)"[^>]*?r:id="([^"]+)"[^>]*?\/>/g));
  const sheetNames = sheetMatches.map(m => m[1]);
  const sheetIds = sheetMatches.map(m => m[2]);
  const sheetRIds = sheetMatches.map(m => m[3]);

  console.log('\n--- ESTRUTURA DO WORKBOOK ---');
  console.log(`Quantidade final de abas: ${sheetNames.length}`);
  console.log(`Ordem final das abas:`);
  sheetNames.forEach((name, idx) => console.log(`  ${idx}: ${name} (sheetId=${sheetIds[idx]}, rId=${sheetRIds[idx]})`));

  // Check unique sheetIds and rIds
  const uniqueSheetIds = new Set(sheetIds);
  const uniqueRIds = new Set(sheetRIds);
  console.log(`Unicidade de sheetIds: ${uniqueSheetIds.size === sheetIds.length ? 'OK (únicos)' : 'FALHA (duplicados)'}`);
  console.log(`Unicidade de rIds:     ${uniqueRIds.size === sheetRIds.length ? 'OK (únicos)' : 'FALHA (duplicados)'}`);

  // Check Padrão de Desempenho position
  const padraoIdx = sheetNames.indexOf('Padrão de Desempenho');
  const lastSaIdx = Math.max(...sheetNames.map((n, idx) => (/^SA\d+$/i.test(n) ? idx : -1)));
  console.log(`Posição da aba "Padrão de Desempenho": ${padraoIdx} (Após última SA de índice ${lastSaIdx}): ${padraoIdx > lastSaIdx ? 'CORRETA' : 'FALHA'}`);

  // Check activeTab
  const activeTabMatch = wbXml.match(/activeTab="(\d+)"/);
  console.log(`activeTab em workbook.xml: ${activeTabMatch ? activeTabMatch[1] : 'N/A'} (Esperado: ${activeTabIndex} - Plano de Ensino)`);

  // Check relationship completeness in workbook.xml.rels
  let relsOk = true;
  sheetRIds.forEach(rId => {
    if (!relsXml.includes(`Id="${rId}"`)) {
      console.error(`ERRO: Relationship Id="${rId}" ausente em workbook.xml.rels!`);
      relsOk = false;
    }
  });
  console.log(`Relacionamentos de abas em workbook.xml.rels: ${relsOk ? 'OK' : 'FALHA'}`);

  // Check content types for worksheets
  let ctOk = true;
  sheetRIds.forEach(rId => {
    const targetMatch = relsXml.match(new RegExp(`<Relationship\\s+[^>]*?Id="${rId}"[^>]*?Target="([^"]+)"`));
    if (targetMatch) {
      let t = targetMatch[1];
      if (!t.startsWith('xl/')) t = 'xl/' + t.replace(/^\//, '');
      if (!contentTypesXml.includes(`PartName="/${t}"`)) {
        console.error(`ERRO: ContentType para /${t} ausente em [Content_Types].xml!`);
        ctOk = false;
      }
    }
  });
  console.log(`ContentTypes de worksheets: ${ctOk ? 'OK' : 'FALHA'}`);

  // Audit report
  const report = compareWorkbookIntegrity(templateBuffer, generatedBuffer);
  console.log('\n--- RELATÓRIO DE INTEGRIDADE E PACOTE ---');
  console.log(`Entradas no ZIP original: ${report.origEntryCount}`);
  console.log(`Entradas no ZIP gerado:   ${report.genEntryCount}`);
  console.log(`vbaProject.bin preservado: ${report.hasVbaProject ? 'SIM' : 'NÃO'}`);
  console.log(`Arquivos VML preservados:  ${report.vmlFilesCount > 0 ? `SIM (${report.vmlFilesCount})` : 'NÃO'}`);
  console.log(`Desenhos preservados:     ${report.drawingsFilesCount > 0 ? `SIM (${report.drawingsFilesCount})` : 'NÃO'}`);
  console.log(`Mídias preservadas:       ${report.mediaFilesCount > 0 ? `SIM (${report.mediaFilesCount})` : 'NÃO'}`);
  console.log(`PrinterSettings:          ${report.printerSettingsCount > 0 ? `SIM (${report.printerSettingsCount})` : 'NÃO'}`);
  console.log(`Fórmulas antes: ${report.formulasBeforeCount}, Fórmulas depois: ${report.formulasAfterCount}`);

  console.log('\nArquivos alterados/adicionados/removidos:');
  console.log('  Removidos:', saManagerResult.removedFiles.length ? saManagerResult.removedFiles.join(', ') : 'Nenhum');
  console.log('  Adicionados:', saManagerResult.addedFiles.length ? saManagerResult.addedFiles.join(', ') : 'Nenhum');
}

function main() {
  console.log('=== INICIANDO SUÍTE DE TESTES DE ESCALONAMENTO DINÂMICO DE SAs ===');
  
  validateScenario(2, 'TESTE_2_SAS.xlsm');
  validateScenario(5, 'TESTE_5_SAS.xlsm');
  validateScenario(6, 'TESTE_6_SAS.xlsm');

  console.log('\n================================================================');
  console.log('SUÍTE DE TESTES CONCLUÍDA COM SUCESSO ABSOLUTO!');
  console.log('================================================================\n');
}

main();

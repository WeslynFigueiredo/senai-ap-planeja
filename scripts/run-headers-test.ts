import fs from 'node:fs';
import path from 'node:path';
import { generateWorkbookWithHeaders, type PlanoGeralData } from '../lib/excel/openxml.ts';
import { compareWorkbookIntegrity } from '../lib/excel/integrity-checker.ts';
import { loadZipPackage, readZipFileText } from '../lib/excel/zip-package.ts';
import { getWorksheetPathByName } from '../lib/excel/workbook-map.ts';

const TEMPLATE_PATH = path.resolve(process.cwd(), 'templates/PLANO_DE_ENSINO_MODELO.xlsm');
const OUTPUT_PATH = path.resolve(process.cwd(), 'tmp/TESTE_CABECALHOS_V2.xlsm');

const TEST_FIXTURE: PlanoGeralData = {
  curso: 'TÉCNICO EM DESENVOLVIMENTO DE SISTEMAS',
  cargaHorariaCurso: 1200,
  modalidade: 'Presencial',
  unidadeCurricular: 'PROGRAMAÇÃO WEB',
  cargaHorariaUc: 120,
  modulo: 'Módulo II',
  situacoesAprendizagem: [
    {
      numero: 1,
      nome: 'SA01',
      estrategia: 'Situação-Problema',
      contextualizacao: 'Desenvolvimento de portal web responsivo para gestão acadêmica.',
      qtdCapacidadesTecnicas: 4,
      qtdCapacidadesSocioemocionais: 2,
    },
    {
      numero: 2,
      nome: 'SA02',
      estrategia: 'Estudo de Caso',
      contextualizacao: 'Análise de vazamento de dados e auditoria de APIs REST.',
      qtdCapacidadesTecnicas: 3,
      qtdCapacidadesSocioemocionais: 2,
    },
    {
      numero: 3,
      nome: 'SA03',
      estrategia: 'Projeto Integrador',
      contextualizacao: 'Construção de plataforma SaaS com banco de dados relacional.',
      qtdCapacidadesTecnicas: 5,
      qtdCapacidadesSocioemocionais: 3,
    },
  ],
};

interface CellDetails {
  ref: string;
  tag: string;
  formula: string;
  typeAttr: string;
  cachedValue: string;
}

function inspectCellDetails(sheetXml: string, cellRef: string): CellDetails {
  const m = sheetXml.match(new RegExp(`<c\\s+[^>]*?\\br="${cellRef}"[\\s\\S]*?(?:\\/>|>([\\s\\S]*?)<\\/c>)`));
  if (!m) {
    return { ref: cellRef, tag: '[INEXISTENTE]', formula: 'N/A', typeAttr: 'N/A', cachedValue: 'N/A' };
  }
  const tag = m[0];
  const tMatch = tag.match(/t="([^"]+)"/);
  const fMatch = tag.match(/<f[^>]*?>([\s\S]*?)<\/f>/);
  const vMatch = tag.match(/<v>([\s\S]*?)<\/v>/);

  return {
    ref: cellRef,
    tag,
    formula: fMatch ? fMatch[1] : 'N/A',
    typeAttr: tMatch ? tMatch[1] : '(ausente/numérico)',
    cachedValue: vMatch ? vMatch[1] : 'N/A',
  };
}

function main() {
  console.log('=== TESTE DE PREENCHIMENTO DE CABEÇALHOS E SA01-SA05 (VERSÃO 2) ===');
  console.log('Template:', TEMPLATE_PATH);

  if (!fs.existsSync(TEMPLATE_PATH)) {
    console.error('ERRO: Template original não encontrado em:', TEMPLATE_PATH);
    process.exit(1);
  }

  const templateBuffer = fs.readFileSync(TEMPLATE_PATH);

  console.log('Aplicando dados do fixture (3 SAs utilizadas):');
  console.log(`  Curso: ${TEST_FIXTURE.curso}`);
  console.log(`  UC: ${TEST_FIXTURE.unidadeCurricular}`);
  console.log(`  Modalidade: ${TEST_FIXTURE.modalidade}`);
  console.log(`  Módulo: ${TEST_FIXTURE.modulo}`);
  console.log(`  SAs no fixture: ${TEST_FIXTURE.situacoesAprendizagem.length} (SA01, SA02, SA03)`);

  const { buffer: generatedBuffer, activeTabIndex, recalcInfo } = generateWorkbookWithHeaders(
    templateBuffer,
    TEST_FIXTURE
  );

  const tmpDir = path.dirname(OUTPUT_PATH);
  if (!fs.existsSync(tmpDir)) {
    fs.mkdirSync(tmpDir, { recursive: true });
  }

  fs.writeFileSync(OUTPUT_PATH, generatedBuffer);
  console.log(`\nArquivo de teste gerado com sucesso em:\n  ${OUTPUT_PATH}`);

  const genZip = loadZipPackage(generatedBuffer);

  console.log('\n================ VALIDAÇÃO PROGRAMÁTICA DO CACHE DAS FÓRMULAS ================');

  // Validate Plano de Ensino
  const planoPath = getWorksheetPathByName(genZip, 'Plano de Ensino');
  const planoXml = readZipFileText(genZip, planoPath);
  console.log('--- Aba "Plano de Ensino" ---');
  console.log('  C4 (Curso):             ', inspectCellDetails(planoXml, 'C4').cachedValue || inspectCellDetails(planoXml, 'C4').tag);
  console.log('  H4 (Carga Horária Curso):', inspectCellDetails(planoXml, 'H4').cachedValue);
  console.log('  J4 (Modalidade):        ', inspectCellDetails(planoXml, 'J4').cachedValue);
  console.log('  D5 (Unidade Curricular): ', inspectCellDetails(planoXml, 'D5').cachedValue);
  console.log('  H5 (Carga Horária UC):   ', inspectCellDetails(planoXml, 'H5').cachedValue);
  console.log('  J5 (Módulo):            ', inspectCellDetails(planoXml, 'J5').cachedValue);

  // Validate SA sheets (SA01..SA03)
  [1, 2, 3].forEach((num) => {
    const saName = `SA${String(num).padStart(2, '0')}`;
    const saPath = getWorksheetPathByName(genZip, saName);
    const saXml = readZipFileText(genZip, saPath);

    console.log(`\n--- Aba "${saName}" (UTILIZADA) ---`);
    ['E4', 'I4', 'L4', 'E5', 'I5', 'L5'].forEach((ref) => {
      const details = inspectCellDetails(saXml, ref);
      console.log(`  ${ref}:`);
      console.log(`    - Fórmula: =${details.formula}`);
      console.log(`    - Tipo t:  ${details.typeAttr}`);
      console.log(`    - Cache:   "${details.cachedValue}"`);
    });
  });

  // Validate Recalculation settings
  const wbXml = readZipFileText(genZip, 'xl/workbook.xml');
  const activeTabMatch = wbXml.match(/activeTab="(\d+)"/);
  console.log('\n================ CONFIGURAÇÃO DE CÁLCULO E ABERTURA (WORKBOOK.XML) ================');
  console.log(`  calcMode:       ${recalcInfo.calcMode}`);
  console.log(`  fullCalcOnLoad: ${recalcInfo.fullCalcOnLoad}`);
  console.log(`  forceFullCalc:  ${recalcInfo.forceFullCalc}`);
  console.log(`  calcId:         ${recalcInfo.calcId}`);
  console.log(`  activeTab:      ${activeTabMatch ? activeTabMatch[1] : 'N/A'} (Esperado: ${activeTabIndex} - Plano de Ensino)`);

  // Run integrity checker
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
  report.modifiedInternalFiles.forEach((f) => console.log(`  - ${f}`));

  console.log('\n================ VERIFICAÇÃO DE FÓRMULAS ================');
  console.log(`Fórmulas antes:                  ${report.formulasBeforeCount}`);
  console.log(`Fórmulas depois:                 ${report.formulasAfterCount}`);
  console.log(`Fórmulas alteradas inadvertidamente: ${report.inadvertentlyModifiedFormulasCount}`);
  console.log('=========================================================\n');
}

main();

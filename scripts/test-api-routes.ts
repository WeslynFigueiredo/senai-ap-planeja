import fs from 'fs';
import path from 'path';
import { validatePlanoEnsinoJson, mapJsonToPlanoGeralData } from '../lib/excel/plan-validator.ts';
import { generateWorkbookWithHeaders } from '../lib/excel/openxml.ts';
import { loadZipPackage, readZipFileText } from '../lib/excel/zip-package.ts';
import { getWorksheetPathByName } from '../lib/excel/workbook-map.ts';
import { compareWorkbookIntegrity } from '../lib/excel/integrity-checker.ts';

const TEMPLATE_PATH = path.resolve('templates/PLANO_DE_ENSINO_MODELO.xlsm');
const FIXTURE_PATH = path.resolve('tests/fixtures/plano-completo.json');
const DOWNLOAD_PATH = path.resolve('tmp/DOWNLOAD_API_TESTE.xlsm');

function sanitizeFilename(ucName: string): string {
  const cleanUc = (ucName || 'OFICIAL')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[\\/*?:"<>|]/g, '_')
    .replace(/\s+/g, '_')
    .toUpperCase();

  return `PLANO_DE_ENSINO_${cleanUc}.xlsm`;
}

function testSuccessFlow() {
  console.log('================ 1. TESTE DA ROTA DA API (FLUXO DE SUCESSO) ================');
  const jsonInput = JSON.parse(fs.readFileSync(FIXTURE_PATH, 'utf-8'));

  // 1. Validation check
  const valResult = validatePlanoEnsinoJson(jsonInput);
  console.log('Validação do JSON:', valResult.valid ? '✅ VÁLIDO' : '❌ INVÁLIDO');
  if (!valResult.valid) {
    console.error('Mensagem:', valResult.message);
    process.exit(1);
  }

  // 2. Generate Excel Buffer
  const templateBuffer = fs.readFileSync(TEMPLATE_PATH);
  const planoData = mapJsonToPlanoGeralData(jsonInput);
  const { buffer: excelBuffer } = generateWorkbookWithHeaders(templateBuffer, planoData);

  // 3. Filename sanitization
  const filename = sanitizeFilename(planoData.unidadeCurricular);
  console.log('Nome do arquivo gerado:', filename);

  fs.writeFileSync(DOWNLOAD_PATH, excelBuffer);
  console.log(`Arquivo salvo com sucesso para download em:\n  ${DOWNLOAD_PATH}`);

  // 4. Verify downloaded XLSM structure
  const genZip = loadZipPackage(excelBuffer);
  const wbXml = readZipFileText(genZip, 'xl/workbook.xml');
  const sheetMatches = Array.from(wbXml.matchAll(/<sheet [^>]*name="([^"]+)"/g)).map(m => m[1]);
  console.log('Abas no arquivo baixado:', sheetMatches.join(', '));

  const hasSa04 = sheetMatches.includes('SA04');
  const hasSa05 = sheetMatches.includes('SA05');
  console.log(`SA04/SA05 removidas? ${!hasSa04 && !hasSa05 ? 'SIM (OK)' : 'NÃO (ERRO)'}`);

  // Opening tab check
  const activeTabMatch = wbXml.match(/<workbookView [^>]*activeTab="(\d+)"/);
  const activeTabIdx = activeTabMatch ? parseInt(activeTabMatch[1], 10) : -1;
  const activeSheetName = activeTabIdx >= 0 ? sheetMatches[activeTabIdx] : 'desconhecida';
  console.log(`Aba ativa ao abrir: "${activeSheetName}" (index ${activeTabIdx})`);

  // Package integrity
  const report = compareWorkbookIntegrity(templateBuffer, excelBuffer);
  console.log('vbaProject.bin (macros) preservado:', report.hasVbaProject ? 'SIM' : 'NÃO');
}

function testErrorFlows() {
  console.log('\n================ 2. TESTE DE BLOQUEIOS DE ERRO NA API ================');
  const baseJson = JSON.parse(fs.readFileSync(FIXTURE_PATH, 'utf-8'));

  // Test 1: Plano não aprovado
  console.log('\n[Caso 1] Plano não aprovado (status.plano_aprovado = false)');
  const unapprovedPlan = JSON.parse(JSON.stringify(baseJson));
  unapprovedPlan.status = { plano_aprovado: false, padrao_desempenho_aprovado: true };
  const res1 = validatePlanoEnsinoJson(unapprovedPlan);
  console.log(res1.valid ? '❌ ERRO: Deveria ter sido bloqueado' : `✅ BLOQUEADO: "${res1.message}"`);

  // Test 2: Padrão de desempenho não aprovado
  console.log('\n[Caso 2] Padrão não aprovado (status.padrao_desempenho_aprovado = false)');
  const unapprovedPadrao = JSON.parse(JSON.stringify(baseJson));
  unapprovedPadrao.status = { plano_aprovado: true, padrao_desempenho_aprovado: false };
  const res2 = validatePlanoEnsinoJson(unapprovedPadrao);
  console.log(res2.valid ? '❌ ERRO: Deveria ter sido bloqueado' : `✅ BLOQUEADO: "${res2.message}"`);

  // Test 3: Peso != 100%
  console.log('\n[Caso 3] Pesos totalizam 90%');
  const badPeso = JSON.parse(JSON.stringify(baseJson));
  badPeso.padrao_desempenho.criterios[0].peso = 5;
  const res3 = validatePlanoEnsinoJson(badPeso);
  console.log(res3.valid ? '❌ ERRO: Deveria ter sido bloqueado' : `✅ BLOQUEADO: "${res3.message}"`);

  // Test 4: Arquivo estruturalmente inválido / sem schema_version
  console.log('\n[Caso 4] Schema version ausente / incompatível');
  const badSchema = JSON.parse(JSON.stringify(baseJson));
  badSchema.schema_version = 'invalido-1.0';
  const res4 = validatePlanoEnsinoJson(badSchema);
  console.log(res4.valid ? '❌ ERRO: Deveria ter sido bloqueado' : `✅ BLOQUEADO: "${res4.message}"`);
}

function main() {
  testSuccessFlow();
  testErrorFlows();
  console.log('\n================ TODOS OS TESTES DA API E FLUXO CONCLUÍDOS ================');
}

main();

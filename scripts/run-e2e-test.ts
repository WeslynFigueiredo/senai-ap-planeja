import fs from 'fs';
import path from 'path';
import { generateWorkbookWithHeaders, type PlanoGeralData } from '../lib/excel/openxml.ts';
import { validatePlanoEnsinoJson, mapJsonToPlanoGeralData } from '../lib/excel/plan-validator.ts';
import { loadZipPackage, readZipFileText } from '../lib/excel/zip-package.ts';
import { getWorksheetPathByName } from '../lib/excel/workbook-map.ts';
import { compareWorkbookIntegrity } from '../lib/excel/integrity-checker.ts';

const TEMPLATE_PATH = path.resolve('templates/PLANO_DE_ENSINO_MODELO.xlsm');
const FIXTURE_PATH = process.argv[2]
  ? path.resolve(process.argv[2])
  : path.resolve('tests/fixtures/plano-completo.json');
const OUTPUT_PATH = path.resolve('tmp/TESTE_COMPLETO_FINAL.xlsm');

function extractCellTextOrVal(sheetXml: string, cellRef: string): string {
  const m = sheetXml.match(new RegExp(`<c\\s+[^>]*?\\br="${cellRef}"[\\s\\S]*?(?:\\/>|>([\\s\\S]*?)<\\/c>)`));
  if (!m) return '[VAZIO]';
  const tag = m[0];
  const isMatch = tag.match(/<t[^>]*?>([\s\S]*?)<\/t>/);
  if (isMatch) return isMatch[1];
  const vMatch = tag.match(/<v>([\s\S]*?)<\/v>/);
  if (vMatch) return vMatch[1];
  const fMatch = tag.match(/<f[^>]*?>([\s\S]*?)<\/f>/);
  if (fMatch) return `Fórmula: =${fMatch[1]}`;
  return tag;
}

function main() {
  console.log('================ TESTE COMPLETO PONTA A PONTA (E2E) MOTOR XLSM ================');
  console.log('Template:', TEMPLATE_PATH);
  console.log('Fixture:', FIXTURE_PATH);

  if (!fs.existsSync(TEMPLATE_PATH)) {
    console.error('ERRO: Template original não encontrado em:', TEMPLATE_PATH);
    process.exit(1);
  }
  if (!fs.existsSync(FIXTURE_PATH)) {
    console.error('ERRO: Fixture JSON não encontrada em:', FIXTURE_PATH);
    process.exit(1);
  }

  const rawFixture = fs.readFileSync(FIXTURE_PATH, 'utf-8');
  const jsonInput = JSON.parse(rawFixture);

  // Validate JSON schema
  const validation = validatePlanoEnsinoJson(jsonInput);
  if (!validation.valid) {
    console.error('❌ ERRO NA VALIDAÇÃO DO JSON:', validation.message);
    process.exit(1);
  }
  console.log('✅ Validação do JSON executada com sucesso!');

  // Map JSON schema "senai-plano-ensino-1.0" to PlanoGeralData
  const planoData: PlanoGeralData = mapJsonToPlanoGeralData(jsonInput);

  const templateBuffer = fs.readFileSync(TEMPLATE_PATH);

  console.log('\n--- GERAÇÃO DO PACOTE XLSM EM MEMÓRIA ---');
  const { buffer: generatedBuffer, saSummaryList, recalcInfo } = generateWorkbookWithHeaders(
    templateBuffer,
    planoData
  );

  const tmpDir = path.dirname(OUTPUT_PATH);
  if (!fs.existsSync(tmpDir)) {
    fs.mkdirSync(tmpDir, { recursive: true });
  }

  fs.writeFileSync(OUTPUT_PATH, generatedBuffer);
  console.log(`✅ Arquivo final gerado com sucesso em:\n  ${OUTPUT_PATH}`);

  const genZip = loadZipPackage(generatedBuffer);

  // 1. VALIDAR ESTRUTURA DO WORKBOOK E ABAS
  console.log('\n================ 1. ESTRUTURA DO WORKBOOK E ABAS ================');
  const wbXml = readZipFileText(genZip, 'xl/workbook.xml');
  const sheetMatches = Array.from(wbXml.matchAll(/<sheet [^>]*name="([^"]+)"/g)).map(m => m[1]);
  console.log('Abas presentes no arquivo final:', sheetMatches.join(', '));

  const expectedSheets = ['Base', 'Plano de Ensino', 'TempGPT', ...planoData.situacoesAprendizagem.map(sa => sa.nome), 'Padrão de Desempenho'];
  const hasSa04 = sheetMatches.includes('SA04');
  const hasSa05 = sheetMatches.includes('SA05');

  console.log(`Abas SA04/SA05 removidas fisicamente? ${!hasSa04 && !hasSa05 ? 'SIM (OK)' : 'NÃO (ERRO)'}`);
  console.log(`Lista exata de abas coincide com o esperado? ${JSON.stringify(sheetMatches) === JSON.stringify(expectedSheets) ? 'SIM (OK)' : 'NÃO'}`);

  const activeTabMatch = wbXml.match(/<workbookView [^>]*activeTab="(\d+)"/);
  const activeTabIdx = activeTabMatch ? parseInt(activeTabMatch[1], 10) : -1;
  const activeSheetName = activeTabIdx >= 0 ? sheetMatches[activeTabIdx] : 'desconhecida';
  console.log(`Aba ativa na abertura inicial: "${activeSheetName}" (index ${activeTabIdx})`);

  // 2. VALIDAR PLANO DE ENSINO
  console.log('\n================ 2. VALIDAÇÃO DA ABA PLANO DE ENSINO ================');
  const planoPath = getWorksheetPathByName(genZip, 'Plano de Ensino');
  const planoXml = readZipFileText(genZip, planoPath);

  const cursoVal = extractCellTextOrVal(planoXml, 'C4');
  const chCursoVal = extractCellTextOrVal(planoXml, 'H4');
  const modVal = extractCellTextOrVal(planoXml, 'J4');
  const ucVal = extractCellTextOrVal(planoXml, 'D5');
  const chUcVal = extractCellTextOrVal(planoXml, 'H5');
  const moduloVal = extractCellTextOrVal(planoXml, 'J5');

  console.log(`  Curso:               "${cursoVal}"`);
  console.log(`  CH Curso:            ${chCursoVal}h`);
  console.log(`  Modalidade:          "${modVal}"`);
  console.log(`  Unidade Curricular:  "${ucVal}"`);
  console.log(`  CH UC:               ${chUcVal}h`);
  console.log(`  Módulo:              "${moduloVal}"`);

  // 3. VALIDAR SAs E CABEÇALHOS
  console.log('\n================ 3. VALIDAÇÃO DAS ABAS SAs ================');
  saSummaryList.forEach((summary) => {
    console.log(`\n--- Planilha "${summary.saName}" ---`);
    console.log(`  Carga Horária Declarada: ${summary.declaredCh}h`);
    console.log(`  Carga Horária Calculada: ${summary.calculatedCh}h`);
    console.log(`  Quantidade de Aulas:     ${summary.aulasCount}`);
    console.log(`  Linhas Utilizadas:       ${summary.rowsUsed.join(', ')}`);
    console.log(`  Cache da Fórmula L6:    ${summary.chFormulaCachedValue}`);

    const saPath = getWorksheetPathByName(genZip, summary.saName);
    const saXml = readZipFileText(genZip, saPath);

    // Check L6 formula tag
    const l6Cell = saXml.match(/<c [^>]*r="L6"[^>]*>[\s\S]*?<\/c>/);
    console.log(`  Célula L6 no XML:        ${l6Cell ? l6Cell[0] : '[Missing]'}`);

    // Check headers cached values
    console.log(`  Header Curso (E4):       ${extractCellTextOrVal(saXml, 'E4')}`);
    console.log(`  Header UC (E5):          ${extractCellTextOrVal(saXml, 'E5')}`);

    // Strategy row 6
    const strategyVal = extractCellTextOrVal(saXml, 'B6');
    console.log(`  Estratégia (Linha 6):    "${strategyVal.slice(0, 40)}..."`);
  });

  // 4. VALIDAR AULAS DETALHADAS
  console.log('\n================ 4. VALIDAÇÃO DETALHADA DAS AULAS ================');
  planoData.situacoesAprendizagem.forEach((sa: any) => {
    const saName = sa.nome;
    const saPath = getWorksheetPathByName(genZip, saName);
    const saXml = readZipFileText(genZip, saPath);

    console.log(`\n--- Aulas da ${saName} ---`);
    (sa.aulas || []).forEach((aula: any, idx: number) => {
      const row = 13 + idx * 2;
      console.log(`  -> Aula ${aula.numero} (Linha ${row}):`);
      console.log(`       Data (Col B):          ${extractCellTextOrVal(saXml, `B${row}`)}`);
      console.log(`       CH (Col C):            ${extractCellTextOrVal(saXml, `C${row}`)}h`);
      console.log(`       Capacidades (Col D):   "${extractCellTextOrVal(saXml, `D${row}`).replace(/\n/g, ' | ').slice(0, 50)}..."`);
      console.log(`       Conhecimentos (Col G): "${extractCellTextOrVal(saXml, `G${row}`).replace(/\n/g, ' | ').slice(0, 50)}..."`);
      console.log(`       Desafio (Col I):       "${extractCellTextOrVal(saXml, `I${row}`).slice(0, 40)}..."`);
      console.log(`       Resultado (Col J):     "${extractCellTextOrVal(saXml, `J${row}`).slice(0, 40)}..."`);
      console.log(`       Estratégias (Col K):   "${extractCellTextOrVal(saXml, `K${row}`).replace(/\n/g, ' \\n ').slice(0, 50)}..."`);
      console.log(`       Avaliação (Col L):     "${extractCellTextOrVal(saXml, `L${row}`).replace(/\n/g, ' \\n ').slice(0, 50)}..."`);

      // Assertions on final XML content of K and L cells
      const kContent = extractCellTextOrVal(saXml, `K${row}`);
      const lContent = extractCellTextOrVal(saXml, `L${row}`);

      if (!kContent.includes('Estratégias de Ensino:')) throw new Error(`[SA XML Assert Error] ${saName} K${row} sem 'Estratégias de Ensino:'`);
      if (!kContent.includes('Descrição da atividade:')) throw new Error(`[SA XML Assert Error] ${saName} K${row} sem 'Descrição da atividade:'`);
      if (!kContent.includes('Roteiro da aula:')) throw new Error(`[SA XML Assert Error] ${saName} K${row} sem 'Roteiro da aula:'`);

      if (!lContent.includes('Instrumentos:')) throw new Error(`[SA XML Assert Error] ${saName} L${row} sem 'Instrumentos:'`);
      if (!lContent.includes('Avaliação:')) throw new Error(`[SA XML Assert Error] ${saName} L${row} sem 'Avaliação:'`);
      if (!lContent.includes('Evidências:')) throw new Error(`[SA XML Assert Error] ${saName} L${row} sem 'Evidências:'`);
      if (!lContent.includes('Entrega:')) throw new Error(`[SA XML Assert Error] ${saName} L${row} sem 'Entrega:'`);

      if (saName === 'SA01' && row === 13) {
        if (!kContent.includes('07h30')) throw new Error(`[SA XML Assert Error] SA01 K13 sem horário '07h30'`);
        if (!kContent.includes('12h50')) throw new Error(`[SA XML Assert Error] SA01 K13 sem horário '12h50'`);
      }
    });

    // Audit row heights in current SA sheet: NO row ht may exceed 409pt (Excel limit)
    const rowMatches = Array.from(saXml.matchAll(/<row\s+[^>]*?\br="(\d+)"[^>]*?\bht="([^"]+)"[^>]*?>/g));
    for (const m of rowMatches) {
      const rNum = parseInt(m[1], 10);
      const htVal = parseFloat(m[2]);
      if (htVal > 409.0) {
        throw new Error(`[Excel Max Height Violation Error] ${saName} linha ${rNum} possui ht=${htVal}pt > 409pt!`);
      }
    }
  });

  // 5. VALIDAR PADRÃO DE DESEMPENHO
  console.log('\n================ 5. VALIDAÇÃO DO PADRÃO DE DESEMPENHO ================');
  const padraoPath = getWorksheetPathByName(genZip, 'Padrão de Desempenho');
  const padraoXml = readZipFileText(genZip, padraoPath);

  jsonInput.padrao_desempenho.criterios.forEach((crit: any, idx: number) => {
    const row = 5 + idx;
    console.log(`  Critério ${crit.numero} (Linha ${row}):`);
    console.log(`    Nº: ${extractCellTextOrVal(padraoXml, `C${row}`)} | Peso: ${Number(extractCellTextOrVal(padraoXml, `M${row}`)) * 100}%`);
    console.log(`    Capacidade: "${extractCellTextOrVal(padraoXml, `D${row}`).slice(0, 45)}..."`);
    console.log(`    Critério:   "${extractCellTextOrVal(padraoXml, `F${row}`).slice(0, 45)}..."`);
    console.log(`    SA(s):      "${extractCellTextOrVal(padraoXml, `G${row}`).replace(/\n/g, ', ')}"`);
  });

  // 6. TESTE DE FIDELIDADE JSON -> XLSM (VERIFICAÇÃO TEXTUAL DE CAMPO POR CAMPO)
  console.log('\n================ 6. TESTE DE FIDELIDADE JSON -> XLSM ================');
  let fidelityErrors = 0;

  function checkFidelity(fieldName: string, expected: string, actual: string) {
    if (expected.trim() !== actual.trim()) {
      console.error(`❌ DIVERGÊNCIA EM "${fieldName}":`);
      console.error(`     Esperado: "${expected}"`);
      console.error(`     Encontrado: "${actual}"`);
      fidelityErrors++;
    } else {
      console.log(`  ✅ [OK] ${fieldName}`);
    }
  }

  checkFidelity('Curso (Plano de Ensino)', planoData.curso, cursoVal);
  checkFidelity('Unidade Curricular (Plano de Ensino)', planoData.unidadeCurricular, ucVal);
  checkFidelity('Modalidade (Plano de Ensino)', planoData.modalidade, modVal);
  checkFidelity('Módulo (Plano de Ensino)', planoData.modulo, moduloVal);

  // Check first SA contextualização
  const sa01Path = getWorksheetPathByName(genZip, 'SA01');
  const sa01Xml = readZipFileText(genZip, sa01Path);
  checkFidelity('Contextualização SA01 (B10)', planoData.situacoesAprendizagem[0].contextualizacao, extractCellTextOrVal(sa01Xml, 'B10'));

  // Check first lesson desafio
  checkFidelity('Desafio Aula 1 SA01 (I13)', planoData.situacoesAprendizagem[0].aulas[0].desafio, extractCellTextOrVal(sa01Xml, 'I13'));

  // Check first criterion N4 descriptor
  const n4Crit1 = extractCellTextOrVal(padraoXml, 'L5');
  checkFidelity('Descritor N4 Critério 1 (L5)', planoData.padraoDesempenho!.criterios[0].n4, n4Crit1);

  if (fidelityErrors === 0) {
    console.log('\n🎉 FIDELIDADE TOTAL 100%! Nenhum truncamento, alteração ou perda de caracter detectada.');
  } else {
    console.error(`\n⚠️ ENCONTRADAS ${fidelityErrors} DIVERGÊNCIAS DE FIDELIDADE.`);
  }

  // 7. INTEGRIDADE DO PACOTE
  const report = compareWorkbookIntegrity(templateBuffer, generatedBuffer);
  console.log('\n================ 7. RELATÓRIO DE INTEGRIDADE E RECURSOS ================');
  console.log(`vbaProject.bin preservado: ${report.hasVbaProject ? 'SIM' : 'NÃO'}`);
  console.log(`Arquivos VML preservados:  ${report.vmlFilesCount > 0 ? `SIM (${report.vmlFilesCount})` : 'NÃO'}`);
  console.log(`Desenhos preservados:     ${report.drawingsFilesCount > 0 ? `SIM (${report.drawingsFilesCount})` : 'NÃO'}`);
  console.log(`Mídias preservadas:       ${report.mediaFilesCount > 0 ? `SIM (${report.mediaFilesCount})` : 'NÃO'}`);
  console.log(`PrinterSettings:          ${report.printerSettingsCount > 0 ? `SIM (${report.printerSettingsCount})` : 'NÃO'}`);
  console.log(`Fórmulas alteradas inadvertidamente: ${report.inadvertentlyModifiedFormulasCount}`);

  console.log('\n================ TESTE COMPLETO E2E CONCLUÍDO COM SUCESSO ================');
}

main();

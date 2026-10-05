import fs from 'node:fs';
import path from 'node:path';
import { generateWorkbookWithHeaders, type PlanoGeralData } from '../lib/excel/openxml.ts';
import { compareWorkbookIntegrity } from '../lib/excel/integrity-checker.ts';
import { loadZipPackage, readZipFileText } from '../lib/excel/zip-package.ts';
import { getWorksheetPathByName } from '../lib/excel/workbook-map.ts';

const TEMPLATE_PATH = path.resolve(process.cwd(), 'templates/PLANO_DE_ENSINO_MODELO.xlsm');
const OUTPUT_PATH = path.resolve(process.cwd(), 'tmp/TESTE_AULAS_SA.xlsm');

const LESSONS_FIXTURE: PlanoGeralData = {
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
      carga_horaria: 12,
      estrategia: 'Situação-Problema',
      contextualizacao: 'Desenvolvimento de portal web responsivo para gestão acadêmica.',
      qtdCapacidadesTecnicas: 3,
      qtdCapacidadesSocioemocionais: 2,
      aulas: [
        {
          numero: 1,
          sa_id: 'SA01',
          data: '2026-03-15',
          carga_horaria: 4,
          capacidades_tecnicas: ['Identificar requisitos de software', 'Mapear leiaute de páginas web'],
          capacidades_socioemocionais: ['Demonstrar atitude colaborativa em equipe'],
          conhecimentos: ['Conceitos de HTML5 e semântica web', 'Estrutura de documentos digitais'],
          desafio: 'Elaborar protótipo navegável em HTML5 conforme especificação do cliente.',
          resultado_esperado: 'Protótipo funcional validado sem erros de validação W3C.',
          estrategias_ensino: 'Aula expositiva dialogada e prática em laboratório.',
          descricao_atividade: 'Alunos desenvolvem individualmente a estrutura HTML5 da homepage.',
          roteiro_temporal: '00:00-01:00 — Apresentação dos conceitos;\n01:00-04:00 — Prática guiada.',
          instrumentos_avaliacao: 'Ficha de observação de desempenho prático.',
          descricao_avaliacao: 'Avaliação formativa da estrutura de código.',
          evidencias_aprendizagem: 'Arquivo index.html estruturado.',
          entrega_estudante: 'Envio do código no repositório institucional.',
        },
        {
          numero: 2,
          sa_id: 'SA01',
          data: '2026-03-16',
          carga_horaria: 4,
          capacidades_tecnicas: ['Estilizar interfaces digitais com CSS3', 'Aplicar regras de responsividade'],
          capacidades_socioemocionais: ['Trabalhar com perseverança na resolução de problemas'],
          conhecimentos: ['CSS3 Flexbox e Grid Layout', 'Media Queries'],
          desafio: 'Implementar a estilização responsiva do portal acadêmico.',
          resultado_esperado: 'Layout adaptável para dispositivos móveis e desktop.',
          estrategias_ensino: 'Estudo dirigido de codificação CSS.',
          descricao_atividade: 'Construção de folhas de estilo modularizadas.',
          roteiro_temporal: '00:00-00:40 — Conceitos de Flexbox;\n00:40-04:00 — Codificação.',
          instrumentos_avaliacao: 'Checklist de critérios de layout.',
          descricao_avaliacao: 'Verificação da fidelidade ao protótipo visual.',
          evidencias_aprendizagem: 'Folha de estilo style.css completa.',
          entrega_estudante: 'Push do projeto estilizado no Git.',
        },
        {
          numero: 3,
          sa_id: 'SA01',
          data: '2026-03-17',
          carga_horaria: 4,
          capacidades_tecnicas: ['Programar interatividade client-side com JavaScript'],
          capacidades_socioemocionais: ['Comunicar ideias de forma clara e assertiva'],
          conhecimentos: ['Manipulação do DOM com JavaScript', 'Eventos do navegador'],
          desafio: 'Adicionar validação de formulários e efeitos interativos em JS.',
          resultado_esperado: 'Formulário com feedback visual dinâmico ao usuário.',
          estrategias_ensino: 'Coding session em pares.',
          descricao_atividade: 'Implementação de scripts de validação de dados de entrada.',
          roteiro_temporal: '00:00-01:00 — DOM manipulation;\n01:00-04:00 — Prática.',
          instrumentos_avaliacao: 'Rubrica analítica de código JavaScript.',
          descricao_avaliacao: 'Avaliação da lógica e tratamento de exceções.',
          evidencias_aprendizagem: 'Script app.js funcionando.',
          entrega_estudante: 'Link do repositório publicado.',
        },
      ],
    },
    {
      numero: 2,
      nome: 'SA02',
      carga_horaria: 8,
      estrategia: 'Estudo de Caso',
      contextualizacao: 'Análise de falhas de segurança e otimização de performance em APIs.',
      qtdCapacidadesTecnicas: 2,
      qtdCapacidadesSocioemocionais: 1,
      aulas: [
        {
          numero: 1,
          sa_id: 'SA02',
          data: '2026-03-18',
          carga_horaria: 4,
          capacidades_tecnicas: ['Identificar vulnerabilidades em requisições HTTP'],
          capacidades_socioemocionais: ['Pensamento crítico e investigativo'],
          conhecimentos: ['Segurança Web', 'OWASP Top 10'],
          desafio: 'Analisar e documentar brechas de segurança em API legado.',
          resultado_esperado: 'Relatório técnico de auditoria com plano de mitigação.',
          estrategias_ensino: 'Análise orientada de estudo de caso.',
          descricao_atividade: 'Auditoria de endpoints HTTP utilizando inspeção de rede.',
          roteiro_temporal: '00:00-01:30 — Apresentação do caso;\n01:30-04:00 — Análise.',
          instrumentos_avaliacao: 'Relatório técnico formatado.',
          descricao_avaliacao: 'Clareza na identificação dos riscos de segurança.',
          evidencias_aprendizagem: 'Documento PDF com diagnóstico de segurança.',
          entrega_estudante: 'Upload do relatório auditado.',
        },
        {
          numero: 2,
          sa_id: 'SA02',
          data: '2026-03-19',
          carga_horaria: 4,
          capacidades_tecnicas: ['Refatorar rotas e implementar sanitização de entradas'],
          capacidades_socioemocionais: ['Comprometimento com a qualidade do software'],
          conhecimentos: ['Sanitização de dados', 'Tokens de autenticação JWT'],
          desafio: 'Corrigir as vulnerabilidades identificadas no estudo de caso.',
          resultado_esperado: 'API refatorada e protegida contra ataques XSS e Injection.',
          estrategias_ensino: 'Hands-on de refatoração de código.',
          descricao_atividade: 'Implementação de middlewares de segurança e validação.',
          roteiro_temporal: '00:00-01:00 — Revisão do código;\n01:00-04:00 — Refatoração.',
          instrumentos_avaliacao: 'Bateria de testes automatizados de segurança.',
          descricao_avaliacao: 'Aprovação em 100% dos testes de penetração.',
          evidencias_aprendizagem: 'Código-fonte refatorado com middleware.',
          entrega_estudante: 'Merge request no repositório.',
        },
      ],
    },
    {
      numero: 3,
      nome: 'SA03',
      carga_horaria: 12,
      estrategia: 'Projeto Integrador',
      contextualizacao: 'Construção de plataforma SaaS com banco de dados relacional.',
      qtdCapacidadesTecnicas: 3,
      qtdCapacidadesSocioemocionais: 2,
      aulas: [
        {
          numero: 1,
          sa_id: 'SA03',
          data: '2026-03-20',
          carga_horaria: 4,
          capacidades_tecnicas: ['Modelar banco de dados relacional'],
          capacidades_socioemocionais: ['Organização e planejamento do tempo'],
          conhecimentos: ['Modelagem MER e DER', 'Normalização até 3FN'],
          desafio: 'Criar o diagrama entidade-relacionamento da plataforma SaaS.',
          resultado_esperado: 'Modelo DER validado e normalizado.',
          estrategias_ensino: 'Modelagem orientada no quadro e software CASE.',
          descricao_atividade: 'Criação do esquema de tabelas e relacionamentos.',
          roteiro_temporal: '00:00-01:00 — Regras de negócio;\n01:00-04:00 — Diagramação.',
          instrumentos_avaliacao: 'Avaliação de diagrama conceitual.',
          descricao_avaliacao: 'Consistência das chaves primárias e estrangeiras.',
          evidencias_aprendizagem: 'Arquivo .png/.sql do DER.',
          entrega_estudante: 'Envio da modelagem no projeto.',
        },
        {
          numero: 2,
          sa_id: 'SA03',
          data: '2026-03-21',
          carga_horaria: 4,
          capacidades_tecnicas: ['Escrever scripts SQL de DDL e DML'],
          capacidades_socioemocionais: ['Foco nos resultados do cliente'],
          conhecimentos: ['Comandos SQL (CREATE, INSERT, SELECT, JOIN)'],
          desafio: 'Construir o banco de dados e popular com dados de teste.',
          resultado_esperado: 'Banco de dados criado e populado com scripts autônomos.',
          estrategias_ensino: 'Prática de consultas SQL em SGBD.',
          descricao_atividade: 'Execução de scripts de migração de banco de dados.',
          roteiro_temporal: '00:00-01:00 — DDL e DML;\n01:00-04:00 — Criação do SGBD.',
          instrumentos_avaliacao: 'Script SQL executável.',
          descricao_avaliacao: 'Verificação da integridade referencial do SGBD.',
          evidencias_aprendizagem: 'Script schema.sql testado.',
          entrega_estudante: 'Commit do script SQL no repositório.',
        },
        {
          numero: 3,
          sa_id: 'SA03',
          data: '2026-03-22',
          carga_horaria: 4,
          capacidades_tecnicas: ['Integrar aplicação web ao banco de dados relacional'],
          capacidades_socioemocionais: ['Visão holística de sistemas'],
          conhecimentos: ['Driver de conexão de BD', 'Operações CRUD via API'],
          desafio: 'Finalizar a integração completa da aplicação SaaS com persistência.',
          resultado_esperado: 'Aplicação realizando cadastro, leitura e edição no BD em tempo real.',
          estrategias_ensino: 'Desenvolvimento integrado de backend e banco.',
          descricao_atividade: 'Conexão da camada de controle ao banco de dados.',
          roteiro_temporal: '00:00-01:00 — Arquitetura de persistência;\n01:00-04:00 — Integração.',
          instrumentos_avaliacao: 'Demonstração ao vivo do sistema em funcionamento.',
          descricao_avaliacao: 'Avaliação da persistência real de dados.',
          evidencias_aprendizagem: 'Sistema SaaS completo operando.',
          entrega_estudante: 'Apresentação e entrega final do produto.',
        },
      ],
    },
  ],
};

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
  console.log('=== TESTE DE PREENCHIMENTO DA GRADE DE AULAS DAS SAs ===');
  console.log('Template:', TEMPLATE_PATH);

  if (!fs.existsSync(TEMPLATE_PATH)) {
    console.error('ERRO: Template original não encontrado em:', TEMPLATE_PATH);
    process.exit(1);
  }

  const templateBuffer = fs.readFileSync(TEMPLATE_PATH);

  console.log('\nExecutando geração com fixture de aulas...');
  const { buffer: generatedBuffer, saSummaryList } = generateWorkbookWithHeaders(
    templateBuffer,
    LESSONS_FIXTURE
  );

  const tmpDir = path.dirname(OUTPUT_PATH);
  if (!fs.existsSync(tmpDir)) {
    fs.mkdirSync(tmpDir, { recursive: true });
  }

  fs.writeFileSync(OUTPUT_PATH, generatedBuffer);
  console.log(`\nArquivo de teste gerado com sucesso em:\n  ${OUTPUT_PATH}`);

  const genZip = loadZipPackage(generatedBuffer);

  console.log('\n================ VALIDAÇÃO PROGRAMÁTICA DAS AULAS POR SA ================');

  saSummaryList.forEach((summary) => {
    console.log(`\n--- Planilha "${summary.saName}" ---`);
    console.log(`  Carga Horária Declarada: ${summary.declaredCh}h`);
    console.log(`  Carga Horária Calculada: ${summary.calculatedCh}h`);
    console.log(`  Quantidade de Aulas:     ${summary.aulasCount}`);
    console.log(`  Linhas Utilizadas:       ${summary.rowsUsed.join(', ')}`);
    console.log(`  Cache do \`=SUM(C13:C54)\`: ${summary.chFormulaCachedValue}`);

    const saPath = getWorksheetPathByName(genZip, summary.saName);
    const saXml = readZipFileText(genZip, saPath);

    // Check L6 formula
    const l6Cell = saXml.match(/<c [^>]*r="L6"[^>]*>[\s\S]*?<\/c>/);
    console.log(`  Célula L6 no XML:        ${l6Cell ? l6Cell[0] : '[Missing]'}`);

    // Spot-check each lesson row
    summary.rowsUsed.forEach((r, idx) => {
      console.log(`  -> Aula ${idx + 1} (Linha ${r}):`);
      console.log(`       Data (Col B):             ${extractCellTextOrVal(saXml, `B${r}`)}`);
      console.log(`       CH (Col C):               ${extractCellTextOrVal(saXml, `C${r}`)}`);
      console.log(`       Capacidades (Col D):      "${extractCellTextOrVal(saXml, `D${r}`).split('\n')[0]}..."`);
      console.log(`       Conhecimentos (Col G):    "${extractCellTextOrVal(saXml, `G${r}`).split('\n')[0]}..."`);
      console.log(`       Desafio (Col I):          "${extractCellTextOrVal(saXml, `I${r}`).slice(0, 30)}..."`);
      console.log(`       Resultado (Col J):        "${extractCellTextOrVal(saXml, `J${r}`).slice(0, 30)}..."`);
    });
  });

  // Package integrity check
  const report = compareWorkbookIntegrity(templateBuffer, generatedBuffer);
  console.log('\n================ RELATÓRIO DE INTEGRIDADE E PACOTE ================');
  console.log(`Entradas no ZIP original: ${report.origEntryCount}`);
  console.log(`Entradas no ZIP gerado:   ${report.genEntryCount}`);
  console.log(`vbaProject.bin preservado: ${report.hasVbaProject ? 'SIM' : 'NÃO'}`);
  console.log(`Arquivos VML preservados:  ${report.vmlFilesCount > 0 ? `SIM (${report.vmlFilesCount})` : 'NÃO'}`);
  console.log(`Desenhos preservados:     ${report.drawingsFilesCount > 0 ? `SIM (${report.drawingsFilesCount})` : 'NÃO'}`);
  console.log(`Mídias preservadas:       ${report.mediaFilesCount > 0 ? `SIM (${report.mediaFilesCount})` : 'NÃO'}`);
  console.log(`PrinterSettings:          ${report.printerSettingsCount > 0 ? `SIM (${report.printerSettingsCount})` : 'NÃO'}`);
  console.log(`Fórmulas antes: ${report.formulasBeforeCount}, Fórmulas depois: ${report.formulasAfterCount}`);
  console.log(`Fórmulas alteradas inadvertidamente: ${report.inadvertentlyModifiedFormulasCount}`);
}

main();

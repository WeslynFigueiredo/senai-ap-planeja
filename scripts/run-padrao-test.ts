import fs from 'fs';
import path from 'path';
import { generateWorkbookWithHeaders, type PlanoGeralData } from '../lib/excel/openxml.ts';
import { PADRAO_START_ROW, MAX_PADRAO_CRITERIA_CAPACITY, type CriterioDesempenhoData } from '../lib/excel/padrao-writer.ts';
import { loadZipPackage, readZipFileText } from '../lib/excel/zip-package.ts';
import { getWorksheetPathByName } from '../lib/excel/workbook-map.ts';
import { compareWorkbookIntegrity } from '../lib/excel/integrity-checker.ts';

const TEMPLATE_PATH = path.resolve('templates/PLANO_DE_ENSINO_MODELO.xlsm');
const OUTPUT_PATH = path.resolve('tmp/TESTE_PADRAO_DESEMPENHO.xlsm');

const OFFICIAL_CAPACITIES = [
  'Identificar requisitos de software e regras de negócio para aplicações web',
  'Estilizar interfaces digitais com CSS3 e frameworks responsivos',
  'Programar interatividade client-side com JavaScript moderno',
  'Identificar vulnerabilidades em requisições HTTP e rotas de API',
  'Modelar banco de dados relacional e aplicar rotinas SQL',
  'Trabalhar em equipe com comunicação assertiva e proatividade'
];

const VALID_CRITERIA_FIXTURE: CriterioDesempenhoData[] = [
  {
    numero: 1,
    capacidade_oficial: 'Identificar requisitos de software e regras de negócio para aplicações web',
    conhecimentos_pertinentes: ['Conceitos de HTML5 e semântica web', 'Arquitetura Client-Server'],
    criterio_desempenho: 'Documenta os requisitos de software conforme especificações do cliente.',
    sa_ids: ['SA01'],
    peso: 15,
    n0: 'Não identifica requisitos de software.',
    n1: 'Identifica parcialmente os requisitos de software com auxílio.',
    n2: 'Identifica os requisitos principais de software com pouca ajuda.',
    n3: 'Identifica autonomamente todos os requisitos de software.',
    n4: 'Identifica autonomamente e otimiza a documentação de requisitos de software.'
  },
  {
    numero: 2,
    capacidade_oficial: 'Estilizar interfaces digitais com CSS3 e frameworks responsivos',
    conhecimentos_pertinentes: ['CSS3 Flexbox e Grid Layout', 'Design Responsivo'],
    criterio_desempenho: 'Desenvolve estilos CSS3 adaptáveis para múltiplos dispositivos.',
    sa_ids: ['SA01', 'SA02'],
    peso: 20,
    n0: 'Não aplica estilos CSS.',
    n1: 'Aplica estilos CSS básicos sem suporte a dispositivos móveis.',
    n2: 'Aplica estilos responsivos com pequenos ajustes pendentes.',
    n3: 'Aplica estilos responsivos perfeitos em múltiplos layouts.',
    n4: 'Aplica estilos responsivos avançados com animações e otimização CSS.'
  },
  {
    numero: 3,
    capacidade_oficial: 'Programar interatividade client-side com JavaScript moderno',
    conhecimentos_pertinentes: ['Manipulação do DOM', 'Eventos e Validação de Formulários'],
    criterio_desempenho: 'Implementa scripts em JavaScript para validação dinâmica de dados no navegador.',
    sa_ids: ['SA01'],
    peso: 15,
    n0: 'Não desenvolve scripts em JavaScript.',
    n1: 'Desenvolve scripts com erros de sintaxe ou execução.',
    n2: 'Desenvolve scripts funcionais com pequenos desvios de validação.',
    n3: 'Desenvolve scripts robustos com tratamento de erros.',
    n4: 'Desenvolve scripts modulares, eficientes e desacoplados em JS.'
  },
  {
    numero: 4,
    capacidade_oficial: 'Identificar vulnerabilidades em requisições HTTP e rotas de API',
    conhecimentos_pertinentes: ['Segurança Web (OWASP Top 10)', 'Sanitização de Dados'],
    criterio_desempenho: 'Audita e corrige falhas de segurança em rotas e formulários web.',
    sa_ids: ['SA02'],
    peso: 15,
    n0: 'Não reconhece falhas de segurança.',
    n1: 'Reconhece falhas com suporte constante.',
    n2: 'Identifica falhas OWASP básicas e aplica correções simples.',
    n3: 'Identifica autonomamente falhas críticas e sanitiza dados.',
    n4: 'Realiza auditoria completa e implementa políticas de segurança avançadas.'
  },
  {
    numero: 5,
    capacidade_oficial: 'Modelar banco de dados relacional e aplicar rotinas SQL',
    conhecimentos_pertinentes: ['Modelagem MER e DER', 'Comandos SQL (DDL e DML)'],
    criterio_desempenho: 'Cria diagramas DER e executa scripts SQL para persistência de dados.',
    sa_ids: ['SA03'],
    peso: 20,
    n0: 'Não elabora modelagem de banco de dados.',
    n1: 'Elabora modelagem com inconsistências nas chaves.',
    n2: 'Elabora modelagem correta e scripts SQL básicos.',
    n3: 'Elabora modelagem normalizada (3FN) e rotinas SQL completas.',
    n4: 'Elabora modelagem otimizada, transações e índices SQL.'
  },
  {
    numero: 6,
    capacidade_oficial: 'Trabalhar em equipe com comunicação assertiva e proatividade',
    conhecimentos_pertinentes: ['Metodologias Ágeis (Scrum/Kanban)', 'Comunicação Interpessoal'],
    criterio_desempenho: 'Demonstra postura colaborativa e cumpre entregas nos prazos estipulados.',
    sa_ids: ['SA01', 'SA02', 'SA03'],
    peso: 15,
    n0: 'Não colabora com a equipe.',
    n1: 'Colabora apenas quando solicitado.',
    n2: 'Participa ativamente das rotinas da equipe.',
    n3: 'Demonstra proatividade e liderança nas entregas.',
    n4: 'Atua como facilitador, apoiando membros e superando metas.'
  }
];

const PADRAO_TEST_DATA: PlanoGeralData = {
  curso: 'Técnico em Desenvolvimento de Sistemas',
  cargaHorariaCurso: 1200,
  modalidade: 'Presencial',
  unidadeCurricular: 'Desenvolvimento Web Front-End',
  cargaHorariaUc: 80,
  modulo: 'Módulo II',
  capacidadesTecnicasOficiais: OFFICIAL_CAPACITIES,
  situacoesAprendizagem: [
    {
      numero: 1,
      nome: 'SA01',
      carga_horaria: 12,
      estrategia: 'Estudo de Caso com Metodologia Ágil',
      contextualizacao: 'Desenvolvimento de uma plataforma web responsiva para e-commerce.',
      qtdCapacidadesTecnicas: 3,
      qtdCapacidadesSocioemocionais: 1,
      aulas: [
        { sa_id: 'SA01', numero: 1, data: '2026-03-15', carga_horaria: 4, capacidades_tecnicas: ['Identificar requisitos de software'], conhecimentos: ['HTML5'], desafio: 'Protótipo', resultado_esperado: 'Protótipo funcional', estrategias_ensino: ['Aula prática'], descricao_atividade: ['Codificação'], roteiro_temporal: ['00:00–04:00 — Lab'], instrumentos_avaliacao: ['Observação'], descricao_avaliacao: ['Checklist'], evidencias_aprendizagem: ['Código fonte'], entrega_estudante: ['Repositório Git'] },
        { sa_id: 'SA01', numero: 2, data: '2026-03-16', carga_horaria: 4, capacidades_tecnicas: ['Estilizar interfaces'], conhecimentos: ['CSS3'], desafio: 'Estilização', resultado_esperado: 'Layout responsivo', estrategias_ensino: ['Demonstração'], descricao_atividade: ['CSS Grid'], roteiro_temporal: ['00:00–04:00 — CSS'], instrumentos_avaliacao: ['Ficha de avaliação'], descricao_avaliacao: ['Rubrica'], evidencias_aprendizagem: ['CSS'], entrega_estudante: ['Arquivo CSS'] },
        { sa_id: 'SA01', numero: 3, data: '2026-03-17', carga_horaria: 4, capacidades_tecnicas: ['Programar interatividade'], conhecimentos: ['JS DOM'], desafio: 'Validação', resultado_esperado: 'Formulário dinâmico', estrategias_ensino: ['Coding Session'], descricao_atividade: ['Event Listeners'], roteiro_temporal: ['00:00–04:00 — JS'], instrumentos_avaliacao: ['Teste prático'], descricao_avaliacao: ['Execução'], evidencias_aprendizagem: ['Scripts JS'], entrega_estudante: ['JS App'] }
      ]
    },
    {
      numero: 2,
      nome: 'SA02',
      carga_horaria: 8,
      estrategia: 'Resolução de Problemas de Segurança Web',
      contextualizacao: 'Auditoria de código e correção de vulnerabilidades em API.',
      qtdCapacidadesTecnicas: 2,
      qtdCapacidadesSocioemocionais: 1,
      aulas: [
        { sa_id: 'SA02', numero: 1, data: '2026-03-18', carga_horaria: 4, capacidades_tecnicas: ['Identificar vulnerabilidades'], conhecimentos: ['OWASP'], desafio: 'Auditoria', resultado_esperado: 'Relatório de segurança', estrategias_ensino: ['Estudo dirigido'], descricao_atividade: ['Análise de código'], roteiro_temporal: ['00:00–04:00 — Sec'], instrumentos_avaliacao: ['Relatório'], descricao_avaliacao: ['Análise técnica'], evidencias_aprendizagem: ['Documento PDF'], entrega_estudante: ['Relatório de Vulnerabilidades'] },
        { sa_id: 'SA02', numero: 2, data: '2026-03-19', carga_horaria: 4, capacidades_tecnicas: ['Refatorar rotas'], conhecimentos: ['Sanitização'], desafio: 'Sanitização', resultado_esperado: 'API protegida', estrategias_ensino: ['Lab de Segurança'], descricao_atividade: ['Fixing OWASP'], roteiro_temporal: ['00:00–04:00 — Refactor'], instrumentos_avaliacao: ['Validação de código'], descricao_avaliacao: ['Pentest básico'], evidencias_aprendizagem: ['Código corrigido'], entrega_estudante: ['Patch de Segurança'] }
      ]
    },
    {
      numero: 3,
      nome: 'SA03',
      carga_horaria: 12,
      estrategia: 'Projeto Integrador de Banco de Dados',
      contextualizacao: 'Criação da camada de dados de um sistema de gestão.',
      qtdCapacidadesTecnicas: 2,
      qtdCapacidadesSocioemocionais: 1,
      aulas: [
        { sa_id: 'SA03', numero: 1, data: '2026-03-20', carga_horaria: 4, capacidades_tecnicas: ['Modelar banco de dados'], conhecimentos: ['MER/DER'], desafio: 'Modelagem DER', resultado_esperado: 'Diagrama validado', estrategias_ensino: ['Modelagem em dupla'], descricao_atividade: ['Diagramação'], roteiro_temporal: ['00:00–04:00 — DER'], instrumentos_avaliacao: ['Diagrama'], descricao_avaliacao: ['Normalização'], evidencias_aprendizagem: ['DER em PNG'], entrega_estudante: ['Modelo DER'] },
        { sa_id: 'SA03', numero: 2, data: '2026-03-21', carga_horaria: 4, capacidades_tecnicas: ['Escrever scripts SQL'], conhecimentos: ['DDL/DML'], desafio: 'Scripts SQL', resultado_esperado: 'BD populado', estrategias_ensino: ['Lab SQL'], descricao_atividade: ['DDL/DML'], roteiro_temporal: ['00:00–04:00 — SQL'], instrumentos_avaliacao: ['Execução de Script'], descricao_avaliacao: ['Queries SQL'], evidencias_aprendizagem: ['Script .sql'], entrega_estudante: ['Banco .sql'] },
        { sa_id: 'SA03', numero: 3, data: '2026-03-22', carga_horaria: 4, capacidades_tecnicas: ['Integrar aplicação web'], conhecimentos: ['Driver BD'], desafio: 'Conexão DB-Web', resultado_esperado: 'Sistema integrado', estrategias_ensino: ['Integração prática'], descricao_atividade: ['Conexão BD'], roteiro_temporal: ['00:00–04:00 — Integration'], instrumentos_avaliacao: ['Demonstração ao vivo'], descricao_avaliacao: ['CRUD completo'], evidencias_aprendizagem: ['Sistema rodando'], entrega_estudante: ['Aplicação Final'] }
      ]
    }
  ],
  padraoDesempenho: {
    criterios: VALID_CRITERIA_FIXTURE
  }
};

function runPositiveTest() {
  console.log('\n================ 1. TESTE POSITIVO: GERANDO PADRÃO DE DESEMPENHO ================');
  const templateBuffer = fs.readFileSync(TEMPLATE_PATH);
  const { buffer: genBuffer } = generateWorkbookWithHeaders(templateBuffer, PADRAO_TEST_DATA);

  fs.writeFileSync(OUTPUT_PATH, genBuffer);
  console.log(`Workbook de teste salvo em:\n  ${OUTPUT_PATH}`);

  const genZip = loadZipPackage(genBuffer);
  const padraoPath = getWorksheetPathByName(genZip, 'Padrão de Desempenho');
  const padraoXml = readZipFileText(genZip, padraoPath);

  console.log('\n--- VALIDAÇÃO PROGRAMÁTICA DOS CRITÉRIOS EM PADRÃO DE DESEMPENHO ---');
  VALID_CRITERIA_FIXTURE.forEach((crit, idx) => {
    const row = PADRAO_START_ROW + idx;
    const rMatch = padraoXml.match(new RegExp(`<row [^>]*r="${row}"[\\s\\S]*?<\\/row>`));
    if (!rMatch) {
      console.error(`ERRO: Linha ${row} não encontrada no XML!`);
      return;
    }
    const rXml = rMatch[0];

    const getVal = (col: string) => {
      const cMatch = rXml.match(new RegExp(`<c [^>]*r="${col}${row}"[^>]*>(?:<v>([\\s\\S]*?)<\\/v>|<is><t[^>]*?>([\\s\\S]*?)<\\/t><\\/is>)`));
      if (!cMatch) return '[VAZIO]';
      return cMatch[1] || cMatch[2] || '[VAZIO]';
    };

    console.log(`\nCritério ${crit.numero} (Linha ${row}):`);
    console.log(`  Col C (Nº):           ${getVal('C')}`);
    console.log(`  Col D (Capacidade):   "${getVal('D').slice(0, 45)}..."`);
    console.log(`  Col E (Conhecimentos):"${getVal('E').replace(/\n/g, ' | ')}"`);
    console.log(`  Col F (Critério):     "${getVal('F').slice(0, 45)}..."`);
    console.log(`  Col G (SA Vinculada): "${getVal('G').replace(/\n/g, ', ')}"`);
    console.log(`  Col H (N0):           "${getVal('H').slice(0, 30)}..."`);
    console.log(`  Col I (N1):           "${getVal('I').slice(0, 30)}..."`);
    console.log(`  Col J (N2):           "${getVal('J').slice(0, 30)}..."`);
    console.log(`  Col K (N3):           "${getVal('K').slice(0, 30)}..."`);
    console.log(`  Col L (N4):           "${getVal('L').slice(0, 30)}..."`);
    console.log(`  Col M (Peso armazenado): ${getVal('M')} (${Number(getVal('M')) * 100}%)`);
  });

  // Verify unused slots 7 to 75 are empty in columns C..M
  console.log('\n--- VALIDAÇÃO DE SLOTS EXCEDENTES / LIMPEZA PRÉVIA ---');
  let unusedSlotsClean = true;
  for (let idx = VALID_CRITERIA_FIXTURE.length; idx < MAX_PADRAO_CRITERIA_CAPACITY; idx++) {
    const r = PADRAO_START_ROW + idx;
    const rMatch = padraoXml.match(new RegExp(`<row [^>]*r="${r}"[\\s\\S]*?<\\/row>`));
    if (rMatch) {
      const cMatch = rMatch[0].match(/<c [^>]*r="[C-M]\d+"[^>]*>(?:<v>[\s\S]*?<\/v>|<is>[\s\S]*?<\/is>)/);
      if (cMatch) {
        console.error(`ERRO: Slot na linha ${r} contém valor residual: ${cMatch[0]}`);
        unusedSlotsClean = false;
      }
    }
  }
  if (unusedSlotsClean) {
    console.log(`✅ Todos os slots não utilizados (${VALID_CRITERIA_FIXTURE.length + 1} até ${MAX_PADRAO_CRITERIA_CAPACITY}, linhas ${PADRAO_START_ROW + VALID_CRITERIA_FIXTURE.length} a ${PADRAO_START_ROW + MAX_PADRAO_CRITERIA_CAPACITY - 1}) estão perfeitamente limpos.`);
  }

  // Integrity report
  const report = compareWorkbookIntegrity(templateBuffer, genBuffer);
  console.log('\n================ RELATÓRIO DE INTEGRIDADE DO PACOTE ================');
  console.log(`vbaProject.bin preservado: ${report.hasVbaProject ? 'SIM' : 'NÃO'}`);
  console.log(`Arquivos VML preservados:  ${report.vmlFilesCount > 0 ? `SIM (${report.vmlFilesCount})` : 'NÃO'}`);
  console.log(`Desenhos preservados:     ${report.drawingsFilesCount > 0 ? `SIM (${report.drawingsFilesCount})` : 'NÃO'}`);
  console.log(`Mídias preservadas:       ${report.mediaFilesCount > 0 ? `SIM (${report.mediaFilesCount})` : 'NÃO'}`);
  console.log(`PrinterSettings:          ${report.printerSettingsCount > 0 ? `SIM (${report.printerSettingsCount})` : 'NÃO'}`);
  console.log(`Fórmulas alteradas inadvertidamente: ${report.inadvertentlyModifiedFormulasCount}`);
}

function runNegativeTests() {
  console.log('\n================ 2. EXECUÇÃO DE TESTES NEGATIVOS (BLOQUEIOS) ================');
  const templateBuffer = fs.readFileSync(TEMPLATE_PATH);

  // Test Negativo 1: Peso Total != 100%
  console.log('\n[Teste Negativo 1] Soma dos pesos = 95%');
  const invalidPesoData: PlanoGeralData = JSON.parse(JSON.stringify(PADRAO_TEST_DATA));
  invalidPesoData.padraoDesempenho!.criterios[0].peso = 10; // Total 95%
  try {
    generateWorkbookWithHeaders(templateBuffer, invalidPesoData);
    console.error('❌ ERRO: A geração deveria ter sido bloqueada por soma de peso incorreta!');
  } catch (err: any) {
    console.log('✅ BLOQUEADO COM SUCESSO. Mensagem capturada:\n  ', err.message);
  }

  // Test Negativo 2: Capacidade sem critério
  console.log('\n[Teste Negativo 2] Capacidade oficial não representada nos critérios');
  const missingCapData: PlanoGeralData = JSON.parse(JSON.stringify(PADRAO_TEST_DATA));
  missingCapData.capacidadesTecnicasOficiais!.push('Configurar servidores Nginx e SSL');
  try {
    generateWorkbookWithHeaders(templateBuffer, missingCapData);
    console.error('❌ ERRO: A geração deveria ter sido bloqueada por falta de cobertura de capacidade!');
  } catch (err: any) {
    console.log('✅ BLOQUEADO COM SUCESSO. Mensagem capturada:\n  ', err.message);
  }

  // Test Negativo 3: SA Inexistente
  console.log('\n[Teste Negativo 3] Critério aponta para SA07 (inexistente)');
  const invalidSaData: PlanoGeralData = JSON.parse(JSON.stringify(PADRAO_TEST_DATA));
  invalidSaData.padraoDesempenho!.criterios[0].sa_ids = ['SA07'];
  try {
    generateWorkbookWithHeaders(templateBuffer, invalidSaData);
    console.error('❌ ERRO: A geração deveria ter sido bloqueada por SA inexistente!');
  } catch (err: any) {
    console.log('✅ BLOQUEADO COM SUCESSO. Mensagem capturada:\n  ', err.message);
  }

  // Test Negativo 4: Mais critérios do que o limite da matriz (76 critérios)
  console.log('\n[Teste Negativo 4] 76 critérios (limite físico = 75)');
  const overflowData: PlanoGeralData = JSON.parse(JSON.stringify(PADRAO_TEST_DATA));
  const manyCriteria: CriterioDesempenhoData[] = [];
  for (let i = 1; i <= 76; i++) {
    manyCriteria.push({
      numero: i,
      capacidade_oficial: OFFICIAL_CAPACITIES[0],
      conhecimentos_pertinentes: ['Test'],
      criterio_desempenho: `Critério ${i}`,
      sa_ids: ['SA01'],
      peso: 100 / 76,
      n0: 'N0', n1: 'N1', n2: 'N2', n3: 'N3', n4: 'N4'
    });
  }
  overflowData.padraoDesempenho!.criterios = manyCriteria;
  try {
    generateWorkbookWithHeaders(templateBuffer, overflowData);
    console.error('❌ ERRO: A geração deveria ter sido bloqueada por excesso de critérios!');
  } catch (err: any) {
    console.log('✅ BLOQUEADO COM SUCESSO. Mensagem capturada:\n  ', err.message);
  }
}

function main() {
  runPositiveTest();
  runNegativeTests();
  console.log('\n================ TODOS OS TESTES CONCLUÍDOS COM SUCESSO ================');
}

main();

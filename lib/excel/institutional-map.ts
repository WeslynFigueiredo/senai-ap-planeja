/**
 * Institutional Cell Mapping
 * Maps SENAI Plano de Ensino, Situacoes de Aprendizagem (SA), and Padrao de Desempenho
 * fields to specific worksheet names and cell references.
 */

export interface CellMapping {
  sheetName: string;
  cellRef: string;
  label: string;
}

export const PLANO_DE_ENSINO_CELL_MAP = {
  curso: { sheetName: 'Plano de Ensino', cellRef: 'C4', label: 'Curso' },
  cargaHoraria: { sheetName: 'Plano de Ensino', cellRef: 'H4', label: 'Carga Horária do Curso' },
  modalidade: { sheetName: 'Plano de Ensino', cellRef: 'J4', label: 'Modalidade' },
  unidadeCurricular: { sheetName: 'Plano de Ensino', cellRef: 'D5', label: 'Unidade Curricular' },
  cargaHorariaUc: { sheetName: 'Plano de Ensino', cellRef: 'H5', label: 'Carga Horária da UC' },
  modulo: { sheetName: 'Plano de Ensino', cellRef: 'J5', label: 'Módulo' },
  dataInicio: { sheetName: 'Plano de Ensino', cellRef: 'D6', label: 'Data Início da UC' },
  dataFim: { sheetName: 'Plano de Ensino', cellRef: 'F6', label: 'Data Término da UC' },
  turma: { sheetName: 'Plano de Ensino', cellRef: 'H6', label: 'Turma' },
  objetivoUc: { sheetName: 'Plano de Ensino', cellRef: 'D7', label: 'Objetivo da UC' },
  ambientePedagogico: { sheetName: 'Plano de Ensino', cellRef: 'J7', label: 'Ambiente Pedagógico' },
  situacoesCount: { sheetName: 'Plano de Ensino', cellRef: 'D8', label: 'No. Situações de Aprendizagem' },
  aulasCount: { sheetName: 'Plano de Ensino', cellRef: 'G8', label: 'Número de Aulas' },
  capacidadesTecnicasCount: { sheetName: 'Plano de Ensino', cellRef: 'I8', label: 'Nº Total de Capacidades Básicas / Técnicas' },
  capacidadesSocioemocionaisCount: { sheetName: 'Plano de Ensino', cellRef: 'N8', label: 'Nº Total de Capacidades Socioemocionais' },
  capacidadesTecnicasList: { sheetName: 'Plano de Ensino', cellRef: 'B12', label: 'Capacidades Básicas / Técnicas' },
  conhecimentosList: { sheetName: 'Plano de Ensino', cellRef: 'I12', label: 'Conhecimentos' },
  capacidadesSocioemocionaisList: { sheetName: 'Plano de Ensino', cellRef: 'B45', label: 'Capacidades Socioemocionais' },
} as const;

export const SA_CELL_MAP = {
  header: {
    curso: 'E4',               // Formula: 'Plano de Ensino'!C4
    cargaHorariaCurso: 'I4',   // Formula: 'Plano de Ensino'!H4
    modalidade: 'L4',          // Formula: 'Plano de Ensino'!J4
    unidadeCurricular: 'E5',   // Formula: 'Plano de Ensino'!D5
    cargaHorariaUc: 'I5',      // Formula: 'Plano de Ensino'!H5
    modulo: 'L5',              // Formula: 'Plano de Ensino'!J5
  },
  specific: {
    estrategiaRegion: 'E6',
    contextualizacao: 'B10',
    capacidadesTecnicasCount: 'I10',
    capacidadesSocioemocionaisCount: 'K10',
  },
} as const;

export const PADRAO_DE_DESEMPENHO_CELL_MAP = {
  generalHeaderFields: [],
} as const;

/**
 * Official Institutional Strategies List (sourced from Base!$A$2:$A$13)
 */
export const INSTITUTIONAL_STRATEGIES = [
  'Exposição Dialogada',
  'Atividade Prática',
  'Trabalho em Grupo',
  'Dinâmica de Grupo',
  'Visita Técnica',
  'Ensaio Tecnológico',
  'Workshop',
  'Seminário',
  'Painel Temático',
  'Gameficação',
  'Sala de Aula Invertida',
  'Design Thinking',
] as const;

/**
 * Official Institutional Instruments List (sourced from Base!$B$2:$B$9)
 */
export const INSTITUTIONAL_INSTRUMENTS = [
  'Fichas de Observação',
  'Relatórios',
  'Portfólios',
  'Provas Objetivas',
  'Provas de Respostas Construídas',
  'Provas Práticas',
  'Autoavaliações',
] as const;

/**
 * Normalizes an educational strategy string to the institutional equivalent when unambiguous.
 */
export function normalizeStrategyName(rawName: string): { normalized: string; isDirectMatch: boolean } {
  const trimmed = rawName.trim();
  const lower = trimmed.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  if (lower === 'exposicao dialogada') return { normalized: 'Exposição Dialogada', isDirectMatch: true };
  if (lower === 'atividade pratica') return { normalized: 'Atividade Prática', isDirectMatch: true };
  if (lower === 'trabalho em grupo') return { normalized: 'Trabalho em Grupo', isDirectMatch: true };
  if (lower === 'dinamica de grupo') return { normalized: 'Dinâmica de Grupo', isDirectMatch: true };
  if (lower === 'visita tecnica') return { normalized: 'Visita Técnica', isDirectMatch: true };
  if (lower === 'ensaio tecnologico') return { normalized: 'Ensaio Tecnológico', isDirectMatch: true };
  if (lower === 'workshop') return { normalized: 'Workshop', isDirectMatch: true };
  if (lower === 'seminario' || lower === 'seminario/apresentacao tecnica') return { normalized: 'Seminário', isDirectMatch: true };
  if (lower === 'painel tematico') return { normalized: 'Painel Temático', isDirectMatch: true };
  if (lower === 'gameficacao' || lower === 'gamificacao') return { normalized: 'Gameficação', isDirectMatch: true };
  if (lower === 'sala de aula invertida') return { normalized: 'Sala de Aula Invertida', isDirectMatch: true };
  if (lower === 'design thinking') return { normalized: 'Design Thinking', isDirectMatch: true };

  return { normalized: trimmed, isDirectMatch: false };
}

/**
 * Normalizes an assessment instrument string to the institutional equivalent when unambiguous.
 */
export function normalizeInstrumentName(rawName: string): { normalized: string; isDirectMatch: boolean } {
  const trimmed = rawName.trim();
  const lower = trimmed.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');

  if (lower === 'ficha de observacao' || lower === 'fichas de observacao' || lower === 'ficha de observacao direta') {
    return { normalized: 'Fichas de Observação', isDirectMatch: true };
  }
  if (lower === 'relatorio' || lower === 'relatorios' || lower === 'relatorio tecnico') {
    return { normalized: 'Relatórios', isDirectMatch: true };
  }
  if (lower === 'portfolio' || lower === 'portfolios') {
    return { normalized: 'Portfólios', isDirectMatch: true };
  }
  if (lower === 'prova objetiva' || lower === 'provas objetivas' || lower === 'prova objetiva (simulado)' || lower === 'prova objetiva (simulado geral)') {
    return { normalized: 'Provas Objetivas', isDirectMatch: true };
  }
  if (lower === 'prova de resposta construida' || lower === 'provas de respostas construidas') {
    return { normalized: 'Provas de Respostas Construídas', isDirectMatch: true };
  }
  if (lower === 'prova pratica' || lower === 'provas praticas') {
    return { normalized: 'Provas Práticas', isDirectMatch: true };
  }
  if (lower === 'autoavaliacao' || lower === 'autoavaliacoes') {
    return { normalized: 'Autoavaliações', isDirectMatch: true };
  }

  return { normalized: trimmed, isDirectMatch: false };
}

/**
 * Formats line 6 SA Strategy options text, marking only the main strategy with ( X ).
 */
export function formatSaStrategy(strategyName: string): string {
  const norm = String(strategyName)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, '');

  let sp = '   ';
  let ec = '   ';
  let pa = '  ';
  let pr = '   ';
  let pi = '   ';

  if (norm.includes('integrador')) {
    pi = ' X ';
  } else if (norm.includes('situacao') || norm.includes('problema')) {
    sp = ' X ';
  } else if (norm.includes('estudo') || norm.includes('caso')) {
    ec = ' X ';
  } else if (norm.includes('pesquisa') || norm.includes('aplicada')) {
    pa = ' X ';
  } else if (norm.includes('projeto')) {
    pr = ' X ';
  } else {
    sp = ' X ';
  }

  return `(${sp}) Situação Problema  (${ec}) Estudo de Caso    (${pa}) Pesquisa Aplicada   (${pr}) Projeto  (${pi}) Projeto Integrador`;
}


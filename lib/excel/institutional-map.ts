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
  curso: {
    sheetName: 'Plano de Ensino',
    cellRef: 'C4',
    label: 'Curso',
  },
  cargaHoraria: {
    sheetName: 'Plano de Ensino',
    cellRef: 'H4',
    label: 'Carga Horária do Curso',
  },
  unidadeCurricular: {
    sheetName: 'Plano de Ensino',
    cellRef: 'D5',
    label: 'Unidade Curricular',
  },
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

/**
 * Padrão de Desempenho Header Map
 * Empirically verified against sheet9.xml: no general identification header fields
 * (such as Curso, Unidade Curricular, Carga Horária, Módulo) exist in the template
 * header of this sheet (it starts directly at Row 1 with evaluation matrix).
 */
export const PADRAO_DE_DESEMPENHO_CELL_MAP = {
  generalHeaderFields: [],
} as const;

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

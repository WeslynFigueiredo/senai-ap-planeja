import { setCellText, setCellNumber, setFormulaCellCachedValue } from './cell-writer.ts';

export interface AulaData {
  numero: number;
  sa_id: string; // e.g. "SA01" or "1"
  data?: string;
  carga_horaria: number;
  capacidades_tecnicas?: string[] | string;
  capacidades_socioemocionais?: string[] | string;
  conhecimentos?: string[] | string;
  desafio?: string;
  resultado_esperado?: string;
  estrategias_ensino?: string;
  descricao_atividade?: string;
  roteiro_temporal?: string;
  instrumentos_avaliacao?: string;
  descricao_avaliacao?: string;
  evidencias_aprendizagem?: string;
  entrega_estudante?: string;
}

export function dateToExcelSerial(dateStr: string): number | null {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) return null;
  const parts = dateStr.split('-').map(Number);
  const year = parts[0];
  const month = parts[1] - 1;
  const day = parts[2];
  const d = new Date(Date.UTC(year, month, day));
  const epoch = Date.UTC(1899, 11, 30);
  return Math.floor((d.getTime() - epoch) / (24 * 60 * 60 * 1000));
}

export function clearLessonsGrid(sheetXml: string): string {
  return sheetXml.replace(
    /<c\s+([^>]*?\br="([A-Z]+)(\d+)"[^>]*?)>(?:[\s\S]*?)<\/c>/g,
    (match, attrStr, col, rowStr) => {
      const row = parseInt(rowStr, 10);
      if (row >= 13 && row <= 54) {
        const cleanAttrStr = attrStr.replace(/\s*t="[^"]*"/g, '').trim();
        return `<c ${cleanAttrStr}/>`;
      }
      return match;
    }
  );
}

/**
 * Populates the lessons grid (rows 13 to 54) for a specific SA sheet.
 * Validates declared CH vs sum of lesson CHs and template slot capacity (max 21 slots).
 */
export function populateSaLessons(
  sheetXml: string,
  saName: string,
  declaredSaCh: number,
  aulas: AulaData[]
): { updatedXml: string; aulasCount: number; calculatedCh: number; rowsUsed: number[] } {
  const sortedAulas = [...aulas].sort((a, b) => a.numero - b.numero);

  // 1. Capacity check: Max 21 lesson slots (rows 13 to 54 in 2-row blocks)
  if (sortedAulas.length > 21) {
    throw new Error(`A ${saName} possui mais aulas (${sortedAulas.length}) do que a capacidade atual do modelo institucional (21 aulas).`);
  }

  // 2. Carga Horária validation check
  const calculatedCh = sortedAulas.reduce((acc, a) => acc + Number(a.carga_horaria || 0), 0);
  if (calculatedCh !== declaredSaCh) {
    throw new Error(`${saName}: carga declarada = ${declaredSaCh}h; carga das aulas = ${calculatedCh}h`);
  }

  // 3. Clear previous variable values in rows 13 to 54
  let updatedXml = clearLessonsGrid(sheetXml);
  const rowsUsed: number[] = [];

  // 4. Populate lessons into grid rows (top-left row = 13 + slotIndex * 2)
  sortedAulas.forEach((aula, slotIdx) => {
    const r = 13 + slotIdx * 2;
    rowsUsed.push(r);

    // Col B: Data
    if (aula.data) {
      const serial = dateToExcelSerial(aula.data);
      if (serial !== null) {
        updatedXml = setCellNumber(updatedXml, `B${r}`, serial);
      } else {
        updatedXml = setCellText(updatedXml, `B${r}`, aula.data);
      }
    }

    // Col C: CH
    updatedXml = setCellNumber(updatedXml, `C${r}`, Number(aula.carga_horaria || 0));

    // Col D: Capacidades
    const tec = Array.isArray(aula.capacidades_tecnicas)
      ? aula.capacidades_tecnicas.join('\n')
      : (aula.capacidades_tecnicas || '');
    const soc = Array.isArray(aula.capacidades_socioemocionais)
      ? aula.capacidades_socioemocionais.join('\n')
      : (aula.capacidades_socioemocionais || '');
    const capText = [tec, soc].filter(Boolean).join('\n');
    if (capText) {
      updatedXml = setCellText(updatedXml, `D${r}`, capText);
    }

    // Col G: Conhecimentos
    const conText = Array.isArray(aula.conhecimentos)
      ? aula.conhecimentos.join('\n')
      : (aula.conhecimentos || '');
    if (conText) {
      updatedXml = setCellText(updatedXml, `G${r}`, conText);
    }

    // Col I: Desafio
    if (aula.desafio) {
      updatedXml = setCellText(updatedXml, `I${r}`, aula.desafio);
    }

    // Col J: Resultado Esperado
    if (aula.resultado_esperado) {
      updatedXml = setCellText(updatedXml, `J${r}`, aula.resultado_esperado);
    }

    // Col K: Estratégias e Atividade
    const kParts = [];
    if (aula.estrategias_ensino) kParts.push(`Estratégias de Ensino:\n${aula.estrategias_ensino}`);
    if (aula.descricao_atividade) kParts.push(`Descrição da Atividade:\n${aula.descricao_atividade}`);
    if (aula.roteiro_temporal) kParts.push(`Roteiro Temporal:\n${aula.roteiro_temporal}`);
    const kText = kParts.join('\n\n');
    if (kText) {
      updatedXml = setCellText(updatedXml, `K${r}`, kText);
    }

    // Col L: Avaliação
    const lParts = [];
    if (aula.instrumentos_avaliacao) lParts.push(`Instrumentos:\n${aula.instrumentos_avaliacao}`);
    if (aula.descricao_avaliacao) lParts.push(`Descrição da Avaliação:\n${aula.descricao_avaliacao}`);
    if (aula.evidencias_aprendizagem) lParts.push(`Evidências:\n${aula.evidencias_aprendizagem}`);
    if (aula.entrega_estudante) lParts.push(`Entrega do Estudante:\n${aula.entrega_estudante}`);
    const lText = lParts.join('\n\n');
    if (lText) {
      updatedXml = setCellText(updatedXml, `L${r}`, lText);
    }
  });

  // 5. Update cached value of total CH formula cell (L6)
  updatedXml = setFormulaCellCachedValue(updatedXml, 'L6', 'number', calculatedCh);

  return {
    updatedXml,
    aulasCount: sortedAulas.length,
    calculatedCh,
    rowsUsed,
  };
}

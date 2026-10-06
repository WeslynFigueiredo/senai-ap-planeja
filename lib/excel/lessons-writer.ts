import { setCellText, setCellNumber, setFormulaCellCachedValue } from './cell-writer.ts';
import { normalizeStrategyName, normalizeInstrumentName } from './institutional-map.ts';

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
  estrategias_ensino?: string[] | string;
  descricao_atividade?: string[] | string;
  roteiro_temporal?: any;
  instrumentos_avaliacao?: string[] | string;
  descricao_avaliacao?: string[] | string;
  evidencias_aprendizagem?: string[] | string;
  entrega_estudante?: string[] | string;
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

    // Col K: Estratégias de Ensino, Descrição da Atividade e Roteiro da Aula
    const kParts = [];
    if (aula.estrategias_ensino) {
      const estList = Array.isArray(aula.estrategias_ensino)
        ? aula.estrategias_ensino
        : [aula.estrategias_ensino];
      const estStr = estList.map(e => normalizeStrategyName(String(e)).normalized).join(' | ');
      kParts.push(`Estratégias de Ensino:\n${estStr}`);
    }
    if (aula.descricao_atividade) {
      const actStr = Array.isArray(aula.descricao_atividade)
        ? aula.descricao_atividade.join('\n')
        : String(aula.descricao_atividade);
      kParts.push(`Descrição da atividade:\n${actStr}`);
    }
    if (aula.roteiro_temporal) {
      let rotStr = '';
      if (Array.isArray(aula.roteiro_temporal)) {
        rotStr = aula.roteiro_temporal.map((item: any) => {
          if (typeof item === 'object' && item !== null) {
            const inicio = item.inicio || '';
            const fim = item.fim || '';
            const desc = item.descricao || '';
            const range = inicio && fim ? `${inicio}–${fim}` : (inicio || fim);
            return range ? `${range} — ${desc}` : desc;
          }
          return String(item);
        }).join('\n');
      } else {
        rotStr = String(aula.roteiro_temporal);
      }
      kParts.push(`Roteiro da aula:\n${rotStr}`);
    }
    const kText = kParts.join('\n\n');
    if (kText) {
      updatedXml = setCellText(updatedXml, `K${r}`, kText);
    }

    // Col L: Instrumentos de Avaliação da Aprendizagem, Avaliação, Evidências e Entrega
    const lParts = [];
    if (aula.instrumentos_avaliacao) {
      const instList = Array.isArray(aula.instrumentos_avaliacao)
        ? aula.instrumentos_avaliacao
        : [aula.instrumentos_avaliacao];
      const instStr = instList.map(i => normalizeInstrumentName(String(i)).normalized).join(', ');
      lParts.push(`Instrumentos:\n${instStr}`);
    }
    if (aula.descricao_avaliacao) {
      const descAvStr = Array.isArray(aula.descricao_avaliacao)
        ? aula.descricao_avaliacao.join('\n')
        : String(aula.descricao_avaliacao);
      lParts.push(`Avaliação:\n${descAvStr}`);
    }
    if (aula.evidencias_aprendizagem) {
      let evidItems: string[] = [];
      if (Array.isArray(aula.evidencias_aprendizagem)) {
        evidItems = aula.evidencias_aprendizagem.map(e => String(e).trim());
      } else {
        evidItems = String(aula.evidencias_aprendizagem).split('\n').map(e => e.trim()).filter(Boolean);
      }
      const formattedEvid = evidItems
        .map(e => e.startsWith('•') ? e : `• ${e}`)
        .join('\n');
      lParts.push(`Evidências:\n${formattedEvid}`);
    }
    if (aula.entrega_estudante) {
      const entStr = Array.isArray(aula.entrega_estudante)
        ? aula.entrega_estudante.join('\n')
        : String(aula.entrega_estudante);
      lParts.push(`Entrega:\n${entStr}`);
    }
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

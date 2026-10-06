import { setCellText, setCellNumber, setFormulaCellCachedValue, setRowHeight } from './cell-writer.ts';
import { normalizeStrategyName, normalizeInstrumentName } from './institutional-map.ts';

function countVisualLines(text: string, colCharWidth: number): number {
  if (!text) return 0;
  const lines = text.split('\n');
  let count = 0;
  for (const line of lines) {
    if (line.length === 0) {
      count += 1;
    } else {
      count += Math.ceil(line.length / colCharWidth);
    }
  }
  return count;
}

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

    // Col K: Estratégias de Ensino (Dropdown na linha r; Detalhamento na linha r+1)
    const estList = aula.estrategias_ensino
      ? (Array.isArray(aula.estrategias_ensino) ? aula.estrategias_ensino : [aula.estrategias_ensino])
      : [];
    const normalizedEstList = estList
      .map(e => normalizeStrategyName(String(e)).normalized)
      .filter(Boolean);

    // K{r}: Valor único para manter a validação do dropdown institucional
    const mainStrategy = normalizedEstList[0] || '';
    if (mainStrategy) {
      updatedXml = setCellText(updatedXml, `K${r}`, mainStrategy);
    }

    // K{r+1}: Bloco detalhado da aula (Estratégias utilizadas, Descrição e Roteiro)
    const kDetailParts = [];
    if (normalizedEstList.length > 0) {
      kDetailParts.push(`Estratégias utilizadas:\n${normalizedEstList.join(' | ')}`);
    }
    if (aula.descricao_atividade) {
      const actStr = Array.isArray(aula.descricao_atividade)
        ? aula.descricao_atividade.join('\n')
        : String(aula.descricao_atividade);
      kDetailParts.push(`Descrição da atividade:\n${actStr}`);
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
      kDetailParts.push(`Roteiro da aula:\n${rotStr}`);
    }
    const kDetailText = kDetailParts.join('\n\n');
    if (kDetailText) {
      updatedXml = setCellText(updatedXml, `K${r + 1}`, kDetailText);
    }

    // Col L: Instrumentos de Avaliação (Dropdown na linha r; Detalhamento na linha r+1)
    const instList = aula.instrumentos_avaliacao
      ? (Array.isArray(aula.instrumentos_avaliacao) ? aula.instrumentos_avaliacao : [aula.instrumentos_avaliacao])
      : [];
    const normalizedInstList = instList
      .map(i => normalizeInstrumentName(String(i)).normalized)
      .filter(Boolean);

    // L{r}: Valor único para manter a validação do dropdown institucional
    const mainInstrument = normalizedInstList[0] || '';
    if (mainInstrument) {
      updatedXml = setCellText(updatedXml, `L${r}`, mainInstrument);
    }

    // L{r+1}: Bloco detalhado (Instrumentos utilizados, Avaliação, Evidências e Entrega)
    const lDetailParts = [];
    if (normalizedInstList.length > 0) {
      lDetailParts.push(`Instrumentos utilizados:\n${normalizedInstList.join(' | ')}`);
    }
    if (aula.descricao_avaliacao) {
      const descAvStr = Array.isArray(aula.descricao_avaliacao)
        ? aula.descricao_avaliacao.join('\n')
        : String(aula.descricao_avaliacao);
      lDetailParts.push(`Avaliação:\n${descAvStr}`);
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
      lDetailParts.push(`Evidências:\n${formattedEvid}`);
    }
    if (aula.entrega_estudante) {
      const entStr = Array.isArray(aula.entrega_estudante)
        ? aula.entrega_estudante.join('\n')
        : String(aula.entrega_estudante);
      lDetailParts.push(`Entrega:\n${entStr}`);
    }
    const lDetailText = lDetailParts.join('\n\n');
    if (lDetailText) {
      updatedXml = setCellText(updatedXml, `L${r + 1}`, lDetailText);
    }

    // Ajuste de altura das linhas do bloco (r: topo para dropdown; r+1: detalhamento)
    // Linha superior (r): altura padrão institucional (24.6pt)
    updatedXml = setRowHeight(updatedXml, r, 24.6);

    // Linha inferior (r+1): altura dinâmica para detalhamento (máximo 409pt por linha)
    const linesKDetail = countVisualLines(kDetailText, 40);
    const linesLDetail = countVisualLines(lDetailText, 48);
    const linesD = countVisualLines(capText, 38);
    const linesG = countVisualLines(conText, 32);
    const linesI = countVisualLines(aula.desafio || '', 32);
    const linesJ = countVisualLines(aula.resultado_esperado || '', 32);

    const maxDetailLines = Math.max(linesKDetail, linesLDetail, linesD, linesG, linesI, linesJ, 1);
    const calculatedDetailHt = Math.max(177.75, Math.ceil(maxDetailLines * 12.5 + 15));
    const detailHeight = Math.min(409.0, calculatedDetailHt);

    updatedXml = setRowHeight(updatedXml, r + 1, detailHeight);
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

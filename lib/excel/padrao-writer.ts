export interface CriterioDesempenhoData {
  numero: number;
  tipo?: string;
  capacidade_oficial: string;
  conhecimentos_pertinentes: string[];
  criterio_desempenho: string;
  sa_ids: string[];
  aulas_relacionadas?: number[];
  evidencia_relacionada?: string;
  peso: number;
  n0: string;
  n1: string;
  n2: string;
  n3: string;
  n4: string;
}

export interface PadraoDesempenhoData {
  criterios: CriterioDesempenhoData[];
}

export const MAX_PADRAO_CRITERIA_CAPACITY = 75;
export const PADRAO_START_ROW = 5;
export const PADRAO_END_ROW = 79;

function escapeXmlText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function normalizeCapString(str: string): string {
  return str.trim().replace(/\s+/g, ' ').toLowerCase();
}

export function validatePadraoDesempenho(
  criterios: CriterioDesempenhoData[],
  officialCapacities: string[],
  validSaNames: string[]
): void {
  // 1. Capacity check against max template limit
  if (criterios.length > MAX_PADRAO_CRITERIA_CAPACITY) {
    throw new Error(
      `O Padrão de Desempenho possui ${criterios.length} critérios, mas o modelo institucional atual comporta até ${MAX_PADRAO_CRITERIA_CAPACITY}. É necessário expandir o modelo antes da geração.`
    );
  }

  // 2. Validate duplicates in criterio.numero
  const seenNumbers = new Set<number>();
  for (const c of criterios) {
    if (seenNumbers.has(c.numero)) {
      throw new Error(`Número de critério duplicado no Padrão de Desempenho: ${c.numero}`);
    }
    seenNumbers.add(c.numero);
  }

  // 3. Validate sum of weights (must equal 100%)
  const totalPeso = criterios.reduce((acc, c) => acc + (c.peso || 0), 0);
  if (Math.abs(totalPeso - 100) > 0.001) {
    throw new Error(
      `Os pesos do Padrão de Desempenho totalizam ${totalPeso}%. O total deve ser exatamente 100%.`
    );
  }

  // 4. Validate coverage of official capacities
  const criterionCapacitiesNormalized = criterios.map(c => normalizeCapString(c.capacidade_oficial || ''));
  for (const offCap of officialCapacities) {
    const normOff = normalizeCapString(offCap);
    if (!normOff) continue;
    const isCovered = criterionCapacitiesNormalized.some(cCap => cCap === normOff);
    if (!isCovered) {
      throw new Error(
        `A capacidade oficial abaixo não está representada no Padrão de Desempenho: ${offCap}`
      );
    }
  }

  // 5. Validate linked SAs
  for (const c of criterios) {
    for (const saId of c.sa_ids || []) {
      if (!validSaNames.includes(saId)) {
        throw new Error(
          `O critério ${c.numero} está vinculado à SA "${saId}", mas ela não existe no plano gerado.`
        );
      }
    }
  }
}

export function populatePadraoDesempenho(
  sheetXml: string,
  criterios: CriterioDesempenhoData[],
  officialCapacities: string[],
  validSaNames: string[]
): string {
  // Validate data rules first
  validatePadraoDesempenho(criterios, officialCapacities, validSaNames);

  let updatedXml = sheetXml;

  // Process rows 5 to 79
  for (let slotIndex = 0; slotIndex < MAX_PADRAO_CRITERIA_CAPACITY; slotIndex++) {
    const targetRow = PADRAO_START_ROW + slotIndex;
    const criterio = criterios[slotIndex];

    const rowRegex = new RegExp(`<row [^>]*r="${targetRow}"[\\s\\S]*?<\\/row>`);
    const rowMatch = updatedXml.match(rowRegex);
    if (!rowMatch) continue;

    let rowXml = rowMatch[0];

    const writeCell = (
      col: string,
      type: 'str' | 'num' | 'clear',
      val: string | number = ''
    ) => {
      const cellRef = `${col}${targetRow}`;
      const cellRegex = new RegExp(`<c [^>]*r="${cellRef}"[^>]*>(?:[\\s\\S]*?<\\/c>|\\/>)`);
      const cellMatch = rowXml.match(cellRegex);
      if (!cellMatch) return;

      const fullCell = cellMatch[0];
      const styleMatch = fullCell.match(/\bs="(\d+)"/);
      const styleAttr = styleMatch ? ` s="${styleMatch[1]}"` : '';

      let newCellXml = '';
      if (type === 'clear') {
        newCellXml = `<c r="${cellRef}"${styleAttr}/>`;
      } else if (type === 'num') {
        newCellXml = `<c r="${cellRef}"${styleAttr}><v>${val}</v></c>`;
      } else if (type === 'str') {
        const textVal = String(val);
        if (!textVal) {
          newCellXml = `<c r="${cellRef}"${styleAttr}/>`;
        } else {
          newCellXml = `<c r="${cellRef}"${styleAttr} t="inlineStr"><is><t xml:space="preserve">${escapeXmlText(
            textVal
          )}</t></is></c>`;
        }
      }

      rowXml = rowXml.replace(fullCell, newCellXml);
    };

    if (criterio) {
      // C: Nº
      writeCell('C', 'num', criterio.numero);
      // D: Capacidade oficial
      writeCell('D', 'str', criterio.capacidade_oficial);
      // E: Conhecimentos (joined by newline)
      const conhecimentosStr = (criterio.conhecimentos_pertinentes || []).join('\n');
      writeCell('E', 'str', conhecimentosStr);
      // F: Critério de Desempenho
      writeCell('F', 'str', criterio.criterio_desempenho);
      // G: SA vinculada (joined by newline)
      const saStr = (criterio.sa_ids || []).join('\n');
      writeCell('G', 'str', saStr);
      // H: N0
      writeCell('H', 'str', criterio.n0);
      // I: N1
      writeCell('I', 'str', criterio.n1);
      // J: N2
      writeCell('J', 'str', criterio.n2);
      // K: N3
      writeCell('K', 'str', criterio.n3);
      // L: N4
      writeCell('L', 'str', criterio.n4);
      // M: Peso (fraction 0.0 - 1.0)
      const pesoFraction = criterio.peso / 100;
      writeCell('M', 'num', pesoFraction);
    } else {
      // Clear unused criteria slot (columns C to M)
      ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M'].forEach(col => {
        writeCell(col, 'clear');
      });
    }

    updatedXml = updatedXml.replace(rowMatch[0], rowXml);
  }

  return updatedXml;
}

import { loadZipPackage, saveZipPackage, readZipFileText, writeZipFileText } from './zip-package';
import { getWorksheetPathByName, setActiveSheet, enableWorkbookRecalculation } from './workbook-map';
import { setCellText, setCellNumber, setFormulaCellCachedValue } from './cell-writer';
import { PLANO_DE_ENSINO_CELL_MAP, SA_CELL_MAP, formatSaStrategy } from './institutional-map';
import { adjustSaSheets } from './sa-manager';
import { populateSaLessons, type AulaData } from './lessons-writer';
import { populatePadraoDesempenho, type CriterioDesempenhoData, type PadraoDesempenhoData } from './padrao-writer';

export type { CriterioDesempenhoData, PadraoDesempenhoData };

export interface CellUpdate {
  sheetName: string;
  cellRef: string;
  type: 'string' | 'number';
  value: string | number;
}

export interface SituacaoAprendizagemData {
  numero: number;
  nome?: string;
  carga_horaria: number;
  estrategia: string;
  contextualizacao: string;
  qtdCapacidadesTecnicas: number;
  qtdCapacidadesSocioemocionais: number;
  aulas: AulaData[];
}

export interface PlanoGeralData {
  curso: string;
  cargaHorariaCurso: number;
  modalidade: string;
  unidadeCurricular: string;
  cargaHorariaUc: number;
  modulo: string;
  capacidadesTecnicasOficiais?: string[];
  capacidadesSocioemocionaisOficiais?: string[];
  situacoesAprendizagem: SituacaoAprendizagemData[];
  padraoDesempenho?: PadraoDesempenhoData;
}

export interface SaPopulationSummary {
  saName: string;
  declaredCh: number;
  calculatedCh: number;
  aulasCount: number;
  rowsUsed: number[];
  chFormulaCachedValue: number;
}

/**
 * Executes a set of explicit cell modifications on an OpenXML workbook in memory.
 * Never modifies the input template buffer.
 */
export function updateWorkbookCells(
  templateBuffer: Buffer,
  updates: CellUpdate[]
): Buffer {
  const zip = loadZipPackage(templateBuffer);

  const updatesBySheet = new Map<string, CellUpdate[]>();
  for (const update of updates) {
    const list = updatesBySheet.get(update.sheetName) || [];
    list.push(update);
    updatesBySheet.set(update.sheetName, list);
  }

  for (const [sheetName, sheetUpdates] of updatesBySheet.entries()) {
    const sheetPath = getWorksheetPathByName(zip, sheetName);
    let sheetXml = readZipFileText(zip, sheetPath);

    for (const update of sheetUpdates) {
      if (update.type === 'string') {
        sheetXml = setCellText(sheetXml, update.cellRef, String(update.value));
      } else if (update.type === 'number') {
        sheetXml = setCellNumber(sheetXml, update.cellRef, Number(update.value));
      }
    }

    writeZipFileText(zip, sheetPath, sheetXml);
  }

  return saveZipPackage(zip);
}

/**
 * Populates Plano de Ensino headers, dynamically scales SA sheets,
 * populates SA headers & lessons grid for each SA with validation,
 * populates Padrão de Desempenho sheet with validations,
 * enables forced workbook recalculation,
 * and sets active opening tab to "Plano de Ensino".
 */
export function generateWorkbookWithHeaders(
  templateBuffer: Buffer,
  data: PlanoGeralData
): {
  buffer: Buffer;
  activeTabIndex: number;
  updatedSheets: string[];
  recalcInfo: { calcMode: string; fullCalcOnLoad: string; forceFullCalc: string; calcId: string };
  saManagerResult: { finalSaCount: number; sheetsList: string[]; removedFiles: string[]; addedFiles: string[] };
  saSummaryList: SaPopulationSummary[];
} {
  const zip = loadZipPackage(templateBuffer);
  const updatedSheets: string[] = [];
  const saSummaryList: SaPopulationSummary[] = [];

  // 1. Dynamically scale SA sheets to match data.situacoesAprendizagem.length
  const saManagerResult = adjustSaSheets(zip, data.situacoesAprendizagem.length);

  // 2. Update "Plano de Ensino" header cells
  const planoPath = getWorksheetPathByName(zip, 'Plano de Ensino');
  let planoXml = readZipFileText(zip, planoPath);
  planoXml = setCellText(planoXml, PLANO_DE_ENSINO_CELL_MAP.curso.cellRef, data.curso);
  planoXml = setCellNumber(planoXml, PLANO_DE_ENSINO_CELL_MAP.cargaHoraria.cellRef, data.cargaHorariaCurso);
  planoXml = setCellText(planoXml, 'J4', data.modalidade);
  planoXml = setCellText(planoXml, PLANO_DE_ENSINO_CELL_MAP.unidadeCurricular.cellRef, data.unidadeCurricular);
  planoXml = setCellNumber(planoXml, 'H5', data.cargaHorariaUc);
  planoXml = setCellText(planoXml, 'J5', data.modulo);
  writeZipFileText(zip, planoPath, planoXml);
  updatedSheets.push('Plano de Ensino');

  // 3. Update headers, cached values, and lessons grid for ALL SA sheets in data.situacoesAprendizagem
  data.situacoesAprendizagem.forEach((sa) => {
    const numStr = String(sa.numero).padStart(2, '0');
    const saSheetName = `SA${numStr}`;

    try {
      const saPath = getWorksheetPathByName(zip, saSheetName);
      let saXml = readZipFileText(zip, saPath);

      // Update cached values for header formulas without touching formulas <f>
      saXml = setFormulaCellCachedValue(saXml, SA_CELL_MAP.header.curso, 'string', data.curso);
      saXml = setFormulaCellCachedValue(saXml, SA_CELL_MAP.header.cargaHorariaCurso, 'number', data.cargaHorariaCurso);
      saXml = setFormulaCellCachedValue(saXml, SA_CELL_MAP.header.modalidade, 'string', data.modalidade);
      saXml = setFormulaCellCachedValue(saXml, SA_CELL_MAP.header.unidadeCurricular, 'string', data.unidadeCurricular);
      saXml = setFormulaCellCachedValue(saXml, SA_CELL_MAP.header.cargaHorariaUc, 'number', data.cargaHorariaUc);
      saXml = setFormulaCellCachedValue(saXml, SA_CELL_MAP.header.modulo, 'string', data.modulo);

      // Strategy line 6
      const strategyFormatted = formatSaStrategy(sa.estrategia);
      saXml = setCellText(saXml, SA_CELL_MAP.specific.estrategiaRegion, strategyFormatted);

      // Contextualização B10
      saXml = setCellText(saXml, SA_CELL_MAP.specific.contextualizacao, sa.contextualizacao);

      // Capacidades Básicas/Técnicas count I10
      saXml = setCellNumber(saXml, SA_CELL_MAP.specific.capacidadesTecnicasCount, sa.qtdCapacidadesTecnicas);

      // Capacidades Socioemocionais count K10
      saXml = setCellNumber(saXml, SA_CELL_MAP.specific.capacidadesSocioemocionaisCount, sa.qtdCapacidadesSocioemocionais);

      // Populate Lessons Grid (rows 13-54) with CH validation and capacity checks
      const lessonResult = populateSaLessons(
        saXml,
        saSheetName,
        sa.carga_horaria,
        sa.aulas || []
      );
      saXml = lessonResult.updatedXml;

      writeZipFileText(zip, saPath, saXml);
      updatedSheets.push(saSheetName);

      saSummaryList.push({
        saName: saSheetName,
        declaredCh: sa.carga_horaria,
        calculatedCh: lessonResult.calculatedCh,
        aulasCount: lessonResult.aulasCount,
        rowsUsed: lessonResult.rowsUsed,
        chFormulaCachedValue: lessonResult.calculatedCh,
      });
    } catch (err) {
      console.error(`ERRO AO PROCESSAR PLANILHA "${saSheetName}":`, err);
      throw err;
    }
  });

  // 4. Populate "Padrão de Desempenho" if provided
  if (data.padraoDesempenho && data.padraoDesempenho.criterios) {
    const padraoPath = getWorksheetPathByName(zip, 'Padrão de Desempenho');
    let padraoXml = readZipFileText(zip, padraoPath);

    // Collect all valid SA sheet names
    const validSaNames = data.situacoesAprendizagem.map(
      sa => `SA${String(sa.numero).padStart(2, '0')}`
    );

    // Collect all official capacities for coverage check
    const officialCapacities = [
      ...(data.capacidadesTecnicasOficiais || []),
      ...(data.capacidadesSocioemocionaisOficiais || []),
    ];

    padraoXml = populatePadraoDesempenho(
      padraoXml,
      data.padraoDesempenho.criterios,
      officialCapacities,
      validSaNames
    );

    writeZipFileText(zip, padraoPath, padraoXml);
    updatedSheets.push('Padrão de Desempenho');
  }

  // 5. Set active opening sheet to "Plano de Ensino"
  const { activeTabIndex } = setActiveSheet(zip, 'Plano de Ensino');

  // 6. Force full recalculation on load in workbook.xml
  const recalcInfo = enableWorkbookRecalculation(zip);

  const resultBuffer = saveZipPackage(zip);
  return {
    buffer: resultBuffer,
    activeTabIndex,
    updatedSheets,
    recalcInfo,
    saManagerResult,
    saSummaryList,
  };
}

export interface PlanValidationResult {
  valid: boolean;
  message?: string;
  summary?: {
    curso: string;
    uc: string;
    cargaHoraria: number;
    situacoes: number;
    aulas: number;
    capacidades: number;
    criterios: number;
  };
}

export function validatePlanoEnsinoJson(data: any): PlanValidationResult {
  if (!data || typeof data !== 'object') {
    return { valid: false, message: 'O arquivo enviado não contém um Plano de Ensino válido.' };
  }

  if (!data.schema_version || !String(data.schema_version).startsWith('senai-plano-ensino-')) {
    return { valid: false, message: 'Este arquivo não foi reconhecido como um Arquivo do Plano de Ensino compatível.' };
  }

  const identificacao = data.identificacao || {};
  const dadosOficiais = data.dados_oficiais || {};
  const status = data.status || {};
  const sas = Array.isArray(data.situacoes_aprendizagem) ? data.situacoes_aprendizagem : [];
  const capacidadesTecnicas = Array.isArray(dadosOficiais.capacidades_tecnicas) ? dadosOficiais.capacidades_tecnicas : [];
  const capacidadesSocio = Array.isArray(dadosOficiais.capacidades_socioemocionais) ? dadosOficiais.capacidades_socioemocionais : [];
  const criterios = Array.isArray(data.padrao_desempenho?.criterios) ? data.padrao_desempenho.criterios : [];

  // Check approval status (if explicit status object exists)
  if (status.plano_aprovado === false) {
    return { valid: false, message: 'O Plano de Ensino precisa estar aprovado antes da geração do Excel.' };
  }
  if (status.padrao_desempenho_aprovado === false) {
    return { valid: false, message: 'O Padrão de Desempenho precisa estar aprovado antes da geração do Excel.' };
  }

  // Check required basic fields
  if (!identificacao.unidade_curricular) {
    return { valid: false, message: 'A Unidade Curricular precisa estar preenchida no planejamento.' };
  }
  if (!identificacao.carga_horaria_uc || Number(identificacao.carga_horaria_uc) <= 0) {
    return { valid: false, message: 'A Carga Horária da Unidade Curricular precisa ser um valor numérico válido.' };
  }

  // Collect all lessons
  let totalAulasCount = 0;
  const lessonsBySa = new Map<string, any[]>();

  for (const sa of sas) {
    const saId = sa.nome || `SA${String(sa.numero).padStart(2, '0')}`;
    let saAulas = Array.isArray(sa.aulas) ? sa.aulas : [];
    if (saAulas.length === 0 && Array.isArray(data.aulas)) {
      saAulas = data.aulas.filter((a: any) => a.sa_id === saId || a.sa_id === `SA${String(sa.numero).padStart(2, '0')}`);
    }
    lessonsBySa.set(saId, saAulas);
    totalAulasCount += saAulas.length;
  }

  if (sas.length === 0) {
    return { valid: false, message: 'O planejamento precisa conter pelo menos uma Situação de Aprendizagem (SA).' };
  }
  if (totalAulasCount === 0) {
    return { valid: false, message: 'O planejamento precisa conter aulas cadastradas nas Situações de Aprendizagem.' };
  }
  if (capacidadesTecnicas.length === 0) {
    return { valid: false, message: 'O planejamento precisa conter capacidades técnicas oficiais.' };
  }
  if (criterios.length === 0) {
    return { valid: false, message: 'O Padrão de Desempenho precisa conter critérios cadastrados.' };
  }

  // Check physical capacity limits
  if (criterios.length > 75) {
    return {
      valid: false,
      message: `O Padrão de Desempenho possui ${criterios.length} critérios, mas o modelo institucional atual comporta até 75. É necessário expandir o modelo antes da geração.`
    };
  }

  // Check SA lessons vs capacity and CH
  for (const sa of sas) {
    const saId = sa.nome || `SA${String(sa.numero).padStart(2, '0')}`;
    const saAulas = lessonsBySa.get(saId) || [];
    if (saAulas.length > 21) {
      return {
        valid: false,
        message: `A ${saId} possui ${saAulas.length} aulas, mas o modelo institucional comporta até 21 aulas por SA.`
      };
    }
    const declaredCh = Number(sa.carga_horaria || 0);
    const sumAulasCh = saAulas.reduce((acc: number, a: any) => acc + Number(a.carga_horaria || 0), 0);
    if (declaredCh > 0 && Math.abs(declaredCh - sumAulasCh) > 0.001) {
      return {
        valid: false,
        message: `Divergência na carga horária da ${saId}: carga declarada = ${declaredCh}h; soma das aulas = ${sumAulasCh}h.`
      };
    }
  }

  // Check sum of weights = 100%
  const totalPeso = criterios.reduce((sum: number, c: any) => sum + Number(c?.peso || 0), 0);
  if (Math.abs(totalPeso - 100) > 0.001) {
    return {
      valid: false,
      message: `Os pesos do Padrão de Desempenho somam ${totalPeso}%. O total precisa ser exatamente 100%.`
    };
  }

  // Check linked SAs in criteria
  const validSaNames = sas.map((sa: any) => sa.nome || `SA${String(sa.numero).padStart(2, '0')}`);
  for (const c of criterios) {
    const saIds = Array.isArray(c.sa_ids) ? c.sa_ids : [];
    for (const saId of saIds) {
      if (!validSaNames.includes(saId)) {
        return {
          valid: false,
          message: `O critério ${c.numero} está vinculado à ${saId}, mas ela não existe nas Situações de Aprendizagem do plano.`
        };
      }
    }
  }

  // Check capacity coverage in criteria
  const officialCapacities = [...capacidadesTecnicas, ...capacidadesSocio];
  const criterionCapsNormalized = criterios.map((c: any) =>
    (c.capacidade_oficial || '').trim().replace(/\s+/g, ' ').toLowerCase()
  );
  for (const offCap of officialCapacities) {
    const normOff = offCap.trim().replace(/\s+/g, ' ').toLowerCase();
    if (!normOff) continue;
    if (!criterionCapsNormalized.includes(normOff)) {
      return {
        valid: false,
        message: `A capacidade oficial abaixo não está representada no Padrão de Desempenho:\n"${offCap}"`
      };
    }
  }

  return {
    valid: true,
    summary: {
      curso: identificacao.curso || '[A DEFINIR PELO DOCENTE]',
      uc: identificacao.unidade_curricular,
      cargaHoraria: Number(identificacao.carga_horaria_uc),
      situacoes: sas.length,
      aulas: totalAulasCount,
      capacidades: capacidadesTecnicas.length,
      criterios: criterios.length,
    }
  };
}

/**
 * Maps raw JSON input to PlanoGeralData structure for generateWorkbookWithHeaders.
 */
export function mapJsonToPlanoGeralData(data: any) {
  const identificacao = data.identificacao || {};
  const dadosOficiais = data.dados_oficiais || {};
  const sas = Array.isArray(data.situacoes_aprendizagem) ? data.situacoes_aprendizagem : [];

  const mappedSas = sas.map((sa: any) => {
    const saId = sa.nome || `SA${String(sa.numero).padStart(2, '0')}`;
    let saAulas = Array.isArray(sa.aulas) ? sa.aulas : [];
    if (saAulas.length === 0 && Array.isArray(data.aulas)) {
      saAulas = data.aulas.filter((a: any) => a.sa_id === saId || a.sa_id === `SA${String(sa.numero).padStart(2, '0')}`);
    }

    return {
      numero: sa.numero,
      nome: sa.nome,
      carga_horaria: Number(sa.carga_horaria || 0),
      estrategia: sa.estrategia || '',
      contextualizacao: sa.contextualizacao || '',
      qtdCapacidadesTecnicas: Number(sa.qtdCapacidadesTecnicas || 0),
      qtdCapacidadesSocioemocionais: Number(sa.qtdCapacidadesSocioemocionais || 0),
      aulas: saAulas,
    };
  });

  return {
    curso: identificacao.curso || '[PREENCHER PELO DOCENTE]',
    cargaHorariaCurso: Number(identificacao.carga_horaria_curso || 0),
    modalidade: identificacao.modalidade || 'Presencial',
    unidadeCurricular: identificacao.unidade_curricular || '',
    cargaHorariaUc: Number(identificacao.carga_horaria_uc || 0),
    modulo: identificacao.modulo || '',
    capacidadesTecnicasOficiais: dadosOficiais.capacidades_tecnicas || [],
    capacidadesSocioemocionaisOficiais: dadosOficiais.capacidades_socioemocionais || [],
    situacoesAprendizagem: mappedSas,
    padraoDesempenho: data.padrao_desempenho,
  };
}

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
  const globalAulas = Array.isArray(data.aulas) ? data.aulas : [];
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

  if (sas.length === 0) {
    return { valid: false, message: 'O planejamento precisa conter pelo menos uma Situação de Aprendizagem (SA).' };
  }

  // Build map of valid SA identifiers
  const validSaIds = new Set<string>();
  sas.forEach((sa: any, index: number) => {
    const num = sa.numero || (sa.id ? parseInt(String(sa.id).replace(/\D/g, ''), 10) : index + 1) || index + 1;
    const formattedId = `SA${String(num).padStart(2, '0')}`;

    if (sa.id) validSaIds.add(String(sa.id));
    if (sa.nome) validSaIds.add(String(sa.nome));
    validSaIds.add(formattedId);
    validSaIds.add(`SA${num}`);
  });

  // Validação 2: Confirmar que todos os sa_id usados nas aulas correspondem a uma SA existente
  for (const aula of globalAulas) {
    if (!aula.sa_id || !validSaIds.has(String(aula.sa_id))) {
      return {
        valid: false,
        message: `A aula ${aula.numero || ''} está vinculada à "${aula.sa_id || 'desconhecida'}", mas essa Situação de Aprendizagem não existe.`
      };
    }
  }

  // Validações por SA
  let totalAulasCount = 0;
  const lessonsBySa = new Map<string, any[]>();

  for (let i = 0; i < sas.length; i++) {
    const sa = sas[i];
    const num = sa.numero || (sa.id ? parseInt(String(sa.id).replace(/\D/g, ''), 10) : i + 1) || i + 1;
    const formattedId = `SA${String(num).padStart(2, '0')}`;
    const rawId = sa.id || sa.nome;

    // Filter aulas for this SA from globalAulas or sa.aulas
    let saAulas = globalAulas.filter((a: any) =>
      a.sa_id === rawId || a.sa_id === formattedId || a.sa_id === sa.id || a.sa_id === sa.nome
    );
    if (saAulas.length === 0 && Array.isArray(sa.aulas)) {
      saAulas = sa.aulas;
    }

    // Validação 1 & 6: Confirmar que existe pelo menos uma aula vinculada para a SA (sem exigir sa.aulas)
    if (saAulas.length === 0) {
      return {
        valid: false,
        message: `O planejamento precisa conter aulas cadastradas nas Situações de Aprendizagem (${formattedId}).`
      };
    }

    // Validação de capacidade máxima de slots no modelo institucional (21 aulas)
    if (saAulas.length > 21) {
      return {
        valid: false,
        message: `A ${formattedId} possui ${saAulas.length} aulas, mas o modelo institucional comporta até 21 aulas por SA.`
      };
    }

    // Validação 3: Confirmar que os números de aula relacionados sejam coerentes com sa.encontros
    if (Array.isArray(sa.encontros) && sa.encontros.length > 0) {
      const lessonNumbers = saAulas.map((a: any) => Number(a.numero)).sort((a: number, b: number) => a - b);
      const encontrosNumbers = [...sa.encontros].map((n: any) => Number(n)).sort((a: number, b: number) => a - b);

      const isCoherent = lessonNumbers.length === encontrosNumbers.length &&
        lessonNumbers.every((val: number, idx: number) => val === encontrosNumbers[idx]);

      if (!isCoherent) {
        return {
          valid: false,
          message: `Os números de aula da ${formattedId} (${lessonNumbers.join(', ')}) não correspondem aos encontros declarados (${encontrosNumbers.join(', ')}).`
        };
      }
    }

    // Validação 4: Confirmar que a soma das cargas horárias das aulas da SA corresponde a sa.carga_horaria
    const declaredCh = Number(sa.carga_horaria || 0);
    const sumAulasCh = saAulas.reduce((acc: number, a: any) => acc + Number(a.carga_horaria || 0), 0);
    if (declaredCh > 0 && Math.abs(declaredCh - sumAulasCh) > 0.001) {
      return {
        valid: false,
        message: `Divergência na carga horária da ${formattedId}: carga declarada = ${declaredCh}h; soma das aulas = ${sumAulasCh}h.`
      };
    }

    lessonsBySa.set(formattedId, saAulas);
    totalAulasCount += saAulas.length;
  }

  // Validação 5: Confirmar que a soma global das aulas corresponde à carga horária planejada
  const plannedCh = Number(data.calendario?.carga_horaria_planejada || identificacao.carga_horaria_uc || 0);
  const allAulas = globalAulas.length > 0
    ? globalAulas
    : sas.flatMap((sa: any) => Array.isArray(sa.aulas) ? sa.aulas : []);
  const totalGlobalAulasCh = allAulas.reduce((acc: number, a: any) => acc + Number(a.carga_horaria || 0), 0);

  if (plannedCh > 0 && Math.abs(plannedCh - totalGlobalAulasCh) > 0.001) {
    return {
      valid: false,
      message: `Divergência na carga horária total das aulas: carga horária planejada = ${plannedCh}h; soma de todas as aulas = ${totalGlobalAulasCh}h.`
    };
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

  // Check sum of weights = 100%
  const totalPeso = criterios.reduce((sum: number, c: any) => sum + Number(c?.peso || 0), 0);
  if (Math.abs(totalPeso - 100) > 0.001) {
    return {
      valid: false,
      message: `Os pesos do Padrão de Desempenho somam ${totalPeso}%. O total precisa ser exatamente 100%.`
    };
  }

  // Check linked SAs in criteria
  for (const c of criterios) {
    const saIds = Array.isArray(c.sa_ids) ? c.sa_ids : [];
    for (const saId of saIds) {
      if (!validSaIds.has(String(saId))) {
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
  const globalAulas = Array.isArray(data.aulas) ? data.aulas : [];

  const totalAulasCount = globalAulas.length > 0
    ? globalAulas.length
    : sas.reduce((sum: number, sa: any) => sum + (Array.isArray(sa.aulas) ? sa.aulas.length : 0), 0);

  const mappedSas = sas.map((sa: any, index: number) => {
    const saNumero = sa.numero || (sa.id ? parseInt(String(sa.id).replace(/\D/g, ''), 10) : index + 1) || index + 1;
    const saId = sa.id || sa.nome || `SA${String(saNumero).padStart(2, '0')}`;
    const formattedSaId = `SA${String(saNumero).padStart(2, '0')}`;

    let saAulas = globalAulas.filter((a: any) =>
      a.sa_id === saId || a.sa_id === formattedSaId || a.sa_id === sa.id || a.sa_id === sa.nome
    );
    if (saAulas.length === 0 && Array.isArray(sa.aulas) && sa.aulas.length > 0) {
      saAulas = sa.aulas;
    }

    return {
      numero: saNumero,
      nome: formattedSaId,
      carga_horaria: Number(sa.carga_horaria || 0),
      estrategia: sa.estrategia_desafiadora_principal || sa.estrategia || '',
      contextualizacao: sa.contexto || sa.contextualizacao || '',
      qtdCapacidadesTecnicas: Number(sa.qtdCapacidadesTecnicas || (Array.isArray(sa.capacidades) ? sa.capacidades.length : 0)),
      qtdCapacidadesSocioemocionais: Number(sa.qtdCapacidadesSocioemocionais || (Array.isArray(sa.capacidades_socioemocionais) ? sa.capacidades_socioemocionais.length : 0)),
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
    dataInicio: identificacao.data_inicio || '',
    dataFim: identificacao.data_fim || '',
    turma: identificacao.turma || '',
    objetivoUc: dadosOficiais.objetivo_uc || '',
    ambientePedagogico: identificacao.ambiente_pedagogico || '',
    conhecimentosOficiais: dadosOficiais.conhecimentos || [],
    capacidadesTecnicasOficiais: dadosOficiais.capacidades_tecnicas || [],
    capacidadesSocioemocionaisOficiais: dadosOficiais.capacidades_socioemocionais || [],
    totalAulasCount: totalAulasCount,
    situacoesAprendizagem: mappedSas,
    padraoDesempenho: data.padrao_desempenho,
  };
}


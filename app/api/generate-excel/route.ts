import { NextResponse } from 'next/server';
import fs from 'fs';
import path from 'path';
import { validatePlanoEnsinoJson, mapJsonToPlanoGeralData } from '@/lib/excel/plan-validator';
import { generateWorkbookWithHeaders } from '@/lib/excel/openxml';

export const runtime = 'nodejs';

function sanitizeFilename(ucName: string): string {
  const cleanUc = (ucName || 'OFICIAL')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '') // remove accents
    .replace(/[\\/*?:"<>|]/g, '_')   // replace illegal filename chars
    .replace(/\s+/g, '_')            // replace spaces with underscores
    .toUpperCase();

  return `PLANO_DE_ENSINO_${cleanUc}.xlsm`;
}

export async function POST(request: Request) {
  try {
    const data = await request.json();

    // 1. Centralized validation check
    const validation = validatePlanoEnsinoJson(data);
    if (!validation.valid) {
      return NextResponse.json(
        {
          success: false,
          message: 'Não foi possível gerar o Excel.',
          details: validation.message || 'Dados do plano de ensino inválidos.',
        },
        { status: 400 }
      );
    }

    // 2. Read institutional template
    const templatePath = path.resolve(process.cwd(), 'templates/PLANO_DE_ENSINO_MODELO.xlsm');
    if (!fs.existsSync(templatePath)) {
      console.error('[GENERATE EXCEL] Template não encontrado em:', templatePath);
      return NextResponse.json(
        {
          success: false,
          message: 'Não foi possível gerar o Excel.',
          details: 'Modelo institucional não foi encontrado no servidor.',
        },
        { status: 500 }
      );
    }

    const templateBuffer = fs.readFileSync(templatePath);

    // 3. Map JSON input to PlanoGeralData and generate workbook in memory
    const planoData = mapJsonToPlanoGeralData(data);
    const { buffer: resultBuffer } = generateWorkbookWithHeaders(templateBuffer, planoData);

    // 4. Sanitize filename for download
    const filename = sanitizeFilename(planoData.unidadeCurricular);

    // 5. Return binary XLSM stream
    const responseBody = new Uint8Array(resultBuffer);

    return new Response(responseBody, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.ms-excel.sheet.macroEnabled.12',
        'Content-Disposition': `attachment; filename="${filename}"`,
        'Cache-Control': 'no-store, no-cache, must-revalidate',
      },
    });
  } catch (error) {
    console.error('[GENERATE EXCEL] Erro ao processar arquivo:', error);
    const detailsMessage = error instanceof Error ? error.message : 'Erro interno durante o processamento do Excel.';
    return NextResponse.json(
      {
        success: false,
        message: 'Não foi possível gerar o Excel.',
        details: detailsMessage,
      },
      { status: 500 }
    );
  }
}

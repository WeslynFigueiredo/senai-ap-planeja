import { NextResponse } from "next/server";
import { validatePlanoEnsinoJson } from "@/lib/excel/plan-validator";

export async function POST(request: Request) {
  try {
    const data = await request.json();
    const result = validatePlanoEnsinoJson(data);

    if (!result.valid) {
      return NextResponse.json(
        { ok: false, message: result.message || "O arquivo não é um Plano de Ensino válido." },
        { status: 400 }
      );
    }

    return NextResponse.json({
      ok: true,
      summary: result.summary,
    });
  } catch {
    return NextResponse.json(
      { ok: false, message: "O arquivo parece estar corrompido ou não está no formato esperado." },
      { status: 400 }
    );
  }
}

"use client";

import { useRef, useState, useEffect } from "react";
import {
  ArrowIcon,
  AwardIcon,
  BookIcon,
  CalendarIcon,
  CheckIcon,
  ClockIcon,
  DownloadIcon,
  FileIcon,
  LayersIcon,
  RefreshIcon,
  SparkIcon,
  SpinnerIcon,
  UploadIcon,
} from "@/components/icons";

type PlanSummary = {
  curso: string;
  uc: string;
  cargaHoraria: number | string;
  situacoes: number;
  aulas: number;
  capacidades: number;
  criterios: number;
};

type AppState =
  | { status: "idle" }
  | { status: "validating"; fileName: string }
  | { status: "validated"; fileName: string; summary: PlanSummary; raw: unknown }
  | { status: "generating"; fileName: string; summary: PlanSummary; raw: unknown; stepIndex: number }
  | { status: "success"; fileName: string; summary: PlanSummary; downloadUrl: string; downloadFilename: string }
  | { status: "error"; fileName?: string; message: string; raw?: unknown };

const GENERATION_STEPS = [
  "Validando o planejamento",
  "Organizando as Situações de Aprendizagem",
  "Montando o Excel institucional",
  "Finalizando o arquivo",
];

export default function HomePage() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [state, setState] = useState<AppState>({ status: "idle" });
  const [dragging, setDragging] = useState(false);

  // Animate sequential steps during generation
  useEffect(() => {
    if (state.status !== "generating") return;

    const interval = setInterval(() => {
      setState((prev) => {
        if (prev.status !== "generating") return prev;
        if (prev.stepIndex < GENERATION_STEPS.length - 1) {
          return { ...prev, stepIndex: prev.stepIndex + 1 };
        }
        return prev;
      });
    }, 600);

    return () => clearInterval(interval);
  }, [state.status]);

  async function processFile(file?: File) {
    if (!file) return;
    setState({ status: "validating", fileName: file.name });
    try {
      if (!file.name.toLowerCase().endsWith(".json")) {
        throw new Error("Selecione o Arquivo do Plano de Ensino baixado ao final do Mestre de Plano de Ensino.");
      }
      const text = await file.text();
      let parsedJson: unknown;
      try {
        parsedJson = JSON.parse(text);
      } catch {
        throw new Error("O arquivo selecionado parece estar corrompido ou malformatado.");
      }

      const response = await fetch("/api/validate-plan", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(parsedJson),
      });

      const result = await response.json();
      if (!response.ok || !result.ok) {
        throw new Error(result.message || "Não foi possível validar este arquivo.");
      }

      setState({
        status: "validated",
        fileName: file.name,
        summary: result.summary,
        raw: parsedJson,
      });
    } catch (error) {
      setState({
        status: "error",
        fileName: file.name,
        message: error instanceof Error ? error.message : "Não foi possível processar o arquivo do plano.",
      });
    }
  }

  async function handleGenerateExcel() {
    if (state.status !== "validated") return;

    const { fileName, summary, raw } = state;
    setState({ status: "generating", fileName, summary, raw, stepIndex: 0 });

    try {
      const response = await fetch("/api/generate-excel", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(raw),
      });

      if (!response.ok) {
        let errorMessage = "Não foi possível gerar o Excel oficial.";
        try {
          const errJson = await response.json();
          errorMessage = errJson.details || errJson.message || errorMessage;
        } catch {
          // Fallback if response is not JSON
        }
        throw new Error(errorMessage);
      }

      // Extract filename from Content-Disposition header if present
      const disposition = response.headers.get("Content-Disposition");
      let downloadFilename = `PLANO_DE_ENSINO_${(summary.uc || "OFICIAL").toUpperCase().replace(/\s+/g, "_")}.xlsm`;

      if (disposition && disposition.includes("filename=")) {
        const match = disposition.match(/filename=["']?([^"';]+)["']?/);
        if (match && match[1]) {
          downloadFilename = match[1];
        }
      }

      const blob = await response.blob();
      const downloadUrl = URL.createObjectURL(blob);

      // Trigger automatic browser download
      triggerBrowserDownload(downloadUrl, downloadFilename);

      setState({
        status: "success",
        fileName,
        summary,
        downloadUrl,
        downloadFilename,
      });
    } catch (error) {
      setState({
        status: "error",
        fileName,
        message: error instanceof Error ? error.message : "Falha durante a geração do arquivo.",
        raw,
      });
    }
  }

  function triggerBrowserDownload(url: string, filename: string) {
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
  }

  function reset() {
    if (state.status === "success" && state.downloadUrl) {
      URL.revokeObjectURL(state.downloadUrl);
    }
    setState({ status: "idle" });
    if (inputRef.current) inputRef.current.value = "";
  }

  return (
    <main className="site-shell">
      <div className="ambient ambient-one" />
      <div className="ambient ambient-two" />
      <div className="ambient ambient-three" />
      <div className="grid-overlay" />

      <header className="topbar">
        <a className="brand" href="#top" aria-label="SENAI AP Planeja">
          <span className="brand-mark"><SparkIcon /></span>
          <span><b>SENAI AP Planeja</b><small>Do plano ao Excel institucional</small></span>
        </a>
        <div className="top-status"><span className="pulse" /> Sistema de exportação</div>
      </header>

      <section className="hero" id="top">
        <h1>Seu planejamento pronto<br /><span>para o Excel institucional.</span></h1>
        <p>Envie o <strong>Arquivo do Plano de Ensino</strong> e gere o documento oficial.</p>

        <div className="journey" aria-label="Etapas do processo">
          <div className={`journey-step ${state.status !== "idle" ? "active" : ""}`}>
            <span>1</span>
            <div><b>Enviar arquivo</b><small>Selecione o plano baixado</small></div>
          </div>
          <i />
          <div className={`journey-step ${state.status === "validated" || state.status === "generating" || state.status === "success" ? "active" : ""}`}>
            <span>2</span>
            <div><b>Conferir</b><small>Validação automática</small></div>
          </div>
          <i />
          <div className={`journey-step ${state.status === "success" ? "active" : ""}`}>
            <span>3</span>
            <div><b>Gerar Excel</b><small>Download do arquivo oficial</small></div>
          </div>
        </div>
      </section>

      <section className="workspace">
        <div className="card upload-card">
          <div className="card-head">
            <div>
              <span className="step-label">PASSO 1</span>
              <h2>Envie o arquivo do seu plano</h2>
            </div>
            <div className="security-badge"><CheckIcon /> Validação automática</div>
          </div>

          <input
            ref={inputRef}
            type="file"
            accept="application/json,.json"
            hidden
            onChange={(e) => processFile(e.target.files?.[0])}
          />

          {/* IDLE / VALIDATING DROPZONE */}
          {(state.status === "idle" || state.status === "validating") && (
            <button
              className={`dropzone ${dragging ? "dragging" : ""}`}
              onClick={() => inputRef.current?.click()}
              onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={(e) => { e.preventDefault(); setDragging(false); processFile(e.dataTransfer.files?.[0]); }}
              aria-label="Área de envio do arquivo do plano de ensino"
            >
              <span className="upload-orbit">
                <span className="upload-core">
                  {state.status === "validating" ? <SpinnerIcon /> : <UploadIcon />}
                </span>
              </span>
              <strong>
                {state.status === "validating" ? "Conferindo seu arquivo..." : "Arraste o arquivo aqui"}
              </strong>
              <small>
                {state.status === "validating" ? state.fileName : "ou selecione no computador"}
              </small>
              {state.status !== "validating" && (
                <span className="select-button">Selecionar arquivo <ArrowIcon /></span>
              )}
            </button>
          )}

          {/* ERROR STATE */}
          {state.status === "error" && (
            <div className="message error-message">
              <FileIcon />
              <div>
                <b>Não foi possível gerar o Excel.</b>
                <p>{state.message}</p>
              </div>
              <button onClick={reset}>Tentar novamente</button>
            </div>
          )}

          {/* VALIDATED STATE */}
          {state.status === "validated" && (
            <div className="success-panel">
              <div className="success-title">
                <span className="success-icon"><CheckIcon /></span>
                <div>
                  <small>ARQUIVO RECONHECIDO</small>
                  <h3>{state.fileName}</h3>
                </div>
                <button className="text-button" onClick={reset}>Trocar arquivo</button>
              </div>

              <div className="summary-grid">
                <Summary label="Curso" value={state.summary.curso} icon={<BookIcon />} wide />
                <Summary label="Unidade Curricular" value={state.summary.uc} icon={<LayersIcon />} wide />
                <Summary label="Carga horária" value={`${state.summary.cargaHoraria}h`} icon={<ClockIcon />} />
                <Summary label="Situações de Aprendizagem" value={state.summary.situacoes} icon={<BookIcon />} />
                <Summary label="Aulas" value={state.summary.aulas} icon={<CalendarIcon />} />
                <Summary label="Capacidades" value={state.summary.capacidades} icon={<AwardIcon />} />
                <Summary label="Critérios" value={state.summary.criterios} icon={<CheckIcon />} />
              </div>

              <div className="validation-line">
                <CheckIcon /> Estrutura validada e pronta para gerar o Excel institucional.
              </div>

              <button className="primary-action" onClick={handleGenerateExcel}>
                Gerar Excel oficial <ArrowIcon />
              </button>
            </div>
          )}

          {/* GENERATING STATE */}
          {state.status === "generating" && (
            <div className="success-panel">
              <div className="success-title">
                <span className="success-icon"><SpinnerIcon /></span>
                <div>
                  <small>PROCESSANDO O PLANO</small>
                  <h3>{state.fileName}</h3>
                </div>
              </div>

              <div className="generation-panel">
                <div className="generation-spinner" />
                <h3>Preparando seu Excel...</h3>
                <p className="coming-note">Montando a estrutura oficial e ajustando a grade de aulas das Situações de Aprendizagem.</p>

                <div className="generation-steps">
                  {GENERATION_STEPS.map((stepText, idx) => {
                    const isDone = idx < state.stepIndex;
                    const isActive = idx === state.stepIndex;
                    return (
                      <div
                        key={stepText}
                        className={`gen-step-item ${isActive ? "active" : ""} ${isDone ? "done" : ""}`}
                      >
                        <span className="gen-step-dot" />
                        <span>{stepText}</span>
                        {isDone && <CheckIcon style={{ width: 16, marginLeft: "auto", color: "var(--accent)" }} />}
                      </div>
                    );
                  })}
                </div>
              </div>

              <button className="primary-action" disabled>
                Preparando seu Excel... <SpinnerIcon />
              </button>
            </div>
          )}

          {/* SUCCESS STATE */}
          {state.status === "success" && (
            <div className="success-panel">
              <div className="success-title">
                <span className="success-icon" style={{ background: "rgba(52,211,153,.2)" }}>
                  <CheckIcon />
                </span>
                <div>
                  <small style={{ color: "var(--accent-2)" }}>CONCLUÍDO COM SUCESSO</small>
                  <h3>Excel gerado com sucesso!</h3>
                </div>
              </div>

              <p style={{ margin: "18px 0 0", color: "#d2e5f8", fontSize: "0.95rem", lineHeight: 1.6 }}>
                Seu Plano de Ensino institucional está pronto para uso e o download foi iniciado automaticamente.
              </p>

              <div className="summary-grid">
                <Summary label="Curso" value={state.summary.curso} icon={<BookIcon />} wide />
                <Summary label="Unidade Curricular" value={state.summary.uc} icon={<LayersIcon />} wide />
                <Summary label="Carga horária" value={`${state.summary.cargaHoraria}h`} icon={<ClockIcon />} />
                <Summary label="Situações de Aprendizagem" value={state.summary.situacoes} icon={<BookIcon />} />
                <Summary label="Aulas" value={state.summary.aulas} icon={<CalendarIcon />} />
                <Summary label="Capacidades" value={state.summary.capacidades} icon={<AwardIcon />} />
                <Summary label="Critérios" value={state.summary.criterios} icon={<CheckIcon />} />
              </div>

              <div className="action-group">
                <button
                  className="primary-action"
                  style={{ marginTop: 0 }}
                  onClick={() => triggerBrowserDownload(state.downloadUrl, state.downloadFilename)}
                >
                  Baixar novamente <DownloadIcon />
                </button>
                <button className="secondary-action" onClick={reset}>
                  <RefreshIcon /> Gerar outro plano
                </button>
              </div>
            </div>
          )}
        </div>

        <aside className="side-card">
          <span className="step-label">COMO FUNCIONA</span>
          <h3>Você não precisa configurar nada.</h3>
          <div className="mini-steps">
            <div><span>01</span><p><b>Finalize seu planejamento</b><small>No Mestre de Plano de Ensino.</small></p></div>
            <div><span>02</span><p><b>Baixe o Arquivo do Plano</b><small>Ao concluir o planejamento.</small></p></div>
            <div><span>03</span><p><b>Gere o Excel</b><small>Envie o arquivo aqui e pronto.</small></p></div>
          </div>
        </aside>
      </section>

      <footer><span>SENAI AP Planeja</span><small>Ferramenta de apoio ao planejamento docente</small></footer>
    </main>
  );
}

function Summary({
  label,
  value,
  icon,
  wide = false,
}: {
  label: string;
  value: string | number;
  icon?: React.ReactNode;
  wide?: boolean;
}) {
  return (
    <div className={`summary-item ${wide ? "wide" : ""}`}>
      <div className="summary-item-head">
        {icon}
        <span>{label}</span>
      </div>
      <strong>{value || "—"}</strong>
    </div>
  );
}

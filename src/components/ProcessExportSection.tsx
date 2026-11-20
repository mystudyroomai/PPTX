import React from 'react';
import {
  AlertCircle,
  AlertTriangle,
  CheckCircle2,
  Download,
  GitCompare,
  Lock,
  Plus,
  RotateCcw,
  ShieldCheck,
  Trash2,
} from 'lucide-react';
import type {
  PptxAnalysis,
  ProcessingOptions,
  ProcessingResult,
  TextReplacementRule,
  VerificationCheck,
} from '../types/pptx';
import { downloadUint8Array, formatBytes, shortHash } from '../utils/formatters';

interface ProcessExportSectionProps {
  originalAnalysis: PptxAnalysis;
  options: ProcessingOptions;
  onChangeOptions: (next: ProcessingOptions) => void;
  onResetOptions: () => void;
  onRunProcess: () => void;
  isProcessing: boolean;
  result: ProcessingResult | null;
  onGoToComparison: () => void;
}

export const ProcessExportSection: React.FC<ProcessExportSectionProps> = ({
  originalAnalysis,
  options,
  onChangeOptions,
  onResetOptions,
  onRunProcess,
  isProcessing,
  result,
  onGoToComparison,
}) => {
  const addReplacementRule = () => {
    const newRule: TextReplacementRule = {
      id: `rule-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`,
      find: '',
      replace: '',
      caseSensitive: false,
    };
    onChangeOptions({
      ...options,
      textReplacements: [...options.textReplacements, newRule],
    });
  };

  const updateRule = (id: string, patch: Partial<TextReplacementRule>) => {
    onChangeOptions({
      ...options,
      textReplacements: options.textReplacements.map((r) =>
        r.id === id ? { ...r, ...patch } : r
      ),
    });
  };

  const removeRule = (id: string) => {
    onChangeOptions({
      ...options,
      textReplacements: options.textReplacements.filter((r) => r.id !== id),
    });
  };

  const toggleSlideHidden = (slideNumber: number, currentlyHidden: boolean) => {
    onChangeOptions({
      ...options,
      slideVisibilityOverrides: {
        ...options.slideVisibilityOverrides,
        [slideNumber]: !currentlyHidden,
      },
    });
  };

  const renderCheckStatus = (status: VerificationCheck['status']) => {
    if (status === 'pass') {
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-semibold text-emerald-800 whitespace-nowrap">
          <CheckCircle2 className="w-4 h-4 text-emerald-700 shrink-0" aria-hidden="true" />
          <span>Verificado</span>
        </span>
      );
    }
    if (status === 'limitation') {
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-semibold text-amber-800 whitespace-nowrap">
          <AlertTriangle className="w-4 h-4 text-amber-700 shrink-0" aria-hidden="true" />
          <span>Verificado com limitação</span>
        </span>
      );
    }
    return (
      <span className="inline-flex items-center gap-1.5 text-xs font-semibold text-red-800 whitespace-nowrap">
        <AlertCircle className="w-4 h-4 text-red-700 shrink-0" aria-hidden="true" />
        <span>Falha — Bloqueado</span>
      </span>
    );
  };

  const slidesWithNotesCount = originalAnalysis.slides.filter(
    (s) => s.hasNotes && s.notesText.trim().length > 0
  ).length;

  return (
    <div className="space-y-8">
      <div className="pb-5 border-b border-slate-200 flex flex-col sm:flex-row sm:items-end justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">
            Processamento seguro, verificação e exportação
          </h1>
          <p className="mt-1.5 text-sm text-slate-600 max-w-3xl">
            Configure operações estritamente não destrutivas sobre uma cópia de trabalho isolada.
            Antes de desbloquear o descarregamento, o motor executa uma verificação automática
            contra o ficheiro original.
          </p>
        </div>

        <button
          type="button"
          onClick={onResetOptions}
          className="inline-flex items-center gap-1.5 px-3 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 self-start sm:self-auto cursor-pointer"
        >
          <RotateCcw className="w-3.5 h-3.5" aria-hidden="true" />
          <span>Repor opções</span>
        </button>
      </div>

      {/* Immutable backup guarantee */}
      <div className="px-4 py-3 bg-white border border-slate-200 rounded-lg flex flex-wrap items-center justify-between gap-3 text-xs text-slate-700">
        <div className="flex items-center gap-2">
          <Lock className="w-4 h-4 text-emerald-700 shrink-0" aria-hidden="true" />
          <span>
            <strong className="font-semibold text-slate-900">Cópia de segurança original ativa:</strong>{' '}
            {originalAnalysis.fileName} ({formatBytes(originalAnalysis.fileSize)})
          </span>
        </div>
        <span className="font-mono text-slate-500">
          SHA-256 {shortHash(originalAnalysis.sha256)} · Bloqueios ativos: Fontes, Hiperligações,
          Animações, Multimédia
        </span>
      </div>

      {/* Non-destructive processing form */}
      <div className="bg-white border border-slate-200 rounded-lg divide-y divide-slate-200">
        {/* 1. Safe text replacement */}
        <section className="p-6 space-y-4">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div>
              <h2 className="text-base font-semibold text-slate-900">
                1. Substituição segura de texto (nós <code className="font-mono text-xs">&lt;a:t&gt;</code>)
              </h2>
              <p className="text-xs text-slate-600 mt-0.5">
                Altera apenas o conteúdo textual sem modificar <code className="font-mono">&lt;a:rPr&gt;</code>,
                mantendo intactos o tipo de letra, peso, cor, tamanho, espaçamento e posicionamento.
              </p>
            </div>
            <button
              type="button"
              onClick={addReplacementRule}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-blue-700 bg-blue-50 border border-blue-200 rounded-lg hover:bg-blue-100 transition-colors whitespace-nowrap shrink-0 self-start sm:self-auto cursor-pointer"
            >
              <Plus className="w-3.5 h-3.5" aria-hidden="true" />
              <span>Adicionar regra</span>
            </button>
          </div>

          {options.textReplacements.length === 0 ? (
            <p className="text-xs text-slate-500">
              Nenhuma substituição configurada. Sem regras ativas, o texto de todos os diapositivos
              permanece byte-a-byte idêntico ao original.
            </p>
          ) : (
            <div className="space-y-3">
              {options.textReplacements.map((rule) => (
                <div
                  key={rule.id}
                  className="grid grid-cols-1 md:grid-cols-[1fr_1fr_auto_auto] gap-3 items-center"
                >
                  <input
                    type="text"
                    value={rule.find}
                    onChange={(e) => updateRule(rule.id, { find: e.target.value })}
                    placeholder="Texto a procurar (ex.: 2026)"
                    className="px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-blue-600"
                  />
                  <input
                    type="text"
                    value={rule.replace}
                    onChange={(e) => updateRule(rule.id, { replace: e.target.value })}
                    placeholder="Substituir por (ex.: 2027)"
                    className="px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-blue-600"
                  />
                  <label className="inline-flex items-center gap-2 text-xs text-slate-700 whitespace-nowrap cursor-pointer">
                    <input
                      type="checkbox"
                      checked={rule.caseSensitive}
                      onChange={(e) => updateRule(rule.id, { caseSensitive: e.target.checked })}
                      className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                    />
                    <span>Maiúsculas/minúsculas</span>
                  </label>
                  <button
                    type="button"
                    onClick={() => removeRule(rule.id)}
                    aria-label="Remover regra de substituição"
                    className="p-2 text-slate-500 hover:text-red-700 hover:bg-red-50 rounded-lg transition-colors cursor-pointer"
                  >
                    <Trash2 className="w-4 h-4" aria-hidden="true" />
                  </button>
                </div>
              ))}
            </div>
          )}

          <label className="flex items-start gap-2.5 pt-2 text-xs text-slate-700 cursor-pointer">
            <input
              type="checkbox"
              checked={options.protectHyperlinkedText}
              onChange={(e) =>
                onChangeOptions({ ...options, protectHyperlinkedText: e.target.checked })
              }
              className="mt-0.5 rounded border-slate-300 text-blue-700 focus:ring-blue-600"
            />
            <span>
              <strong className="font-semibold text-slate-900">
                Proteger estritamente textos com hiperligações (recomendado).
              </strong>{' '}
              Ignora automaticamente qualquer fragmento de texto ligado a um{' '}
              <code className="font-mono">&lt;a:hlinkClick&gt;</code> para não colocar em risco o
              texto âncora nem o comportamento da ligação.
            </span>
          </label>
        </section>

        {/* 2. Safe Document Metadata */}
        <section className="p-6 space-y-4">
          <div className="flex items-center justify-between gap-4">
            <div>
              <h2 className="text-base font-semibold text-slate-900">
                2. Metadados do documento (<code className="font-mono text-xs">docProps/core.xml</code>)
              </h2>
              <p className="text-xs text-slate-600 mt-0.5">
                Atualize propriedades do ficheiro ou anonimize o autor sem tocar em nenhum
                diapositivo.
              </p>
            </div>
            <label className="inline-flex items-center gap-2 text-xs font-medium text-slate-800 whitespace-nowrap cursor-pointer">
              <input
                type="checkbox"
                checked={options.metadataUpdates.enabled}
                onChange={(e) =>
                  onChangeOptions({
                    ...options,
                    metadataUpdates: {
                      ...options.metadataUpdates,
                      enabled: e.target.checked,
                    },
                  })
                }
                className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
              />
              <span>Ativar edição de metadados</span>
            </label>
          </div>

          {options.metadataUpdates.enabled && (
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 pt-1">
              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">
                  Título da apresentação
                </label>
                <input
                  type="text"
                  value={options.metadataUpdates.title}
                  onChange={(e) =>
                    onChangeOptions({
                      ...options,
                      metadataUpdates: { ...options.metadataUpdates, title: e.target.value },
                    })
                  }
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-blue-600"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">Assunto</label>
                <input
                  type="text"
                  value={options.metadataUpdates.subject}
                  onChange={(e) =>
                    onChangeOptions({
                      ...options,
                      metadataUpdates: { ...options.metadataUpdates, subject: e.target.value },
                    })
                  }
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 focus:outline-none focus:border-blue-600"
                />
              </div>

              <div>
                <label className="block text-xs font-medium text-slate-700 mb-1">Autor</label>
                <input
                  type="text"
                  disabled={options.metadataUpdates.anonymizeAuthor}
                  value={
                    options.metadataUpdates.anonymizeAuthor
                      ? ''
                      : options.metadataUpdates.creator
                  }
                  placeholder={
                    options.metadataUpdates.anonymizeAuthor
                      ? '(Será removido por privacidade)'
                      : ''
                  }
                  onChange={(e) =>
                    onChangeOptions({
                      ...options,
                      metadataUpdates: { ...options.metadataUpdates, creator: e.target.value },
                    })
                  }
                  className="w-full px-3 py-2 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 disabled:bg-slate-50 disabled:text-slate-400 focus:outline-none focus:border-blue-600"
                />
              </div>

              <label className="md:col-span-3 inline-flex items-center gap-2 text-xs text-slate-700 cursor-pointer">
                <input
                  type="checkbox"
                  checked={options.metadataUpdates.anonymizeAuthor}
                  onChange={(e) =>
                    onChangeOptions({
                      ...options,
                      metadataUpdates: {
                        ...options.metadataUpdates,
                        anonymizeAuthor: e.target.checked,
                      },
                    })
                  }
                  className="rounded border-slate-300 text-blue-700 focus:ring-blue-600"
                />
                <span>
                  Remover identificação pessoal de autor e último editor antes de partilha externa
                </span>
              </label>
            </div>
          )}
        </section>

        {/* 3. Non-destructive slide visibility & speaker notes */}
        <section className="p-6 space-y-4">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              3. Visibilidade de diapositivos e notas do orador
            </h2>
            <p className="text-xs text-slate-600 mt-0.5">
              Ocultar um diapositivo define apenas o atributo <code className="font-mono">show="0"</code>
              — nenhum diapositivo é eliminado, mantendo intactas todas as hiperligações internas.
            </p>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2.5">
            {originalAnalysis.slides.map((slide) => {
              const isHidden =
                typeof options.slideVisibilityOverrides[slide.slideNumber] === 'boolean'
                  ? options.slideVisibilityOverrides[slide.slideNumber]
                  : slide.hidden;
              const isChanged = isHidden !== slide.hidden;
              return (
                <button
                  key={slide.path}
                  type="button"
                  onClick={() => toggleSlideHidden(slide.slideNumber, isHidden)}
                  className={`flex items-center justify-between gap-2 px-3 py-2 text-left text-xs border rounded-lg transition-colors cursor-pointer ${
                    isHidden
                      ? 'bg-slate-50 border-slate-300 text-slate-600'
                      : 'bg-white border-slate-200 text-slate-900 hover:border-slate-300'
                  }`}
                >
                  <span className="truncate">
                    <span className="font-mono text-slate-500 mr-1.5">
                      #{String(slide.slideNumber).padStart(2, '0')}
                    </span>
                    {slide.title}
                  </span>
                  <span
                    className={`whitespace-nowrap shrink-0 font-medium ${
                      isHidden
                        ? 'text-amber-800'
                        : isChanged
                        ? 'text-blue-700'
                        : 'text-emerald-800'
                    }`}
                  >
                    {isHidden ? 'Oculto' : 'Visível'}
                  </span>
                </button>
              );
            })}
          </div>

          <div className="pt-2 border-t border-slate-100">
            <label className="inline-flex items-start gap-2.5 text-xs text-slate-700 cursor-pointer">
              <input
                type="checkbox"
                disabled={slidesWithNotesCount === 0}
                checked={options.cleanSpeakerNotes}
                onChange={(e) =>
                  onChangeOptions({ ...options, cleanSpeakerNotes: e.target.checked })
                }
                className="mt-0.5 rounded border-slate-300 text-blue-700 focus:ring-blue-600 disabled:opacity-50"
              />
              <span>
                <strong className="font-medium text-slate-900">
                  Limpar texto das notas do orador ({slidesWithNotesCount} diapositivo(s) com notas)
                </strong>
                <span className="block text-slate-500 mt-0.5">
                  Atua apenas sobre <code className="font-mono">ppt/notesSlides/</code>, deixando o
                  conteúdo visual dos diapositivos 100% intocado.
                </span>
              </span>
            </label>
          </div>
        </section>

        {/* Action footer */}
        <div className="px-6 py-4 bg-slate-50 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <p className="text-xs text-slate-600">
            O processamento gera uma cópia de trabalho, reanalisa o PPTX produzido e executa 7
            verificações obrigatórias antes de permitir o descarregamento.
          </p>
          <button
            type="button"
            disabled={isProcessing}
            onClick={onRunProcess}
            className="inline-flex items-center justify-center gap-2 px-5 py-2.5 text-sm font-semibold text-white bg-blue-700 rounded-lg hover:bg-blue-800 disabled:opacity-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            <ShieldCheck className="w-4 h-4" aria-hidden="true" />
            <span>{isProcessing ? 'A processar e verificar…' : 'Processar e verificar PPTX'}</span>
          </button>
        </div>
      </div>

      {/* Automatic Verification Report & Export Gate */}
      {result && (
        <section className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          <div className="p-6 border-b border-slate-200 flex flex-col lg:flex-row lg:items-center justify-between gap-4">
            <div className="flex items-start gap-3">
              {result.aborted || !result.verification.canExport ? (
                <AlertCircle className="w-6 h-6 text-red-700 shrink-0 mt-0.5" aria-hidden="true" />
              ) : result.verification.overallStatus === 'verified_with_limitations' ? (
                <AlertTriangle
                  className="w-6 h-6 text-amber-700 shrink-0 mt-0.5"
                  aria-hidden="true"
                />
              ) : (
                <CheckCircle2
                  className="w-6 h-6 text-emerald-700 shrink-0 mt-0.5"
                  aria-hidden="true"
                />
              )}
              <div>
                <h2 className="text-lg font-bold text-slate-900">
                  {result.aborted || !result.verification.canExport
                    ? 'Exportação cancelada por segurança'
                    : 'Verificação concluída — Ficheiro pronto para exportação'}
                </h2>
                <p className="text-sm text-slate-600 mt-0.5">
                  {result.aborted || !result.verification.canExport
                    ? result.abortReason
                    : `Todas as verificações estruturais, tipográficas, de hiperligações, animações e multimédia foram aprovadas para ${result.processedFileName}.`}
                </p>
              </div>
            </div>

            <div className="flex flex-wrap items-center gap-3 shrink-0">
              {result.comparison && (
                <button
                  type="button"
                  onClick={onGoToComparison}
                  className="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
                >
                  <GitCompare className="w-4 h-4" aria-hidden="true" />
                  <span>Comparar original vs. final</span>
                </button>
              )}

              <button
                type="button"
                disabled={!result.verification.canExport || !result.processedBuffer}
                onClick={() => {
                  if (result.processedBuffer) {
                    downloadUint8Array(result.processedBuffer, result.processedFileName);
                  }
                }}
                className="inline-flex items-center gap-2 px-5 py-2 text-sm font-semibold text-white bg-emerald-700 rounded-lg hover:bg-emerald-800 disabled:bg-slate-300 disabled:text-slate-500 disabled:cursor-not-allowed transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                <Download className="w-4 h-4" aria-hidden="true" />
                <span>Exportar PPTX verificado</span>
              </button>
            </div>
          </div>

          {/* Operations applied & protected runs */}
          <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 space-y-2 text-xs text-slate-700">
            <div>
              <strong className="font-semibold text-slate-900">Operações executadas:</strong>{' '}
              {result.operationsApplied.join(' · ')}
            </div>
            {result.skippedProtectedRuns.length > 0 && (
              <div className="text-amber-900">
                <strong className="font-semibold">
                  Proteção de hiperligações ativada ({result.skippedProtectedRuns.length}{' '}
                  fragmento(s) preservado(s) sem alteração):
                </strong>{' '}
                {result.skippedProtectedRuns
                  .map((s) => `Diap. #${s.slideNumber} ("${s.text.slice(0, 40)}")`)
                  .join(', ')}
              </div>
            )}
          </div>

          {/* 7 Verification checks table */}
          {result.verification.checks.length > 0 && (
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse text-sm">
                <thead>
                  <tr className="border-b border-slate-200 bg-white text-xs font-medium text-slate-500">
                    <th className="py-3 pl-6 pr-3">Verificação automática</th>
                    <th className="py-3 px-3">Original</th>
                    <th className="py-3 px-3">Final</th>
                    <th className="py-3 px-3">Estado</th>
                    <th className="py-3 pl-3 pr-6">Detalhes</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {result.verification.checks.map((check) => (
                    <tr key={check.id} className="align-top">
                      <td className="py-3.5 pl-6 pr-3 font-medium text-slate-900">
                        {check.title}
                      </td>
                      <td className="py-3.5 px-3 font-mono text-xs tabular-nums text-slate-600 whitespace-nowrap">
                        {check.originalSummary}
                      </td>
                      <td className="py-3.5 px-3 font-mono text-xs tabular-nums text-slate-900 whitespace-nowrap">
                        {check.processedSummary}
                      </td>
                      <td className="py-3.5 px-3">{renderCheckStatus(check.status)}</td>
                      <td className="py-3.5 pl-3 pr-6 text-xs text-slate-600 max-w-md leading-relaxed">
                        <div>{check.details}</div>
                        {check.limitationNote && (
                          <div className="mt-1.5 text-amber-800">
                            <strong>Limitação transparente:</strong> {check.limitationNote}
                          </div>
                        )}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* Explicit Limitations summary */}
          {result.verification.limitations.length > 0 && (
            <div className="px-6 py-4 bg-amber-50/70 border-t border-amber-200 text-xs text-amber-950 space-y-1.5">
              <p className="font-semibold">
                Limitações de verificação fora do motor nativo do Microsoft PowerPoint:
              </p>
              <ul className="list-disc pl-5 space-y-1 text-amber-900">
                {result.verification.limitations.map((lim, idx) => (
                  <li key={idx}>{lim}</li>
                ))}
              </ul>
            </div>
          )}
        </section>
      )}
    </div>
  );
};

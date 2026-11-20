import React, { useMemo, useState } from 'react';
import { ArrowRight, CheckCircle2, Download, FileDiff } from 'lucide-react';
import type { PptxAnalysis, ProcessingResult, WorkspaceSection } from '../types/pptx';
import { downloadUint8Array, formatBytes, shortHash } from '../utils/formatters';

interface ComparisonSectionProps {
  original: PptxAnalysis;
  result: ProcessingResult | null;
  onNavigate: (section: WorkspaceSection) => void;
}

export const ComparisonSection: React.FC<ComparisonSectionProps> = ({
  original,
  result,
  onNavigate,
}) => {
  const [viewMode, setViewMode] = useState<'slides' | 'files'>('slides');
  const [fileFilter, setFileFilter] = useState<'all' | 'modified' | 'unchanged'>('all');

  const filteredFileDiffs = useMemo(() => {
    if (!result?.comparison) return [];
    return result.comparison.fileDiffs.filter((f) => {
      if (fileFilter === 'modified') return f.status !== 'unchanged';
      if (fileFilter === 'unchanged') return f.status === 'unchanged';
      return true;
    });
  }, [result, fileFilter]);

  if (!result || !result.processedAnalysis || !result.comparison) {
    return (
      <div className="space-y-6">
        <div className="pb-5 border-b border-slate-200">
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">
            Comparação: Original vs. Ficheiro Final
          </h1>
          <p className="mt-1.5 text-sm text-slate-600">
            Execute o processamento seguro para gerar a comparação estrutural, por diapositivo e
            ficheiro-a-ficheiro entre o original imutável e o PPTX final.
          </p>
        </div>

        <div className="p-10 bg-white border border-slate-200 rounded-lg text-center max-w-xl mx-auto">
          <div className="w-11 h-11 rounded-lg bg-slate-100 border border-slate-200 flex items-center justify-center mx-auto mb-4">
            <FileDiff className="w-5 h-5 text-slate-600" aria-hidden="true" />
          </div>
          <h2 className="text-base font-semibold text-slate-900">
            Ainda não existe um ficheiro processado para comparar
          </h2>
          <p className="mt-2 text-sm text-slate-600">
            Aceda à área de processamento para validar ou aplicar edições seguras e comparar o
            resultado final com a cópia original.
          </p>
          <button
            type="button"
            onClick={() => onNavigate('process_export')}
            className="mt-5 inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-blue-700 rounded-lg hover:bg-blue-800 transition-colors cursor-pointer"
          >
            <span>Ir para Processar e Exportar</span>
            <ArrowRight className="w-4 h-4" aria-hidden="true" />
          </button>
        </div>
      </div>
    );
  }

  const processed = result.processedAnalysis;
  const comp = result.comparison;

  const metricRows = [
    {
      label: 'Número de diapositivos',
      orig: String(original.slideCount),
      proc: String(processed.slideCount),
      match: original.slideCount === processed.slideCount,
    },
    {
      label: 'Fontes referenciadas / incorporadas',
      orig: `${original.fonts.length} (${original.embeddedFontFlags.embeddedFontFiles.length} inc.)`,
      proc: `${processed.fonts.length} (${processed.embeddedFontFlags.embeddedFontFiles.length} inc.)`,
      match:
        original.fonts.length === processed.fonts.length &&
        original.embeddedFontFlags.embeddedFontFiles.length ===
          processed.embeddedFontFlags.embeddedFontFiles.length,
    },
    {
      label: 'Hiperligações totais',
      orig: String(original.hyperlinks.length),
      proc: String(processed.hyperlinks.length),
      match: original.hyperlinks.length === processed.hyperlinks.length,
    },
    {
      label: 'Recursos multimédia e objetos',
      orig: String(original.multimedia.length),
      proc: String(processed.multimedia.length),
      match: original.multimedia.length === processed.multimedia.length,
    },
    {
      label: 'Animações e transições',
      orig: `${original.visualSummary.totalAnimations} anim. · ${original.visualSummary.slidesWithTransitions} trans.`,
      proc: `${processed.visualSummary.totalAnimations} anim. · ${processed.visualSummary.slidesWithTransitions} trans.`,
      match:
        original.visualSummary.totalAnimations === processed.visualSummary.totalAnimations &&
        original.visualSummary.slidesWithTransitions ===
          processed.visualSummary.slidesWithTransitions,
    },
    {
      label: 'Formas, imagens e tabelas',
      orig: `${original.visualSummary.totalShapes} / ${original.visualSummary.totalImages} / ${original.visualSummary.totalTables}`,
      proc: `${processed.visualSummary.totalShapes} / ${processed.visualSummary.totalImages} / ${processed.visualSummary.totalTables}`,
      match:
        original.visualSummary.totalShapes === processed.visualSummary.totalShapes &&
        original.visualSummary.totalImages === processed.visualSummary.totalImages &&
        original.visualSummary.totalTables === processed.visualSummary.totalTables,
    },
    {
      label: 'Tamanho do ficheiro .pptx',
      orig: formatBytes(original.fileSize),
      proc: formatBytes(processed.fileSize),
      match: true,
    },
    {
      label: 'Assinatura SHA-256 do pacote',
      orig: shortHash(original.sha256),
      proc: shortHash(processed.sha256),
      match: true,
    },
  ];

  return (
    <div className="space-y-8">
      <div className="pb-5 border-b border-slate-200 flex flex-col sm:flex-row sm:items-end justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">
            Comparação: Original vs. Ficheiro Final
          </h1>
          <p className="mt-1.5 text-sm text-slate-600">
            <span className="font-mono tabular-nums font-medium text-slate-900">
              {comp.unchangedEntriesCount}
            </span>{' '}
            de{' '}
            <span className="font-mono tabular-nums font-medium text-slate-900">
              {comp.totalEntriesCount}
            </span>{' '}
            ficheiros internos OOXML mantiveram-se byte-a-byte idênticos ·{' '}
            <span className="font-mono tabular-nums font-medium text-blue-700">
              {comp.modifiedEntriesCount}
            </span>{' '}
            ficheiro(s) atualizado(s) de forma controlada.
          </p>
        </div>

        {result.verification.canExport && result.processedBuffer && (
          <button
            type="button"
            onClick={() => downloadUint8Array(result.processedBuffer!, result.processedFileName)}
            className="inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-emerald-700 rounded-lg hover:bg-emerald-800 transition-colors whitespace-nowrap shrink-0 self-start sm:self-auto cursor-pointer"
          >
            <Download className="w-4 h-4" aria-hidden="true" />
            <span>Exportar PPTX verificado</span>
          </button>
        )}
      </div>

      {/* Side-by-side summary table */}
      <section className="bg-white border border-slate-200 rounded-lg overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200">
          <h2 className="text-base font-semibold text-slate-900">
            Quadro comparativo de fidelidade
          </h2>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-sm">
            <thead>
              <tr className="border-b border-slate-200 bg-slate-50 text-xs font-medium text-slate-600">
                <th className="py-3 pl-6 pr-3">Propriedade</th>
                <th className="py-3 px-3">Original ({original.fileName})</th>
                <th className="py-3 px-3">Final ({processed.fileName})</th>
                <th className="py-3 pl-3 pr-6 text-right">Paridade</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {metricRows.map((row) => (
                <tr key={row.label}>
                  <td className="py-3 pl-6 pr-3 font-medium text-slate-900">{row.label}</td>
                  <td className="py-3 px-3 font-mono text-xs tabular-nums text-slate-700">
                    {row.orig}
                  </td>
                  <td className="py-3 px-3 font-mono text-xs tabular-nums text-slate-900">
                    {row.proc}
                  </td>
                  <td className="py-3 pl-3 pr-6 text-right whitespace-nowrap">
                    {row.match ? (
                      <span className="inline-flex items-center gap-1.5 text-xs font-medium text-emerald-800">
                        <CheckCircle2
                          className="w-3.5 h-3.5 text-emerald-700 shrink-0"
                          aria-hidden="true"
                        />
                        <span>Preservado</span>
                      </span>
                    ) : (
                      <span className="text-xs font-semibold text-red-700">Divergência</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {/* Switcher: Slides vs Internal ZIP files */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div className="inline-flex items-center gap-1 p-1 bg-slate-100 border border-slate-200 rounded-lg self-start">
          <button
            type="button"
            onClick={() => setViewMode('slides')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              viewMode === 'slides'
                ? 'bg-white text-slate-900 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Comparação por diapositivo ({comp.slideDiffs.length})
          </button>
          <button
            type="button"
            onClick={() => setViewMode('files')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              viewMode === 'files'
                ? 'bg-white text-slate-900 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Ficheiros OOXML internos ({comp.totalEntriesCount})
          </button>
        </div>

        {viewMode === 'files' && (
          <div className="inline-flex items-center gap-1 p-1 bg-slate-100 border border-slate-200 rounded-lg self-start">
            {(
              [
                ['all', `Todos (${comp.totalEntriesCount})`],
                ['modified', `Modificados (${comp.modifiedEntriesCount})`],
                ['unchanged', `Intactos (${comp.unchangedEntriesCount})`],
              ] as const
            ).map(([key, label]) => (
              <button
                key={key}
                type="button"
                onClick={() => setFileFilter(key)}
                className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                  fileFilter === key
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                {label}
              </button>
            ))}
          </div>
        )}
      </div>

      {viewMode === 'slides' ? (
        <div className="bg-white border border-slate-200 rounded-lg divide-y divide-slate-200">
          {comp.slideDiffs.map((sd) => {
            const hasAnyDiff = !sd.xmlUnchanged || sd.notesChanged;
            return (
              <div key={sd.slideNumber} className="p-5 space-y-3">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <div className="flex items-center gap-2.5 min-w-0">
                    <span className="font-mono text-xs text-slate-500 tabular-nums">
                      #{String(sd.slideNumber).padStart(2, '0')}
                    </span>
                    <h3 className="text-sm font-semibold text-slate-900 truncate">{sd.title}</h3>
                  </div>
                  <span
                    className={`text-xs font-medium whitespace-nowrap ${
                      hasAnyDiff ? 'text-blue-700' : 'text-emerald-800'
                    }`}
                  >
                    {sd.xmlUnchanged && !sd.notesChanged
                      ? 'XML byte-a-byte idêntico ao original'
                      : sd.xmlUnchanged && sd.notesChanged
                      ? 'Diapositivo intacto · Notas atualizadas'
                      : 'Atualizado com preservação de estrutura'}
                  </span>
                </div>

                {sd.visibilityChanged && (
                  <p className="text-xs text-slate-700">
                    Visibilidade:{' '}
                    <span className="font-mono">
                      {sd.originalHidden ? 'Oculto' : 'Visível'} →{' '}
                      {sd.processedHidden ? 'Oculto' : 'Visível'}
                    </span>
                  </p>
                )}

                {sd.notesChanged && (
                  <p className="text-xs text-slate-700">
                    Notas do orador: texto removido em <code className="font-mono">notesSlides</code>{' '}
                    (diapositivo visual inalterado).
                  </p>
                )}

                {sd.textChanges.length > 0 && (
                  <div className="space-y-2 pt-1">
                    {sd.textChanges.map((tc, i) => (
                      <div
                        key={i}
                        className="grid grid-cols-1 md:grid-cols-2 gap-3 text-xs p-3 bg-slate-50 border border-slate-200 rounded-md"
                      >
                        <div>
                          <span className="block text-slate-500 mb-1 font-medium">Original</span>
                          <p className="text-slate-700 break-words">{tc.before}</p>
                        </div>
                        <div>
                          <span className="block text-blue-700 mb-1 font-medium">
                            Final (formatação preservada)
                          </span>
                          <p className="text-slate-900 font-medium break-words">{tc.after}</p>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      ) : (
        <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 text-slate-600 font-medium">
                  <th className="py-2.5 pl-6 pr-3">Ficheiro interno OOXML</th>
                  <th className="py-2.5 px-3">Estado</th>
                  <th className="py-2.5 px-3 text-right">Tamanho original</th>
                  <th className="py-2.5 px-3 text-right">Tamanho final</th>
                  <th className="py-2.5 pl-3 pr-6">Nota técnica</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredFileDiffs.length === 0 ? (
                  <tr>
                    <td colSpan={5} className="py-8 text-center text-slate-500">
                      Nenhum ficheiro corresponde ao filtro selecionado.
                    </td>
                  </tr>
                ) : (
                  filteredFileDiffs.map((fd) => (
                    <tr key={fd.path} className="hover:bg-slate-50/70">
                      <td className="py-2.5 pl-6 pr-3 font-mono text-slate-800">{fd.path}</td>
                      <td className="py-2.5 px-3 whitespace-nowrap">
                        {fd.status === 'unchanged' ? (
                          <span className="text-emerald-800 font-medium">Intacto byte-a-byte</span>
                        ) : fd.status === 'modified' ? (
                          <span className="text-blue-700 font-semibold">Editado</span>
                        ) : fd.status === 'added' ? (
                          <span className="text-amber-800 font-semibold">Adicionado</span>
                        ) : (
                          <span className="text-red-700 font-semibold">Removido</span>
                        )}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-600">
                        {formatBytes(fd.originalSize)}
                      </td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-900">
                        {formatBytes(fd.processedSize)}
                      </td>
                      <td className="py-2.5 pl-3 pr-6 text-slate-600">{fd.reason}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
};

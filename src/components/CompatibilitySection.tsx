import React, { useMemo, useState } from 'react';
import { AlertCircle, AlertTriangle, CheckCircle2, Info, Search } from 'lucide-react';
import type { CompatibilityIssue, PptxAnalysis } from '../types/pptx';
import { formatBytes, shortHash } from '../utils/formatters';

interface CompatibilitySectionProps {
  analysis: PptxAnalysis;
}

export const CompatibilitySection: React.FC<CompatibilitySectionProps> = ({ analysis }) => {
  const [entryFilter, setEntryFilter] = useState<'all' | 'xml' | 'rels' | 'media' | 'fonts'>('all');
  const [entrySearch, setEntrySearch] = useState('');

  const filteredEntries = useMemo(() => {
    const q = entrySearch.trim().toLowerCase();
    return analysis.structure.entries.filter((e) => {
      if (entryFilter === 'xml' && !e.isXml) return false;
      if (entryFilter === 'rels' && !e.isRels) return false;
      if (entryFilter === 'media' && !e.isMedia) return false;
      if (entryFilter === 'fonts' && !e.isFont) return false;
      if (q && !e.path.toLowerCase().includes(q)) return false;
      return true;
    });
  }, [analysis.structure.entries, entryFilter, entrySearch]);

  const renderSeverity = (severity: CompatibilityIssue['severity']) => {
    if (severity === 'critical') {
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-semibold text-red-800">
          <AlertCircle className="w-4 h-4 text-red-700 shrink-0" aria-hidden="true" />
          <span>Crítico</span>
        </span>
      );
    }
    if (severity === 'warning') {
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-semibold text-amber-800">
          <AlertTriangle className="w-4 h-4 text-amber-700 shrink-0" aria-hidden="true" />
          <span>Aviso de compatibilidade</span>
        </span>
      );
    }
    return (
      <span className="inline-flex items-center gap-1.5 text-xs font-medium text-blue-800">
        <Info className="w-4 h-4 text-blue-700 shrink-0" aria-hidden="true" />
        <span>Informação técnica</span>
      </span>
    );
  };

  return (
    <div className="space-y-8">
      <div className="pb-5 border-b border-slate-200">
        <h1 className="text-2xl font-bold tracking-tight text-slate-900">
          Diagnóstico de compatibilidade e integridade OOXML
        </h1>
        <p className="mt-1.5 text-sm text-slate-600 max-w-3xl">
          Identificação preventiva de dependências externas (fontes não incorporadas, ficheiros
          ligados, codecs legados) e inspeção direta da estrutura ZIP/XML do ficheiro .pptx.
        </p>
      </div>

      {/* Diagnostics list */}
      <section className="space-y-3">
        <h2 className="text-base font-semibold text-slate-900">
          Pontos de compatibilidade detetados ({analysis.compatibilityIssues.length})
        </h2>

        {analysis.compatibilityIssues.length === 0 ? (
          <div className="p-6 bg-white border border-slate-200 rounded-lg flex items-center gap-3 text-sm text-emerald-900">
            <CheckCircle2 className="w-5 h-5 text-emerald-700 shrink-0" aria-hidden="true" />
            <span>
              Não foram detetados riscos de compatibilidade externa. Todas as fontes, relações e
              recursos estão autocontidos no pacote PPTX.
            </span>
          </div>
        ) : (
          <div className="bg-white border border-slate-200 rounded-lg divide-y divide-slate-200">
            {analysis.compatibilityIssues.map((issue) => (
              <div key={issue.id} className="p-5 space-y-2">
                <div className="flex flex-wrap items-center justify-between gap-2">
                  <h3 className="text-sm font-semibold text-slate-900">{issue.title}</h3>
                  {renderSeverity(issue.severity)}
                </div>
                <p className="text-sm text-slate-600 leading-relaxed">{issue.description}</p>
                <div className="pt-1 flex flex-wrap items-center gap-x-4 gap-y-1 text-xs text-slate-600">
                  <span>
                    <strong className="font-medium text-slate-800">Estado:</strong>{' '}
                    {issue.preservationStatus}
                  </span>
                  {issue.affectedSlides.length > 0 && (
                    <span className="font-mono tabular-nums">
                      Diapositivos: {issue.affectedSlides.map((n) => `#${n}`).join(', ')}
                    </span>
                  )}
                </div>
                <p className="text-xs text-slate-500">
                  <strong className="font-medium text-slate-700">Recomendação:</strong>{' '}
                  {issue.recommendation}
                </p>
              </div>
            ))}
          </div>
        )}
      </section>

      {/* Internal OOXML package inspector */}
      <section className="space-y-4">
        <div className="flex flex-col lg:flex-row lg:items-end justify-between gap-3">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              Estrutura interna do pacote PPTX ({analysis.structure.totalEntries} ficheiros)
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              A aplicação opera diretamente sobre o pacote OOXML, preservando ficheiros não
              editados byte-a-byte e mantendo elementos XML desconhecidos.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <div className="inline-flex items-center gap-1 p-1 bg-slate-100 border border-slate-200 rounded-lg">
              {(
                [
                  ['all', 'Todos'],
                  ['xml', `XML (${analysis.structure.xmlCount})`],
                  ['rels', `Relações (${analysis.structure.relsCount})`],
                  ['media', `Media (${analysis.structure.mediaCount})`],
                  ['fonts', `Fontes (${analysis.structure.fontFilesCount})`],
                ] as const
              ).map(([key, label]) => (
                <button
                  key={key}
                  type="button"
                  onClick={() => setEntryFilter(key)}
                  className={`px-2.5 py-1 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                    entryFilter === key
                      ? 'bg-white text-slate-900 shadow-xs'
                      : 'text-slate-600 hover:text-slate-900'
                  }`}
                >
                  {label}
                </button>
              ))}
            </div>

            <div className="relative w-56">
              <Search
                className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-1/2 -translate-y-1/2 pointer-events-none"
                aria-hidden="true"
              />
              <input
                type="text"
                value={entrySearch}
                onChange={(e) => setEntrySearch(e.target.value)}
                placeholder="Filtrar caminho interno…"
                className="w-full pl-8 pr-2.5 py-1.5 text-xs bg-white border border-slate-300 rounded-lg text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-blue-600"
              />
            </div>
          </div>
        </div>

        <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          <div className="max-h-96 overflow-y-auto">
            <table className="w-full text-left border-collapse text-xs">
              <thead className="sticky top-0 bg-slate-50 border-b border-slate-200 text-slate-600 font-medium">
                <tr>
                  <th className="py-2.5 pl-6 pr-3">Caminho interno no pacote ZIP</th>
                  <th className="py-2.5 px-3">Tipo</th>
                  <th className="py-2.5 px-3 text-right">Tamanho</th>
                  <th className="py-2.5 pl-3 pr-6 text-right">Assinatura SHA-256</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredEntries.map((entry) => {
                  const typeLabel = entry.isFont
                    ? 'Fonte incorporada'
                    : entry.isMedia
                    ? 'Multimédia'
                    : entry.isRels
                    ? 'Relações OOXML'
                    : entry.isXml
                    ? 'XML OOXML'
                    : 'Recurso binário';
                  return (
                    <tr key={entry.path} className="hover:bg-slate-50/70">
                      <td className="py-2.5 pl-6 pr-3 font-mono text-slate-800">{entry.path}</td>
                      <td className="py-2.5 px-3 text-slate-600 whitespace-nowrap">{typeLabel}</td>
                      <td className="py-2.5 px-3 text-right font-mono tabular-nums text-slate-700 whitespace-nowrap">
                        {formatBytes(entry.size)}
                      </td>
                      <td
                        className="py-2.5 pl-3 pr-6 text-right font-mono text-slate-500 whitespace-nowrap"
                        title={entry.sha256}
                      >
                        {shortHash(entry.sha256)}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </section>
    </div>
  );
};

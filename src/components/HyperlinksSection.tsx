import React, { useMemo, useState } from 'react';
import { AlertTriangle, CheckCircle2, Search } from 'lucide-react';
import type { HyperlinkAuditEntry, PptxAnalysis } from '../types/pptx';

interface HyperlinksSectionProps {
  analysis: PptxAnalysis;
}

type LinkFilter = 'all' | HyperlinkAuditEntry['linkKind'];

const KIND_LABELS: Record<HyperlinkAuditEntry['linkKind'], string> = {
  web: 'Página Web',
  slide: 'Diapositivo interno',
  external_file: 'Ficheiro externo',
  email: 'Email',
  action: 'Ação interativa',
};

const ELEMENT_LABELS: Record<HyperlinkAuditEntry['elementType'], string> = {
  text: 'Texto',
  image: 'Imagem',
  shape: 'Forma / Botão',
  interactive: 'Elemento interativo',
};

export const HyperlinksSection: React.FC<HyperlinksSectionProps> = ({ analysis }) => {
  const [filter, setFilter] = useState<LinkFilter>('all');
  const [query, setQuery] = useState('');

  const counts = useMemo(() => {
    const map = {
      all: analysis.hyperlinks.length,
      web: 0,
      slide: 0,
      external_file: 0,
      email: 0,
      action: 0,
    };
    for (const hl of analysis.hyperlinks) {
      map[hl.linkKind]++;
    }
    return map;
  }, [analysis.hyperlinks]);

  const filteredLinks = useMemo(() => {
    const q = query.trim().toLowerCase();
    return analysis.hyperlinks.filter((hl) => {
      if (filter !== 'all' && hl.linkKind !== filter) return false;
      if (
        q &&
        !hl.target.toLowerCase().includes(q) &&
        !hl.elementLabel.toLowerCase().includes(q) &&
        !hl.tooltip.toLowerCase().includes(q)
      ) {
        return false;
      }
      return true;
    });
  }, [analysis.hyperlinks, filter, query]);

  return (
    <div className="space-y-6">
      <div className="pb-5 border-b border-slate-200">
        <h1 className="text-2xl font-bold tracking-tight text-slate-900">
          Auditoria e preservação de hiperligações
        </h1>
        <p className="mt-1.5 text-sm text-slate-600 max-w-3xl">
          Todas as hiperligações associadas a textos, imagens, formas, botões e elementos
          interativos — bem como as respetivas relações internas (
          <code className="font-mono text-xs">.rels</code>) e ações{' '}
          <code className="font-mono text-xs">ppaction://</code> — são preservadas sem modificação
          nem desativação durante o processamento.
        </p>
      </div>

      {/* Filter & search */}
      <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3">
        <div className="inline-flex items-center gap-1 p-1 bg-slate-100 border border-slate-200 rounded-lg overflow-x-auto">
          {(
            [
              ['all', `Todas (${counts.all})`],
              ['web', `Web (${counts.web})`],
              ['slide', `Diapositivos (${counts.slide})`],
              ['external_file', `Ficheiros (${counts.external_file})`],
              ['email', `Email (${counts.email})`],
              ['action', `Ações (${counts.action})`],
            ] as [LinkFilter, string][]
          ).map(([key, label]) => (
            <button
              key={key}
              type="button"
              onClick={() => setFilter(key)}
              className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
                filter === key
                  ? 'bg-white text-slate-900 shadow-xs'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              {label}
            </button>
          ))}
        </div>

        <div className="relative w-full lg:w-72">
          <Search
            className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none"
            aria-hidden="true"
          />
          <input
            type="text"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Pesquisar URL, destino ou texto…"
            className="w-full pl-9 pr-3 py-1.5 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-blue-600"
          />
        </div>
      </div>

      {/* Table */}
      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-sm">
            <thead>
              <tr className="border-b border-slate-200 bg-slate-50 text-xs font-medium text-slate-600">
                <th className="py-3 pl-6 pr-3 w-20">Diap.</th>
                <th className="py-3 px-3">Elemento associado</th>
                <th className="py-3 px-3">Tipo</th>
                <th className="py-3 px-3">Destino original</th>
                <th className="py-3 px-3">Relação</th>
                <th className="py-3 pl-3 pr-6">Integridade</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredLinks.length === 0 ? (
                <tr>
                  <td colSpan={6} className="py-10 text-center text-sm text-slate-500">
                    {analysis.hyperlinks.length === 0
                      ? 'Esta apresentação não contém hiperligações.'
                      : 'Nenhuma hiperligação corresponde ao filtro aplicado.'}
                  </td>
                </tr>
              ) : (
                filteredLinks.map((hl) => (
                  <tr key={hl.id} className="hover:bg-slate-50/70 align-top">
                    <td className="py-3.5 pl-6 pr-3 font-mono tabular-nums text-slate-700">
                      #{String(hl.slideNumber).padStart(2, '0')}
                    </td>
                    <td className="py-3.5 px-3 max-w-xs">
                      <div className="font-medium text-slate-900 break-words">
                        {hl.elementLabel}
                      </div>
                      <div className="text-xs text-slate-500 mt-0.5">
                        {ELEMENT_LABELS[hl.elementType]}
                        {hl.tooltip ? ` · Dica: ${hl.tooltip}` : ''}
                      </div>
                    </td>
                    <td className="py-3.5 px-3 text-xs text-slate-700 whitespace-nowrap">
                      {KIND_LABELS[hl.linkKind]}
                    </td>
                    <td className="py-3.5 px-3 max-w-sm">
                      <div className="font-mono text-xs text-slate-800 break-all">{hl.target}</div>
                      {hl.action && hl.action !== hl.target && (
                        <div className="font-mono text-[11px] text-slate-500 mt-0.5 break-all">
                          {hl.action}
                        </div>
                      )}
                      {hl.warning && (
                        <p className="text-xs text-amber-800 mt-1 leading-relaxed">{hl.warning}</p>
                      )}
                    </td>
                    <td className="py-3.5 px-3 font-mono text-xs text-slate-600 whitespace-nowrap">
                      {hl.relId || 'Ação direta'}
                    </td>
                    <td className="py-3.5 pl-3 pr-6 whitespace-nowrap">
                      {hl.isValidRel ? (
                        <span className="inline-flex items-center gap-1.5 text-xs font-medium text-emerald-800">
                          <CheckCircle2
                            className="w-3.5 h-3.5 text-emerald-700 shrink-0"
                            aria-hidden="true"
                          />
                          <span>Preservada</span>
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1.5 text-xs font-medium text-red-800">
                          <AlertTriangle
                            className="w-3.5 h-3.5 text-red-700 shrink-0"
                            aria-hidden="true"
                          />
                          <span>Relação ausente</span>
                        </span>
                      )}
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

import React, { useMemo, useState } from 'react';
import { AlertTriangle, CheckCircle2, Info, Search } from 'lucide-react';
import type { FontAuditEntry, PptxAnalysis } from '../types/pptx';
import { formatBytes } from '../utils/formatters';

interface FontsSectionProps {
  analysis: PptxAnalysis;
}

type FontFilter = 'all' | 'external_unembedded' | 'embedded' | 'standard';

export const FontsSection: React.FC<FontsSectionProps> = ({ analysis }) => {
  const [filter, setFilter] = useState<FontFilter>('all');
  const [query, setQuery] = useState('');

  const counts = useMemo(() => {
    return {
      all: analysis.fonts.length,
      external_unembedded: analysis.fonts.filter((f) => f.status === 'external_unembedded').length,
      embedded: analysis.fonts.filter((f) => f.status === 'embedded').length,
      standard: analysis.fonts.filter((f) => f.status === 'standard').length,
    };
  }, [analysis.fonts]);

  const filteredFonts = useMemo(() => {
    return analysis.fonts.filter((f) => {
      if (filter !== 'all' && f.status !== filter) return false;
      if (query.trim() && !f.name.toLowerCase().includes(query.trim().toLowerCase())) {
        return false;
      }
      return true;
    });
  }, [analysis.fonts, filter, query]);

  const renderStatus = (font: FontAuditEntry) => {
    if (font.status === 'embedded') {
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-medium text-emerald-800">
          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-700 shrink-0" aria-hidden="true" />
          <span>Incorporada no ficheiro</span>
        </span>
      );
    }
    if (font.status === 'external_unembedded') {
      return (
        <span className="inline-flex items-center gap-1.5 text-xs font-medium text-amber-800">
          <AlertTriangle className="w-3.5 h-3.5 text-amber-700 shrink-0" aria-hidden="true" />
          <span>Externa (não incorporada)</span>
        </span>
      );
    }
    return (
      <span className="inline-flex items-center gap-1.5 text-xs font-medium text-slate-700">
        <Info className="w-3.5 h-3.5 text-slate-500 shrink-0" aria-hidden="true" />
        <span>Sistema padrão</span>
      </span>
    );
  };

  return (
    <div className="space-y-6">
      <div className="pb-5 border-b border-slate-200">
        <h1 className="text-2xl font-bold tracking-tight text-slate-900">
          Auditoria e preservação de fontes
        </h1>
        <p className="mt-1.5 text-sm text-slate-600 max-w-3xl">
          Todas as declarações tipográficas (<code className="font-mono text-xs">a:latin</code>,{' '}
          <code className="font-mono text-xs">a:ea</code>,{' '}
          <code className="font-mono text-xs">a:cs</code>), ficheiros incorporados em{' '}
          <code className="font-mono text-xs">ppt/fonts/</code> e estilos (negrito, itálico,
          sublinhado e tamanhos) são mantidos intactos. Nenhuma fonte desconhecida é substituída
          automaticamente por Arial ou Calibri.
        </p>
      </div>

      {/* Warning banner if external unembedded fonts exist */}
      {counts.external_unembedded > 0 && (
        <div className="p-4 bg-amber-50 border border-amber-200 rounded-lg flex items-start gap-3 text-sm text-amber-950">
          <AlertTriangle className="w-5 h-5 text-amber-700 shrink-0 mt-0.5" aria-hidden="true" />
          <div>
            <p className="font-semibold">
              {counts.external_unembedded} fonte(s) externa(s) ou personalizada(s) não incorporada(s)
              no ficheiro original
            </p>
            <p className="mt-1 text-amber-900 leading-relaxed">
              A referência original destas fontes será preservada sem qualquer substituição. Para
              evitar alterações na distribuição das linhas e parágrafos noutros equipamentos, poderá
              ser necessário instalar estas fontes no computador de destino.
            </p>
          </div>
        </div>
      )}

      {/* Embedded font configuration summary */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div className="p-4 bg-white border border-slate-200 rounded-lg">
          <p className="text-xs text-slate-500">Incorporação TrueType (PPTX)</p>
          <p className="mt-1.5 text-sm font-semibold text-slate-900">
            {analysis.embeddedFontFlags.embedTrueTypeFonts
              ? 'Ativa no ficheiro original'
              : 'Não ativada no ficheiro original'}
          </p>
          <p className="mt-1 text-xs text-slate-500">
            {analysis.embeddedFontFlags.saveSubsetFonts
              ? 'Subconjunto de carateres guardado (saveSubsetFonts)'
              : 'Configuração original preservada sem alteração'}
          </p>
        </div>

        <div className="p-4 bg-white border border-slate-200 rounded-lg">
          <p className="text-xs text-slate-500">Ficheiros de fontes em ppt/fonts/</p>
          <p className="mt-1.5 text-sm font-semibold text-slate-900 font-mono tabular-nums">
            {analysis.embeddedFontFlags.embeddedFontFiles.length} ficheiro(s)
          </p>
          <p className="mt-1 text-xs text-slate-500 truncate">
            {analysis.embeddedFontFlags.embeddedFontFiles.length > 0
              ? analysis.embeddedFontFlags.embeddedFontFiles
                  .map((f) => `${f.path.split('/').pop()} (${formatBytes(f.size)})`)
                  .join(' · ')
              : 'Nenhum ficheiro binário de fonte embutido'}
          </p>
        </div>

        <div className="p-4 bg-white border border-slate-200 rounded-lg">
          <p className="text-xs text-slate-500">Política de substituição tipográfica</p>
          <p className="mt-1.5 text-sm font-semibold text-emerald-800">
            Substituição silenciosa desativada
          </p>
          <p className="mt-1 text-xs text-slate-500">
            Verificação automática bloqueia a exportação se alguma referência de fonte mudar
          </p>
        </div>
      </div>

      {/* Filter & Search Toolbar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
        <div className="inline-flex items-center gap-1 p-1 bg-slate-100 border border-slate-200 rounded-lg overflow-x-auto">
          {(
            [
              ['all', `Todas (${counts.all})`],
              ['external_unembedded', `Externas (${counts.external_unembedded})`],
              ['embedded', `Incorporadas (${counts.embedded})`],
              ['standard', `Sistema (${counts.standard})`],
            ] as [FontFilter, string][]
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

        <div className="relative w-full sm:w-64">
          <Search
            className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2 pointer-events-none"
            aria-hidden="true"
          />
          <input
            type="text"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            placeholder="Filtrar por nome da fonte…"
            className="w-full pl-9 pr-3 py-1.5 text-sm bg-white border border-slate-300 rounded-lg text-slate-900 placeholder:text-slate-400 focus:outline-none focus:border-blue-600"
          />
        </div>
      </div>

      {/* Fonts table */}
      <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-sm">
            <thead>
              <tr className="border-b border-slate-200 bg-slate-50 text-xs font-medium text-slate-600">
                <th className="py-3 pl-6 pr-3">Tipo de letra</th>
                <th className="py-3 px-3">Estado</th>
                <th className="py-3 px-3">Estilos e tamanhos detetados</th>
                <th className="py-3 px-3">Diapositivos</th>
                <th className="py-3 pl-3 pr-6">Preservação e aviso</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredFonts.length === 0 ? (
                <tr>
                  <td colSpan={5} className="py-10 text-center text-sm text-slate-500">
                    Nenhuma fonte corresponde ao filtro selecionado.
                  </td>
                </tr>
              ) : (
                filteredFonts.map((font) => {
                  const styles: string[] = [];
                  if (font.hasBold) styles.push('Negrito');
                  if (font.hasItalic) styles.push('Itálico');
                  if (font.hasUnderline) styles.push('Sublinhado');
                  if (styles.length === 0) styles.push('Regular');

                  const themeRoles: string[] = [];
                  if (font.isThemeMajor) themeRoles.push('Títulos do tema');
                  if (font.isThemeMinor) themeRoles.push('Corpo do tema');

                  return (
                    <tr key={font.name} className="hover:bg-slate-50/70 align-top">
                      <td className="py-3.5 pl-6 pr-3">
                        <div className="font-semibold text-slate-900">{font.name}</div>
                        {themeRoles.length > 0 && (
                          <div className="text-xs text-slate-500 mt-0.5">
                            {themeRoles.join(' · ')}
                          </div>
                        )}
                        {font.embeddedVariants.length > 0 && (
                          <div className="text-xs text-emerald-700 mt-0.5">
                            Variantes incorporadas: {font.embeddedVariants.join(', ')}
                          </div>
                        )}
                      </td>
                      <td className="py-3.5 px-3 whitespace-nowrap">{renderStatus(font)}</td>
                      <td className="py-3.5 px-3 text-xs text-slate-700">
                        <div>{styles.join(' · ')}</div>
                        {font.sizesPt.length > 0 && (
                          <div className="mt-1 font-mono text-slate-500 tabular-nums">
                            {font.sizesPt.map((s) => `${s} pt`).join(', ')}
                          </div>
                        )}
                      </td>
                      <td className="py-3.5 px-3 font-mono text-xs text-slate-700 tabular-nums">
                        {font.slidesUsed.length > 0
                          ? font.slidesUsed.map((n) => `#${n}`).join(', ')
                          : font.usedInMastersOrLayouts
                          ? 'Modelo global'
                          : 'Tema'}
                      </td>
                      <td className="py-3.5 pl-3 pr-6 text-xs text-slate-600 max-w-sm leading-relaxed">
                        {font.advisory}
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

import React from 'react';
import { ArrowRight, Lock, AlertTriangle, CheckCircle2 } from 'lucide-react';
import type { PptxAnalysis, WorkspaceSection } from '../types/pptx';
import { formatBytes, formatDatePt, shortHash } from '../utils/formatters';

interface OverviewSectionProps {
  analysis: PptxAnalysis;
  hasProcessedResult: boolean;
  onNavigate: (section: WorkspaceSection) => void;
}

export const OverviewSection: React.FC<OverviewSectionProps> = ({
  analysis,
  hasProcessedResult,
  onNavigate,
}) => {
  const unembeddedFontsCount = analysis.fonts.filter(
    (f) => f.status === 'external_unembedded'
  ).length;
  const warningIssuesCount = analysis.compatibilityIssues.filter(
    (i) => i.severity === 'warning' || i.severity === 'critical'
  ).length;

  return (
    <div className="space-y-8">
      {/* Header summary */}
      <div className="flex flex-col lg:flex-row lg:items-end lg:justify-between gap-4 pb-6 border-b border-slate-200">
        <div className="min-w-0">
          <p className="text-xs font-medium text-slate-500 mb-1">
            Análise automática da apresentação original
          </p>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900 truncate">
            {analysis.fileName}
          </h1>
          <div className="mt-2 flex flex-wrap items-center gap-x-2 gap-y-1 text-xs text-slate-600 tabular-nums">
            <span>{analysis.slideCount} diapositivos</span>
            <span aria-hidden="true">·</span>
            <span>{formatBytes(analysis.fileSize)}</span>
            <span aria-hidden="true">·</span>
            <span>
              {analysis.aspectRatioLabel} ({analysis.slideWidthCm} × {analysis.slideHeightCm} cm)
            </span>
            <span aria-hidden="true">·</span>
            <span className="font-mono text-slate-500" title={analysis.sha256}>
              SHA-256 {shortHash(analysis.sha256)}
            </span>
          </div>
        </div>

        <div className="flex items-center gap-3 shrink-0">
          {hasProcessedResult && (
            <button
              type="button"
              onClick={() => onNavigate('comparison')}
              className="px-4 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              Ver comparação
            </button>
          )}
          <button
            type="button"
            onClick={() => onNavigate('process_export')}
            className="inline-flex items-center gap-2 px-4 py-2 text-sm font-semibold text-white bg-blue-700 rounded-lg hover:bg-blue-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            <span>Processar com segurança</span>
            <ArrowRight className="w-4 h-4" aria-hidden="true" />
          </button>
        </div>
      </div>

      {/* Immutable original protection bar */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 px-4 py-3 bg-white border border-slate-200 rounded-lg text-sm">
        <div className="flex items-center gap-2.5 text-slate-800">
          <Lock className="w-4 h-4 text-emerald-700 shrink-0" aria-hidden="true" />
          <span>
            <strong className="font-semibold">Original protegido em memória imutável.</strong>{' '}
            Qualquer operação cria uma cópia de trabalho independente e nunca sobrescreve o ficheiro
            original.
          </span>
        </div>
        <span className="text-xs font-mono text-slate-500 whitespace-nowrap shrink-0">
          {analysis.structure.totalEntries} ficheiros internos no pacote
        </span>
      </div>

      {/* 4-column KPI strip */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        <button
          type="button"
          onClick={() => onNavigate('fonts')}
          className="text-left p-5 bg-white border border-slate-200 rounded-lg hover:border-slate-300 transition-colors cursor-pointer"
        >
          <p className="text-xs font-medium text-slate-500">Fontes detetadas</p>
          <p className="mt-2 text-2xl font-bold text-slate-900 tabular-nums font-mono">
            {analysis.fonts.length}
          </p>
          <p className="mt-1.5 text-xs text-slate-600">
            {analysis.embeddedFontFlags.embeddedFontFiles.length} incorporada(s) ·{' '}
            <span className={unembeddedFontsCount > 0 ? 'text-amber-700 font-medium' : ''}>
              {unembeddedFontsCount} externa(s)
            </span>
          </p>
        </button>

        <button
          type="button"
          onClick={() => onNavigate('hyperlinks')}
          className="text-left p-5 bg-white border border-slate-200 rounded-lg hover:border-slate-300 transition-colors cursor-pointer"
        >
          <p className="text-xs font-medium text-slate-500">Hiperligações</p>
          <p className="mt-2 text-2xl font-bold text-slate-900 tabular-nums font-mono">
            {analysis.hyperlinks.length}
          </p>
          <p className="mt-1.5 text-xs text-slate-600">
            Texto, imagens, formas e saltos internos
          </p>
        </button>

        <button
          type="button"
          onClick={() => onNavigate('visual_media')}
          className="text-left p-5 bg-white border border-slate-200 rounded-lg hover:border-slate-300 transition-colors cursor-pointer"
        >
          <p className="text-xs font-medium text-slate-500">Animações e multimédia</p>
          <p className="mt-2 text-2xl font-bold text-slate-900 tabular-nums font-mono">
            {analysis.visualSummary.totalAnimations + analysis.multimedia.length}
          </p>
          <p className="mt-1.5 text-xs text-slate-600">
            {analysis.visualSummary.slidesWithTransitions} transições · {analysis.multimedia.length}{' '}
            recursos
          </p>
        </button>

        <button
          type="button"
          onClick={() => onNavigate('compatibility')}
          className="text-left p-5 bg-white border border-slate-200 rounded-lg hover:border-slate-300 transition-colors cursor-pointer"
        >
          <p className="text-xs font-medium text-slate-500">Compatibilidade</p>
          <p className="mt-2 text-2xl font-bold text-slate-900 tabular-nums font-mono">
            {warningIssuesCount}
          </p>
          <p className="mt-1.5 text-xs flex items-center gap-1.5">
            {warningIssuesCount > 0 ? (
              <>
                <AlertTriangle className="w-3.5 h-3.5 text-amber-700 shrink-0" aria-hidden="true" />
                <span className="text-amber-800 font-medium">Avisos a rever</span>
              </>
            ) : (
              <>
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-700 shrink-0" aria-hidden="true" />
                <span className="text-emerald-800 font-medium">Sem riscos críticos</span>
              </>
            )}
          </p>
        </button>
      </div>

      {/* Document properties & Visual structure summary */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <section className="bg-white border border-slate-200 rounded-lg p-6">
          <h2 className="text-base font-semibold text-slate-900 mb-4">
            Propriedades do documento
          </h2>
          <dl className="divide-y divide-slate-100 text-sm">
            <div className="py-2.5 flex justify-between gap-4">
              <dt className="text-slate-500">Título</dt>
              <dd className="text-slate-900 font-medium text-right truncate">
                {analysis.coreProperties.title || '—'}
              </dd>
            </div>
            <div className="py-2.5 flex justify-between gap-4">
              <dt className="text-slate-500">Autor / Organização</dt>
              <dd className="text-slate-900 text-right truncate">
                {[analysis.coreProperties.creator, analysis.coreProperties.company]
                  .filter(Boolean)
                  .join(' · ') || '—'}
              </dd>
            </div>
            <div className="py-2.5 flex justify-between gap-4">
              <dt className="text-slate-500">Aplicação de origem</dt>
              <dd className="text-slate-900 text-right">
                {analysis.coreProperties.application || 'Microsoft PowerPoint / OOXML'}
              </dd>
            </div>
            <div className="py-2.5 flex justify-between gap-4">
              <dt className="text-slate-500">Última modificação</dt>
              <dd className="text-slate-900 text-right tabular-nums">
                {formatDatePt(analysis.coreProperties.modified)}
              </dd>
            </div>
            <div className="py-2.5 flex justify-between gap-4">
              <dt className="text-slate-500">Tema e modelos</dt>
              <dd className="text-slate-900 text-right">
                {analysis.visualSummary.themeName} · {analysis.visualSummary.masterCount} mestre(s)
                · {analysis.visualSummary.layoutCount} esquema(s)
              </dd>
            </div>
          </dl>
        </section>

        <section className="bg-white border border-slate-200 rounded-lg p-6">
          <h2 className="text-base font-semibold text-slate-900 mb-4">
            Resumo de elementos preservados
          </h2>
          <dl className="divide-y divide-slate-100 text-sm">
            <div className="py-2.5 flex justify-between gap-4">
              <dt className="text-slate-500">Formas e caixas de texto</dt>
              <dd className="font-mono tabular-nums text-slate-900">
                {analysis.visualSummary.totalShapes}
              </dd>
            </div>
            <div className="py-2.5 flex justify-between gap-4">
              <dt className="text-slate-500">Imagens, tabelas e gráficos</dt>
              <dd className="font-mono tabular-nums text-slate-900">
                {analysis.visualSummary.totalImages} img · {analysis.visualSummary.totalTables} tab
                · {analysis.visualSummary.totalCharts} grf
              </dd>
            </div>
            <div className="py-2.5 flex justify-between gap-4">
              <dt className="text-slate-500">Fundos e efeitos visuais</dt>
              <dd className="font-mono tabular-nums text-slate-900">
                {analysis.visualSummary.slidesWithBackground} fundos ·{' '}
                {analysis.visualSummary.slidesWithEffects} efeitos
              </dd>
            </div>
            <div className="py-2.5 flex justify-between gap-4">
              <dt className="text-slate-500">Ficheiros XML e relações (.rels)</dt>
              <dd className="font-mono tabular-nums text-slate-900">
                {analysis.structure.xmlCount} XML · {analysis.structure.relsCount} rels
              </dd>
            </div>
            <div className="py-2.5 flex justify-between gap-4">
              <dt className="text-slate-500">Extensões OOXML específicas</dt>
              <dd className="text-slate-900 text-right">
                {analysis.structure.unknownExtensions.length > 0
                  ? `${analysis.structure.unknownExtensions.length} bloco(s) mantidos intactos`
                  : 'Nenhuma extensão proprietária extra'}
              </dd>
            </div>
          </dl>
        </section>
      </div>

      {/* Slide inventory table */}
      <section className="bg-white border border-slate-200 rounded-lg overflow-hidden">
        <div className="px-6 py-4 border-b border-slate-200 flex items-center justify-between gap-4">
          <div>
            <h2 className="text-base font-semibold text-slate-900">
              Mapa de diapositivos ({analysis.slideCount})
            </h2>
            <p className="text-xs text-slate-500 mt-0.5">
              Inventário estrutural extraído diretamente dos ficheiros ppt/slides/slide*.xml
            </p>
          </div>
          <button
            type="button"
            onClick={() => onNavigate('visual_media')}
            className="text-xs font-medium text-blue-700 hover:text-blue-800 whitespace-nowrap shrink-0 cursor-pointer"
          >
            Ver detalhes visuais e multimédia
          </button>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-sm">
            <thead>
              <tr className="border-b border-slate-200 bg-slate-50 text-xs font-medium text-slate-600">
                <th className="py-3 pl-6 pr-3 w-16">#</th>
                <th className="py-3 px-3">Título / Conteúdo</th>
                <th className="py-3 px-3">Fontes</th>
                <th className="py-3 px-3 text-right">Links</th>
                <th className="py-3 px-3">Transição / Animações</th>
                <th className="py-3 pl-3 pr-6 text-right">XML</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {analysis.slides.map((slide) => (
                <tr key={slide.path} className="hover:bg-slate-50/70 transition-colors">
                  <td className="py-3 pl-6 pr-3 font-mono tabular-nums text-slate-600">
                    {String(slide.slideNumber).padStart(2, '0')}
                  </td>
                  <td className="py-3 px-3 max-w-md">
                    <div className="font-medium text-slate-900 truncate">
                      {slide.title}
                      {slide.hidden && (
                        <span className="ml-2 text-xs font-normal text-amber-700">
                          (Oculto)
                        </span>
                      )}
                    </div>
                    <div className="text-xs text-slate-500 truncate mt-0.5">
                      {slide.shapeCount} formas · {slide.imageCount} imagens · {slide.tableCount}{' '}
                      tabelas
                      {slide.hasNotes ? ' · Com notas do orador' : ''}
                    </div>
                  </td>
                  <td className="py-3 px-3 text-xs text-slate-700 max-w-[200px] truncate">
                    {slide.fontsUsed.join(', ') || '—'}
                  </td>
                  <td className="py-3 px-3 text-right font-mono tabular-nums text-slate-900">
                    {slide.hyperlinkCount}
                  </td>
                  <td className="py-3 px-3 text-xs text-slate-700">
                    <span>{slide.transition.present ? slide.transition.type : 'Sem transição'}</span>
                    {slide.animations.present && (
                      <span className="text-slate-500">
                        {' '}
                        · {slide.animations.nodeCount} anim.
                      </span>
                    )}
                  </td>
                  <td className="py-3 pl-3 pr-6 text-right font-mono tabular-nums text-xs text-slate-500">
                    {formatBytes(slide.xmlByteSize)}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  );
};

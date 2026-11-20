import React, { useState } from 'react';
import { CheckCircle2, AlertTriangle } from 'lucide-react';
import type { MultimediaAuditEntry, PptxAnalysis } from '../types/pptx';
import { formatBytes } from '../utils/formatters';

interface VisualMediaSectionProps {
  analysis: PptxAnalysis;
}

const MEDIA_KIND_LABELS: Record<MultimediaAuditEntry['kind'], string> = {
  image: 'Imagem incorporada',
  audio: 'Áudio incorporado',
  video: 'Vídeo incorporado',
  external_media: 'Multimédia externa (ligada)',
  chart: 'Gráfico nativo',
  ole_embedded: 'Objeto incorporado (OLE)',
  smartart: 'Diagrama SmartArt',
};

export const VisualMediaSection: React.FC<VisualMediaSectionProps> = ({ analysis }) => {
  const [subTab, setSubTab] = useState<'slides' | 'media'>('slides');

  return (
    <div className="space-y-6">
      <div className="pb-5 border-b border-slate-200 flex flex-col sm:flex-row sm:items-end justify-between gap-4">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">
            Preservação visual, animações e multimédia
          </h1>
          <p className="mt-1.5 text-sm text-slate-600 max-w-3xl">
            Fundos, formas, tabelas, gráficos, sombras, transparências, transições entre
            diapositivos (<code className="font-mono text-xs">p:transition</code>), animações (
            <code className="font-mono text-xs">p:timing</code>) e recursos multimédia são mantidos
            na sua forma original sem simplificação nem recompressão.
          </p>
        </div>

        <div className="inline-flex items-center gap-1 p-1 bg-slate-100 border border-slate-200 rounded-lg shrink-0 self-start sm:self-auto">
          <button
            type="button"
            onClick={() => setSubTab('slides')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              subTab === 'slides'
                ? 'bg-white text-slate-900 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Diapositivos e animações ({analysis.slideCount})
          </button>
          <button
            type="button"
            onClick={() => setSubTab('media')}
            className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors whitespace-nowrap shrink-0 cursor-pointer ${
              subTab === 'media'
                ? 'bg-white text-slate-900 shadow-xs'
                : 'text-slate-600 hover:text-slate-900'
            }`}
          >
            Multimédia e objetos ({analysis.multimedia.length})
          </button>
        </div>
      </div>

      {/* Visual metrics summary */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
        <div className="p-4 bg-white border border-slate-200 rounded-lg">
          <p className="text-xs text-slate-500">Formas e grupos</p>
          <p className="mt-1 text-xl font-bold text-slate-900 font-mono tabular-nums">
            {analysis.visualSummary.totalShapes + analysis.visualSummary.totalGroups}
          </p>
          <p className="mt-0.5 text-xs text-slate-500">
            {analysis.visualSummary.totalTables} tabela(s) · {analysis.visualSummary.totalCharts}{' '}
            gráfico(s)
          </p>
        </div>

        <div className="p-4 bg-white border border-slate-200 rounded-lg">
          <p className="text-xs text-slate-500">Transições</p>
          <p className="mt-1 text-xl font-bold text-slate-900 font-mono tabular-nums">
            {analysis.visualSummary.slidesWithTransitions}
          </p>
          <p className="mt-0.5 text-xs text-slate-500">
            de {analysis.slideCount} diapositivo(s)
          </p>
        </div>

        <div className="p-4 bg-white border border-slate-200 rounded-lg">
          <p className="text-xs text-slate-500">Nós de animação</p>
          <p className="mt-1 text-xl font-bold text-slate-900 font-mono tabular-nums">
            {analysis.visualSummary.totalAnimations}
          </p>
          <p className="mt-0.5 text-xs text-slate-500">Sequências e temporizações</p>
        </div>

        <div className="p-4 bg-white border border-slate-200 rounded-lg">
          <p className="text-xs text-slate-500">Recursos multimédia</p>
          <p className="mt-1 text-xl font-bold text-slate-900 font-mono tabular-nums">
            {analysis.multimedia.length}
          </p>
          <p className="mt-0.5 text-xs text-slate-500">Imagens, áudio, vídeo e objetos</p>
        </div>
      </div>

      {subTab === 'slides' ? (
        <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-sm">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 text-xs font-medium text-slate-600">
                  <th className="py-3 pl-6 pr-3 w-16">#</th>
                  <th className="py-3 px-3">Diapositivo</th>
                  <th className="py-3 px-3">Elementos visuais</th>
                  <th className="py-3 px-3">Aspeto e efeitos</th>
                  <th className="py-3 px-3">Transição</th>
                  <th className="py-3 pl-3 pr-6">Animações</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {analysis.slides.map((slide) => {
                  const visualParts: string[] = [];
                  if (slide.shapeCount > 0) visualParts.push(`${slide.shapeCount} formas`);
                  if (slide.imageCount > 0) visualParts.push(`${slide.imageCount} img`);
                  if (slide.tableCount > 0) visualParts.push(`${slide.tableCount} tab`);
                  if (slide.chartCount > 0) visualParts.push(`${slide.chartCount} grf`);
                  if (slide.smartArtCount > 0) visualParts.push(`${slide.smartArtCount} SmartArt`);
                  if (slide.groupCount > 0) visualParts.push(`${slide.groupCount} grupos`);

                  const styleParts: string[] = [];
                  if (slide.hasBackground) styleParts.push('Fundo próprio');
                  if (slide.hasShadows) styleParts.push('Sombras/Efeitos');
                  if (slide.hasTransparency) styleParts.push('Transparência');
                  if (slide.hasHeaderFooter) styleParts.push('Cabeçalho/Rodapé');

                  return (
                    <tr key={slide.path} className="hover:bg-slate-50/70 align-top">
                      <td className="py-3.5 pl-6 pr-3 font-mono tabular-nums text-slate-600">
                        {String(slide.slideNumber).padStart(2, '0')}
                      </td>
                      <td className="py-3.5 px-3 max-w-xs">
                        <div className="font-medium text-slate-900 truncate">{slide.title}</div>
                        <div className="text-xs font-mono text-slate-500 mt-0.5">{slide.path}</div>
                      </td>
                      <td className="py-3.5 px-3 text-xs text-slate-700">
                        {visualParts.join(' · ') || 'Apenas esquema base'}
                      </td>
                      <td className="py-3.5 px-3 text-xs text-slate-600">
                        {styleParts.join(' · ') || 'Herdado do tema'}
                      </td>
                      <td className="py-3.5 px-3 text-xs text-slate-800 whitespace-nowrap">
                        {slide.transition.present ? (
                          <div>
                            <span className="font-medium">{slide.transition.type}</span>
                            {slide.transition.durationMs && (
                              <span className="text-slate-500 font-mono">
                                {' '}
                                ({slide.transition.durationMs} ms)
                              </span>
                            )}
                          </div>
                        ) : (
                          <span className="text-slate-400">Nenhuma</span>
                        )}
                      </td>
                      <td className="py-3.5 pl-3 pr-6 text-xs text-slate-700">
                        {slide.animations.present ? (
                          <div>
                            <span className="font-mono font-medium text-slate-900 tabular-nums">
                              {slide.animations.nodeCount} efeito(s)
                            </span>
                            <div className="text-slate-500 mt-0.5">
                              {slide.animations.effectTypes.join(' · ')}
                            </div>
                          </div>
                        ) : (
                          <span className="text-slate-400">Sem animações</span>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      ) : (
        <div className="bg-white border border-slate-200 rounded-lg overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-sm">
              <thead>
                <tr className="border-b border-slate-200 bg-slate-50 text-xs font-medium text-slate-600">
                  <th className="py-3 pl-6 pr-3">Ficheiro / Recurso</th>
                  <th className="py-3 px-3">Categoria</th>
                  <th className="py-3 px-3">Formato</th>
                  <th className="py-3 px-3 text-right">Tamanho</th>
                  <th className="py-3 px-3">Diapositivos</th>
                  <th className="py-3 pl-3 pr-6">Preservação</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {analysis.multimedia.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="py-10 text-center text-sm text-slate-500">
                      Nenhum ficheiro multimédia ou gráfico incorporado foi encontrado nesta
                      apresentação.
                    </td>
                  </tr>
                ) : (
                  analysis.multimedia.map((item) => (
                    <tr key={item.path} className="hover:bg-slate-50/70 align-top">
                      <td className="py-3.5 pl-6 pr-3 max-w-xs">
                        <div className="font-medium text-slate-900 break-all">{item.fileName}</div>
                        <div className="font-mono text-xs text-slate-500 break-all mt-0.5">
                          {item.path}
                        </div>
                        {item.compatibilityNote && (
                          <p className="text-xs text-amber-800 mt-1">{item.compatibilityNote}</p>
                        )}
                      </td>
                      <td className="py-3.5 px-3 text-xs text-slate-700 whitespace-nowrap">
                        {MEDIA_KIND_LABELS[item.kind]}
                      </td>
                      <td className="py-3.5 px-3 font-mono text-xs text-slate-600">
                        {item.format}
                      </td>
                      <td className="py-3.5 px-3 text-right font-mono text-xs tabular-nums text-slate-700">
                        {item.isExternal ? 'Externo' : formatBytes(item.size)}
                      </td>
                      <td className="py-3.5 px-3 font-mono text-xs tabular-nums text-slate-700">
                        {item.slidesUsed.length > 0
                          ? item.slidesUsed.map((n) => `#${n}`).join(', ')
                          : 'Modelo/Tema'}
                      </td>
                      <td className="py-3.5 pl-3 pr-6 whitespace-nowrap">
                        {item.isExternal ? (
                          <span className="inline-flex items-center gap-1.5 text-xs font-medium text-amber-800">
                            <AlertTriangle
                              className="w-3.5 h-3.5 text-amber-700 shrink-0"
                              aria-hidden="true"
                            />
                            <span>Ligação externa preservada</span>
                          </span>
                        ) : (
                          <span className="inline-flex items-center gap-1.5 text-xs font-medium text-emerald-800">
                            <CheckCircle2
                              className="w-3.5 h-3.5 text-emerald-700 shrink-0"
                              aria-hidden="true"
                            />
                            <span>Binário byte-a-byte</span>
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
      )}
    </div>
  );
};

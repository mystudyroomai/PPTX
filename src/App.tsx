import React, { useRef, useState } from 'react';
import {
  FileUp,
  LayoutDashboard,
  Type,
  Link2,
  Layers,
  ShieldAlert,
  ShieldCheck,
  GitCompare,
  Menu,
  X,
  Download,
} from 'lucide-react';
import type {
  PptxAnalysis,
  ProcessingOptions,
  ProcessingResult,
  WorkspaceSection,
} from './types/pptx';
import { analyzePptxBuffer } from './engine/pptxAnalyzer';
import { processPptxSafely } from './engine/pptxProcessor';
import { generateSamplePptx } from './engine/samplePptxGenerator';
import { downloadUint8Array, formatBytes } from './utils/formatters';
import { ImportView } from './components/ImportView';
import { OverviewSection } from './components/OverviewSection';
import { FontsSection } from './components/FontsSection';
import { HyperlinksSection } from './components/HyperlinksSection';
import { VisualMediaSection } from './components/VisualMediaSection';
import { CompatibilitySection } from './components/CompatibilitySection';
import { ProcessExportSection } from './components/ProcessExportSection';
import { ComparisonSection } from './components/ComparisonSection';

function createDefaultOptions(analysis: PptxAnalysis): ProcessingOptions {
  return {
    textReplacements: [],
    protectHyperlinkedText: true,
    metadataUpdates: {
      enabled: false,
      title: analysis.coreProperties.title || '',
      subject: analysis.coreProperties.subject || '',
      creator: analysis.coreProperties.creator || '',
      anonymizeAuthor: false,
    },
    slideVisibilityOverrides: {},
    cleanSpeakerNotes: false,
  };
}

const SECTION_LABELS: Record<WorkspaceSection, string> = {
  overview: 'Visão geral',
  fonts: 'Fontes e tipografia',
  hyperlinks: 'Hiperligações',
  visual_media: 'Visual e multimédia',
  compatibility: 'Compatibilidade',
  process_export: 'Processar e exportar',
  comparison: 'Comparação',
};

export default function App() {
  const hiddenFileInputRef = useRef<HTMLInputElement | null>(null);

  const [originalBuffer, setOriginalBuffer] = useState<Uint8Array | null>(null);
  const [analysis, setAnalysis] = useState<PptxAnalysis | null>(null);
  const [options, setOptions] = useState<ProcessingOptions | null>(null);
  const [processingResult, setProcessingResult] = useState<ProcessingResult | null>(null);

  const [activeSection, setActiveSection] = useState<WorkspaceSection>('overview');
  const [isLoadingFile, setIsLoadingFile] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  const loadPptxBytes = async (bytes: Uint8Array, fileName: string) => {
    setIsLoadingFile(true);
    setErrorMessage(null);
    try {
      // Store an isolated immutable copy of the original file
      const immutableCopy = new Uint8Array(bytes.byteLength);
      immutableCopy.set(bytes);

      const parsed = await analyzePptxBuffer(immutableCopy, fileName);
      setOriginalBuffer(immutableCopy);
      setAnalysis(parsed);
      setOptions(createDefaultOptions(parsed));
      setProcessingResult(null);
      setActiveSection('overview');
      setMobileMenuOpen(false);
    } catch (err) {
      const msg =
        err instanceof Error
          ? err.message
          : 'Erro ao analisar o ficheiro. Confirme que se trata de um ficheiro Microsoft PowerPoint (.pptx) válido.';
      setErrorMessage(msg);
    } finally {
      setIsLoadingFile(false);
    }
  };

  const handleFileSelected = async (file: File) => {
    if (!file.name.toLowerCase().endsWith('.pptx')) {
      setErrorMessage(
        'Formato não suportado. Selecione um ficheiro Microsoft PowerPoint com extensão .pptx.'
      );
      return;
    }
    const arrBuf = await file.arrayBuffer();
    await loadPptxBytes(new Uint8Array(arrBuf), file.name);
  };

  const handleLoadSample = async () => {
    setIsLoadingFile(true);
    setErrorMessage(null);
    try {
      const { buffer, fileName } = await generateSamplePptx();
      await loadPptxBytes(buffer, fileName);
    } catch (err) {
      setErrorMessage(
        err instanceof Error ? err.message : 'Erro ao gerar a apresentação de demonstração.'
      );
      setIsLoadingFile(false);
    }
  };

  const handleRunProcess = async () => {
    if (!originalBuffer || !analysis || !options) return;
    setIsProcessing(true);
    try {
      const res = await processPptxSafely(originalBuffer, analysis, options);
      setProcessingResult(res);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleResetOptions = () => {
    if (!analysis) return;
    setOptions(createDefaultOptions(analysis));
    setProcessingResult(null);
  };

  const handleClosePresentation = () => {
    setOriginalBuffer(null);
    setAnalysis(null);
    setOptions(null);
    setProcessingResult(null);
    setErrorMessage(null);
    setActiveSection('overview');
    setMobileMenuOpen(false);
  };

  const triggerHiddenFilePicker = () => {
    hiddenFileInputRef.current?.click();
  };

  const hiddenFileInput = (
    <input
      ref={hiddenFileInputRef}
      type="file"
      accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
      className="hidden"
      onChange={(e) => {
        const file = e.target.files?.[0];
        if (file) handleFileSelected(file);
        e.target.value = '';
      }}
    />
  );

  // Initial empty state (Import view)
  if (!analysis || !options) {
    return (
      <div className="min-h-screen flex flex-col bg-slate-50 text-slate-900">
        {hiddenFileInput}
        <header className="bg-white border-b border-slate-200 px-6 h-16 flex items-center justify-between gap-8">
          <span className="text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0">
            PreservaPPT
          </span>

          <nav className="hidden md:flex items-center gap-6 text-sm font-medium text-slate-600">
            <button
              type="button"
              onClick={triggerHiddenFilePicker}
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              Importar PPTX
            </button>
            <button
              type="button"
              onClick={handleLoadSample}
              className="hover:text-slate-900 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              Demonstração
            </button>
          </nav>

          <button
            type="button"
            disabled={isLoadingFile}
            onClick={triggerHiddenFilePicker}
            className="px-4 py-2 text-xs font-semibold text-white bg-blue-700 rounded-lg hover:bg-blue-800 disabled:opacity-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
          >
            Selecionar .pptx
          </button>
        </header>

        <main className="flex-1">
          <ImportView
            onFileSelected={handleFileSelected}
            onLoadSample={handleLoadSample}
            isLoading={isLoadingFile}
            errorMessage={errorMessage}
          />
        </main>
      </div>
    );
  }

  const warningCount = analysis.compatibilityIssues.filter(
    (i) => i.severity === 'warning' || i.severity === 'critical'
  ).length;

  const navItems: {
    id: WorkspaceSection;
    label: string;
    icon: React.ComponentType<{ className?: string }>;
    meta?: string;
  }[] = [
    {
      id: 'overview',
      label: 'Visão geral',
      icon: LayoutDashboard,
      meta: `${analysis.slideCount}`,
    },
    {
      id: 'fonts',
      label: 'Fontes e tipografia',
      icon: Type,
      meta: `${analysis.fonts.length}`,
    },
    {
      id: 'hyperlinks',
      label: 'Hiperligações',
      icon: Link2,
      meta: `${analysis.hyperlinks.length}`,
    },
    {
      id: 'visual_media',
      label: 'Visual e multimédia',
      icon: Layers,
      meta: `${analysis.multimedia.length}`,
    },
    {
      id: 'compatibility',
      label: 'Compatibilidade',
      icon: ShieldAlert,
      meta: warningCount > 0 ? `${warningCount}` : '0',
    },
    {
      id: 'process_export',
      label: 'Processar e exportar',
      icon: ShieldCheck,
    },
    {
      id: 'comparison',
      label: 'Comparação',
      icon: GitCompare,
    },
  ];

  const canQuickExport = Boolean(
    processingResult &&
      processingResult.verification.canExport &&
      processingResult.processedBuffer
  );

  return (
    <div className="min-h-screen flex bg-slate-50 text-slate-900">
      {hiddenFileInput}

      {/* Mobile sidebar backdrop */}
      {mobileMenuOpen && (
        <div
          className="fixed inset-0 z-30 bg-slate-900/40 lg:hidden"
          onClick={() => setMobileMenuOpen(false)}
          aria-hidden="true"
        />
      )}

      {/* Left Sidebar Navigation */}
      <aside
        className={`fixed lg:sticky top-0 left-0 z-40 h-screen w-64 bg-white border-r border-slate-200 flex flex-col transition-transform duration-200 lg:translate-x-0 ${
          mobileMenuOpen ? 'translate-x-0' : '-translate-x-full'
        }`}
      >
        <div className="h-16 px-5 border-b border-slate-200 flex items-center justify-between">
          <button
            type="button"
            onClick={handleClosePresentation}
            className="text-lg font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0 cursor-pointer"
          >
            PreservaPPT
          </button>
          <button
            type="button"
            onClick={() => setMobileMenuOpen(false)}
            aria-label="Fechar menu"
            className="p-1.5 text-slate-500 hover:text-slate-900 lg:hidden cursor-pointer"
          >
            <X className="w-5 h-5" aria-hidden="true" />
          </button>
        </div>

        <nav className="flex-1 px-3 py-4 space-y-1 overflow-y-auto" aria-label="Secções da análise">
          {navItems.map((item) => {
            const Icon = item.icon;
            const isActive = activeSection === item.id;
            return (
              <button
                key={item.id}
                type="button"
                onClick={() => {
                  setActiveSection(item.id);
                  setMobileMenuOpen(false);
                }}
                className={`w-full flex items-center justify-between gap-2 px-3 py-2.5 text-sm rounded-lg transition-colors cursor-pointer ${
                  isActive
                    ? 'bg-blue-50 text-blue-800 font-semibold'
                    : 'text-slate-700 hover:bg-slate-100 hover:text-slate-900 font-medium'
                }`}
              >
                <span className="flex items-center gap-2.5 min-w-0">
                  <Icon
                    className={`w-4 h-4 shrink-0 ${
                      isActive ? 'text-blue-700' : 'text-slate-500'
                    }`}
                  />
                  <span className="truncate">{item.label}</span>
                </span>
                {item.meta !== undefined && (
                  <span className="font-mono text-xs tabular-nums text-slate-500 shrink-0">
                    {item.meta}
                  </span>
                )}
              </button>
            );
          })}
        </nav>

        <div className="p-4 border-t border-slate-200 space-y-3">
          <div className="min-w-0">
            <p className="text-xs font-semibold text-slate-900 truncate" title={analysis.fileName}>
              {analysis.fileName}
            </p>
            <p className="text-xs text-slate-500 tabular-nums mt-0.5">
              {analysis.slideCount} diapositivos · {formatBytes(analysis.fileSize)}
            </p>
          </div>
          <button
            type="button"
            disabled={isLoadingFile}
            onClick={triggerHiddenFilePicker}
            className="w-full inline-flex items-center justify-center gap-2 px-3 py-2 text-xs font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 transition-colors whitespace-nowrap cursor-pointer"
          >
            <FileUp className="w-3.5 h-3.5" aria-hidden="true" />
            <span>Importar outro .pptx</span>
          </button>
        </div>
      </aside>

      {/* Main Content Area */}
      <div className="flex-1 min-w-0 flex flex-col">
        {/* Top Workspace Bar */}
        <header className="sticky top-0 z-20 h-16 bg-white border-b border-slate-200 px-6 flex items-center justify-between gap-6">
          <div className="flex items-center gap-3 min-w-0">
            <button
              type="button"
              onClick={() => setMobileMenuOpen(true)}
              aria-label="Abrir navegação"
              className="p-1.5 -ml-1.5 text-slate-600 hover:text-slate-900 lg:hidden cursor-pointer"
            >
              <Menu className="w-5 h-5" aria-hidden="true" />
            </button>
            <div className="flex items-center gap-2 text-sm min-w-0">
              <span className="text-slate-500 truncate max-w-[180px] sm:max-w-xs">
                {analysis.fileName}
              </span>
              <span className="text-slate-300" aria-hidden="true">
                /
              </span>
              <span className="font-semibold text-slate-900 whitespace-nowrap truncate">
                {SECTION_LABELS[activeSection]}
              </span>
            </div>
          </div>

          <div className="shrink-0">
            {canQuickExport ? (
              <button
                type="button"
                onClick={() => {
                  if (processingResult?.processedBuffer) {
                    downloadUint8Array(
                      processingResult.processedBuffer,
                      processingResult.processedFileName
                    );
                  }
                }}
                className="inline-flex items-center gap-2 px-4 py-2 text-xs font-semibold text-white bg-emerald-700 rounded-lg hover:bg-emerald-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" aria-hidden="true" />
                <span>Exportar PPTX verificado</span>
              </button>
            ) : (
              <button
                type="button"
                onClick={() => setActiveSection('process_export')}
                className="px-4 py-2 text-xs font-semibold text-white bg-blue-700 rounded-lg hover:bg-blue-800 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
              >
                Processar e verificar
              </button>
            )}
          </div>
        </header>

        <main className="flex-1 max-w-6xl w-full mx-auto px-6 py-8">
          {errorMessage && (
            <div
              role="alert"
              className="mb-6 p-4 bg-red-50 border border-red-200 rounded-lg text-sm text-red-900 flex items-center justify-between gap-4"
            >
              <span>{errorMessage}</span>
              <button
                type="button"
                onClick={() => setErrorMessage(null)}
                className="text-xs font-semibold text-red-800 hover:underline whitespace-nowrap cursor-pointer"
              >
                Fechar
              </button>
            </div>
          )}

          {activeSection === 'overview' && (
            <OverviewSection
              analysis={analysis}
              hasProcessedResult={Boolean(processingResult?.comparison)}
              onNavigate={setActiveSection}
            />
          )}

          {activeSection === 'fonts' && <FontsSection analysis={analysis} />}

          {activeSection === 'hyperlinks' && <HyperlinksSection analysis={analysis} />}

          {activeSection === 'visual_media' && <VisualMediaSection analysis={analysis} />}

          {activeSection === 'compatibility' && <CompatibilitySection analysis={analysis} />}

          {activeSection === 'process_export' && (
            <ProcessExportSection
              originalAnalysis={analysis}
              options={options}
              onChangeOptions={(next) => {
                setOptions(next);
              }}
              onResetOptions={handleResetOptions}
              onRunProcess={handleRunProcess}
              isProcessing={isProcessing}
              result={processingResult}
              onGoToComparison={() => setActiveSection('comparison')}
            />
          )}

          {activeSection === 'comparison' && (
            <ComparisonSection
              original={analysis}
              result={processingResult}
              onNavigate={setActiveSection}
            />
          )}
        </main>
      </div>
    </div>
  );
}

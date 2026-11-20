import React, { useRef, useState } from 'react';
import {
  FileUp,
  ShieldCheck,
  Type,
  Link2,
  Layers,
  AlertCircle,
  FileSpreadsheet,
} from 'lucide-react';

interface ImportViewProps {
  onFileSelected: (file: File) => void;
  onLoadSample: () => void;
  isLoading: boolean;
  errorMessage: string | null;
}

export const ImportView: React.FC<ImportViewProps> = ({
  onFileSelected,
  onLoadSample,
  isLoading,
  errorMessage,
}) => {
  const inputRef = useRef<HTMLInputElement | null>(null);
  const [isDragging, setIsDragging] = useState(false);

  const handleFiles = (files: FileList | null) => {
    if (!files || files.length === 0) return;
    onFileSelected(files[0]);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    handleFiles(e.dataTransfer.files);
  };

  return (
    <div className="max-w-5xl mx-auto px-6 py-12">
      <div className="mb-10">
        <p className="text-sm font-medium text-blue-700 mb-2">
          Processamento não destrutivo de ficheiros Microsoft PowerPoint (.pptx)
        </p>
        <h1 className="text-3xl sm:text-4xl font-bold tracking-tight text-slate-900 max-w-3xl">
          Preservação integral de apresentações PowerPoint com verificação estrutural
        </h1>
        <p className="mt-3 text-base text-slate-600 max-w-2xl leading-relaxed">
          Importe uma apresentação para auditar fontes incorporadas e externas, hiperligações,
          animações, transições e multimédia. Todas as operações trabalham sobre uma cópia isolada
          em XML direto, mantendo o ficheiro original intacto.
        </p>
      </div>

      {errorMessage && (
        <div
          role="alert"
          className="mb-6 flex items-start gap-3 p-4 bg-red-50 border border-red-200 rounded-lg text-sm text-red-900"
        >
          <AlertCircle className="w-5 h-5 text-red-700 shrink-0 mt-0.5" aria-hidden="true" />
          <div>
            <p className="font-semibold">Não foi possível importar o ficheiro</p>
            <p className="mt-1 text-red-800">{errorMessage}</p>
          </div>
        </div>
      )}

      <div
        onDragOver={(e) => {
          e.preventDefault();
          setIsDragging(true);
        }}
        onDragLeave={() => setIsDragging(false)}
        onDrop={handleDrop}
        className={`bg-white border rounded-xl p-10 transition-colors duration-150 ${
          isDragging
            ? 'border-blue-600 bg-blue-50/40'
            : 'border-slate-200 hover:border-slate-300'
        }`}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
          className="hidden"
          onChange={(e) => {
            handleFiles(e.target.files);
            e.target.value = '';
          }}
        />

        <div className="flex flex-col items-center text-center max-w-xl mx-auto">
          <div className="w-12 h-12 rounded-lg bg-slate-100 border border-slate-200 flex items-center justify-center mb-5">
            <FileUp className="w-6 h-6 text-slate-700" aria-hidden="true" />
          </div>

          <h2 className="text-lg font-semibold text-slate-900">
            {isLoading
              ? 'A analisar estrutura interna OOXML…'
              : 'Arraste um ficheiro PowerPoint (.pptx) ou selecione do computador'}
          </h2>
          <p className="mt-2 text-sm text-slate-600">
            O processamento decorre localmente no navegador. O ficheiro original é guardado em
            memória imutável com assinatura SHA-256 e nunca é sobrescrito.
          </p>

          <div className="mt-6 flex flex-wrap items-center justify-center gap-3">
            <button
              type="button"
              disabled={isLoading}
              onClick={() => inputRef.current?.click()}
              className="px-5 py-2.5 text-sm font-semibold text-white bg-blue-700 rounded-lg hover:bg-blue-800 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-blue-700 disabled:opacity-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              Selecionar ficheiro .pptx
            </button>
            <button
              type="button"
              disabled={isLoading}
              onClick={onLoadSample}
              className="px-5 py-2.5 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 hover:text-slate-900 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-500 disabled:opacity-50 transition-colors whitespace-nowrap shrink-0 cursor-pointer"
            >
              Carregar apresentação de demonstração
            </button>
          </div>

          <p className="mt-5 text-xs text-slate-500">
            Formato suportado: Microsoft PowerPoint Open XML (.pptx) · Sem conversão em imagem ·
            Sem reconstrução destrutiva
          </p>
        </div>
      </div>

      <div className="mt-12 pt-8 border-t border-slate-200 grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-8">
        <div>
          <div className="flex items-center gap-2 text-slate-900 font-semibold text-sm mb-2">
            <Type className="w-4 h-4 text-blue-700 shrink-0" aria-hidden="true" />
            <span>Preservação de fontes</span>
          </div>
          <p className="text-sm text-slate-600 leading-relaxed">
            Mantém fontes incorporadas, personalizadas e pouco comuns sem nunca substituir por
            Arial ou Calibri, alertando quando a fonte é externa.
          </p>
        </div>

        <div>
          <div className="flex items-center gap-2 text-slate-900 font-semibold text-sm mb-2">
            <Link2 className="w-4 h-4 text-blue-700 shrink-0" aria-hidden="true" />
            <span>Hiperligações intactas</span>
          </div>
          <p className="text-sm text-slate-600 leading-relaxed">
            Audita e preserva ligações web, saltos internos entre diapositivos, ficheiros externos
            e ações em texto, imagens e botões.
          </p>
        </div>

        <div>
          <div className="flex items-center gap-2 text-slate-900 font-semibold text-sm mb-2">
            <Layers className="w-4 h-4 text-blue-700 shrink-0" aria-hidden="true" />
            <span>XML e multimédia brutos</span>
          </div>
          <p className="text-sm text-slate-600 leading-relaxed">
            Preserva animações, transições, áudio, vídeo, tabelas, gráficos e blocos XML
            desconhecidos byte-a-byte em todos os componentes intocados.
          </p>
        </div>

        <div>
          <div className="flex items-center gap-2 text-slate-900 font-semibold text-sm mb-2">
            <ShieldCheck className="w-4 h-4 text-blue-700 shrink-0" aria-hidden="true" />
            <span>Verificação obrigatória</span>
          </div>
          <p className="text-sm text-slate-600 leading-relaxed">
            Antes de desbloquear a exportação, compara o ficheiro final com o original e cancela
            automaticamente caso detete qualquer perda.
          </p>
        </div>
      </div>

      <div className="mt-8 pt-6 border-t border-slate-200 flex items-center gap-2 text-xs text-slate-500">
        <FileSpreadsheet className="w-4 h-4 text-slate-400 shrink-0" aria-hidden="true" />
        <span>
          A demonstração integrada gera um ficheiro .pptx válido com fontes incorporadas e externas,
          hiperligações em texto/botão/imagem, tabela, áudio incorporado, transição Morph e notas.
        </span>
      </div>
    </div>
  );
};

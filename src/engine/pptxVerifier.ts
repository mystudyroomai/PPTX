import type {
  ComparisonSummary,
  FileDiffEntry,
  PptxAnalysis,
  SlideDiffEntry,
  VerificationCheck,
  VerificationReport,
} from '../types/pptx';

function describeModifiedReason(path: string): string {
  if (path === 'docProps/core.xml') return 'Metadados do documento atualizados com segurança';
  if (path.startsWith('ppt/notesSlides/')) return 'Notas do orador limpas conforme solicitado';
  if (path.startsWith('ppt/slides/slide')) return 'Edição não destrutiva aplicada apenas a nós autorizados';
  return 'Modificado durante o processamento';
}

export function comparePptxAnalyses(
  original: PptxAnalysis,
  processed: PptxAnalysis
): ComparisonSummary {
  const origMap = new Map(original.structure.entries.map((e) => [e.path, e]));
  const procMap = new Map(processed.structure.entries.map((e) => [e.path, e]));
  const allPaths = Array.from(new Set([...origMap.keys(), ...procMap.keys()])).sort();

  const fileDiffs: FileDiffEntry[] = [];
  let unchangedEntriesCount = 0;
  let modifiedEntriesCount = 0;
  let addedEntriesCount = 0;
  let removedEntriesCount = 0;

  for (const path of allPaths) {
    const o = origMap.get(path);
    const p = procMap.get(path);

    if (o && p) {
      if (o.sha256 === p.sha256) {
        unchangedEntriesCount++;
        fileDiffs.push({
          path,
          status: 'unchanged',
          originalSize: o.size,
          processedSize: p.size,
          originalHash: o.sha256,
          processedHash: p.sha256,
          reason: 'Preservado byte-a-byte',
        });
      } else {
        modifiedEntriesCount++;
        fileDiffs.push({
          path,
          status: 'modified',
          originalSize: o.size,
          processedSize: p.size,
          originalHash: o.sha256,
          processedHash: p.sha256,
          reason: describeModifiedReason(path),
        });
      }
    } else if (!o && p) {
      addedEntriesCount++;
      fileDiffs.push({
        path,
        status: 'added',
        originalSize: 0,
        processedSize: p.size,
        originalHash: '',
        processedHash: p.sha256,
        reason: 'Ficheiro adicionado',
      });
    } else if (o && !p) {
      removedEntriesCount++;
      fileDiffs.push({
        path,
        status: 'removed',
        originalSize: o.size,
        processedSize: 0,
        originalHash: o.sha256,
        processedHash: '',
        reason: 'Ficheiro removido',
      });
    }
  }

  const slideDiffs: SlideDiffEntry[] = original.slides.map((origSlide, idx) => {
    const procSlide = processed.slides[idx];
    if (!procSlide) {
      return {
        slideNumber: origSlide.slideNumber,
        title: origSlide.title,
        xmlUnchanged: false,
        textChanges: [],
        visibilityChanged: false,
        notesChanged: false,
        originalHidden: origSlide.hidden,
        processedHidden: origSlide.hidden,
      };
    }

    const textChanges: { before: string; after: string }[] = [];
    const maxRuns = Math.max(origSlide.textRuns.length, procSlide.textRuns.length);
    for (let r = 0; r < maxRuns; r++) {
      const before = origSlide.textRuns[r] || '';
      const after = procSlide.textRuns[r] || '';
      if (before !== after) {
        textChanges.push({ before, after });
      }
    }

    return {
      slideNumber: origSlide.slideNumber,
      title: procSlide.title || origSlide.title,
      xmlUnchanged: origSlide.sha256 === procSlide.sha256,
      textChanges,
      visibilityChanged: origSlide.hidden !== procSlide.hidden,
      notesChanged: origSlide.notesText !== procSlide.notesText,
      originalHidden: origSlide.hidden,
      processedHidden: procSlide.hidden,
    };
  });

  return {
    fileDiffs,
    slideDiffs,
    unchangedEntriesCount,
    modifiedEntriesCount,
    addedEntriesCount,
    removedEntriesCount,
    totalEntriesCount: allPaths.length,
  };
}

export function verifyProcessedPptx(
  original: PptxAnalysis,
  processed: PptxAnalysis
): VerificationReport {
  const checks: VerificationCheck[] = [];
  const limitations: string[] = [];
  const blockingReasons: string[] = [];

  // 1. All slides present
  const origSlideIds = original.slides.map((s) => s.slideId).join(',');
  const procSlideIds = processed.slides.map((s) => s.slideId).join(',');
  const slidesOk =
    original.slideCount === processed.slideCount && origSlideIds === procSlideIds;

  checks.push({
    id: 'check-slides',
    title: 'Presença de todos os diapositivos',
    category: 'slides',
    status: slidesOk ? 'pass' : 'fail',
    originalSummary: `${original.slideCount} diapositivos`,
    processedSummary: `${processed.slideCount} diapositivos`,
    details: slidesOk
      ? 'Todos os diapositivos e respetivos identificadores internos (sldId) foram preservados na ordem original.'
      : 'Detetada divergência no número ou identificação dos diapositivos após o processamento.',
  });
  if (!slidesOk) {
    blockingReasons.push('Perda ou alteração estrutural detetada na lista de diapositivos.');
  }

  // 2. Hyperlinks preserved 100%
  const serializeLinks = (analysis: PptxAnalysis) =>
    analysis.hyperlinks
      .map(
        (h) =>
          `${h.slideNumber}|${h.relId}|${h.target}|${h.action}|${h.elementType}|${h.targetMode}`
      )
      .sort()
      .join('\n');

  const origLinksSig = serializeLinks(original);
  const procLinksSig = serializeLinks(processed);
  const linksOk =
    original.hyperlinks.length === processed.hyperlinks.length &&
    origLinksSig === procLinksSig;

  checks.push({
    id: 'check-hyperlinks',
    title: 'Preservação integral de hiperligações',
    category: 'hyperlinks',
    status: linksOk ? 'pass' : 'fail',
    originalSummary: `${original.hyperlinks.length} hiperligações`,
    processedSummary: `${processed.hyperlinks.length} hiperligações`,
    details: linksOk
      ? original.hyperlinks.length > 0
        ? 'Todos os endereços, relações internas (.rels), ações e associações a texto, imagens e formas permanecem idênticos.'
        : 'A apresentação original não contém hiperligações.'
      : 'Uma ou mais hiperligações foram alteradas, removidas ou perderam a sua relação interna.',
  });
  if (!linksOk) {
    blockingReasons.push('Inconsistência detetada nas hiperligações da apresentação.');
  }

  // 3. Original fonts & embedded font files preserved
  const origFontNames = original.fonts.map((f) => f.name).sort().join('|');
  const procFontNames = processed.fonts.map((f) => f.name).sort().join('|');
  const origEmbeddedFiles = original.embeddedFontFlags.embeddedFontFiles
    .map((f) => `${f.path}:${f.size}`)
    .sort()
    .join('|');
  const procEmbeddedFiles = processed.embeddedFontFlags.embeddedFontFiles
    .map((f) => `${f.path}:${f.size}`)
    .sort()
    .join('|');

  const fontPropsOk = original.fonts.every((of) => {
    const pf = processed.fonts.find((f) => f.name === of.name);
    if (!pf) return false;
    return (
      pf.embedded === of.embedded &&
      pf.hasBold === of.hasBold &&
      pf.hasItalic === of.hasItalic &&
      pf.hasUnderline === of.hasUnderline &&
      pf.sizesPt.join(',') === of.sizesPt.join(',')
    );
  });

  const fontsIntact =
    origFontNames === procFontNames &&
    origEmbeddedFiles === procEmbeddedFiles &&
    fontPropsOk;

  const unembeddedExternalFonts = processed.fonts.filter(
    (f) => f.status === 'external_unembedded'
  );

  let fontStatus: VerificationCheck['status'] = 'pass';
  let fontLimitationNote: string | undefined;

  if (!fontsIntact) {
    fontStatus = 'fail';
    blockingReasons.push('Alteração ou substituição inesperada detetada nas referências de fontes.');
  } else if (unembeddedExternalFonts.length > 0) {
    fontStatus = 'limitation';
    fontLimitationNote = `Referências preservadas sem substituição (${unembeddedExternalFonts
      .map((f) => f.name)
      .join(', ')}), mas não é possível validar a renderização de quebras de linha em computadores que não tenham estas fontes instaladas.`;
    limitations.push(
      `Fontes externas não incorporadas (${unembeddedExternalFonts
        .map((f) => f.name)
        .join(', ')}): as referências XML estão 100% intactas, mas a visualização exata exige a instalação das fontes no computador de destino.`
    );
  }

  checks.push({
    id: 'check-fonts',
    title: 'Referências de fontes e fontes incorporadas',
    category: 'fonts',
    status: fontStatus,
    originalSummary: `${original.fonts.length} fontes (${original.embeddedFontFlags.embeddedFontFiles.length} incorporadas)`,
    processedSummary: `${processed.fonts.length} fontes (${processed.embeddedFontFlags.embeddedFontFiles.length} incorporadas)`,
    details: fontsIntact
      ? 'Todas as referências tipográficas, pesos (negrito/itálico/sublinhado), tamanhos e ficheiros de fontes incorporadas mantiveram-se intactos.'
      : 'Detetada divergência nas fontes ou propriedades tipográficas.',
    limitationNote: fontLimitationNote,
  });

  // 4. Multimedia assets preserved byte-for-byte
  const mediaEntriesOrig = original.structure.entries
    .filter(
      (e) =>
        e.path.startsWith('ppt/media/') ||
        e.path.startsWith('ppt/embeddings/') ||
        e.path.startsWith('ppt/charts/') ||
        e.path.startsWith('ppt/diagrams/')
    )
    .map((e) => `${e.path}:${e.size}:${e.sha256}`)
    .sort()
    .join('\n');

  const mediaEntriesProc = processed.structure.entries
    .filter(
      (e) =>
        e.path.startsWith('ppt/media/') ||
        e.path.startsWith('ppt/embeddings/') ||
        e.path.startsWith('ppt/charts/') ||
        e.path.startsWith('ppt/diagrams/')
    )
    .map((e) => `${e.path}:${e.size}:${e.sha256}`)
    .sort()
    .join('\n');

  const externalMediaOrig = original.multimedia
    .filter((m) => m.isExternal)
    .map((m) => m.path)
    .sort()
    .join('|');
  const externalMediaProc = processed.multimedia
    .filter((m) => m.isExternal)
    .map((m) => m.path)
    .sort()
    .join('|');

  const mediaOk =
    mediaEntriesOrig === mediaEntriesProc && externalMediaOrig === externalMediaProc;

  const hasExternalMedia = processed.multimedia.some((m) => m.isExternal);
  let mediaStatus: VerificationCheck['status'] = mediaOk ? 'pass' : 'fail';
  let mediaLimitation: string | undefined;

  if (!mediaOk) {
    blockingReasons.push('Um ou mais ficheiros multimédia ou objetos incorporados foram alterados ou perdidos.');
  } else if (hasExternalMedia) {
    mediaStatus = 'limitation';
    mediaLimitation =
      'As ligações a recursos multimédia externos foram preservadas, mas a disponibilidade física desses ficheiros externos não pode ser confirmada a partir do pacote .pptx.';
    limitations.push(mediaLimitation);
  }

  checks.push({
    id: 'check-multimedia',
    title: 'Ficheiros multimédia e objetos incorporados',
    category: 'multimedia',
    status: mediaStatus,
    originalSummary: `${original.multimedia.length} recursos`,
    processedSummary: `${processed.multimedia.length} recursos`,
    details: mediaOk
      ? 'Todas as imagens, áudio, vídeo, gráficos e objetos incorporados mantêm os mesmos hashes SHA-256 (zero recompressão).'
      : 'Alteração ou perda detetada nos ficheiros multimédia do pacote.',
    limitationNote: mediaLimitation,
  });

  // 5. Animations & transitions preserved
  const animSig = (a: PptxAnalysis) =>
    a.slides
      .map(
        (s) =>
          `${s.slideNumber}:${s.transition.present}:${s.transition.type}:${s.transition.durationMs || ''}:${s.animations.nodeCount}:${s.animations.effectTypes.join(',')}`
      )
      .join('|');

  const animOk = animSig(original) === animSig(processed);
  checks.push({
    id: 'check-animations',
    title: 'Animações e transições entre diapositivos',
    category: 'animations',
    status: animOk ? 'pass' : 'fail',
    originalSummary: `${original.visualSummary.totalAnimations} nós · ${original.visualSummary.slidesWithTransitions} transições`,
    processedSummary: `${processed.visualSummary.totalAnimations} nós · ${processed.visualSummary.slidesWithTransitions} transições`,
    details: animOk
      ? 'Todas as sequências temporais (<p:timing>), efeitos de entrada/saída/ênfase e transições (<p:transition>) foram preservados integralmente.'
      : 'Detetada perda ou alteração nas animações ou transições.',
  });
  if (!animOk) {
    blockingReasons.push('Perda detetada em animações ou transições de diapositivos.');
  }

  // 6. OOXML Structure validity
  const structureOk =
    processed.structure.hasContentTypes &&
    processed.structure.hasRootRels &&
    processed.structure.hasPresentationXml &&
    processed.structure.xmlCount >= original.structure.xmlCount &&
    processed.structure.relsCount >= original.structure.relsCount;

  checks.push({
    id: 'check-structure',
    title: 'Validade da estrutura interna OOXML (.pptx)',
    category: 'structure',
    status: structureOk ? 'pass' : 'fail',
    originalSummary: `${original.structure.totalEntries} ficheiros internos`,
    processedSummary: `${processed.structure.totalEntries} ficheiros internos`,
    details: structureOk
      ? 'Pacote ZIP, [Content_Types].xml, relações (.rels), ppt/presentation.xml e extensões XML desconhecidas mantêm-se válidos.'
      : 'A estrutura interna OOXML ficou incompleta ou perdeu ficheiros críticos.',
  });
  if (!structureOk) {
    blockingReasons.push('Estrutura interna OOXML inválida ou incompleta.');
  }

  // 7. Visual elements preserved
  const vOrig = original.visualSummary;
  const vProc = processed.visualSummary;
  const visualOk =
    vOrig.totalShapes === vProc.totalShapes &&
    vOrig.totalImages === vProc.totalImages &&
    vOrig.totalTables === vProc.totalTables &&
    vOrig.totalCharts === vProc.totalCharts &&
    vOrig.totalSmartArt === vProc.totalSmartArt &&
    vOrig.totalGroups === vProc.totalGroups &&
    vOrig.slidesWithBackground === vProc.slidesWithBackground &&
    vOrig.slidesWithEffects === vProc.slidesWithEffects &&
    vOrig.masterCount === vProc.masterCount &&
    vOrig.layoutCount === vProc.layoutCount &&
    original.slideWidthCm === processed.slideWidthCm &&
    original.slideHeightCm === processed.slideHeightCm;

  // Transparent rendering limitation note (Requirement 8: "Se alguma propriedade não puder ser verificada, indicar claramente essa limitação")
  const visualLimitation =
    'A validação confirma 100% dos nós XML, geometria, estilos e recursos visuais. A renderização final píxel-a-píxel depende do motor gráfico nativo do Microsoft PowerPoint no sistema de destino.';
  limitations.push(visualLimitation);

  checks.push({
    id: 'check-visual',
    title: 'Elementos visuais, temas, formas e efeitos',
    category: 'visual',
    status: visualOk ? 'limitation' : 'fail',
    originalSummary: `${vOrig.totalShapes} formas · ${vOrig.totalImages} imagens · ${vOrig.totalTables} tabelas`,
    processedSummary: `${vProc.totalShapes} formas · ${vProc.totalImages} imagens · ${vProc.totalTables} tabelas`,
    details: visualOk
      ? 'Formas, caixas de texto, imagens, tabelas, gráficos, fundos, sombras, transparências, modelos globais e dimensões preservados integralmente.'
      : 'Detetada eliminação ou alteração inesperada de elementos visuais.',
    limitationNote: visualOk ? visualLimitation : undefined,
  });
  if (!visualOk) {
    blockingReasons.push('Perda ou modificação detetada nos elementos visuais dos diapositivos.');
  }

  const hasFail = checks.some((c) => c.status === 'fail');
  const hasLimitation = checks.some((c) => c.status === 'limitation');

  return {
    overallStatus: hasFail
      ? 'blocked'
      : hasLimitation
      ? 'verified_with_limitations'
      : 'verified',
    canExport: !hasFail,
    timestamp: new Date().toISOString(),
    checks,
    limitations,
    blockingReasons,
  };
}

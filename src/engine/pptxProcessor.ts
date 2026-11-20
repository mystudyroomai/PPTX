import JSZip from 'jszip';
import type {
  PptxAnalysis,
  ProcessingOptions,
  ProcessingResult,
} from '../types/pptx';
import {
  analyzePptxBuffer,
  getElementsByLocalName,
  parseXmlSafe,
} from './pptxAnalyzer';
import { comparePptxAnalyses, verifyProcessedPptx } from './pptxVerifier';

function escapeRegExp(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function serializeXmlPreservingDecl(doc: Document, originalXml: string): string {
  const serializer = new XMLSerializer();
  let serialized = serializer.serializeToString(doc);
  const origDeclMatch = originalXml.match(/^<\?xml[^?]*\?>/);
  const hasDecl = /^<\?xml[^?]*\?>/.test(serialized);
  if (origDeclMatch && !hasDecl) {
    serialized = `${origDeclMatch[0]}\n${serialized}`;
  }
  return serialized;
}

function updateCorePropertyElement(
  doc: Document,
  localName: string,
  value: string
): boolean {
  const els = getElementsByLocalName(doc, localName);
  if (els.length > 0) {
    if (els[0].textContent !== value) {
      els[0].textContent = value;
      return true;
    }
    return false;
  }
  return false;
}

function buildOutputFileName(originalName: string): string {
  const clean = originalName.replace(/\.pptx$/i, '');
  return `${clean}_preservado.pptx`;
}

export async function processPptxSafely(
  originalBuffer: Uint8Array,
  originalAnalysis: PptxAnalysis,
  options: ProcessingOptions
): Promise<ProcessingResult> {
  const outputFileName = buildOutputFileName(originalAnalysis.fileName);
  const operationsApplied: string[] = [];
  const skippedProtectedRuns: { slideNumber: number; text: string; reason: string }[] = [];

  try {
    // 1. Create an isolated working copy from the immutable original buffer.
    // The original Uint8Array is NEVER modified.
    const workingBytes = new Uint8Array(originalBuffer.byteLength);
    workingBytes.set(originalBuffer);
    const zip = await JSZip.loadAsync(workingBytes);

    const activeReplacements = options.textReplacements.filter(
      (r) => r.find.trim().length > 0 && r.find !== r.replace
    );

    // 2. Apply strictly non-destructive slide operations ONLY to slides that need them
    let totalTextReplacements = 0;
    let totalVisibilityUpdates = 0;

    for (const slideInfo of originalAnalysis.slides) {
      const targetHidden = options.slideVisibilityOverrides[slideInfo.slideNumber];
      const wantsVisibilityChange =
        typeof targetHidden === 'boolean' && targetHidden !== slideInfo.hidden;

      if (activeReplacements.length === 0 && !wantsVisibilityChange) {
        // Leave untouched slide XML bytes 100% verbatim in the ZIP
        continue;
      }

      const slideFile = zip.file(slideInfo.path);
      if (!slideFile) {
        throw new Error(`Diapositivo ${slideInfo.slideNumber} (${slideInfo.path}) inacessível na cópia de trabalho.`);
      }

      const originalSlideXml = await slideFile.async('string');
      const { doc: sDoc, error: parseErr } = parseXmlSafe(originalSlideXml);
      if (!sDoc) {
        throw new Error(
          `Erro de leitura XML no diapositivo ${slideInfo.slideNumber}: ${parseErr}`
        );
      }

      let slideModified = false;

      // 2a. Non-destructive visibility attribute toggle (`show="0"`)
      if (wantsVisibilityChange) {
        const root = sDoc.documentElement;
        if (targetHidden) {
          root.setAttribute('show', '0');
        } else {
          root.removeAttribute('show');
        }
        slideModified = true;
        totalVisibilityUpdates++;
      }

      // 2b. Safe text replacement inside <a:t> ONLY (never touching <a:rPr>, fonts, sizes, or <a:hlinkClick>)
      if (activeReplacements.length > 0) {
        const runs = getElementsByLocalName(sDoc, 'r');
        for (const r of runs) {
          const tNodes = getElementsByLocalName(r, 't');
          if (tNodes.length === 0) continue;

          // Check if this run carries a hyperlink
          const rPr = getElementsByLocalName(r, 'rPr')[0];
          const hasHyperlink =
            rPr &&
            (getElementsByLocalName(rPr, 'hlinkClick').length > 0 ||
              getElementsByLocalName(rPr, 'hlinkHover').length > 0);

          for (const tNode of tNodes) {
            const currentText = tNode.textContent || '';
            if (!currentText) continue;

            let updatedText = currentText;
            let matched = false;

            for (const rule of activeReplacements) {
              const regex = new RegExp(
                escapeRegExp(rule.find),
                rule.caseSensitive ? 'g' : 'gi'
              );
              if (regex.test(updatedText)) {
                matched = true;
                if (hasHyperlink && options.protectHyperlinkedText) {
                  skippedProtectedRuns.push({
                    slideNumber: slideInfo.slideNumber,
                    text: currentText,
                    reason:
                      'Protegido automaticamente porque o texto possui uma hiperligação associada.',
                  });
                } else {
                  updatedText = updatedText.replace(regex, rule.replace);
                }
              }
            }

            if (matched && updatedText !== currentText) {
              tNode.textContent = updatedText;
              slideModified = true;
              totalTextReplacements++;
            }
          }
        }
      }

      if (slideModified) {
        const updatedXml = serializeXmlPreservingDecl(sDoc, originalSlideXml);
        // Validate newly serialized XML before writing back to working copy
        const validation = parseXmlSafe(updatedXml);
        if (!validation.doc) {
          throw new Error(
            `Abortado: a serialização XML do diapositivo ${slideInfo.slideNumber} falhou na validação de integridade.`
          );
        }
        zip.file(slideInfo.path, updatedXml);
      }
    }

    if (totalTextReplacements > 0) {
      operationsApplied.push(
        `${totalTextReplacements} ocorrência(s) de texto substituída(s) preservando <a:rPr>, fontes e hiperligações.`
      );
    }
    if (totalVisibilityUpdates > 0) {
      operationsApplied.push(
        `Visibilidade atualizada em ${totalVisibilityUpdates} diapositivo(s) sem remoção de conteúdo.`
      );
    }

    // 3. Safe Document Metadata update (docProps/core.xml only)
    if (options.metadataUpdates.enabled && zip.file('docProps/core.xml')) {
      const origCoreXml = await zip.file('docProps/core.xml')!.async('string');
      const { doc: coreDoc } = parseXmlSafe(origCoreXml);
      if (coreDoc) {
        let coreChanged = false;
        if (options.metadataUpdates.title !== originalAnalysis.coreProperties.title) {
          if (updateCorePropertyElement(coreDoc, 'title', options.metadataUpdates.title)) {
            coreChanged = true;
          }
        }
        if (options.metadataUpdates.subject !== originalAnalysis.coreProperties.subject) {
          if (updateCorePropertyElement(coreDoc, 'subject', options.metadataUpdates.subject)) {
            coreChanged = true;
          }
        }
        if (options.metadataUpdates.anonymizeAuthor) {
          if (updateCorePropertyElement(coreDoc, 'creator', '')) coreChanged = true;
          if (updateCorePropertyElement(coreDoc, 'lastModifiedBy', '')) coreChanged = true;
        } else if (options.metadataUpdates.creator !== originalAnalysis.coreProperties.creator) {
          if (updateCorePropertyElement(coreDoc, 'creator', options.metadataUpdates.creator)) {
            coreChanged = true;
          }
        }

        if (coreChanged) {
          const updatedCore = serializeXmlPreservingDecl(coreDoc, origCoreXml);
          if (!parseXmlSafe(updatedCore).doc) {
            throw new Error('Abortado: falha na validação XML de docProps/core.xml.');
          }
          zip.file('docProps/core.xml', updatedCore);
          operationsApplied.push('Metadados do documento (docProps/core.xml) atualizados sem tocar nos diapositivos.');
        }
      }
    }

    // 4. Optional Speaker Notes cleaning (modifies ONLY notesSlides XML, never slide XML)
    if (options.cleanSpeakerNotes) {
      let cleanedNotesCount = 0;
      for (const slideInfo of originalAnalysis.slides) {
        if (!slideInfo.notesPath || !slideInfo.notesText) continue;
        const nFile = zip.file(slideInfo.notesPath);
        if (!nFile) continue;
        const origNotesXml = await nFile.async('string');
        const { doc: nDoc } = parseXmlSafe(origNotesXml);
        if (!nDoc) continue;

        let changed = false;
        const spNodes = getElementsByLocalName(nDoc, 'sp');
        for (const sp of spNodes) {
          const ph = getElementsByLocalName(sp, 'ph')[0];
          if (ph?.getAttribute('type') === 'sldNum') continue;
          const tNodes = getElementsByLocalName(sp, 't');
          for (const t of tNodes) {
            if (t.textContent && t.textContent.trim().length > 0) {
              t.textContent = '';
              changed = true;
            }
          }
        }

        if (changed) {
          const updatedNotes = serializeXmlPreservingDecl(nDoc, origNotesXml);
          if (!parseXmlSafe(updatedNotes).doc) {
            throw new Error(`Abortado: falha na validação XML de ${slideInfo.notesPath}.`);
          }
          zip.file(slideInfo.notesPath, updatedNotes);
          cleanedNotesCount++;
        }
      }
      if (cleanedNotesCount > 0) {
        operationsApplied.push(
          `Texto das notas do orador limpo em ${cleanedNotesCount} diapositivo(s) mantendo os diapositivos intactos.`
        );
      }
    }

    if (operationsApplied.length === 0) {
      operationsApplied.push(
        'Reempacotamento de preservação integral concluído (nenhum diapositivo ou recurso foi modificado).'
      );
    }

    // 5. Generate the processed PPTX archive
    const processedBuffer = await zip.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });

    // 6. Independent full re-analysis of the output PPTX from scratch
    const processedAnalysis = await analyzePptxBuffer(processedBuffer, outputFileName);

    // 7. Run strict Verification & Comparison Gate
    const verification = verifyProcessedPptx(originalAnalysis, processedAnalysis);
    const comparison = comparePptxAnalyses(originalAnalysis, processedAnalysis);

    // 8. Enforce Regra Final & Requirement 9: If ANY check failed, CANCEL export!
    if (!verification.canExport) {
      return {
        success: false,
        aborted: true,
        abortReason:
          verification.blockingReasons.join(' ') ||
          'A verificação automática detetou risco de perda de elementos. A exportação foi cancelada para proteger a fidelidade da apresentação.',
        processedFileName: outputFileName,
        processedAnalysis,
        verification,
        comparison,
        operationsApplied,
        skippedProtectedRuns,
      };
    }

    return {
      success: true,
      aborted: false,
      processedBuffer,
      processedFileName: outputFileName,
      processedAnalysis,
      verification,
      comparison,
      operationsApplied,
      skippedProtectedRuns,
    };
  } catch (err) {
    const message =
      err instanceof Error ? err.message : 'Erro inesperado durante o processamento do ficheiro PPTX.';
    return {
      success: false,
      aborted: true,
      abortReason: message,
      processedFileName: outputFileName,
      verification: {
        overallStatus: 'blocked',
        canExport: false,
        timestamp: new Date().toISOString(),
        checks: [],
        limitations: [],
        blockingReasons: [message],
      },
      operationsApplied,
      skippedProtectedRuns,
    };
  }
}

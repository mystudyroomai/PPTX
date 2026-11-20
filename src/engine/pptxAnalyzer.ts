import JSZip from 'jszip';
import type {
  CompatibilityIssue,
  CoreProperties,
  FontAuditEntry,
  HyperlinkAuditEntry,
  MultimediaAuditEntry,
  PptxAnalysis,
  SlideAnimationInfo,
  SlideAuditEntry,
  SlideTransitionInfo,
  VisualSummary,
  ZipEntryInfo,
} from '../types/pptx';

const STANDARD_OFFICE_FONTS = new Set([
  'arial',
  'arial black',
  'arial narrow',
  'calibri',
  'calibri light',
  'cambria',
  'candara',
  'century gothic',
  'comic sans ms',
  'consolas',
  'constantia',
  'corbel',
  'courier new',
  'franklin gothic medium',
  'garamond',
  'georgia',
  'helvetica',
  'impact',
  'lucida console',
  'lucida sans unicode',
  'palatino linotype',
  'segoe ui',
  'segoe ui light',
  'segoe ui semibold',
  'symbol',
  'tahoma',
  'times new roman',
  'trebuchet ms',
  'verdana',
  'wingdings',
  'aptos',
  'aptos display',
  'aptos narrow',
]);

const TRANSITION_LABELS: Record<string, string> = {
  fade: 'Desvanecer (Fade)',
  push: 'Empurrar (Push)',
  wipe: 'Apagar (Wipe)',
  split: 'Dividir (Split)',
  reveal: 'Revelar (Reveal)',
  cut: 'Cortar (Cut)',
  cover: 'Cobrir (Cover)',
  pull: 'Puxar (Pull)',
  randomBar: 'Barras Aleatórias',
  blinds: 'Persianas',
  checker: 'Xadrez',
  circle: 'Círculo',
  dissolve: 'Dissolver',
  comb: 'Pente',
  diamond: 'Losango',
  newsflash: 'Flash',
  plus: 'Mais',
  random: 'Aleatória',
  strips: 'Tiras',
  wedge: 'Cunha',
  wheel: 'Roda',
  zoom: 'Zoom',
  morph: 'Transformar (Morph)',
  flythrough: 'Voo Direto',
  glitter: 'Brilho',
  honeycomb: 'Favo de Mel',
  ripple: 'Ondulação',
  shred: 'Triturar',
  vortex: 'Vórtice',
  switch: 'Alternar',
  flip: 'Virar',
  gallery: 'Galeria',
  cube: 'Cubo',
  doors: 'Portas',
  box: 'Caixa',
  pan: 'Panorâmica',
  ferris: 'Roda Gigante',
  conveyor: 'Tapete Rolante',
  rotate: 'Rodar',
  window: 'Janela',
  orbit: 'Órbita',
  drape: 'Cortina',
  curtains: 'Cortinas',
  wind: 'Vento',
  prestige: 'Prestígio',
  fracture: 'Fratura',
  crush: 'Esmagar',
  peelOff: 'Descolar',
  pageCurlDouble: 'Virar Página',
  airplane: 'Avião',
  origami: 'Origami',
};

export async function sha256Hex(data: Uint8Array | ArrayBuffer | string): Promise<string> {
  let bytes: Uint8Array;
  if (typeof data === 'string') {
    bytes = new TextEncoder().encode(data);
  } else if (data instanceof ArrayBuffer) {
    bytes = new Uint8Array(data);
  } else {
    bytes = data;
  }
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  const hashBuffer = await crypto.subtle.digest('SHA-256', copy.buffer);
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  return hashArray.map((b) => b.toString(16).padStart(2, '0')).join('');
}

export function parseXmlSafe(xmlText: string): { doc: Document | null; error?: string } {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, 'application/xml');
  const parseError = doc.getElementsByTagName('parsererror');
  if (parseError.length > 0) {
    return {
      doc: null,
      error: parseError[0].textContent || 'Erro de sintaxe XML',
    };
  }
  return { doc };
}

export function getElementsByLocalName(root: Document | Element, localName: string): Element[] {
  const all = root.getElementsByTagName('*');
  const result: Element[] = [];
  for (let i = 0; i < all.length; i++) {
    const el = all[i];
    if (el.localName === localName) {
      result.push(el);
    }
  }
  return result;
}

export function getAttrByLocalName(el: Element, localName: string): string | null {
  if (el.hasAttribute(localName)) return el.getAttribute(localName);
  for (let i = 0; i < el.attributes.length; i++) {
    const attr = el.attributes[i];
    if (attr.localName === localName) {
      return attr.value;
    }
  }
  return null;
}

interface RelEntry {
  id: string;
  type: string;
  target: string;
  targetMode: string;
}

export function parseRelationships(relsXml: string | null): Map<string, RelEntry> {
  const map = new Map<string, RelEntry>();
  if (!relsXml) return map;
  const { doc } = parseXmlSafe(relsXml);
  if (!doc) return map;
  const rels = getElementsByLocalName(doc, 'Relationship');
  for (const rel of rels) {
    const id = rel.getAttribute('Id') || '';
    const type = rel.getAttribute('Type') || '';
    const target = rel.getAttribute('Target') || '';
    const targetMode = rel.getAttribute('TargetMode') || 'Internal';
    if (id) {
      map.set(id, { id, type, target, targetMode });
    }
  }
  return map;
}

export function resolveZipPath(baseFilePath: string, relativeTarget: string): string {
  if (!relativeTarget) return '';
  if (/^[a-zA-Z]+:\/\//.test(relativeTarget) || relativeTarget.startsWith('mailto:')) {
    return relativeTarget;
  }
  if (relativeTarget.startsWith('/')) {
    return relativeTarget.replace(/^\/+/, '');
  }
  const baseParts = baseFilePath.split('/');
  baseParts.pop(); // remove filename
  const targetParts = relativeTarget.split('/');
  for (const part of targetParts) {
    if (part === '' || part === '.') continue;
    if (part === '..') {
      baseParts.pop();
    } else {
      baseParts.push(part);
    }
  }
  return baseParts.join('/');
}

function getRelsPathForPart(partPath: string): string {
  const idx = partPath.lastIndexOf('/');
  if (idx === -1) return `_rels/${partPath}.rels`;
  const dir = partPath.slice(0, idx);
  const file = partPath.slice(idx + 1);
  return `${dir}/_rels/${file}.rels`;
}

function detectAspectRatio(widthCm: number, heightCm: number): string {
  if (!widthCm || !heightCm) return 'Personalizado';
  const ratio = widthCm / heightCm;
  if (Math.abs(ratio - 16 / 9) < 0.05) return '16:9 (Panorâmico)';
  if (Math.abs(ratio - 4 / 3) < 0.05) return '4:3 (Clássico)';
  if (Math.abs(ratio - 16 / 10) < 0.05) return '16:10';
  return `${ratio.toFixed(2)}:1`;
}

function classifyMediaKind(path: string, relType?: string): MultimediaAuditEntry['kind'] {
  const lower = path.toLowerCase();
  if (lower.startsWith('ppt/charts/')) return 'chart';
  if (lower.startsWith('ppt/diagrams/')) return 'smartart';
  if (lower.startsWith('ppt/embeddings/')) return 'ole_embedded';
  if (relType && relType.includes('audio')) return 'audio';
  if (relType && relType.includes('video')) return 'video';
  if (/\.(mp3|wav|m4a|wma|ogg|flac|aac)$/i.test(lower)) return 'audio';
  if (/\.(mp4|mov|wmv|avi|mkv|webm|m4v|mpg|mpeg)$/i.test(lower)) return 'video';
  return 'image';
}

function getExtension(path: string): string {
  const match = path.match(/\.([a-zA-Z0-9]+)$/);
  return match ? match[1].toLowerCase() : 'desconhecido';
}

function findAncestorByLocalName(el: Element | null, names: string[]): Element | null {
  let curr: Node | null = el?.parentNode || null;
  while (curr && curr.nodeType === 1) {
    const elem = curr as Element;
    if (names.includes(elem.localName)) return elem;
    curr = curr.parentNode;
  }
  return null;
}

function extractTextFromElement(el: Element): string {
  const texts = getElementsByLocalName(el, 't')
    .map((t) => t.textContent || '')
    .join('');
  return texts.trim();
}

export async function analyzePptxBuffer(
  buffer: Uint8Array,
  fileName: string
): Promise<PptxAnalysis> {
  const zip = await JSZip.loadAsync(buffer);
  const fileSha256 = await sha256Hex(buffer);

  // Inventory all ZIP entries with sizes and hashes
  const entryNames = Object.keys(zip.files).filter((k) => !zip.files[k].dir).sort();
  const entries: ZipEntryInfo[] = [];
  const unknownExtensionSet = new Set<string>();

  for (const path of entryNames) {
    const fileObj = zip.files[path];
    const bytes = await fileObj.async('uint8array');
    const hash = await sha256Hex(bytes);
    const lower = path.toLowerCase();
    const isXml = lower.endsWith('.xml');
    const isRels = lower.endsWith('.rels');
    const isMedia = lower.startsWith('ppt/media/');
    const isFont = lower.startsWith('ppt/fonts/') || /\.(fntdata|odttf|ttf|otf)$/i.test(lower);

    entries.push({
      path,
      size: bytes.byteLength,
      sha256: hash,
      isXml,
      isRels,
      isMedia,
      isFont,
    });
  }

  const readText = async (path: string): Promise<string | null> => {
    const f = zip.file(path);
    if (!f) return null;
    return f.async('string');
  };

  const hasContentTypes = Boolean(zip.file('[Content_Types].xml'));
  const hasRootRels = Boolean(zip.file('_rels/.rels'));
  const presentationXmlText = await readText('ppt/presentation.xml');
  const hasPresentationXml = Boolean(presentationXmlText);

  if (!hasContentTypes || !hasPresentationXml) {
    throw new Error(
      'O ficheiro selecionado não possui uma estrutura Microsoft PowerPoint (.pptx / OOXML) válida.'
    );
  }

  // Parse core.xml and app.xml
  const coreProperties: CoreProperties = {
    title: '',
    subject: '',
    creator: '',
    lastModifiedBy: '',
    created: '',
    modified: '',
    company: '',
    application: '',
  };

  const coreXmlText = await readText('docProps/core.xml');
  if (coreXmlText) {
    const { doc } = parseXmlSafe(coreXmlText);
    if (doc) {
      coreProperties.title = getElementsByLocalName(doc, 'title')[0]?.textContent?.trim() || '';
      coreProperties.subject = getElementsByLocalName(doc, 'subject')[0]?.textContent?.trim() || '';
      coreProperties.creator = getElementsByLocalName(doc, 'creator')[0]?.textContent?.trim() || '';
      coreProperties.lastModifiedBy =
        getElementsByLocalName(doc, 'lastModifiedBy')[0]?.textContent?.trim() || '';
      coreProperties.created = getElementsByLocalName(doc, 'created')[0]?.textContent?.trim() || '';
      coreProperties.modified =
        getElementsByLocalName(doc, 'modified')[0]?.textContent?.trim() || '';
    }
  }

  const appXmlText = await readText('docProps/app.xml');
  if (appXmlText) {
    const { doc } = parseXmlSafe(appXmlText);
    if (doc) {
      coreProperties.company = getElementsByLocalName(doc, 'Company')[0]?.textContent?.trim() || '';
      coreProperties.application =
        getElementsByLocalName(doc, 'Application')[0]?.textContent?.trim() || '';
    }
  }

  // Parse presentation.xml & presentation.xml.rels
  const { doc: presDoc, error: presErr } = parseXmlSafe(presentationXmlText!);
  if (!presDoc) {
    throw new Error(`Falha ao ler ppt/presentation.xml: ${presErr}`);
  }
  const presRelsText = await readText('ppt/_rels/presentation.xml.rels');
  const presRels = parseRelationships(presRelsText);

  // Slide size
  let slideWidthCm = 33.867;
  let slideHeightCm = 19.05;
  const sldSz = getElementsByLocalName(presDoc, 'sldSz')[0];
  if (sldSz) {
    const cx = Number(sldSz.getAttribute('cx') || '0');
    const cy = Number(sldSz.getAttribute('cy') || '0');
    if (cx > 0 && cy > 0) {
      // 1 cm = 360000 EMU
      slideWidthCm = Number((cx / 360000).toFixed(2));
      slideHeightCm = Number((cy / 360000).toFixed(2));
    }
  }

  // Embedded font settings & list
  const presRoot = presDoc.documentElement;
  const embedTrueTypeFonts = presRoot?.getAttribute('embedTrueTypeFonts') === '1' ||
    presRoot?.getAttribute('embedTrueTypeFonts') === 'true';
  const saveSubsetFonts = presRoot?.getAttribute('saveSubsetFonts') === '1' ||
    presRoot?.getAttribute('saveSubsetFonts') === 'true';

  const embeddedFontMap = new Map<
    string,
    { variants: Set<string>; files: Set<string> }
  >();

  const embeddedFonts = getElementsByLocalName(presDoc, 'embeddedFont');
  for (const ef of embeddedFonts) {
    const fontEl = getElementsByLocalName(ef, 'font')[0];
    const typeface = fontEl?.getAttribute('typeface')?.trim();
    if (!typeface) continue;
    if (!embeddedFontMap.has(typeface)) {
      embeddedFontMap.set(typeface, { variants: new Set(), files: new Set() });
    }
    const rec = embeddedFontMap.get(typeface)!;
    const variantTags: [string, string][] = [
      ['regular', 'Regular'],
      ['bold', 'Negrito'],
      ['italic', 'Itálico'],
      ['boldItalic', 'Negrito Itálico'],
    ];
    for (const [tag, label] of variantTags) {
      const vEl = getElementsByLocalName(ef, tag)[0];
      if (vEl) {
        rec.variants.add(label);
        const rId = getAttrByLocalName(vEl, 'id');
        if (rId && presRels.has(rId)) {
          const resolved = resolveZipPath('ppt/presentation.xml', presRels.get(rId)!.target);
          rec.files.add(resolved);
        }
      }
    }
  }

  const embeddedFontFiles = entries
    .filter((e) => e.isFont)
    .map((e) => ({ path: e.path, size: e.size }));

  // Parse Themes for major / minor fonts & theme name
  let themeMajorLatin = '';
  let themeMinorLatin = '';
  let themeName = 'Tema Predefinido';
  const themePaths = entryNames.filter((p) => /^ppt\/theme\/theme\d*\.xml$/i.test(p));
  for (const tp of themePaths) {
    const tXml = await readText(tp);
    if (!tXml) continue;
    const { doc: tDoc } = parseXmlSafe(tXml);
    if (!tDoc) continue;
    const tRoot = tDoc.documentElement;
    if (tRoot?.getAttribute('name') && themeName === 'Tema Predefinido') {
      themeName = tRoot.getAttribute('name')!;
    }
    const majorFont = getElementsByLocalName(tDoc, 'majorFont')[0];
    const minorFont = getElementsByLocalName(tDoc, 'minorFont')[0];
    const mjLat = majorFont ? getElementsByLocalName(majorFont, 'latin')[0]?.getAttribute('typeface') : '';
    const mnLat = minorFont ? getElementsByLocalName(minorFont, 'latin')[0]?.getAttribute('typeface') : '';
    if (mjLat && !themeMajorLatin) themeMajorLatin = mjLat;
    if (mnLat && !themeMinorLatin) themeMinorLatin = mnLat;
  }

  // Font accumulator
  const fontAccum = new Map<
    string,
    {
      slides: Set<number>;
      usedInMastersOrLayouts: boolean;
      hasBold: boolean;
      hasItalic: boolean;
      hasUnderline: boolean;
      sizesPt: Set<number>;
    }
  >();

  const ensureFont = (name: string) => {
    const clean = name.trim();
    if (!clean) return null;
    if (!fontAccum.has(clean)) {
      fontAccum.set(clean, {
        slides: new Set(),
        usedInMastersOrLayouts: false,
        hasBold: false,
        hasItalic: false,
        hasUnderline: false,
        sizesPt: new Set(),
      });
    }
    return fontAccum.get(clean)!;
  };

  const resolveFontName = (raw: string): string => {
    if (!raw) return '';
    if (raw === '+mj-lt' || raw === '+mj-ea' || raw === '+mj-cs') {
      return themeMajorLatin || raw;
    }
    if (raw === '+mn-lt' || raw === '+mn-ea' || raw === '+mn-cs') {
      return themeMinorLatin || raw;
    }
    return raw;
  };

  if (themeMajorLatin) ensureFont(themeMajorLatin);
  if (themeMinorLatin) ensureFont(themeMinorLatin);
  for (const efName of embeddedFontMap.keys()) {
    ensureFont(efName);
  }

  // Scan masters & layouts for fonts and count them
  const masterPaths = entryNames.filter((p) => /^ppt\/slideMasters\/slideMaster\d+\.xml$/i.test(p));
  const layoutPaths = entryNames.filter((p) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/i.test(p));

  const scanXmlForExtensions = (doc: Document) => {
    const altContents = getElementsByLocalName(doc, 'AlternateContent');
    if (altContents.length > 0) {
      unknownExtensionSet.add('mc:AlternateContent (Extensões PowerPoint)');
    }
    const extLst = getElementsByLocalName(doc, 'ext');
    for (const ext of extLst) {
      const uri = ext.getAttribute('uri');
      if (uri) {
        unknownExtensionSet.add(`ext ${uri}`);
      }
    }
  };

  for (const mp of [...masterPaths, ...layoutPaths]) {
    const xml = await readText(mp);
    if (!xml) continue;
    const { doc } = parseXmlSafe(xml);
    if (!doc) continue;
    scanXmlForExtensions(doc);
    for (const tag of ['latin', 'ea', 'cs', 'sym']) {
      const nodes = getElementsByLocalName(doc, tag);
      for (const n of nodes) {
        const tf = resolveFontName(n.getAttribute('typeface') || '');
        const rec = ensureFont(tf);
        if (rec) rec.usedInMastersOrLayouts = true;
      }
    }
  }

  // Ordered slides from presentation.xml
  const sldIdEls = getElementsByLocalName(presDoc, 'sldId');
  const slideTargets: { slideId: string; path: string }[] = [];

  for (const sldIdEl of sldIdEls) {
    const slideId = sldIdEl.getAttribute('id') || '';
    const rId = getAttrByLocalName(sldIdEl, 'id');
    // Note: sldId has both `id="256"` and `r:id="rId2"`. Find the relationship id attribute carefully.
    let relId = '';
    for (let i = 0; i < sldIdEl.attributes.length; i++) {
      const a = sldIdEl.attributes[i];
      if (a.localName === 'id' && (a.prefix === 'r' || a.namespaceURI?.includes('relationships'))) {
        relId = a.value;
        break;
      }
    }
    if (!relId && rId && presRels.has(rId)) relId = rId;
    if (relId && presRels.has(relId)) {
      const resolved = resolveZipPath('ppt/presentation.xml', presRels.get(relId)!.target);
      if (zip.file(resolved)) {
        slideTargets.push({ slideId, path: resolved });
      }
    }
  }

  // Fallback if sldIdLst had unusual namespace bindings
  if (slideTargets.length === 0) {
    const fallbackSlides = entryNames
      .filter((p) => /^ppt\/slides\/slide\d+\.xml$/i.test(p))
      .sort((a, b) => {
        const na = Number(a.match(/slide(\d+)\.xml/i)?.[1] || '0');
        const nb = Number(b.match(/slide(\d+)\.xml/i)?.[1] || '0');
        return na - nb;
      });
    fallbackSlides.forEach((p, idx) => {
      slideTargets.push({ slideId: String(256 + idx), path: p });
    });
  }

  // Media tracking map
  const mediaMap = new Map<string, MultimediaAuditEntry>();
  for (const entry of entries) {
    if (
      entry.path.startsWith('ppt/media/') ||
      entry.path.startsWith('ppt/charts/') ||
      entry.path.startsWith('ppt/embeddings/') ||
      entry.path.startsWith('ppt/diagrams/data')
    ) {
      const kind = classifyMediaKind(entry.path);
      const format = getExtension(entry.path).toUpperCase();
      let compatibilityNote: string | undefined;
      if (format === 'EMF' || format === 'WMF') {
        compatibilityNote = 'Formato vetorial legado Windows (EMF/WMF); pode variar em macOS/Web.';
      } else if (format === 'WMV' || format === 'AVI') {
        compatibilityNote = 'Codec de vídeo legado; preservado intacto, recomenda-se verificar reprodução em macOS.';
      }
      mediaMap.set(entry.path, {
        path: entry.path,
        fileName: entry.path.split('/').pop() || entry.path,
        kind,
        format,
        size: entry.size,
        isExternal: false,
        slidesUsed: [],
        compatibilityNote,
      });
    }
  }

  const slides: SlideAuditEntry[] = [];
  const hyperlinks: HyperlinkAuditEntry[] = [];
  let hyperlinkCounter = 0;

  let totalShapes = 0;
  let totalImages = 0;
  let totalTables = 0;
  let totalCharts = 0;
  let totalSmartArt = 0;
  let totalGroups = 0;
  let totalAnimations = 0;
  let slidesWithTransitions = 0;
  let slidesWithBackground = 0;
  let slidesWithEffects = 0;

  const brokenRelsFound: { slideNumber: number; relId: string }[] = [];

  for (let i = 0; i < slideTargets.length; i++) {
    const slideNumber = i + 1;
    const { slideId, path: slidePath } = slideTargets[i];
    const relsPath = getRelsPathForPart(slidePath);
    const slideBytes = await zip.file(slidePath)!.async('uint8array');
    const slideHash = await sha256Hex(slideBytes);
    const slideXml = new TextDecoder('utf-8').decode(slideBytes);
    const { doc: sDoc } = parseXmlSafe(slideXml);
    const relsXml = await readText(relsPath);
    const sRels = parseRelationships(relsXml);

    const fontsOnSlide = new Set<string>();
    const mediaOnSlide = new Set<string>();
    const textRuns: string[] = [];
    let title = '';
    let hidden = false;
    let shapeCount = 0;
    let imageCount = 0;
    let tableCount = 0;
    let chartCount = 0;
    let smartArtCount = 0;
    let groupCount = 0;
    let connectorCount = 0;
    let hasBackground = false;
    let hasShadows = false;
    let hasTransparency = false;
    let hasHeaderFooter = false;
    let slideHyperlinkCount = 0;

    const transition: SlideTransitionInfo = { present: false, type: 'Nenhuma' };
    const animations: SlideAnimationInfo = { present: false, nodeCount: 0, effectTypes: [] };

    if (sDoc) {
      scanXmlForExtensions(sDoc);
      const sRoot = sDoc.documentElement;
      hidden = sRoot?.getAttribute('show') === '0' || sRoot?.getAttribute('show') === 'false';

      // Shapes & titles
      const spEls = getElementsByLocalName(sDoc, 'sp');
      shapeCount = spEls.length;
      for (const sp of spEls) {
        const ph = getElementsByLocalName(sp, 'ph')[0];
        const phType = ph?.getAttribute('type') || '';
        if (['title', 'ctrTitle'].includes(phType) && !title) {
          const t = extractTextFromElement(sp);
          if (t) title = t;
        }
        if (['ftr', 'dt', 'sldNum', 'hdr'].includes(phType)) {
          hasHeaderFooter = true;
        }
      }

      // Extract text runs & typography properties
      const rEls = getElementsByLocalName(sDoc, 'r');
      for (const r of rEls) {
        const tEl = getElementsByLocalName(r, 't')[0];
        const txt = tEl?.textContent || '';
        if (txt.trim()) {
          textRuns.push(txt);
          if (!title && txt.trim().length > 2) {
            title = txt.trim().slice(0, 80);
          }
        }
        const rPr = getElementsByLocalName(r, 'rPr')[0];
        if (rPr) {
          const isBold = rPr.getAttribute('b') === '1' || rPr.getAttribute('b') === 'true';
          const isItalic = rPr.getAttribute('i') === '1' || rPr.getAttribute('i') === 'true';
          const uAttr = rPr.getAttribute('u');
          const isUnderline = Boolean(uAttr && uAttr !== 'none');
          const szAttr = rPr.getAttribute('sz');
          const szPt = szAttr ? Math.round((Number(szAttr) / 100) * 10) / 10 : 0;

          let explicitFontFound = false;
          for (const tag of ['latin', 'ea', 'cs', 'sym']) {
            const fNode = getElementsByLocalName(rPr, tag)[0];
            if (fNode) {
              const rawTf = fNode.getAttribute('typeface') || '';
              const tf = resolveFontName(rawTf);
              if (tf) {
                explicitFontFound = true;
                fontsOnSlide.add(tf);
                const rec = ensureFont(tf);
                if (rec) {
                  rec.slides.add(slideNumber);
                  if (isBold) rec.hasBold = true;
                  if (isItalic) rec.hasItalic = true;
                  if (isUnderline) rec.hasUnderline = true;
                  if (szPt > 0) rec.sizesPt.add(szPt);
                }
              }
            }
          }
          if (!explicitFontFound && themeMinorLatin) {
            fontsOnSlide.add(themeMinorLatin);
            const rec = ensureFont(themeMinorLatin);
            if (rec) {
              rec.slides.add(slideNumber);
              if (isBold) rec.hasBold = true;
              if (isItalic) rec.hasItalic = true;
              if (isUnderline) rec.hasUnderline = true;
              if (szPt > 0) rec.sizesPt.add(szPt);
            }
          }
        }
      }

      // Also scan any remaining font declarations on slide (defRPr, endParaRPr)
      for (const tag of ['latin', 'ea', 'cs', 'sym']) {
        const nodes = getElementsByLocalName(sDoc, tag);
        for (const n of nodes) {
          const tf = resolveFontName(n.getAttribute('typeface') || '');
          if (tf) {
            fontsOnSlide.add(tf);
            const rec = ensureFont(tf);
            if (rec) rec.slides.add(slideNumber);
          }
        }
      }

      // Visual elements
      imageCount = getElementsByLocalName(sDoc, 'pic').length;
      tableCount = getElementsByLocalName(sDoc, 'tbl').length;
      chartCount = getElementsByLocalName(sDoc, 'chart').length;
      smartArtCount = getElementsByLocalName(sDoc, 'relIds').length;
      groupCount = getElementsByLocalName(sDoc, 'grpSp').length;
      connectorCount = getElementsByLocalName(sDoc, 'cxnSp').length;
      hasBackground = getElementsByLocalName(sDoc, 'bg').length > 0;
      hasShadows =
        getElementsByLocalName(sDoc, 'outerShdw').length > 0 ||
        getElementsByLocalName(sDoc, 'innerShdw').length > 0 ||
        getElementsByLocalName(sDoc, 'effectLst').length > 0;
      hasTransparency = getElementsByLocalName(sDoc, 'alpha').length > 0;

      // Transitions
      const transEls = getElementsByLocalName(sDoc, 'transition');
      if (transEls.length > 0) {
        const tEl = transEls[0];
        transition.present = true;
        const childTags: string[] = [];
        for (let c = 0; c < tEl.children.length; c++) {
          const child = tEl.children[c];
          if (child.localName !== 'sndAc') {
            childTags.push(child.localName);
          }
        }
        const rawType = childTags[0] || 'fade';
        transition.type = TRANSITION_LABELS[rawType] || rawType;
        const durAttr = getAttrByLocalName(tEl, 'dur');
        if (durAttr && !isNaN(Number(durAttr))) {
          transition.durationMs = Number(durAttr);
        }
        const advClick = tEl.getAttribute('advClick');
        transition.advanceOnClick = advClick !== '0' && advClick !== 'false';
        const advTm = tEl.getAttribute('advTm');
        if (advTm && !isNaN(Number(advTm))) {
          transition.advanceAfterMs = Number(advTm);
        }
      }

      // Animations (<p:timing>)
      const timingEls = getElementsByLocalName(sDoc, 'timing');
      if (timingEls.length > 0) {
        const timing = timingEls[0];
        const animTags: [string, string][] = [
          ['animEffect', 'Efeito de Entrada/Saída'],
          ['animMotion', 'Trajetória de Movimento'],
          ['animScale', 'Escala / Zoom'],
          ['animRot', 'Rotação'],
          ['animClr', 'Ênfase de Cor'],
          ['anim', 'Propriedade Animada'],
          ['set', 'Alteração de Estado'],
          ['audio', 'Reprodução de Áudio'],
          ['video', 'Reprodução de Vídeo'],
        ];
        let count = 0;
        const types = new Set<string>();
        for (const [tag, label] of animTags) {
          const found = getElementsByLocalName(timing, tag);
          if (found.length > 0) {
            count += found.length;
            types.add(label);
          }
        }
        if (count > 0) {
          animations.present = true;
          animations.nodeCount = count;
          animations.effectTypes = Array.from(types);
        }
      }

      // Hyperlinks (<a:hlinkClick> & <a:hlinkHover>)
      const hlinkEls = [
        ...getElementsByLocalName(sDoc, 'hlinkClick'),
        ...getElementsByLocalName(sDoc, 'hlinkHover'),
      ];

      for (const hl of hlinkEls) {
        hyperlinkCounter++;
        slideHyperlinkCount++;
        const relId = getAttrByLocalName(hl, 'id') || '';
        const action = hl.getAttribute('action') || '';
        const tooltip = hl.getAttribute('tooltip') || '';
        const rel = relId ? sRels.get(relId) : undefined;
        const isValidRel = relId ? Boolean(rel) : Boolean(action);

        if (relId && !rel) {
          brokenRelsFound.push({ slideNumber, relId });
        }

        const target = rel?.target || action || '(Ação sem destino explícito)';
        const targetMode = rel?.targetMode || (action ? 'Action' : 'Internal');

        // Determine element type and label
        let elementType: HyperlinkAuditEntry['elementType'] = 'interactive';
        let elementLabel = 'Elemento interativo';

        const rParent = findAncestorByLocalName(hl, ['r']);
        const picParent = findAncestorByLocalName(hl, ['pic']);
        const spParent = findAncestorByLocalName(hl, ['sp']);
        const grpParent = findAncestorByLocalName(hl, ['grpSp']);

        if (rParent) {
          elementType = 'text';
          const textVal = extractTextFromElement(rParent);
          elementLabel = textVal || 'Texto sem conteúdo visível';
        } else if (picParent) {
          elementType = 'image';
          const cNvPr = getElementsByLocalName(picParent, 'cNvPr')[0];
          elementLabel = cNvPr?.getAttribute('descr') || cNvPr?.getAttribute('name') || 'Imagem';
        } else if (spParent) {
          elementType = 'shape';
          const cNvPr = getElementsByLocalName(spParent, 'cNvPr')[0];
          const spText = extractTextFromElement(spParent);
          const spName = cNvPr?.getAttribute('name') || 'Forma';
          elementLabel = spText ? `${spName} ("${spText.slice(0, 40)}")` : spName;
        } else if (grpParent) {
          elementType = 'interactive';
          const cNvPr = getElementsByLocalName(grpParent, 'cNvPr')[0];
          elementLabel = cNvPr?.getAttribute('name') || 'Grupo interativo';
        }

        // Determine link kind
        let linkKind: HyperlinkAuditEntry['linkKind'] = 'action';
        let warning: string | undefined;

        if (/^mailto:/i.test(target)) {
          linkKind = 'email';
        } else if (/^https?:\/\//i.test(target)) {
          linkKind = 'web';
        } else if (
          action.includes('hlinksldjump') ||
          action.includes('hlinkshowjump') ||
          (rel && rel.type.includes('/slide') && targetMode !== 'External')
        ) {
          linkKind = 'slide';
        } else if (
          action.includes('hlinkfile') ||
          targetMode === 'External' ||
          /^file:\/\//i.test(target)
        ) {
          linkKind = 'external_file';
          warning =
            'Hiperligação aponta para um ficheiro externo/local; preservada intacta, mas pode exigir que o ficheiro exista no destino.';
        } else if (!isValidRel) {
          warning = `Relação interna (${relId}) ausente em ${relsPath}.`;
        }

        hyperlinks.push({
          id: `hl-${hyperlinkCounter}`,
          slideNumber,
          slidePath,
          relId,
          target,
          action,
          tooltip,
          elementType,
          elementLabel,
          linkKind,
          targetMode,
          isValidRel,
          warning,
        });
      }
    }

    // Resolve media & notes from slide relationships
    let hasNotes = false;
    let notesPath: string | undefined;
    let notesText = '';

    for (const rel of sRels.values()) {
      const lowerType = rel.type.toLowerCase();
      if (lowerType.includes('notesslide')) {
        const np = resolveZipPath(slidePath, rel.target);
        const nXml = await readText(np);
        if (nXml) {
          hasNotes = true;
          notesPath = np;
          const { doc: nDoc } = parseXmlSafe(nXml);
          if (nDoc) {
            // Extract notes body excluding slide number placeholder
            const spNotes = getElementsByLocalName(nDoc, 'sp');
            const parts: string[] = [];
            for (const sp of spNotes) {
              const ph = getElementsByLocalName(sp, 'ph')[0];
              if (ph?.getAttribute('type') === 'sldNum') continue;
              const txt = extractTextFromElement(sp);
              if (txt) parts.push(txt);
            }
            notesText = parts.join('\n').trim();
          }
        }
      } else if (
        lowerType.includes('image') ||
        lowerType.includes('audio') ||
        lowerType.includes('video') ||
        lowerType.includes('media') ||
        lowerType.includes('chart') ||
        lowerType.includes('diagramdata') ||
        lowerType.includes('oleobject')
      ) {
        if (rel.targetMode === 'External') {
          const extKey = `external:${rel.target}`;
          mediaOnSlide.add(rel.target);
          if (!mediaMap.has(extKey)) {
            mediaMap.set(extKey, {
              path: rel.target,
              fileName: rel.target.split(/[\\/]/).pop() || rel.target,
              kind: 'external_media',
              format: getExtension(rel.target).toUpperCase(),
              size: 0,
              isExternal: true,
              slidesUsed: [slideNumber],
              compatibilityNote:
                'Recurso multimédia externo (ligado, não incorporado); referência preservada, mas requer acesso ao ficheiro original.',
            });
          } else {
            const existing = mediaMap.get(extKey)!;
            if (!existing.slidesUsed.includes(slideNumber)) {
              existing.slidesUsed.push(slideNumber);
            }
          }
        } else {
          const resolved = resolveZipPath(slidePath, rel.target);
          mediaOnSlide.add(resolved);
          const existing = mediaMap.get(resolved);
          if (existing && !existing.slidesUsed.includes(slideNumber)) {
            existing.slidesUsed.push(slideNumber);
          }
        }
      }
    }

    totalShapes += shapeCount;
    totalImages += imageCount;
    totalTables += tableCount;
    totalCharts += chartCount;
    totalSmartArt += smartArtCount;
    totalGroups += groupCount;
    totalAnimations += animations.nodeCount;
    if (transition.present) slidesWithTransitions++;
    if (hasBackground) slidesWithBackground++;
    if (hasShadows || hasTransparency) slidesWithEffects++;

    slides.push({
      index: i,
      slideNumber,
      path: slidePath,
      relsPath,
      slideId,
      title: title || `Diapositivo ${slideNumber}`,
      hidden,
      textRuns,
      fullTextPreview: textRuns.join(' ').replace(/\s+/g, ' ').trim().slice(0, 240),
      shapeCount,
      imageCount,
      tableCount,
      chartCount,
      smartArtCount,
      groupCount,
      connectorCount,
      hasBackground,
      hasShadows,
      hasTransparency,
      hasHeaderFooter,
      hasNotes,
      notesPath,
      notesText,
      transition,
      animations,
      hyperlinkCount: slideHyperlinkCount,
      fontsUsed: Array.from(fontsOnSlide).sort(),
      mediaRefs: Array.from(mediaOnSlide),
      xmlByteSize: slideBytes.byteLength,
      sha256: slideHash,
    });
  }

  // Finalize Font Audit entries
  const fonts: FontAuditEntry[] = [];
  for (const [name, data] of fontAccum.entries()) {
    const emb = embeddedFontMap.get(name);
    const isEmbedded = Boolean(emb);
    const isStandard = STANDARD_OFFICE_FONTS.has(name.toLowerCase());
    const status: FontAuditEntry['status'] = isEmbedded
      ? 'embedded'
      : isStandard
      ? 'standard'
      : 'external_unembedded';

    let advisory = '';
    if (status === 'embedded') {
      advisory =
        'Fonte incorporada diretamente no ficheiro PPTX. Os ficheiros de fonte e referências XML são preservados integralmente.';
    } else if (status === 'standard') {
      advisory =
        'Fonte de sistema amplamente disponível. A referência original, pesos e estilos são preservados sem alteração.';
    } else {
      advisory =
        'Fonte externa/personalizada não incorporada no ficheiro. A referência original será preservada sem substituição; poderá ser necessário instalar esta fonte no computador de destino para renderização exata.';
    }

    fonts.push({
      name,
      embedded: isEmbedded,
      embeddedVariants: emb ? Array.from(emb.variants) : [],
      embeddedFiles: emb ? Array.from(emb.files) : [],
      isThemeMajor: name === themeMajorLatin,
      isThemeMinor: name === themeMinorLatin,
      isStandardSystem: isStandard,
      slidesUsed: Array.from(data.slides).sort((a, b) => a - b),
      usedInMastersOrLayouts: data.usedInMastersOrLayouts,
      hasBold: data.hasBold,
      hasItalic: data.hasItalic,
      hasUnderline: data.hasUnderline,
      sizesPt: Array.from(data.sizesPt).sort((a, b) => a - b),
      status,
      advisory,
    });
  }

  // Sort fonts: external_unembedded first (most important to inspect), then embedded, then standard
  const fontOrder: Record<FontAuditEntry['status'], number> = {
    external_unembedded: 0,
    embedded: 1,
    standard: 2,
  };
  fonts.sort((a, b) => fontOrder[a.status] - fontOrder[b.status] || a.name.localeCompare(b.name));

  const multimedia = Array.from(mediaMap.values()).sort((a, b) => a.path.localeCompare(b.path));

  // Build Compatibility Issues report
  const compatibilityIssues: CompatibilityIssue[] = [];

  const unembeddedCustomFonts = fonts.filter((f) => f.status === 'external_unembedded');
  for (const uf of unembeddedCustomFonts) {
    compatibilityIssues.push({
      id: `font-unembedded-${uf.name}`,
      severity: 'warning',
      category: 'fonts',
      title: `Fonte personalizada não incorporada: ${uf.name}`,
      description: `A fonte "${uf.name}" é referenciada na apresentação mas não está incorporada em ppt/fonts/. Nunca será substituída automaticamente por Arial ou Calibri.`,
      affectedSlides: uf.slidesUsed,
      recommendation:
        'Instale esta fonte no computador de destino ou ative "Incorporar tipos de letra no ficheiro" no Microsoft PowerPoint se precisar de portabilidade total.',
      preservationStatus: 'Referência tipográfica original preservada integralmente',
    });
  }

  const externalFileLinks = hyperlinks.filter((h) => h.linkKind === 'external_file');
  if (externalFileLinks.length > 0) {
    compatibilityIssues.push({
      id: 'hlink-external-files',
      severity: 'warning',
      category: 'hyperlinks',
      title: `${externalFileLinks.length} hiperligação(ões) para ficheiros externos/locais`,
      description:
        'Existem ligações que apontam para ficheiros fora do pacote PPTX. Ao abrir noutro computador, o caminho relativo ou local pode não estar acessível.',
      affectedSlides: Array.from(new Set(externalFileLinks.map((h) => h.slideNumber))).sort(
        (a, b) => a - b
      ),
      recommendation:
        'Mantenha os ficheiros externos na mesma estrutura de pastas junto à apresentação final.',
      preservationStatus: 'Endereços e relações originais preservados intactos',
    });
  }

  if (brokenRelsFound.length > 0) {
    compatibilityIssues.push({
      id: 'hlink-broken-rels',
      severity: 'critical',
      category: 'hyperlinks',
      title: `${brokenRelsFound.length} referência(s) de relação em falta no ficheiro original`,
      description:
        'O ficheiro original contém identificadores r:id de hiperligação sem entrada correspondente no respetivo ficheiro .rels.',
      affectedSlides: Array.from(new Set(brokenRelsFound.map((b) => b.slideNumber))).sort(
        (a, b) => a - b
      ),
      recommendation:
        'Reveja estas hiperligações no PowerPoint original. A aplicação manterá os nós XML intactos sem eliminação destrutiva.',
      preservationStatus: 'Estrutura XML mantida sem remoção',
    });
  }

  const externalMedia = multimedia.filter((m) => m.isExternal);
  if (externalMedia.length > 0) {
    compatibilityIssues.push({
      id: 'media-external',
      severity: 'warning',
      category: 'multimedia',
      title: `${externalMedia.length} recurso(s) multimédia ligados externamente`,
      description:
        'Os ficheiros de áudio ou vídeo ligados externamente não residem dentro do pacote .pptx.',
      affectedSlides: Array.from(new Set(externalMedia.flatMap((m) => m.slidesUsed))).sort(
        (a, b) => a - b
      ),
      recommendation:
        'Transporte os ficheiros multimédia originais juntamente com a apresentação ou incorpore-os no PowerPoint.',
      preservationStatus: 'Relações externas preservadas integralmente',
    });
  }

  const legacyMedia = multimedia.filter((m) => m.compatibilityNote && !m.isExternal);
  if (legacyMedia.length > 0) {
    compatibilityIssues.push({
      id: 'media-legacy',
      severity: 'info',
      category: 'multimedia',
      title: `${legacyMedia.length} formato(s) multimédia legado(s) detetado(s)`,
      description: legacyMedia.map((m) => `${m.fileName} (${m.format})`).join(', '),
      affectedSlides: Array.from(new Set(legacyMedia.flatMap((m) => m.slidesUsed))).sort(
        (a, b) => a - b
      ),
      recommendation:
        'Nenhum ficheiro é reconvertido ou recomprimido. Verifique a reprodução caso apresente em macOS ou PowerPoint Online.',
      preservationStatus: 'Ficheiros binários preservados byte-a-byte',
    });
  }

  if (unknownExtensionSet.size > 0) {
    compatibilityIssues.push({
      id: 'xml-extensions',
      severity: 'info',
      category: 'xml',
      title: 'Extensões avançadas OOXML / PowerPoint detetadas',
      description: `Encontrados blocos de extensão específicos do Microsoft PowerPoint (${Array.from(
        unknownExtensionSet
      )
        .slice(0, 3)
        .join(', ')}).`,
      affectedSlides: [],
      recommendation:
        'O motor não destrutivo mantém todas as tags XML desconhecidas e blocos mc:AlternateContent sem simplificação.',
      preservationStatus: 'Elementos XML desconhecidos preservados',
    });
  }

  const visualSummary: VisualSummary = {
    totalShapes,
    totalImages,
    totalTables,
    totalCharts,
    totalSmartArt,
    totalGroups,
    totalAnimations,
    slidesWithTransitions,
    slidesWithBackground,
    slidesWithEffects,
    themeName,
    masterCount: masterPaths.length,
    layoutCount: layoutPaths.length,
  };

  return {
    fileName,
    fileSize: buffer.byteLength,
    sha256: fileSha256,
    analyzedAt: new Date().toISOString(),
    slideCount: slides.length,
    slideWidthCm,
    slideHeightCm,
    aspectRatioLabel: detectAspectRatio(slideWidthCm, slideHeightCm),
    coreProperties,
    structure: {
      totalEntries: entries.length,
      xmlCount: entries.filter((e) => e.isXml).length,
      relsCount: entries.filter((e) => e.isRels).length,
      mediaCount: entries.filter((e) => e.isMedia).length,
      fontFilesCount: embeddedFontFiles.length,
      hasContentTypes,
      hasRootRels,
      hasPresentationXml,
      unknownExtensions: Array.from(unknownExtensionSet),
      entries,
    },
    fonts,
    embeddedFontFlags: {
      embedTrueTypeFonts,
      saveSubsetFonts,
      embeddedFontFiles,
    },
    hyperlinks,
    slides,
    multimedia,
    compatibilityIssues,
    visualSummary,
  };
}

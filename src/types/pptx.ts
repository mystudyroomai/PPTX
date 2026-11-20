export interface ZipEntryInfo {
  path: string;
  size: number;
  sha256: string;
  isXml: boolean;
  isRels: boolean;
  isMedia: boolean;
  isFont: boolean;
}

export interface CoreProperties {
  title: string;
  subject: string;
  creator: string;
  lastModifiedBy: string;
  created: string;
  modified: string;
  company: string;
  application: string;
}

export interface FontAuditEntry {
  name: string;
  embedded: boolean;
  embeddedVariants: string[];
  embeddedFiles: string[];
  isThemeMajor: boolean;
  isThemeMinor: boolean;
  isStandardSystem: boolean;
  slidesUsed: number[];
  usedInMastersOrLayouts: boolean;
  hasBold: boolean;
  hasItalic: boolean;
  hasUnderline: boolean;
  sizesPt: number[];
  status: 'embedded' | 'standard' | 'external_unembedded';
  advisory: string;
}

export interface HyperlinkAuditEntry {
  id: string;
  slideNumber: number;
  slidePath: string;
  relId: string;
  target: string;
  action: string;
  tooltip: string;
  elementType: 'text' | 'image' | 'shape' | 'interactive';
  elementLabel: string;
  linkKind: 'web' | 'slide' | 'external_file' | 'email' | 'action';
  targetMode: string;
  isValidRel: boolean;
  warning?: string;
}

export interface SlideTransitionInfo {
  present: boolean;
  type: string;
  durationMs?: number;
  advanceOnClick?: boolean;
  advanceAfterMs?: number;
}

export interface SlideAnimationInfo {
  present: boolean;
  nodeCount: number;
  effectTypes: string[];
}

export interface SlideAuditEntry {
  index: number;
  slideNumber: number;
  path: string;
  relsPath: string;
  slideId: string;
  title: string;
  hidden: boolean;
  textRuns: string[];
  fullTextPreview: string;
  shapeCount: number;
  imageCount: number;
  tableCount: number;
  chartCount: number;
  smartArtCount: number;
  groupCount: number;
  connectorCount: number;
  hasBackground: boolean;
  hasShadows: boolean;
  hasTransparency: boolean;
  hasHeaderFooter: boolean;
  hasNotes: boolean;
  notesPath?: string;
  notesText: string;
  transition: SlideTransitionInfo;
  animations: SlideAnimationInfo;
  hyperlinkCount: number;
  fontsUsed: string[];
  mediaRefs: string[];
  xmlByteSize: number;
  sha256: string;
}

export interface MultimediaAuditEntry {
  path: string;
  fileName: string;
  kind: 'image' | 'audio' | 'video' | 'external_media' | 'chart' | 'ole_embedded' | 'smartart';
  format: string;
  size: number;
  isExternal: boolean;
  slidesUsed: number[];
  compatibilityNote?: string;
}

export interface CompatibilityIssue {
  id: string;
  severity: 'info' | 'warning' | 'critical';
  category: 'fonts' | 'hyperlinks' | 'multimedia' | 'xml' | 'visual';
  title: string;
  description: string;
  affectedSlides: number[];
  recommendation: string;
  preservationStatus: string;
}

export interface VisualSummary {
  totalShapes: number;
  totalImages: number;
  totalTables: number;
  totalCharts: number;
  totalSmartArt: number;
  totalGroups: number;
  totalAnimations: number;
  slidesWithTransitions: number;
  slidesWithBackground: number;
  slidesWithEffects: number;
  themeName: string;
  masterCount: number;
  layoutCount: number;
}

export interface PptxAnalysis {
  fileName: string;
  fileSize: number;
  sha256: string;
  analyzedAt: string;
  slideCount: number;
  slideWidthCm: number;
  slideHeightCm: number;
  aspectRatioLabel: string;
  coreProperties: CoreProperties;
  structure: {
    totalEntries: number;
    xmlCount: number;
    relsCount: number;
    mediaCount: number;
    fontFilesCount: number;
    hasContentTypes: boolean;
    hasRootRels: boolean;
    hasPresentationXml: boolean;
    unknownExtensions: string[];
    entries: ZipEntryInfo[];
  };
  fonts: FontAuditEntry[];
  embeddedFontFlags: {
    embedTrueTypeFonts: boolean;
    saveSubsetFonts: boolean;
    embeddedFontFiles: { path: string; size: number }[];
  };
  hyperlinks: HyperlinkAuditEntry[];
  slides: SlideAuditEntry[];
  multimedia: MultimediaAuditEntry[];
  compatibilityIssues: CompatibilityIssue[];
  visualSummary: VisualSummary;
}

export interface TextReplacementRule {
  id: string;
  find: string;
  replace: string;
  caseSensitive: boolean;
}

export interface ProcessingOptions {
  textReplacements: TextReplacementRule[];
  protectHyperlinkedText: boolean;
  metadataUpdates: {
    enabled: boolean;
    title: string;
    subject: string;
    creator: string;
    anonymizeAuthor: boolean;
  };
  slideVisibilityOverrides: Record<number, boolean>;
  cleanSpeakerNotes: boolean;
}

export interface VerificationCheck {
  id: string;
  title: string;
  category: 'slides' | 'hyperlinks' | 'fonts' | 'multimedia' | 'animations' | 'structure' | 'visual';
  status: 'pass' | 'limitation' | 'fail';
  originalSummary: string;
  processedSummary: string;
  details: string;
  limitationNote?: string;
}

export interface VerificationReport {
  overallStatus: 'verified' | 'verified_with_limitations' | 'blocked';
  canExport: boolean;
  timestamp: string;
  checks: VerificationCheck[];
  limitations: string[];
  blockingReasons: string[];
}

export interface FileDiffEntry {
  path: string;
  status: 'unchanged' | 'modified' | 'added' | 'removed';
  originalSize: number;
  processedSize: number;
  originalHash: string;
  processedHash: string;
  reason: string;
}

export interface SlideDiffEntry {
  slideNumber: number;
  title: string;
  xmlUnchanged: boolean;
  textChanges: { before: string; after: string }[];
  visibilityChanged: boolean;
  notesChanged: boolean;
  originalHidden: boolean;
  processedHidden: boolean;
}

export interface ComparisonSummary {
  fileDiffs: FileDiffEntry[];
  slideDiffs: SlideDiffEntry[];
  unchangedEntriesCount: number;
  modifiedEntriesCount: number;
  addedEntriesCount: number;
  removedEntriesCount: number;
  totalEntriesCount: number;
}

export interface ProcessingResult {
  success: boolean;
  aborted: boolean;
  abortReason?: string;
  processedBuffer?: Uint8Array;
  processedFileName: string;
  processedAnalysis?: PptxAnalysis;
  verification: VerificationReport;
  comparison?: ComparisonSummary;
  operationsApplied: string[];
  skippedProtectedRuns: { slideNumber: number; text: string; reason: string }[];
}

export type WorkspaceSection =
  | 'overview'
  | 'fonts'
  | 'hyperlinks'
  | 'visual_media'
  | 'compatibility'
  | 'process_export'
  | 'comparison';

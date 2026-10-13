export interface PptxRun {
  id: string;
  text: string;
  isBold?: boolean;
  isItalic?: boolean;
  fontSizePt?: number;
  colorHex?: string;
  fontFamily?: string;
  hyperlink?: {
    url: string;
    tooltip?: string;
    rId?: string;
  };
}

export interface PptxParagraph {
  id: string;
  slideNumber: number;
  shapeId: string;
  paragraphIndex: number;
  shapeName: string;
  shapeRole: 'title' | 'subtitle' | 'body' | 'table-cell' | 'shape' | 'footer' | 'header' | 'quote' | 'list-item';
  originalText: string;
  currentText: string;
  runs: PptxRun[];
  isModified: boolean;
  alignment?: 'left' | 'center' | 'right' | 'justify';
  bulletLevel?: number;
  colorHex?: string;
  hyperlink?: {
    url: string;
    tooltip?: string;
    rId?: string;
  };
  aiNotes?: string;
  lengthFitWarning?: string;
  verificationStatus?: 'ok' | 'warning' | 'suggestion';
  verificationIssue?: string;
  verificationSuggestion?: string;
  alternatives?: string[];
}

export interface PptxShape {
  id: string;
  name: string;
  role: 'title' | 'subtitle' | 'body' | 'table-cell' | 'shape' | 'footer' | 'header' | 'quote' | 'list-item';
  bounds: {
    x: number; // percentage 0-100
    y: number; // percentage 0-100
    width: number; // percentage 0-100
    height: number; // percentage 0-100
  };
  backgroundColor?: string;
  borderColor?: string;
  paragraphs: PptxParagraph[];
  isTable?: boolean;
  tableDimensions?: { rows: number; cols: number };
}

export interface PptxImage {
  id: string;
  name: string;
  bounds: {
    x: number;
    y: number;
    width: number;
    height: number;
  };
  dataUrl: string;
}

export interface PptxSlide {
  slideNumber: number;
  title: string;
  xmlPath: string;
  shapes: PptxShape[];
  images?: PptxImage[];
  speakerNotes?: {
    text: string;
    originalText: string;
    isModified: boolean;
    xmlPath?: string;
  };
  widthEmu: number;
  heightEmu: number;
  aspectRatio: number; // e.g. 16/9
  backgroundColor?: string;
  backgroundImageDataUrl?: string;
  renderedImageDataUrl?: string;
  originalRenderedImageDataUrl?: string;
}

export interface PptxDocument {
  docType?: 'pptx' | 'docx';
  fileName: string;
  fileSize: number;
  slides: PptxSlide[];
  rawZip: any; // JSZip instance
  originalBuffer?: ArrayBuffer;
  totalWords: number;
  totalParagraphs: number;
  modifiedCount: number;
}

export interface TranslationItem {
  id: string;
  slideNumber: number;
  role: string;
  originalText: string;
  hyperlinkUrl?: string;
}

export interface VerificationItem {
  id: string;
  slideNumber: number;
  originalText: string;
  currentText: string;
}

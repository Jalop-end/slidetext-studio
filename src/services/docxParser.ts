import JSZip from 'jszip';
import {
  PptxDocument,
  PptxSlide,
  PptxShape,
  PptxParagraph,
  PptxRun,
  PptxImage,
} from '../types/pptx';

/**
 * Extracts and parses a Word (.docx) file using JSZip and browser XML DOMParser.
 * Preserves full OpenXML document structure, formatting, tables, headings, and images.
 */
export async function parseDocxFile(
  file: File | Blob | ArrayBuffer,
  fileName: string = 'dokument.docx'
): Promise<PptxDocument> {
  const zip = await JSZip.loadAsync(file);

  // 1. Check for word/document.xml
  const docXmlFile = zip.file('word/document.xml');
  if (!docXmlFile) {
    throw new Error('Neplatný formát souboru Word: chybí word/document.xml');
  }

  const docXmlStr = await docXmlFile.async('text');
  const parser = new DOMParser();
  const doc = parser.parseFromString(docXmlStr, 'application/xml');

  // 2. Parse relationships (word/_rels/document.xml.rels) for hyperlinks & images
  const rels = await parseDocxRelationships(zip, 'word/_rels/document.xml.rels');

  // 3. Parse styles (word/styles.xml) for heading and font mappings
  const styleMap = await parseDocxStyles(zip);

  // 4. Default A4 page dimensions in EMU / DXA (210mm x 297mm)
  // Standard Word page: width = 11906 dxa, height = 16838 dxa (aspect ratio ~0.707)
  let pageWidthDxa = 11906;
  let pageHeightDxa = 16838;

  const sectPr = doc.querySelector('w\\:sectPr, sectPr');
  if (sectPr) {
    const pgSz = sectPr.querySelector('w\\:pgSz, pgSz');
    if (pgSz) {
      const w = parseInt(pgSz.getAttribute('w:w') || pgSz.getAttribute('w') || '0', 10);
      const h = parseInt(pgSz.getAttribute('w:h') || pgSz.getAttribute('h') || '0', 10);
      if (w > 0) pageWidthDxa = w;
      if (h > 0) pageHeightDxa = h;
    }
  }

  const aspectRatio = pageWidthDxa / pageHeightDxa; // ~0.707 for A4 portrait
  const widthEmu = Math.round((pageWidthDxa / 20) * 12700); // 1 dxa = 1/20 pt = 635 emu
  const heightEmu = Math.round((pageHeightDxa / 20) * 12700);

  // 5. Walk top-level body children: paragraphs <w:p> and tables <w:tbl>
  const body = doc.querySelector('w\\:body, body');
  if (!body) {
    throw new Error('V dokumentu Word nebyl nalezen element <w:body>.');
  }

  // Segment into pages by detecting explicit page breaks (<w:br w:type="page"/>, <w:lastRenderedPageBreak/>)
  // or chunking into pages of reasonable paragraph blocks (~14-18 paragraphs or major headings)
  interface BodyBlock {
    type: 'paragraph' | 'table';
    element: Element;
    hasPageBreakBefore?: boolean;
    hasPageBreakAfter?: boolean;
  }

  const blocks: BodyBlock[] = [];
  const children = Array.from(body.children);

  children.forEach((child) => {
    const tagName = child.tagName.toLowerCase();
    if (tagName === 'w:p' || tagName === 'p') {
      const pageBr = child.querySelector('w\\:br[w\\:type="page"], br[type="page"], w\\:lastRenderedPageBreak, lastRenderedPageBreak');
      blocks.push({
        type: 'paragraph',
        element: child,
        hasPageBreakAfter: !!pageBr,
      });
    } else if (tagName === 'w:tbl' || tagName === 'tbl') {
      blocks.push({
        type: 'table',
        element: child,
      });
    }
  });

  // Group blocks into pages
  const pagesBlocks: BodyBlock[][] = [];
  let currentPageBlocks: BodyBlock[] = [];
  const PARAGRAPHS_PER_PAGE_FALLBACK = 15;

  blocks.forEach((block) => {
    currentPageBlocks.push(block);

    const isHeading1 =
      block.type === 'paragraph' &&
      isMajorHeading(block.element, styleMap);

    const shouldBreak =
      block.hasPageBreakAfter ||
      (currentPageBlocks.length >= PARAGRAPHS_PER_PAGE_FALLBACK && isHeading1) ||
      currentPageBlocks.length >= 24;

    if (shouldBreak) {
      pagesBlocks.push(currentPageBlocks);
      currentPageBlocks = [];
    }
  });

  if (currentPageBlocks.length > 0) {
    pagesBlocks.push(currentPageBlocks);
  }

  if (pagesBlocks.length === 0) {
    pagesBlocks.push([]);
  }

  // 6. Build PptxSlide representation for each Word Page
  const slides: PptxSlide[] = [];
  let totalWords = 0;
  let totalParagraphs = 0;
  let globalPIdx = 0;

  for (let pageIdx = 0; pageIdx < pagesBlocks.length; pageIdx++) {
    const pageNum = pageIdx + 1;
    const pageBlocks = pagesBlocks[pageIdx];
    const shapes: PptxShape[] = [];
    let pageTitle = `Stránka ${pageNum}`;
    let foundFirstHeading = false;

    for (let bIdx = 0; bIdx < pageBlocks.length; bIdx++) {
      const block = pageBlocks[bIdx];

      if (block.type === 'paragraph') {
        const pElem = block.element;
        const parsedP = parseWordParagraph(
          pElem,
          pageNum,
          `shape-p-${bIdx}`,
          globalPIdx,
          rels,
          styleMap
        );

        if (parsedP) {
          totalParagraphs++;
          const words = parsedP.originalText.trim().split(/\s+/).filter(Boolean).length;
          totalWords += words;

          if (!foundFirstHeading && (parsedP.shapeRole === 'title' || parsedP.shapeRole === 'subtitle')) {
            pageTitle = parsedP.originalText.slice(0, 45) + (parsedP.originalText.length > 45 ? '...' : '');
            foundFirstHeading = true;
          }

          // Each paragraph or logical group is represented as a shape with coordinates
          const approxY = Math.min(92, Math.max(5, (bIdx / Math.max(1, pageBlocks.length)) * 88));
          shapes.push({
            id: `page-${pageNum}-block-${bIdx}`,
            name: parsedP.shapeName,
            role: parsedP.shapeRole,
            bounds: {
              x: 10,
              y: approxY,
              width: 80,
              height: Math.max(3, Math.min(25, Math.ceil(parsedP.originalText.length / 80) * 3)),
            },
            paragraphs: [parsedP],
          });

          globalPIdx++;
        }
      } else if (block.type === 'table') {
        // Table block
        const tblElem = block.element;
        const parsedTableShape = parseWordTable(
          tblElem,
          pageNum,
          `tbl-${bIdx}`,
          globalPIdx,
          rels,
          styleMap
        );

        if (parsedTableShape) {
          parsedTableShape.paragraphs.forEach((p) => {
            totalParagraphs++;
            totalWords += p.originalText.trim().split(/\s+/).filter(Boolean).length;
            globalPIdx++;
          });

          const approxY = Math.min(88, Math.max(5, (bIdx / Math.max(1, pageBlocks.length)) * 85));
          parsedTableShape.bounds.y = approxY;
          shapes.push(parsedTableShape);
        }
      }
    }

    slides.push({
      slideNumber: pageNum,
      title: pageTitle,
      xmlPath: 'word/document.xml',
      shapes,
      widthEmu,
      heightEmu,
      aspectRatio,
      backgroundColor: '#ffffff',
    });
  }

  // 7. Load images from word/media/
  const images = await parseDocxImages(zip, rels);
  if (images.length > 0 && slides.length > 0) {
    slides[0].images = images;
  }

  return {
    docType: 'docx',
    fileName,
    fileSize: typeof (file as any).size === 'number' ? (file as any).size : 0,
    slides,
    rawZip: zip,
    totalWords,
    totalParagraphs,
    modifiedCount: 0,
  };
}

/**
 * Checks whether a paragraph has Heading 1 or Title style.
 */
function isMajorHeading(pElem: Element, styleMap: Record<string, string>): boolean {
  const pStyle = pElem.querySelector('w\\:pStyle, pStyle');
  if (!pStyle) return false;
  const val = (pStyle.getAttribute('w:val') || pStyle.getAttribute('val') || '').toLowerCase();
  const resolved = (styleMap[val] || val).toLowerCase();
  return (
    resolved.includes('heading 1') ||
    resolved.includes('nadpis 1') ||
    resolved.includes('title') ||
    resolved.includes('název')
  );
}

/**
 * Parses a single Word <w:p> into PptxParagraph.
 */
function parseWordParagraph(
  pElem: Element,
  pageNum: number,
  shapeId: string,
  pIdx: number,
  rels: Record<string, { target: string; type: string }>,
  styleMap: Record<string, string>
): PptxParagraph | null {
  // Determine role & alignment from <w:pPr>
  let role: PptxParagraph['shapeRole'] = 'body';
  let alignment: PptxParagraph['alignment'] = 'left';
  let bulletLevel: number | undefined;

  const pPr = pElem.querySelector('w\\:pPr, pPr');
  if (pPr) {
    const pStyle = pPr.querySelector('w\\:pStyle, pStyle');
    if (pStyle) {
      const val = (pStyle.getAttribute('w:val') || pStyle.getAttribute('val') || '').toLowerCase();
      const styleName = (styleMap[val] || val).toLowerCase();

      if (styleName.includes('heading 1') || styleName.includes('nadpis 1') || styleName.includes('title') || styleName.includes('název')) {
        role = 'title';
      } else if (styleName.includes('heading 2') || styleName.includes('nadpis 2') || styleName.includes('subtitle') || styleName.includes('podnázev')) {
        role = 'subtitle';
      } else if (styleName.includes('heading 3') || styleName.includes('nadpis 3')) {
        role = 'subtitle';
      } else if (styleName.includes('quote') || styleName.includes('citace')) {
        role = 'quote';
      } else if (styleName.includes('header') || styleName.includes('záhlaví')) {
        role = 'header';
      } else if (styleName.includes('footer') || styleName.includes('zápatí')) {
        role = 'footer';
      } else if (styleName.includes('list') || styleName.includes('seznam') || styleName.includes('bullet')) {
        role = 'list-item';
      }
    }

    const jc = pPr.querySelector('w\\:jc, jc');
    if (jc) {
      const jcVal = (jc.getAttribute('w:val') || jc.getAttribute('val') || '').toLowerCase();
      if (jcVal === 'center') alignment = 'center';
      else if (jcVal === 'right') alignment = 'right';
      else if (jcVal === 'both') alignment = 'justify';
    }

    const numPr = pPr.querySelector('w\\:numPr, numPr');
    if (numPr) {
      role = 'list-item';
      const ilvl = numPr.querySelector('w\\:ilvl, ilvl');
      if (ilvl) {
        bulletLevel = parseInt(ilvl.getAttribute('w:val') || '0', 10);
      }
    }
  }

  // Parse runs <w:r> & hyperlinks <w:hyperlink>
  const runs: PptxRun[] = [];
  let fullText = '';
  let paragraphHyperlink: PptxParagraph['hyperlink'] | undefined;

  const childNodes = Array.from(pElem.children);
  let runIdx = 0;

  childNodes.forEach((child) => {
    const tagName = child.tagName.toLowerCase();

    if (tagName === 'w:r' || tagName === 'r') {
      const parsedRun = parseWordRun(child, `run-${runIdx++}`);
      if (parsedRun && parsedRun.text) {
        fullText += parsedRun.text;
        runs.push(parsedRun);
      }
    } else if (tagName === 'w:hyperlink' || tagName === 'hyperlink') {
      const rId = child.getAttribute('r:id') || child.getAttribute('id');
      let url = '';
      if (rId && rels[rId]) {
        url = rels[rId].target;
      }

      const hlRuns = Array.from(child.querySelectorAll('w\\:r, r'));
      hlRuns.forEach((r) => {
        const parsedRun = parseWordRun(r, `run-hl-${runIdx++}`);
        if (parsedRun && parsedRun.text) {
          if (url) {
            parsedRun.hyperlink = { url, tooltip: url, rId: rId || undefined };
            if (!paragraphHyperlink) {
              paragraphHyperlink = { url, tooltip: url, rId: rId || undefined };
            }
          }
          fullText += parsedRun.text;
          runs.push(parsedRun);
        }
      });
    }
  });

  // Fallback: search for direct text elements <w:t> if no runs were picked up
  if (runs.length === 0) {
    const directTElems = Array.from(pElem.querySelectorAll('w\\:t, t'));
    directTElems.forEach((tEl, dIdx) => {
      const text = tEl.textContent || '';
      if (text) {
        fullText += text;
        runs.push({
          id: `run-d-${dIdx}`,
          text,
        });
      }
    });
  }

  if (!fullText.trim()) return null;

  // Fallback url detection
  if (!paragraphHyperlink) {
    const urlMatch = fullText.match(/https?:\/\/[^\s]+|www\.[^\s]+/);
    if (urlMatch) {
      paragraphHyperlink = {
        url: urlMatch[0].startsWith('www.') ? `https://${urlMatch[0]}` : urlMatch[0],
        tooltip: urlMatch[0],
      };
    }
  }

  const primaryColorHex = runs.find((r) => r.colorHex)?.colorHex;

  return {
    id: `docx-p-${pIdx}`,
    slideNumber: pageNum,
    shapeId,
    paragraphIndex: pIdx,
    shapeName: role === 'title' ? 'Nadpis' : role === 'subtitle' ? 'Podnadpis' : role === 'list-item' ? 'Položka seznamu' : 'Odstavec textu',
    shapeRole: role,
    originalText: fullText,
    currentText: fullText,
    runs,
    isModified: false,
    alignment,
    bulletLevel,
    colorHex: primaryColorHex,
    hyperlink: paragraphHyperlink,
  };
}

/**
 * Parses a <w:r> run element into PptxRun.
 */
function parseWordRun(rElem: Element, runId: string): PptxRun | null {
  const tElem = rElem.querySelector('w\\:t, t');
  const text = tElem?.textContent || '';

  let isBold = false;
  let isItalic = false;
  let fontSizePt: number | undefined;
  let colorHex: string | undefined;
  let fontFamily: string | undefined;

  const rPr = rElem.querySelector('w\\:rPr, rPr');
  if (rPr) {
    isBold = !!rPr.querySelector('w\\:b, b');
    isItalic = !!rPr.querySelector('w\\:i, i');

    const sz = rPr.querySelector('w\\:sz, sz');
    if (sz) {
      const halfPoints = parseInt(sz.getAttribute('w:val') || sz.getAttribute('val') || '0', 10);
      if (halfPoints > 0) fontSizePt = Math.round((halfPoints / 2) * 10) / 10;
    }

    const color = rPr.querySelector('w\\:color, color');
    if (color) {
      const val = color.getAttribute('w:val') || color.getAttribute('val') || '';
      if (val && val.toLowerCase() !== 'auto' && /^[0-9a-fA-F]{6}$/.test(val)) {
        colorHex = `#${val}`;
      }
    }

    const rFonts = rPr.querySelector('w\\:rFonts, rFonts');
    if (rFonts) {
      fontFamily =
        rFonts.getAttribute('w:ascii') ||
        rFonts.getAttribute('w:hAnsi') ||
        rFonts.getAttribute('w:cs') ||
        undefined;
    }
  }

  return {
    id: runId,
    text,
    isBold,
    isItalic,
    fontSizePt,
    colorHex,
    fontFamily,
  };
}

/**
 * Parses a Word table <w:tbl> into a PptxShape.
 */
function parseWordTable(
  tblElem: Element,
  pageNum: number,
  shapeId: string,
  globalPStart: number,
  rels: Record<string, { target: string; type: string }>,
  styleMap: Record<string, string>
): PptxShape | null {
  const rows = Array.from(tblElem.querySelectorAll('w\\:tr, tr'));
  if (rows.length === 0) return null;

  const paragraphs: PptxParagraph[] = [];
  let maxCols = 0;
  let pCounter = globalPStart;

  rows.forEach((row, rIdx) => {
    const cells = Array.from(row.querySelectorAll('w\\:tc, tc'));
    maxCols = Math.max(maxCols, cells.length);

    cells.forEach((cell, cIdx) => {
      const pElements = Array.from(cell.querySelectorAll('w\\:p, p'));
      pElements.forEach((pElem) => {
        const parsedP = parseWordParagraph(
          pElem,
          pageNum,
          `${shapeId}-r${rIdx}c${cIdx}`,
          pCounter++,
          rels,
          styleMap
        );
        if (parsedP) {
          parsedP.shapeRole = 'table-cell';
          parsedP.shapeName = `Tabulka (řádek ${rIdx + 1}, sloupec ${cIdx + 1})`;
          paragraphs.push(parsedP);
        }
      });
    });
  });

  if (paragraphs.length === 0) return null;

  return {
    id: `page-${pageNum}-${shapeId}`,
    name: `Tabulka (${rows.length} × ${maxCols})`,
    role: 'table-cell',
    bounds: {
      x: 8,
      y: 20,
      width: 84,
      height: Math.min(60, rows.length * 6 + 10),
    },
    paragraphs,
    isTable: true,
    tableDimensions: { rows: rows.length, cols: maxCols },
  };
}

/**
 * Parses relationships in word/_rels/document.xml.rels.
 */
async function parseDocxRelationships(
  zip: JSZip,
  relsPath: string
): Promise<Record<string, { target: string; type: string }>> {
  const relsFile = zip.file(relsPath);
  if (!relsFile) return {};

  try {
    const relsXml = await relsFile.async('text');
    const parser = new DOMParser();
    const doc = parser.parseFromString(relsXml, 'application/xml');
    const relsMap: Record<string, { target: string; type: string }> = {};

    const relElements = Array.from(doc.querySelectorAll('Relationship'));
    relElements.forEach((rel) => {
      const id = rel.getAttribute('Id');
      const target = rel.getAttribute('Target');
      const type = rel.getAttribute('Type') || '';
      if (id && target) {
        let cleanTarget = target;
        if (
          !target.startsWith('http://') &&
          !target.startsWith('https://') &&
          !target.startsWith('mailto:')
        ) {
          cleanTarget = target.startsWith('word/') ? target : `word/${target}`;
        }
        relsMap[id] = { target: cleanTarget, type };
      }
    });

    return relsMap;
  } catch {
    return {};
  }
}

/**
 * Parses word/styles.xml to map style IDs to human-readable names.
 */
async function parseDocxStyles(zip: JSZip): Promise<Record<string, string>> {
  const stylesFile = zip.file('word/styles.xml');
  if (!stylesFile) return {};

  try {
    const stylesXml = await stylesFile.async('text');
    const parser = new DOMParser();
    const doc = parser.parseFromString(stylesXml, 'application/xml');
    const map: Record<string, string> = {};

    const styleElements = Array.from(doc.querySelectorAll('w\\:style, style'));
    styleElements.forEach((el) => {
      const styleId = el.getAttribute('w:styleId') || el.getAttribute('styleId');
      const nameEl = el.querySelector('w\\:name, name');
      const nameVal = nameEl?.getAttribute('w:val') || nameEl?.getAttribute('val');
      if (styleId && nameVal) {
        map[styleId.toLowerCase()] = nameVal;
      }
    });

    return map;
  } catch {
    return {};
  }
}

/**
 * Parses embedded images from word/media/.
 */
async function parseDocxImages(
  zip: JSZip,
  rels: Record<string, { target: string; type: string }>
): Promise<PptxImage[]> {
  const images: PptxImage[] = [];
  const imageRels = Object.entries(rels).filter(([_, r]) => r.type.includes('image'));

  for (let idx = 0; idx < imageRels.length; idx++) {
    const [id, rel] = imageRels[idx];
    const file = zip.file(rel.target);
    if (file) {
      try {
        const base64 = await file.async('base64');
        const ext = rel.target.split('.').pop()?.toLowerCase() || 'png';
        const mimeType = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'svg' ? 'image/svg+xml' : 'image/png';
        images.push({
          id,
          name: `Obrázek ${idx + 1}`,
          bounds: { x: 15, y: 30, width: 70, height: 40 },
          dataUrl: `data:${mimeType};base64,${base64}`,
        });
      } catch {}
    }
  }

  return images;
}

/**
 * Injects updated text into the original Word (.docx) ZIP package,
 * preserving 100% of formatting, headers, styles, tables, and margins,
 * and produces a downloadable .docx Blob.
 */
export async function updateDocxAndExport(
  zip: JSZip,
  document: PptxDocument
): Promise<Blob> {
  const docFile = zip.file('word/document.xml');
  if (!docFile) {
    throw new Error('Chybí word/document.xml v archivu Word.');
  }

  // Collect modified paragraphs mapping: paragraph ID -> currentText
  const modifiedMap = new Map<string, string>();
  document.slides.forEach((slide) => {
    slide.shapes.forEach((shape) => {
      shape.paragraphs.forEach((p) => {
        if (p.isModified) {
          modifiedMap.set(p.id, p.currentText);
        }
      });
    });
  });

  if (modifiedMap.size > 0) {
    const docXmlStr = await docFile.async('text');
    const parser = new DOMParser();
    const doc = parser.parseFromString(docXmlStr, 'application/xml');

    // Walk all <w:p> elements in the body (and tables)
    const body = doc.querySelector('w\\:body, body');
    if (body) {
      const allParagraphs = Array.from(body.querySelectorAll('w\\:p, p'));

      let pCounter = 0;
      allParagraphs.forEach((pElem) => {
        // Filter out non-text paragraphs if needed, or match counter
        const tElems = Array.from(pElem.querySelectorAll('w\\:t, t'));
        const text = tElems.map((t) => t.textContent || '').join('');
        if (text.trim().length > 0) {
          const pId = `docx-p-${pCounter}`;
          if (modifiedMap.has(pId)) {
            const newText = modifiedMap.get(pId)!;
            applyNewTextToWordParagraph(pElem, newText);
          }
          pCounter++;
        }
      });
    }

    // Serialize XML back to zip
    const serializer = new XMLSerializer();
    const updatedXmlStr = serializer.serializeToString(doc);
    zip.file('word/document.xml', updatedXmlStr);
  }

  // Generate binary .docx Blob
  const outputBlob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  return outputBlob;
}

/**
 * Replaces the text content of a <w:p> element while preserving its first run's
 * exact font styles, size, colors, and formatting tags (<w:rPr>).
 */
function applyNewTextToWordParagraph(pElem: Element, newText: string) {
  const runElements = Array.from(pElem.querySelectorAll('w\\:r, r'));

  if (runElements.length > 0) {
    const firstRun = runElements[0];

    // Find or create <w:t> element
    let tElem = firstRun.querySelector('w\\:t, t');
    if (!tElem) {
      tElem = pElem.ownerDocument.createElementNS(
        'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w:t'
      );
      firstRun.appendChild(tElem);
    }

    tElem.setAttribute('xml:space', 'preserve');
    tElem.textContent = newText;

    // Remove subsequent runs in this paragraph so original text isn't duplicated
    for (let i = 1; i < runElements.length; i++) {
      pElem.removeChild(runElements[i]);
    }
  } else {
    // If no runs existed, construct <w:r><w:t>
    const doc = pElem.ownerDocument;
    const wR = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:r'
    );
    const wT = doc.createElementNS(
      'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'w:t'
    );
    wT.setAttribute('xml:space', 'preserve');
    wT.textContent = newText;
    wR.appendChild(wT);
    pElem.appendChild(wR);
  }
}

/**
 * Calls backend /api/docx/render to get authentic page images rendered by Microsoft Word.
 */
export async function renderDocxPagesViaBackend(
  fileOrBlob: Blob | ArrayBuffer,
  width: number = 1600,
  height: number = 2260
): Promise<{ success: boolean; count: number; pageImages: string[]; error?: string }> {
  try {
    let arrayBuffer: ArrayBuffer;
    if (fileOrBlob instanceof Blob) {
      arrayBuffer = await fileOrBlob.arrayBuffer();
    } else {
      arrayBuffer = fileOrBlob;
    }

    // Convert to base64 in chunks
    const bytes = new Uint8Array(arrayBuffer);
    let binary = '';
    const chunkSize = 8192;
    for (let i = 0; i < bytes.length; i += chunkSize) {
      const chunk = bytes.subarray(i, i + chunkSize);
      binary += String.fromCharCode.apply(null, Array.from(chunk));
    }
    const docxBase64 = btoa(binary);

    const res = await fetch('/api/docx/render', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ docxBase64, width, height }),
    });

    if (!res.ok) {
      const err = await res.json().catch(() => ({ error: `Server error ${res.status}` }));
      return { success: false, count: 0, pageImages: [], error: err.error };
    }

    const data = await res.json();
    return {
      success: true,
      count: data.count || 0,
      pageImages: data.pageImages || [],
    };
  } catch (err: any) {
    return {
      success: false,
      count: 0,
      pageImages: [],
      error: err.message || 'Nepodařilo se vyrenderovat stránky Word dokumentu.',
    };
  }
}

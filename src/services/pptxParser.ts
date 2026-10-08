import JSZip from 'jszip';
import { PptxDocument, PptxSlide, PptxShape, PptxParagraph, PptxRun, PptxImage } from '../types/pptx';

/**
 * Extracts and parses a PPTX file using JSZip and browser XML DOMParser.
 */
export async function parsePptxFile(file: File | Blob | ArrayBuffer, fileName: string = 'prezentace.pptx'): Promise<PptxDocument> {
  const zip = await JSZip.loadAsync(file);

  // 1. Get slide dimensions from ppt/presentation.xml
  let slideWidthEmu = 12192000; // default 16:9 (13.33 inches)
  let slideHeightEmu = 6858000; // default 16:9 (7.5 inches)

  const presentationXmlFile = zip.file('ppt/presentation.xml');
  const slidePaths: string[] = [];

  if (presentationXmlFile) {
    const presXmlStr = await presentationXmlFile.async('text');
    const parser = new DOMParser();
    const presDoc = parser.parseFromString(presXmlStr, 'application/xml');

    const sldSz = presDoc.getElementsByTagName('p:sldSz')[0] || presDoc.getElementsByTagName('sldSz')[0];
    if (sldSz) {
      const cx = sldSz.getAttribute('cx');
      const cy = sldSz.getAttribute('cy');
      if (cx) slideWidthEmu = parseInt(cx, 10);
      if (cy) slideHeightEmu = parseInt(cy, 10);
    }

    // Try reading relationships to preserve exact slide order
    const presRelsFile = zip.file('ppt/_rels/presentation.xml.rels');
    if (presRelsFile) {
      const relsXmlStr = await presRelsFile.async('text');
      const relsDoc = parser.parseFromString(relsXmlStr, 'application/xml');
      const relMap: Record<string, string> = {};

      const relationships = relsDoc.getElementsByTagName('Relationship');
      for (let i = 0; i < relationships.length; i++) {
        const rel = relationships[i];
        const id = rel.getAttribute('Id');
        const target = rel.getAttribute('Target');
        if (id && target && target.includes('slides/slide')) {
          // Normalize target path to ppt/slides/...
          const cleanPath = target.startsWith('/') ? target.slice(1) : target.startsWith('ppt/') ? target : `ppt/${target}`;
          relMap[id] = cleanPath;
        }
      }

      // Follow p:sldIdLst
      const sldIds = presDoc.getElementsByTagName('p:sldId');
      for (let i = 0; i < sldIds.length; i++) {
        const rId = sldIds[i].getAttribute('r:id') || sldIds[i].getAttribute('id');
        if (rId && relMap[rId]) {
          slidePaths.push(relMap[rId]);
        }
      }
    }
  }

  // Fallback: If no slide paths found via rels, scan zip files matching ppt/slides/slide*.xml
  if (slidePaths.length === 0) {
    const rawSlideFiles = Object.keys(zip.files).filter(
      (path) => path.startsWith('ppt/slides/slide') && path.endsWith('.xml') && !path.includes('_rels')
    );
    // Sort numerically slide1.xml, slide2.xml, slide10.xml
    rawSlideFiles.sort((a, b) => {
      const numA = parseInt(a.replace(/[^0-9]/g, '') || '0', 10);
      const numB = parseInt(b.replace(/[^0-9]/g, '') || '0', 10);
      return numA - numB;
    });
    slidePaths.push(...rawSlideFiles);
  }

  const slides: PptxSlide[] = [];
  let totalWords = 0;
  let totalParagraphs = 0;

  // 2. Parse each slide
  for (let idx = 0; idx < slidePaths.length; idx++) {
    const slidePath = slidePaths[idx];
    const slideFile = zip.file(slidePath);
    if (!slideFile) continue;

    const slideNum = idx + 1;
    const slideXmlStr = await slideFile.async('text');
    const parser = new DOMParser();
    const slideDoc = parser.parseFromString(slideXmlStr, 'application/xml');

    // Parse slide relationships (for images, hyperlinks, and speaker notes)
    const relsPath = slidePath.replace('slides/', 'slides/_rels/') + '.rels';
    const rels = await parseSlideRelationships(zip, relsPath);

    // Parse shapes
    const shapes: PptxShape[] = [];
    let slideTitle = `Snímek ${slideNum}`;

    // Collect all shape-like elements: p:sp, p:graphicFrame (tables), p:pic
    const spElements = Array.from(slideDoc.querySelectorAll('p\\:sp, sp'));
    const gfElements = Array.from(slideDoc.querySelectorAll('p\\:graphicFrame, graphicFrame'));

    // Process normal shapes and textboxes
    spElements.forEach((sp, spIdx) => {
      const shape = parseShapeElement(sp, spIdx, slideNum, slideWidthEmu, slideHeightEmu, rels);
      if (shape) {
        shapes.push(shape);
        if (shape.role === 'title' && shape.paragraphs.length > 0) {
          const titleText = shape.paragraphs.map((p) => p.currentText).join(' ').trim();
          if (titleText) {
            slideTitle = titleText.length > 45 ? titleText.slice(0, 42) + '...' : titleText;
          }
        }
      }
    });

    // Process tables
    gfElements.forEach((gf, gfIdx) => {
      const tableShape = parseTableElement(gf, spElements.length + gfIdx, slideNum, slideWidthEmu, slideHeightEmu, rels);
      if (tableShape) {
        shapes.push(tableShape);
      }
    });

    // Extract pictures <p:pic>
    const images = await parseSlidePictures(slideDoc, rels, zip, slideWidthEmu, slideHeightEmu);

    // Extract speaker notes if present
    const speakerNotes = await parseSpeakerNotes(zip, rels);

    // Count words & paragraphs
    shapes.forEach((s) => {
      s.paragraphs.forEach((p) => {
        totalParagraphs++;
        const words = p.currentText.trim().split(/\s+/).filter(Boolean);
        totalWords += words.length;
      });
    });

    if (speakerNotes?.text) {
      const notesWords = speakerNotes.text.trim().split(/\s+/).filter(Boolean);
      totalWords += notesWords.length;
    }

    // Parse slide background if defined (<p:bg>)
    const bgElem = slideDoc.querySelector('p\\:bg, bg');
    const bgClr = bgElem?.querySelector('a\\:srgbClr, srgbClr');
    const slideBackgroundColor = bgClr?.getAttribute('val') ? `#${bgClr.getAttribute('val')}` : '#ffffff';

    slides.push({
      slideNumber: slideNum,
      title: slideTitle,
      xmlPath: slidePath,
      shapes,
      images,
      speakerNotes,
      widthEmu: slideWidthEmu,
      heightEmu: slideHeightEmu,
      aspectRatio: slideWidthEmu / slideHeightEmu,
      backgroundColor: slideBackgroundColor,
    });
  }

  const fileSize = file instanceof Blob ? file.size : file.byteLength;

  return {
    fileName,
    fileSize,
    slides,
    rawZip: zip,
    totalWords,
    totalParagraphs,
    modifiedCount: 0,
  };
}

/**
 * Parses a <p:sp> XML element into a PptxShape.
 */
function parseShapeElement(
  sp: Element,
  index: number,
  slideNum: number,
  slideW: number,
  slideH: number,
  rels?: Record<string, { target: string; type: string }>
): PptxShape | null {
  // Shape ID & Name
  const cNvPr = sp.querySelector('p\\:cNvPr, cNvPr');
  const rawId = cNvPr?.getAttribute('id') || `sp-${index}`;
  const shapeName = cNvPr?.getAttribute('name') || `Prvek ${index + 1}`;

  // Check shape-level hyperlink
  const shapeHlinkClick = cNvPr?.querySelector('a\\:hlinkClick, hlinkClick');
  let shapeHlinkUrl: string | undefined;
  let shapeHlinkRId: string | undefined;
  if (shapeHlinkClick) {
    shapeHlinkRId = shapeHlinkClick.getAttribute('r:id') || shapeHlinkClick.getAttribute('id') || undefined;
    if (shapeHlinkRId && rels && rels[shapeHlinkRId]) {
      shapeHlinkUrl = rels[shapeHlinkRId].target;
    }
  }

  // Role detection
  const ph = sp.querySelector('p\\:ph, ph');
  const phType = ph?.getAttribute('type') || '';
  let role: PptxShape['role'] = 'body';
  if (phType === 'title' || phType === 'ctrTitle') role = 'title';
  else if (phType === 'subTitle') role = 'subtitle';
  else if (phType === 'dt' || phType === 'ftr' || phType === 'sldNum') role = 'footer';
  else if (shapeName.toLowerCase().includes('title') || shapeName.toLowerCase().includes('název')) role = 'title';

  // Bounds <a:xfrm>
  const xfrm = sp.querySelector('p\\:spPr > a\\:xfrm, spPr > xfrm, a\\:xfrm, xfrm');
  let bounds = { x: 5, y: 5 + index * 12, width: 90, height: 10 };

  if (xfrm) {
    const off = xfrm.querySelector('a\\:off, off');
    const ext = xfrm.querySelector('a\\:ext, ext');
    if (off && ext) {
      const offX = parseInt(off.getAttribute('x') || '0', 10);
      const offY = parseInt(off.getAttribute('y') || '0', 10);
      const extCX = parseInt(ext.getAttribute('cx') || '1', 10);
      const extCY = parseInt(ext.getAttribute('cy') || '1', 10);

      bounds = {
        x: Math.max(0, Math.min(100, (offX / slideW) * 100)),
        y: Math.max(0, Math.min(100, (offY / slideH) * 100)),
        width: Math.max(2, Math.min(100, (extCX / slideW) * 100)),
        height: Math.max(2, Math.min(100, (extCY / slideH) * 100)),
      };
    }
  }

  // Extract Shape Background & Border Colors
  const spPr = sp.querySelector('p\\:spPr, spPr');
  const solidFillClr = spPr?.querySelector('a\\:solidFill > a\\:srgbClr, solidFill > srgbClr');
  const backgroundColor = solidFillClr?.getAttribute('val') ? `#${solidFillClr.getAttribute('val')}` : undefined;

  const lnClr = spPr?.querySelector('a\\:ln a\\:srgbClr, ln srgbClr');
  const borderColor = lnClr?.getAttribute('val') ? `#${lnClr.getAttribute('val')}` : undefined;

  // Parse text body <p:txBody>
  const txBody = sp.querySelector('p\\:txBody, txBody');
  const paragraphs: PptxParagraph[] = [];

  if (txBody) {
    const pElements = Array.from(txBody.querySelectorAll('a\\:p, p'));
    pElements.forEach((pElem, pIdx) => {
      const paragraph = parseParagraphElement(pElem, pIdx, rawId, shapeName, role, slideNum, rels, shapeHlinkUrl ? { url: shapeHlinkUrl, rId: shapeHlinkRId } : undefined);
      if (paragraph && paragraph.currentText.trim().length > 0) {
        paragraphs.push(paragraph);
      }
    });
  }

  // If shape has no text, skip it from the text editor
  if (paragraphs.length === 0) {
    return null;
  }

  return {
    id: rawId,
    name: shapeName,
    role,
    bounds,
    backgroundColor,
    borderColor,
    paragraphs,
  };
}

/**
 * Parses a table inside <p:graphicFrame>.
 */
function parseTableElement(
  gf: Element,
  index: number,
  slideNum: number,
  slideW: number,
  slideH: number,
  rels?: Record<string, { target: string; type: string }>
): PptxShape | null {
  const tbl = gf.querySelector('a\\:tbl, tbl');
  if (!tbl) return null;

  const cNvPr = gf.querySelector('p\\:cNvPr, cNvPr');
  const rawId = cNvPr?.getAttribute('id') || `tbl-${index}`;
  const shapeName = cNvPr?.getAttribute('name') || `Tabulka ${index + 1}`;

  // Bounds
  const xfrm = gf.querySelector('p\\:xfrm, xfrm');
  let bounds = { x: 10, y: 30, width: 80, height: 50 };
  if (xfrm) {
    const off = xfrm.querySelector('a\\:off, off');
    const ext = xfrm.querySelector('a\\:ext, ext');
    if (off && ext) {
      const offX = parseInt(off.getAttribute('x') || '0', 10);
      const offY = parseInt(off.getAttribute('y') || '0', 10);
      const extCX = parseInt(ext.getAttribute('cx') || '1', 10);
      const extCY = parseInt(ext.getAttribute('cy') || '1', 10);
      bounds = {
        x: Math.max(0, Math.min(100, (offX / slideW) * 100)),
        y: Math.max(0, Math.min(100, (offY / slideH) * 100)),
        width: Math.max(5, Math.min(100, (extCX / slideW) * 100)),
        height: Math.max(5, Math.min(100, (extCY / slideH) * 100)),
      };
    }
  }

  const paragraphs: PptxParagraph[] = [];
  const rows = Array.from(tbl.querySelectorAll('a\\:tr, tr'));

  rows.forEach((row, rIdx) => {
    const cells = Array.from(row.querySelectorAll('a\\:tc, tc'));
    cells.forEach((cell, cIdx) => {
      const txBody = cell.querySelector('a\\:txBody, txBody');
      if (txBody) {
        const pElements = Array.from(txBody.querySelectorAll('a\\:p, p'));
        pElements.forEach((pElem, pIdx) => {
          const para = parseParagraphElement(
            pElem,
            pIdx,
            `${rawId}-r${rIdx}c${cIdx}`,
            `Buňka [ř. ${rIdx + 1}, sl. ${cIdx + 1}]`,
            'table-cell',
            slideNum,
            rels
          );
          if (para && para.currentText.trim().length > 0) {
            paragraphs.push(para);
          }
        });
      }
    });
  });

  if (paragraphs.length === 0) return null;

  return {
    id: rawId,
    name: shapeName,
    role: 'table-cell',
    bounds,
    paragraphs,
    isTable: true,
  };
}

/**
 * Parses an <a:p> element into a PptxParagraph with its text runs.
 */
function parseParagraphElement(
  pElem: Element,
  pIdx: number,
  shapeId: string,
  shapeName: string,
  role: PptxParagraph['shapeRole'],
  slideNum: number,
  rels?: Record<string, { target: string; type: string }>,
  defaultHlink?: { url: string; tooltip?: string; rId?: string }
): PptxParagraph | null {
  const pPr = pElem.querySelector('a\\:pPr, pPr');
  const algn = pPr?.getAttribute('algn');
  let alignment: PptxParagraph['alignment'] = 'left';
  if (algn === 'ctr') alignment = 'center';
  else if (algn === 'r') alignment = 'right';
  else if (algn === 'just') alignment = 'justify';

  const lvl = pPr?.getAttribute('lvl');
  const bulletLevel = lvl ? parseInt(lvl, 10) : 0;

  const runs: PptxRun[] = [];
  let fullText = '';
  let paragraphHyperlink = defaultHlink;

  // Scan child elements: <a:r>, <a:fld>, etc.
  const runElements = Array.from(pElem.querySelectorAll('a\\:r, r, a\\:fld, fld'));

  runElements.forEach((rElem, rIdx) => {
    const tElem = rElem.querySelector('a\\:t, t');
    const text = tElem?.textContent || '';
    if (text) {
      fullText += text;

      const rPr = rElem.querySelector('a\\:rPr, rPr');
      const isBold = rPr?.getAttribute('b') === '1' || rPr?.getAttribute('b') === 'true';
      const isItalic = rPr?.getAttribute('i') === '1' || rPr?.getAttribute('i') === 'true';
      const sz = rPr?.getAttribute('sz');
      const fontSizePt = sz ? parseInt(sz, 10) / 100 : undefined;

      const srgbClr = rPr?.querySelector('a\\:srgbClr, srgbClr');
      const colorHex = srgbClr?.getAttribute('val') ? `#${srgbClr.getAttribute('val')}` : undefined;

      // Extract run hyperlink
      const hlinkClick = rPr?.querySelector('a\\:hlinkClick, hlinkClick');
      let runHlink = defaultHlink;

      if (hlinkClick) {
        const rId = hlinkClick.getAttribute('r:id') || hlinkClick.getAttribute('id') || undefined;
        const tooltip = hlinkClick.getAttribute('tooltip') || undefined;
        let url = rId && rels && rels[rId] ? rels[rId].target : undefined;
        if (!url) {
          url = hlinkClick.getAttribute('action') || hlinkClick.getAttribute('invalidUrl') || undefined;
        }
        if (url) {
          runHlink = { url, tooltip, rId };
          if (!paragraphHyperlink) paragraphHyperlink = runHlink;
        }
      }

      // Detect URL in text if no XML hyperlink
      if (!runHlink) {
        const urlMatch = text.match(/https?:\/\/[^\s]+|www\.[^\s]+/);
        if (urlMatch) {
          runHlink = {
            url: urlMatch[0].startsWith('www.') ? `https://${urlMatch[0]}` : urlMatch[0],
            tooltip: urlMatch[0],
          };
          if (!paragraphHyperlink) paragraphHyperlink = runHlink;
        }
      }

      runs.push({
        id: `run-${rIdx}`,
        text,
        isBold,
        isItalic,
        fontSizePt,
        colorHex,
        hyperlink: runHlink,
      });
    }
  });

  // Fallback: check direct <a:t> elements inside <a:p> if no <a:r> runs matched
  if (runs.length === 0) {
    const directTElems = Array.from(pElem.querySelectorAll('a\\:t, t'));
    directTElems.forEach((tEl, dIdx) => {
      const text = tEl.textContent || '';
      if (text) {
        fullText += text;
        runs.push({
          id: `run-d-${dIdx}`,
          text,
          hyperlink: defaultHlink,
        });
      }
    });
  }

  if (!fullText) return null;

  // Fallback: check full text for URL
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
    id: `slide-${slideNum}-shape-${shapeId}-p-${pIdx}`,
    slideNumber: slideNum,
    shapeId,
    paragraphIndex: pIdx,
    shapeName,
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
 * Injects updated text into the original PPTX zip XML, preserving 100% of formatting,
 * themes, masters, and shapes, and outputs a downloadable .pptx Blob.
 */
export async function updatePptxAndExport(
  zip: JSZip,
  slides: PptxSlide[]
): Promise<Blob> {
  // Create a copy of zip or work on the existing zip instance
  for (const slide of slides) {
    const slideFile = zip.file(slide.xmlPath);
    if (!slideFile) continue;

    // Map modified paragraphs for fast lookup
    const modifiedParagraphs = new Map<string, string>();
    slide.shapes.forEach((shape) => {
      shape.paragraphs.forEach((p) => {
        if (p.isModified) {
          modifiedParagraphs.set(p.id, p.currentText);
        }
      });
    });

    if (modifiedParagraphs.size === 0) continue;

    const xmlStr = await slideFile.async('text');
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlStr, 'application/xml');

    // Walk through shapes & paragraphs matching IDs
    const spElements = Array.from(doc.querySelectorAll('p\\:sp, sp'));
    spElements.forEach((sp, spIdx) => {
      const cNvPr = sp.querySelector('p\\:cNvPr, cNvPr');
      const shapeId = cNvPr?.getAttribute('id') || `sp-${spIdx}`;

      const txBody = sp.querySelector('p\\:txBody, txBody');
      if (txBody) {
        const pElements = Array.from(txBody.querySelectorAll('a\\:p, p'));
        let nonBlankCount = 0;
        pElements.forEach((pElem, pIdx) => {
          const pIdExact = `slide-${slide.slideNumber}-shape-${shapeId}-p-${pIdx}`;
          const pIdCounter = `slide-${slide.slideNumber}-shape-${shapeId}-p-${nonBlankCount}`;

          if (modifiedParagraphs.has(pIdExact)) {
            const newText = modifiedParagraphs.get(pIdExact)!;
            applyNewTextToParagraph(pElem, newText);
          } else if (modifiedParagraphs.has(pIdCounter)) {
            const newText = modifiedParagraphs.get(pIdCounter)!;
            applyNewTextToParagraph(pElem, newText);
          }

          const tElems = Array.from(pElem.querySelectorAll('a\\:t, t'));
          const existingText = tElems.map((t) => t.textContent || '').join('');
          if (existingText.trim().length > 0) {
            nonBlankCount++;
          }
        });
      }
    });

    // Walk through tables
    const gfElements = Array.from(doc.querySelectorAll('p\\:graphicFrame, graphicFrame'));
    gfElements.forEach((gf, gfIdx) => {
      const tbl = gf.querySelector('a\\:tbl, tbl');
      if (!tbl) return;

      const cNvPr = gf.querySelector('p\\:cNvPr, cNvPr');
      const shapeId = cNvPr?.getAttribute('id') || `tbl-${gfIdx}`;

      const rows = Array.from(tbl.querySelectorAll('a\\:tr, tr'));
      rows.forEach((row, rIdx) => {
        const cells = Array.from(row.querySelectorAll('a\\:tc, tc'));
        cells.forEach((cell, cIdx) => {
          const txBody = cell.querySelector('a\\:txBody, txBody');
          if (txBody) {
            const pElements = Array.from(txBody.querySelectorAll('a\\:p, p'));
            let nonBlankCount = 0;
            pElements.forEach((pElem, pIdx) => {
              const pIdExact = `slide-${slide.slideNumber}-shape-${shapeId}-r${rIdx}c${cIdx}-p-${pIdx}`;
              const pIdCounter = `slide-${slide.slideNumber}-shape-${shapeId}-r${rIdx}c${cIdx}-p-${nonBlankCount}`;

              if (modifiedParagraphs.has(pIdExact)) {
                const newText = modifiedParagraphs.get(pIdExact)!;
                applyNewTextToParagraph(pElem, newText);
              } else if (modifiedParagraphs.has(pIdCounter)) {
                const newText = modifiedParagraphs.get(pIdCounter)!;
                applyNewTextToParagraph(pElem, newText);
              }

              const tElems = Array.from(pElem.querySelectorAll('a\\:t, t'));
              const existingText = tElems.map((t) => t.textContent || '').join('');
              if (existingText.trim().length > 0) {
                nonBlankCount++;
              }
            });
          }
        });
      });
    });

    // Serialize back to XML
    const serializer = new XMLSerializer();
    const updatedXmlStr = serializer.serializeToString(doc);
    zip.file(slide.xmlPath, updatedXmlStr);
  }

  // Update speaker notes if modified
  for (const slide of slides) {
    if (slide.speakerNotes?.isModified && slide.speakerNotes.xmlPath) {
      const notesFile = zip.file(slide.speakerNotes.xmlPath);
      if (notesFile) {
        const notesXmlStr = await notesFile.async('text');
        const parser = new DOMParser();
        const notesDoc = parser.parseFromString(notesXmlStr, 'application/xml');

        // Find body shape in notes
        const bodyPh = notesDoc.querySelector('p\\:ph[type="body"], ph[type="body"]');
        const sp = bodyPh?.closest('p\\:sp, sp');
        if (sp) {
          const txBody = sp.querySelector('p\\:txBody, txBody');
          const pElem = txBody?.querySelector('a\\:p, p');
          if (pElem) {
            applyNewTextToParagraph(pElem, slide.speakerNotes.text);
            const serializer = new XMLSerializer();
            zip.file(slide.speakerNotes.xmlPath, serializer.serializeToString(notesDoc));
          }
        }
      }
    }
  }

  // Generate binary PPTX blob
  const outputBlob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });

  return outputBlob;
}

/**
 * Parses slide .rels file to map relationship IDs to targets.
 */
async function parseSlideRelationships(
  zip: JSZip,
  relsPath: string
): Promise<Record<string, { target: string; type: string }>> {
  const relsFile = zip.file(relsPath);
  if (!relsFile) return {};

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
      // Resolve target path relative to ppt/slides/
      let cleanTarget = target;
      if (
        target.startsWith('http://') ||
        target.startsWith('https://') ||
        target.startsWith('mailto:') ||
        target.startsWith('www.')
      ) {
        cleanTarget = target;
      } else if (target.startsWith('../')) {
        cleanTarget = `ppt/${target.slice(3)}`;
      } else if (!target.startsWith('ppt/')) {
        cleanTarget = `ppt/slides/${target}`;
      }
      relsMap[id] = { target: cleanTarget, type };
    }
  });

  return relsMap;
}

/**
 * Extracts pictures <p:pic> from slide XML and loads their image data from zip.
 */
async function parseSlidePictures(
  slideDoc: Document,
  rels: Record<string, { target: string; type: string }>,
  zip: JSZip,
  slideW: number,
  slideH: number
): Promise<PptxImage[]> {
  const images: PptxImage[] = [];
  const picElements = Array.from(slideDoc.querySelectorAll('p\\:pic, pic'));

  for (let i = 0; i < picElements.length; i++) {
    const pic = picElements[i];
    const cNvPr = pic.querySelector('p\\:cNvPr, cNvPr');
    const name = cNvPr?.getAttribute('name') || `Obrázek ${i + 1}`;
    const id = cNvPr?.getAttribute('id') || `pic-${i}`;

    const blip = pic.querySelector('a\\:blip, blip');
    const rEmbed = blip?.getAttribute('r:embed') || blip?.getAttribute('embed');

    if (rEmbed && rels[rEmbed]) {
      const mediaPath = rels[rEmbed].target;
      const mediaFile = zip.file(mediaPath);
      if (mediaFile) {
        try {
          const base64 = await mediaFile.async('base64');
          const ext = mediaPath.split('.').pop()?.toLowerCase() || 'png';
          const mimeType = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'svg' ? 'image/svg+xml' : 'image/png';
          const dataUrl = `data:${mimeType};base64,${base64}`;

          // Bounds
          const xfrm = pic.querySelector('p\\:spPr > a\\:xfrm, spPr > xfrm, a\\:xfrm, xfrm');
          let bounds = { x: 10, y: 10, width: 30, height: 30 };
          if (xfrm) {
            const off = xfrm.querySelector('a\\:off, off');
            const extEl = xfrm.querySelector('a\\:ext, ext');
            if (off && extEl) {
              const offX = parseInt(off.getAttribute('x') || '0', 10);
              const offY = parseInt(off.getAttribute('y') || '0', 10);
              const extCX = parseInt(extEl.getAttribute('cx') || '1', 10);
              const extCY = parseInt(extEl.getAttribute('cy') || '1', 10);
              bounds = {
                x: Math.max(0, Math.min(100, (offX / slideW) * 100)),
                y: Math.max(0, Math.min(100, (offY / slideH) * 100)),
                width: Math.max(1, Math.min(100, (extCX / slideW) * 100)),
                height: Math.max(1, Math.min(100, (extCY / slideH) * 100)),
              };
            }
          }

          images.push({
            id,
            name,
            bounds,
            dataUrl,
          });
        } catch {
          // ignore corrupted media
        }
      }
    }
  }

  return images;
}

/**
 * Extracts speaker notes for the slide if present in rels.
 */
async function parseSpeakerNotes(
  zip: JSZip,
  rels: Record<string, { target: string; type: string }>
): Promise<{ text: string; originalText: string; isModified: boolean; xmlPath?: string } | undefined> {
  const notesRel = Object.values(rels).find((r) => r.type.includes('notesSlide'));
  if (!notesRel) return undefined;

  const notesFile = zip.file(notesRel.target);
  if (!notesFile) return undefined;

  try {
    const notesXml = await notesFile.async('text');
    const parser = new DOMParser();
    const doc = parser.parseFromString(notesXml, 'application/xml');

    const bodyPh = doc.querySelector('p\\:ph[type="body"], ph[type="body"]');
    const sp = bodyPh?.closest('p\\:sp, sp');
    if (sp) {
      const txBody = sp.querySelector('p\\:txBody, txBody');
      if (txBody) {
        const textRuns = Array.from(txBody.querySelectorAll('a\\:t, t'));
        const fullNotesText = textRuns.map((t) => t.textContent || '').join('');
        if (fullNotesText.trim()) {
          return {
            text: fullNotesText,
            originalText: fullNotesText,
            isModified: false,
            xmlPath: notesRel.target,
          };
        }
      }
    }
  } catch {
    // ignore notes parsing issues
  }

  return undefined;
}

/**
 * Exports a bilingual comparison report of the presentation.
 */
export function exportBilingualReport(
  document: PptxDocument,
  format: 'markdown' | 'csv' | 'html' = 'html'
): string {
  if (format === 'markdown') {
    let md = `# Bilingvní přehled prezentace: ${document.fileName}\n\n`;
    md += `| Snímek | Typ prvku | Původní text | Upravený / Přeložený text | Změněno |\n`;
    md += `|---|---|---|---|---|\n`;

    document.slides.forEach((slide) => {
      slide.shapes.forEach((shape) => {
        shape.paragraphs.forEach((p) => {
          const orig = p.originalText.replace(/\|/g, '\\|').replace(/\n/g, ' ');
          const curr = p.currentText.replace(/\|/g, '\\|').replace(/\n/g, ' ');
          md += `| #${slide.slideNumber} | ${p.shapeRole} | ${orig} | ${curr} | ${p.isModified ? 'ANO' : 'NE'} |\n`;
        });
      });
    });

    return md;
  }

  if (format === 'csv') {
    let csv = `Snímek;Typ prvku;Původní text;Upravený text;Změněno\n`;
    document.slides.forEach((slide) => {
      slide.shapes.forEach((shape) => {
        shape.paragraphs.forEach((p) => {
          const orig = `"${p.originalText.replace(/"/g, '""')}"`;
          const curr = `"${p.currentText.replace(/"/g, '""')}"`;
          csv += `#${slide.slideNumber};${p.shapeRole};${orig};${curr};${p.isModified ? 'ANO' : 'NE'}\n`;
        });
      });
    });
    return csv;
  }

  // HTML format
  let rows = '';
  document.slides.forEach((slide) => {
    slide.shapes.forEach((shape) => {
      shape.paragraphs.forEach((p) => {
        rows += `
        <tr class="${p.isModified ? 'bg-amber-50 font-medium' : ''}">
          <td style="padding: 10px; border: 1px solid #e2e8f0; font-weight: bold;">Snímek #${slide.slideNumber}</td>
          <td style="padding: 10px; border: 1px solid #e2e8f0; color: #64748b; font-size: 12px;">${p.shapeRole}</td>
          <td style="padding: 10px; border: 1px solid #e2e8f0; color: #475569;">${escapeXml(p.originalText)}</td>
          <td style="padding: 10px; border: 1px solid #e2e8f0; color: #0f172a; font-weight: ${p.isModified ? 'bold' : 'normal'};">${escapeXml(p.currentText)}</td>
          <td style="padding: 10px; border: 1px solid #e2e8f0; text-align: center;">${p.isModified ? '<span style="color: #d97706; font-weight: bold;">ANO</span>' : 'NE'}</td>
        </tr>`;
      });
    });
  });

  return `<!DOCTYPE html>
<html lang="cs">
<head>
  <meta charset="UTF-8">
  <title>Bilingvní srovnání - ${escapeXml(document.fileName)}</title>
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 30px; background: #f8fafc; color: #0f172a; }
    h1 { font-size: 22px; margin-bottom: 5px; }
    p { color: #64748b; font-size: 13px; margin-top: 0; margin-bottom: 20px; }
    table { width: 100%; border-collapse: collapse; background: white; box-shadow: 0 1px 3px rgba(0,0,0,0.1); border-radius: 8px; overflow: hidden; font-size: 13px; }
    th { background: #0f172a; color: white; padding: 12px; text-align: left; font-size: 12px; text-transform: uppercase; }
  </style>
</head>
<body>
  <h1>Bilingvní srovnávací přehled prezentace</h1>
  <p>Soubor: ${escapeXml(document.fileName)} • Vygenerováno ze SlideText Studio</p>
  <table>
    <thead>
      <tr>
        <th style="width: 100px;">Snímek</th>
        <th style="width: 120px;">Role</th>
        <th style="width: 40%;">Původní text</th>
        <th style="width: 40%;">Upravený / Přeložený text</th>
        <th style="width: 80px; text-align: center;">Změna</th>
      </tr>
    </thead>
    <tbody>
      ${rows}
    </tbody>
  </table>
</body>
</html>`;
}

function escapeXml(unsafe: string): string {
  return unsafe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Replaces the text content of an <a:p> element while preserving its first run's
 * exact font styles, size, colors, and formatting tags (<a:rPr>).
 */
function applyNewTextToParagraph(pElem: Element, newText: string) {
  const runElements = Array.from(pElem.querySelectorAll('a\\:r, r'));

  if (runElements.length > 0) {
    // Put new text in the first run
    const firstRun = runElements[0];

    // Check if any run in this paragraph had a hyperlink
    let existingHlinkNode: Element | null = null;
    for (const r of runElements) {
      const hl = r.querySelector('a\\:hlinkClick, hlinkClick');
      if (hl) {
        existingHlinkNode = hl;
        break;
      }
    }

    let rPr = firstRun.querySelector('a\\:rPr, rPr');
    if (existingHlinkNode) {
      if (!rPr) {
        rPr = pElem.ownerDocument.createElementNS(
          'http://schemas.openxmlformats.org/drawingml/2006/main',
          'a:rPr'
        );
        firstRun.insertBefore(rPr, firstRun.firstChild);
      }
      if (!rPr.querySelector('a\\:hlinkClick, hlinkClick')) {
        rPr.appendChild(existingHlinkNode.cloneNode(true));
      }
    }

    let tElem = firstRun.querySelector('a\\:t, t');
    if (!tElem) {
      tElem = pElem.ownerDocument.createElementNS(
        'http://schemas.openxmlformats.org/drawingml/2006/main',
        'a:t'
      );
      firstRun.appendChild(tElem);
    }
    tElem.textContent = newText;

    // Remove remaining runs so old text does not duplicate
    for (let i = 1; i < runElements.length; i++) {
      pElem.removeChild(runElements[i]);
    }
  } else {
    // If no runs existed, create <a:r><a:t>
    const doc = pElem.ownerDocument;
    const aR = doc.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:r');
    const aT = doc.createElementNS('http://schemas.openxmlformats.org/drawingml/2006/main', 'a:t');
    aT.textContent = newText;
    aR.appendChild(aT);
    pElem.appendChild(aR);
  }
}

/**
 * Calls backend /api/pptx/render to get authentic 1080p slide images
 * rendered directly by Microsoft PowerPoint.
 */
export async function renderPptxSlidesViaBackend(
  fileOrBlob: Blob | ArrayBuffer,
  width = 1920,
  height = 1080
): Promise<string[]> {
  try {
    let buffer: ArrayBuffer;
    if (fileOrBlob instanceof Blob) {
      buffer = await fileOrBlob.arrayBuffer();
    } else {
      buffer = fileOrBlob;
    }

    let base64: string;
    if (typeof FileReader !== 'undefined') {
      base64 = await new Promise<string>((resolve, reject) => {
        const blob = fileOrBlob instanceof Blob ? fileOrBlob : new Blob([buffer]);
        const reader = new FileReader();
        reader.onload = () => {
          const dataUrl = reader.result as string;
          resolve(dataUrl.split(',')[1] || '');
        };
        reader.onerror = reject;
        reader.readAsDataURL(blob);
      });
    } else {
      base64 = Buffer.from(buffer).toString('base64');
    }

    const res = await fetch('/api/pptx/render', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ pptxBase64: base64, width, height }),
    });

    if (!res.ok) {
      console.warn('Backend /api/pptx/render vrátil kód', res.status);
      return [];
    }

    const data = await res.json();
    if (data.success && Array.isArray(data.slideImages)) {
      return data.slideImages;
    }
    return [];
  } catch (err) {
    console.warn('Chyba při renderování přes PowerPoint backend:', err);
    return [];
  }
}


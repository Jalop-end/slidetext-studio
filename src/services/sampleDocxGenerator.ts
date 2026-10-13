import JSZip from 'jszip';

/**
 * Generates an authentic sample Microsoft Word (.docx) document in pure JavaScript using JSZip.
 * Contains headings, styled paragraphs, bullet lists, formatting, and a structured table.
 */
export async function generateSampleDocx(): Promise<Blob> {
  const zip = new JSZip();

  // 1. [Content_Types].xml
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`
  );

  // 2. _rels/.rels
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  // 3. word/_rels/document.xml.rels
  zip.file(
    'word/_rels/document.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`
  );

  // 4. word/styles.xml
  zip.file(
    'word/styles.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:color w:val="1E293B"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="1E40AF"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/>
    <w:pPr><w:spacing w:before="180" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="2563EB"/></w:rPr>
  </w:style>
</w:styles>`
  );

  // 5. word/document.xml
  zip.file(
    'word/document.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <!-- Document Title -->
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Heading1"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:b/>
          <w:sz w:val="44"/>
          <w:color w:val="1E3A8A"/>
        </w:rPr>
        <w:t xml:space="preserve">Výroční strategická zpráva &amp; Plán rozvoje</w:t>
      </w:r>
    </w:p>

    <!-- Subtitle / Meta -->
    <w:p>
      <w:pPr>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:i/>
          <w:color w:val="64748B"/>
          <w:sz w:val="22"/>
        </w:rPr>
        <w:t xml:space="preserve">Připraveno pro vedení společnosti • Verze 2.4 • Kvartál Q4</w:t>
      </w:r>
    </w:p>

    <!-- Heading 1 -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t xml:space="preserve">1. Manažerské shrnutí a hlavní milníky</w:t></w:r>
    </w:p>

    <!-- Body Paragraph -->
    <w:p>
      <w:r>
        <w:t xml:space="preserve">V uplynulém období dosáhla naše společnost mimořádného růstu v klíčových oblastech. Úspěšně jsme modernizovali infrastrukturu a zavedli automatizované podpůrné služby napříč všemi pobočkami. Celková spokojenost klientů vzrostla o </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:b/><w:color w:val="15803D"/></w:rPr>
        <w:t xml:space="preserve">28 %</w:t>
      </w:r>
      <w:r>
        <w:t xml:space="preserve"> díky zkrácení doby odezvy na požadavky.</w:t>
      </w:r>
    </w:p>

    <!-- Heading 2 -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t xml:space="preserve">2. Strategické priority a inovace</w:t></w:r>
    </w:p>

    <!-- Bullet 1 -->
    <w:p>
      <w:pPr>
        <w:ind w:left="720"/>
      </w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">• Digitální transformace: </w:t></w:r>
      <w:r><w:t xml:space="preserve">Zavedení pokročilých AI nástrojů a zrychlení zpracování datových toků.</w:t></w:r>
    </w:p>

    <!-- Bullet 2 -->
    <w:p>
      <w:pPr>
        <w:ind w:left="720"/>
      </w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">• Mezinárodní expanze: </w:t></w:r>
      <w:r><w:t xml:space="preserve">Otevření obchodních center ve střední Evropě a adaptace lokálních procesů.</w:t></w:r>
    </w:p>

    <!-- Bullet 3 -->
    <w:p>
      <w:pPr>
        <w:ind w:left="720"/>
      </w:pPr>
      <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">• Udržitelnost a efektivita: </w:t></w:r>
      <w:r><w:t xml:space="preserve">Snížení provozních nákladů o 15 % optimalizací cloudových kapacit.</w:t></w:r>
    </w:p>

    <!-- Heading 2 -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading2"/></w:pPr>
      <w:r><w:t xml:space="preserve">3. Finanční ukazatele a výkonnost</w:t></w:r>
    </w:p>

    <!-- Table -->
    <w:tbl>
      <w:tblPr>
        <w:tblW w:w="9600" w:type="dxa"/>
        <w:jc w:val="center"/>
      </w:tblPr>
      <!-- Row 1: Header -->
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="3200" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="1E3A8A"/></w:tcPr>
          <w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t xml:space="preserve">Ukazatel</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="3200" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="1E3A8A"/></w:tcPr>
          <w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t xml:space="preserve">Rok 2025 (Skutečnost)</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="3200" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="1E3A8A"/></w:tcPr>
          <w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t xml:space="preserve">Cíl 2026</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <!-- Row 2 -->
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="3200" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Obrat (v mil. Kč)</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="3200" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t xml:space="preserve">142,5</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="3200" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t xml:space="preserve">180,0 (+26 %)</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
      <!-- Row 3 -->
      <w:tr>
        <w:tc>
          <w:tcPr><w:tcW w:w="3200" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Počet aktivních uživatelů</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="3200" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t xml:space="preserve">45 200</w:t></w:r></w:p>
        </w:tc>
        <w:tc>
          <w:tcPr><w:tcW w:w="3200" w:type="dxa"/></w:tcPr>
          <w:p><w:r><w:t xml:space="preserve">65 000</w:t></w:r></w:p>
        </w:tc>
      </w:tr>
    </w:tbl>

    <!-- Page Break to test multi-page Word documents -->
    <w:p>
      <w:r>
        <w:br w:type="page"/>
      </w:r>
    </w:p>

    <!-- Page 2 Content -->
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t xml:space="preserve">4. Závěrečná doporučení a další kroky</w:t></w:r>
    </w:p>

    <w:p>
      <w:r>
        <w:t xml:space="preserve">Doporučujeme schválit investiční plán pro rozvoj lidských zdrojů a posílení technické podpory. Tým projektového řízení předloží detailní harmonogram do konce příštího měsíce.</w:t>
      </w:r>
    </w:p>

    <!-- Section properties: A4 Portrait -->
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`
  );

  return await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
}

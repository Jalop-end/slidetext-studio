import JSZip from 'jszip';

/**
 * Generates a valid standard OpenXML PPTX file with rich sample slides
 * so users can test immediately without uploading a local file.
 */
export async function generateSamplePptx(sampleType: 'czech-strategy' | 'english-pitch' = 'czech-strategy'): Promise<{ blob: Blob; fileName: string }> {
  const isCzech = sampleType === 'czech-strategy';
  const fileName = isCzech ? 'Inovace_Strategie_2026.pptx' : 'Global_Tech_Pitch_2026.pptx';
  const publicPath = isCzech ? '/sample_czech.pptx' : '/sample_english.pptx';

  try {
    const res = await fetch(publicPath);
    if (res.ok) {
      const blob = await res.blob();
      return { blob, fileName };
    }
  } catch (_) {}

  const zip = new JSZip();

  // [Content_Types].xml
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide4.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`
  );

  // _rels/.rels
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`
  );

  // ppt/presentation.xml
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
    <p:sldId id="258" r:id="rId3"/>
    <p:sldId id="259" r:id="rId4"/>
  </p:sldIdLst>
</p:presentation>`
  );

  // ppt/_rels/presentation.xml.rels
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide4.xml"/>
</Relationships>`
  );

  // Content definitions depending on sample type
  const slidesContent = isCzech
    ? [
        {
          title: 'Strategie digitální transformace a umělé inteligence 2026',
          subtitle: 'Vize pro automatizaci procesů, expanzi na evropský trh a zvýšení efektivity týmů',
          note: 'Přednáší: Výkonná rada pro inovace a technologický rozvoj',
          items: [],
        },
        {
          title: 'Tři klíčové pilíře našeho technologického růstu',
          subtitle: 'Zaměření na udržitelné technologie a spokojenost koncových klientů',
          items: [
            '1. Automatizace zákaznické podpory s využitím moderních AI modelů',
            '2. Modernizace cloudové infrastruktury a snížení provozních nákladů o 35 %',
            '3. Expanze lokalizovaných služeb do 12 nových evropských zemí v průběhu roku',
          ],
        },
        {
          title: 'Analýza tržního potenciálu a návratnost investic',
          subtitle: 'Očekávaný meziroční růst a optimalizace kapitálových výdajů',
          items: [
            'Očekávané navýšení hrubých tržeb: +145 % do konce třetího kvartálu',
            'Snížení chybovosti manuálního zpracování dat na méně než 0,2 %',
            'Zkrácení doby uvedení nových produktů na trh (Time-to-Market) z 6 měsíců na 3 týdny',
          ],
        },
        {
          title: 'Závěr a následující kroky realizačního plánu',
          subtitle: 'Harmonogram spuštění pilotního projektu od 1. listopadu 2026',
          items: [
            'Schválení finálního rozpočtu a alokace klíčových inženýrských kapacit',
            'Zahájení školení zaměstnanců a spuštění interní testovací fáze',
            'Webová platforma a oficiální projektová dokumentace: https://innovation2026.example.com',
          ],
        },
      ]
    : [
        {
          title: 'NextGen Cloud Architecture & Enterprise AI Platform',
          subtitle: 'Revolutionizing scalable distributed systems and intelligent workflow automation',
          note: 'Presented by: Executive Innovation Board & Engineering Leadership',
          items: [],
        },
        {
          title: 'Three Pillars of Our 2026 Growth Strategy',
          subtitle: 'Focusing on enterprise reliability, cost reduction, and global scale',
          items: [
            '1. Autonomous operations powered by state-of-the-art generative models',
            '2. Multi-region low-latency infrastructure slashing operating expenses by 40%',
            '3. Seamless localization and multi-language rollout across 15 European markets',
          ],
        },
        {
          title: 'Market Opportunity and Projected Financial Metrics',
          subtitle: 'Consistent compounding growth with verified enterprise unit economics',
          items: [
            'Projected Annual Recurring Revenue (ARR) growth: +180% year-over-year',
            'Enterprise customer retention rate exceeding 96.5% across all segments',
            'Accelerated customer onboarding cycle from 45 days down to under 4 hours',
          ],
        },
        {
          title: 'Key Milestones and Next Action Steps',
          subtitle: 'Roadmap execution timeline starting Q4 2026',
          items: [
            'Finalize strategic enterprise contracts and engineering capacity allocation',
            'Deploy secure sandbox environments for tier-1 alpha partners',
            'Official project documentation & API developer portal: https://nextgen-enterprise.example.com',
          ],
        },
      ];

  // Helper to generate slide XML
  slidesContent.forEach((data, i) => {
    const slideNum = i + 1;
    let itemsXml = '';
    if (data.items.length > 0) {
      itemsXml = data.items
        .map(
          (item) => `
        <a:p>
          <a:pPr lvl="0">
            <a:buFont typeface="Arial"/>
            <a:buChar char="•"/>
          </a:pPr>
          <a:r>
            <a:rPr lang="cs-CZ" sz="1900">
              <a:solidFill><a:srgbClr val="334155"/></a:solidFill>
            </a:rPr>
            <a:t>${escapeXml(item)}</a:t>
          </a:r>
        </a:p>`
        )
        .join('');
    }

    const subtitleXml = data.subtitle
      ? `
        <a:p>
          <a:r>
            <a:rPr lang="cs-CZ" sz="1800">
              <a:solidFill><a:srgbClr val="64748B"/></a:solidFill>
            </a:rPr>
            <a:t>${escapeXml(data.subtitle)}</a:t>
          </a:r>
        </a:p>`
      : '';

    const noteXml = (data as any).note
      ? `
        <a:p>
          <a:r>
            <a:rPr lang="cs-CZ" sz="1500" i="1">
              <a:solidFill><a:srgbClr val="94A3B8"/></a:solidFill>
            </a:rPr>
            <a:t>${escapeXml((data as any).note)}</a:t>
          </a:r>
        </a:p>`
      : '';

    const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>

      <!-- Title Shape -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="title"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="838200" y="800000"/>
            <a:ext cx="10515600" cy="1325563"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="cs-CZ" sz="3200" b="1">
                <a:solidFill><a:srgbClr val="0F172A"/></a:solidFill>
              </a:rPr>
              <a:t>${escapeXml(data.title)}</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>

      <!-- Content / Subtitle Shape -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Content 1"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="body" idx="1"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="838200" y="2400000"/>
            <a:ext cx="10515600" cy="3800000"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          ${subtitleXml}
          ${noteXml}
          ${itemsXml}
        </p:txBody>
      </p:sp>

      <!-- Footer Shape -->
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="4" name="Slide Number"/>
          <p:cNvSpPr/>
          <p:nvPr>
            <p:ph type="sldNum" sz="quarter" idx="12"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="10500000" y="6300000"/>
            <a:ext cx="850000" cy="300000"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="r"/>
            <a:r>
              <a:rPr sz="1200">
                <a:solidFill><a:srgbClr val="94A3B8"/></a:solidFill>
              </a:rPr>
              <a:t>Slide ${slideNum} / 4</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`;

    zip.file(`ppt/slides/slide${slideNum}.xml`, slideXml);
  });

  const blob = await zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  });

  return { blob, fileName };
}

function escapeXml(unsafe: string): string {
  return unsafe
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

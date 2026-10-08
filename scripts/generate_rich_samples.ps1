param()

$ppt = New-Object -ComObject PowerPoint.Application
try {
  # -------------------------------------------------------------------------
  # 1. Czech Strategy Presentation
  # -------------------------------------------------------------------------
  $presCz = $ppt.Presentations.Add(0)
  
  # Slide 1: Title
  $s1 = $presCz.Slides.Add(1, 1)
  $s1.Shapes.Title.TextFrame.TextRange.Text = "Strategie digitální transformace 2026"
  $s1.Shapes.Item(2).TextFrame.TextRange.Text = "Vize pro automatizaci procesů, expanzi na trh a moderní AI integraci`nPřednáší: Výkonná rada pro inovace"
  # Add background color card
  $rect1 = $s1.Shapes.AddShape(1, 40, 20, 60, 500) # msoShapeRectangle
  $rect1.Fill.ForeColor.RGB = 0x2563EB # Blue
  $rect1.Line.Visible = 0

  # Slide 2: Pillars
  $s2 = $presCz.Slides.Add(2, 2)
  $s2.Shapes.Title.TextFrame.TextRange.Text = "Tři klíčové pilíře technologického růstu"
  $s2.Shapes.Item(2).TextFrame.TextRange.Text = "1. Automatizace zákaznické podpory s využitím moderních AI modelů`n2. Modernizace cloudové infrastruktury a snížení nákladů o 35 %`n3. Expanze lokalizovaných služeb do 12 nových evropských zemí"

  # Slide 3: Analysis
  $s3 = $presCz.Slides.Add(3, 2)
  $s3.Shapes.Title.TextFrame.TextRange.Text = "Analýza tržního potenciálu a návratnost investic"
  $s3.Shapes.Item(2).TextFrame.TextRange.Text = "Očekávané navýšení hrubých tržeb: +145 % do konce třetího kvartálu`nSnížení chybovosti manuálního zpracování na méně než 0,2 %`nZkrácení doby uvedení nových produktů (Time-to-Market) na 3 týdny"

  # Slide 4: Summary & Roadmap
  $s4 = $presCz.Slides.Add(4, 2)
  $s4.Shapes.Title.TextFrame.TextRange.Text = "Závěr a následující kroky realizačního plánu"
  $s4.Shapes.Item(2).TextFrame.TextRange.Text = "Schválení finálního rozpočtu a alokace klíčových inženýrských kapacit`nZahájení školení zaměstnanců a spuštění interní testovací fáze`nOficiální projektová dokumentace: https://innovation2026.example.com"

  $czPath = Join-Path $PSScriptRoot "..\public\sample_czech.pptx"
  $presCz.SaveAs((Resolve-Path (Join-Path $PSScriptRoot "..\public")).Path + "\sample_czech.pptx")
  $presCz.Close()
  Write-Output "Created sample_czech.pptx"

  # -------------------------------------------------------------------------
  # 2. English Pitch Presentation
  # -------------------------------------------------------------------------
  $presEn = $ppt.Presentations.Add(0)

  # Slide 1: Title
  $se1 = $presEn.Slides.Add(1, 1)
  $se1.Shapes.Title.TextFrame.TextRange.Text = "NextGen Cloud Architecture & Enterprise AI 2026"
  $se1.Shapes.Item(2).TextFrame.TextRange.Text = "Scaling intelligent workflows, distributed resilience, and autonomous operations`nPresented by: Executive Engineering Leadership"
  $rectE1 = $se1.Shapes.AddShape(1, 40, 20, 60, 500)
  $rectE1.Fill.ForeColor.RGB = 0x059669 # Emerald
  $rectE1.Line.Visible = 0

  # Slide 2: Pillars
  $se2 = $presEn.Slides.Add(2, 2)
  $se2.Shapes.Title.TextFrame.TextRange.Text = "Three Strategic Pillars of Global Scalability"
  $se2.Shapes.Item(2).TextFrame.TextRange.Text = "1. Autonomous workflow engines powered by state-of-the-art models`n2. Multi-region cloud architecture slashing latency and operating costs`n3. Native multilingual localization across 15 European markets"

  # Slide 3: Growth Metrics
  $se3 = $presEn.Slides.Add(3, 2)
  $se3.Shapes.Title.TextFrame.TextRange.Text = "Market Opportunity and Projected Enterprise Metrics"
  $se3.Shapes.Item(2).TextFrame.TextRange.Text = "Projected Annual Recurring Revenue (ARR) growth: +180% year-over-year`nEnterprise retention rate exceeding 96.5% across all cohorts`nAccelerated client onboarding cycle reduced from 45 days down to 4 hours"

  # Slide 4: Roadmap
  $se4 = $presEn.Slides.Add(4, 2)
  $se4.Shapes.Title.TextFrame.TextRange.Text = "Execution Roadmap and Immediate Next Steps"
  $se4.Shapes.Item(2).TextFrame.TextRange.Text = "Finalize strategic tier-1 enterprise partnerships and engineering allocation`nDeploy secure zero-trust developer sandbox environments`nOfficial developer documentation & platform portal: https://nextgen-enterprise.example.com"

  $presEn.SaveAs((Resolve-Path (Join-Path $PSScriptRoot "..\public")).Path + "\sample_english.pptx")
  $presEn.Close()
  Write-Output "Created sample_english.pptx"

} finally {
  $ppt.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

import React, { useState, useRef, useEffect } from 'react';
import {
  FileText,
  ExternalLink,
  RotateCcw,
  Sparkles,
  Languages,
  AlertTriangle,
  Lightbulb,
  ChevronDown,
  ChevronRight,
  ChevronLeft,
  Layers,
  Link as LinkIcon,
  Eye,
  CheckCircle2,
  Loader2,
} from 'lucide-react';
import { PptxDocument, PptxParagraph, PptxSlide } from '../types/pptx';
import { useLanguage } from '../context/LanguageContext';
import { fetchTranslationAlternatives } from '../services/aiService';

interface DocumentEditorProps {
  document: PptxDocument;
  selectedParagraphId: string | null;
  onSelectParagraph: (paragraphId: string, slideNumber: number) => void;
  onUpdateParagraphText: (paragraphId: string, newText: string) => void;
  onResetParagraphText: (paragraphId: string) => void;
  currentSlideNumber: number;
  onSlideChange: (slideNumber: number) => void;
  onSingleTranslate: (paragraph: PptxParagraph) => void;
  onTranslateSlide?: (slideNumber: number) => void;
  onSinglePolish: (paragraph: PptxParagraph) => void;
  onUpdateSpeakerNotes: (slideNumber: number, newNotes: string) => void;
  onUpdateParagraphHyperlink: (paragraphId: string, newUrl: string) => void;
}

export const DocumentEditor: React.FC<DocumentEditorProps> = ({
  document,
  selectedParagraphId,
  onSelectParagraph,
  onUpdateParagraphText,
  onResetParagraphText,
  currentSlideNumber,
  onSlideChange,
  onSingleTranslate,
  onTranslateSlide,
  onSinglePolish,
  onUpdateSpeakerNotes,
  onUpdateParagraphHyperlink,
}) => {
  const { t } = useLanguage();
  // View mode: 'single' (only current slide), 'modified' (only slides with changes), 'all' (all slides)
  const [viewFilter, setViewFilter] = useState<'single' | 'modified' | 'all'>('single');
  const [collapsedSlides, setCollapsedSlides] = useState<Record<number, boolean>>({});
  const [showDiffMap, setShowDiffMap] = useState<Record<string, boolean>>({});
  const itemRefs = useRef<Record<string, HTMLDivElement | null>>({});
  const [loadingAltsForId, setLoadingAltsForId] = useState<string | null>(null);
  const [dynamicAlternatives, setDynamicAlternatives] = useState<Record<string, string[]>>({});

  const handleFetchMoreAlternatives = async (p: PptxParagraph) => {
    setLoadingAltsForId(p.id);
    try {
      const alts = await fetchTranslationAlternatives(p.originalText, p.currentText, 'en');
      if (alts && alts.length > 0) {
        setDynamicAlternatives((prev) => ({
          ...prev,
          [p.id]: Array.from(new Set([...(prev[p.id] || p.alternatives || []), ...alts])),
        }));
      }
    } finally {
      setLoadingAltsForId(null);
    }
  };

  // Auto-scroll when a paragraph is selected from the Slide Visualizer
  useEffect(() => {
    if (selectedParagraphId && itemRefs.current[selectedParagraphId]) {
      itemRefs.current[selectedParagraphId]?.scrollIntoView({
        behavior: 'smooth',
        block: 'nearest',
      });
    }
  }, [selectedParagraphId]);

  const toggleSlideCollapse = (slideNum: number) => {
    setCollapsedSlides((prev) => ({
      ...prev,
      [slideNum]: !prev[slideNum],
    }));
  };

  const toggleDiff = (paragraphId: string) => {
    setShowDiffMap((prev) => ({
      ...prev,
      [paragraphId]: !prev[paragraphId],
    }));
  };

  const handlePrevSlide = () => {
    if (currentSlideNumber > 1) {
      onSlideChange(currentSlideNumber - 1);
    }
  };

  const handleNextSlide = () => {
    if (currentSlideNumber < document.slides.length) {
      onSlideChange(currentSlideNumber + 1);
    }
  };

  const modifiedSlides = document.slides.filter(
    (s) => s.shapes.some((sh) => sh.paragraphs.some((p) => p.isModified)) || !!s.speakerNotes?.isModified
  );
  const totalModifiedSlidesCount = modifiedSlides.length;

  const areAllSlideDiffsOpen = (slide: PptxSlide) => {
    const modParas = slide.shapes.flatMap((sh) => sh.paragraphs.filter((p) => p.isModified));
    return modParas.length > 0 && modParas.every((p) => !!showDiffMap[p.id]);
  };

  const toggleAllSlideDiffs = (slide: PptxSlide) => {
    const currentlyOpen = areAllSlideDiffsOpen(slide);
    const modParas = slide.shapes.flatMap((sh) => sh.paragraphs.filter((p) => p.isModified));
    setShowDiffMap((prev) => {
      const next = { ...prev };
      modParas.forEach((p) => {
        next[p.id] = !currentlyOpen;
      });
      return next;
    });
  };

  const displayedSlides =
    viewFilter === 'single'
      ? document.slides.filter((s) => s.slideNumber === currentSlideNumber)
      : viewFilter === 'modified'
      ? modifiedSlides
      : document.slides;

  return (
    <div className="flex-1 overflow-y-auto bg-slate-100/70 p-4 sm:p-6 lg:p-8 font-sans">
      <div className="max-w-4xl mx-auto space-y-6">
        {/* Slide Selector & Navigation Toolbar */}
        <div className="bg-white rounded-xl p-4 shadow-xs border border-slate-200/80 sticky top-0 z-20 space-y-3">
          {/* Row 1: Mode toggle & Slide pagination */}
          <div className="flex flex-wrap items-center justify-between gap-3">
            {/* Mode pills */}
            <div className="flex items-center gap-1 bg-slate-100 p-1 rounded-xl border border-slate-200">
              <button
                type="button"
                onClick={() => setViewFilter('single')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                  viewFilter === 'single'
                    ? 'bg-white text-blue-700 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
                title={t('singleSlideViewTooltip', 'Zobrazit pouze jeden vybraný snímek bez nutnosti skrolování')}
              >
                <Eye className="w-3.5 h-3.5" />
                <span>{t('singleSlideView', 'Vybraný snímek')}</span>
              </button>

              <button
                type="button"
                onClick={() => setViewFilter('modified')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                  viewFilter === 'modified'
                    ? 'bg-white text-amber-700 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
                title={t('onlyModifiedTooltip', 'Zobrazit pouze snímky, na kterých byly provedeny změny')}
              >
                <Sparkles className="w-3.5 h-3.5 text-amber-500" />
                <span>{t('onlyModifiedSlides', 'Pouze se změnami')}</span>
                {totalModifiedSlidesCount > 0 && (
                  <span className="px-1.5 py-0.2 rounded-full text-[10px] bg-amber-100 text-amber-800 font-extrabold">
                    {totalModifiedSlidesCount}
                  </span>
                )}
              </button>

              <button
                type="button"
                onClick={() => setViewFilter('all')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                  viewFilter === 'all'
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
                title={t('allSlidesTooltip', 'Zobrazit celou prezentaci jako souvislý dokument')}
              >
                <Layers className="w-3.5 h-3.5" />
                <span>{t('allSlidesView', 'Všechny snímky')}</span>
              </button>
            </div>

            {/* Prev / Next & Slide Dropdown */}
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={handlePrevSlide}
                disabled={currentSlideNumber <= 1}
                className="p-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-100 disabled:opacity-30 disabled:pointer-events-none transition-colors"
                title="Předchozí snímek"
              >
                <ChevronLeft className="w-4 h-4" />
              </button>

              {/* Select dropdown */}
              <select
                value={currentSlideNumber}
                onChange={(e) => {
                  const num = Number(e.target.value);
                  onSlideChange(num);
                }}
                className="text-xs font-bold bg-slate-50 border border-slate-200 text-slate-800 rounded-lg px-2.5 py-1.5 focus:outline-none focus:ring-2 focus:ring-blue-500/20 max-w-[200px] truncate"
              >
                {document.slides.map((s) => {
                  const modCount = s.shapes.reduce(
                    (acc, sh) => acc + sh.paragraphs.filter((p) => p.isModified).length,
                    0
                  );
                  return (
                    <option key={`opt-slide-${s.slideNumber}`} value={s.slideNumber}>
                      {`#${s.slideNumber}: ${s.title || 'Snímek ' + s.slideNumber}${
                        modCount > 0 ? ` (${modCount} změn)` : ''
                      }`}
                    </option>
                  );
                })}
              </select>

              <button
                type="button"
                onClick={handleNextSlide}
                disabled={currentSlideNumber >= document.slides.length}
                className="p-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-100 disabled:opacity-30 disabled:pointer-events-none transition-colors"
                title="Další snímek"
              >
                <ChevronRight className="w-4 h-4" />
              </button>
            </div>
          </div>

          {/* Row 2: Slide Pills Scroll */}
          <div className="flex items-center gap-1.5 overflow-x-auto pb-1 pt-0.5">
            {document.slides.map((s) => {
              const isSelected = s.slideNumber === currentSlideNumber;
              const modCount = s.shapes.reduce(
                (acc, sh) => acc + sh.paragraphs.filter((p) => p.isModified).length,
                0
              );
              return (
                <button
                  key={`pill-slide-${s.slideNumber}`}
                  type="button"
                  onClick={() => {
                    onSlideChange(s.slideNumber);
                    if (viewFilter === 'all') {
                      const el = window.document.getElementById(`slide-page-${s.slideNumber}`);
                      el?.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    }
                  }}
                  className={`px-2.5 py-1 rounded-lg text-xs font-bold transition-all shrink-0 flex items-center gap-1.5 ${
                    isSelected
                      ? 'bg-blue-600 text-white shadow-xs ring-2 ring-blue-500/30'
                      : 'bg-slate-100 text-slate-700 hover:bg-slate-200'
                  }`}
                >
                  <span>#{s.slideNumber}</span>
                  {modCount > 0 && (
                    <span
                      className={`w-2 h-2 rounded-full ${
                        isSelected ? 'bg-amber-300' : 'bg-emerald-500'
                      }`}
                      title={`${modCount} upravených textů`}
                    />
                  )}
                </button>
              );
            })}
          </div>
        </div>

        {/* Empty state when filtering for modified slides but none found */}
        {viewFilter === 'modified' && displayedSlides.length === 0 && (
          <div className="bg-white rounded-xl p-8 border border-slate-200/80 text-center shadow-xs space-y-3">
            <div className="w-12 h-12 rounded-full bg-slate-100 text-slate-500 mx-auto flex items-center justify-center">
              <CheckCircle2 className="w-6 h-6 text-emerald-500" />
            </div>
            <h3 className="font-bold text-slate-900 text-sm">
              {t('noModifiedSlidesTitle', 'Zatím nebyly provedeny žádné úpravy')}
            </h3>
            <p className="text-xs text-slate-500 max-w-sm mx-auto">
              {t(
                'noModifiedSlidesDesc',
                'Všechny texty jsou v původním stavu. Zvolte "Vybraný snímek" pro editaci nebo spusťte překlad prezentace.'
              )}
            </p>
            <div className="flex items-center justify-center gap-2 pt-2">
              <button
                type="button"
                onClick={() => setViewFilter('single')}
                className="px-3.5 py-1.5 rounded-lg bg-blue-600 text-white text-xs font-semibold hover:bg-blue-700 transition-colors"
              >
                {t('showCurrentSlide', 'Zobrazit vybraný snímek')}
              </button>
            </div>
          </div>
        )}

        {/* Slides rendering */}
        {displayedSlides.map((slide) => {
          const isCurrent = slide.slideNumber === currentSlideNumber;
          const isCollapsed = !!collapsedSlides[slide.slideNumber];

          const slideModifiedCount = slide.shapes.reduce(
            (acc, shape) => acc + shape.paragraphs.filter((p) => p.isModified).length,
            0
          );

          return (
            <div
              key={`slide-doc-${slide.slideNumber}`}
              id={`slide-page-${slide.slideNumber}`}
              className={`bg-white rounded-xl shadow-md border transition-all duration-200 overflow-hidden ${
                isCurrent
                  ? 'ring-2 ring-blue-500/80 border-blue-300'
                  : 'border-slate-200 hover:border-slate-300'
              }`}
            >
              {/* Word Document Slide Page Header */}
              <div
                onClick={() => {
                  onSlideChange(slide.slideNumber);
                }}
                className={`px-6 py-4 flex items-center justify-between cursor-pointer select-none transition-colors ${
                  isCurrent
                    ? 'bg-blue-50/60 border-b border-blue-100'
                    : 'bg-slate-50/80 border-b border-slate-200/70 hover:bg-slate-100/60'
                }`}
              >
                <div className="flex items-center gap-3 min-w-0">
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      toggleSlideCollapse(slide.slideNumber);
                    }}
                    className="p-1 rounded text-slate-400 hover:text-slate-700 hover:bg-slate-200/60 transition-colors"
                  >
                    {isCollapsed ? (
                      <ChevronRight className="w-4 h-4" />
                    ) : (
                      <ChevronDown className="w-4 h-4" />
                    )}
                  </button>

                  <span className="w-7 h-7 rounded-lg bg-slate-800 text-white font-bold text-xs flex items-center justify-center shrink-0 shadow-xs">
                    #{slide.slideNumber}
                  </span>

                  <div className="min-w-0">
                    <h3 className="font-bold text-slate-900 text-sm truncate flex items-center gap-2">
                      {slide.title || `${t('slide', 'Snímek')} ${slide.slideNumber}`}
                      {isCurrent && (
                        <span className="text-[10px] uppercase font-bold tracking-wider px-1.5 py-0.5 rounded bg-blue-100 text-blue-700 border border-blue-200/60">
                          {t('activePreview', 'Aktivní náhled')}
                        </span>
                      )}
                    </h3>
                    <p className="text-xs text-slate-500 truncate">
                      {slide.shapes.reduce((acc, s) => acc + s.paragraphs.length, 0)}{' '}
                      {t('textBlocksCount', 'textových bloků')}
                      {slideModifiedCount > 0 &&
                        ` • ${slideModifiedCount} ${t('modified', 'upraveno')}`}
                    </p>
                  </div>
                </div>

                <div className="flex items-center gap-2 shrink-0">
                  {onTranslateSlide && (
                    <button
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        onTranslateSlide(slide.slideNumber);
                      }}
                      className="text-xs font-semibold px-2.5 py-1 rounded-md text-indigo-700 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 flex items-center gap-1.5 transition-colors"
                      title={t('translateSlideTooltip', 'Přeložit všechny texty na tomto snímku')}
                    >
                      <Languages className="w-3.5 h-3.5 text-indigo-600" />
                      <span>{t('translateSlide', 'Přeložit snímek')}</span>
                    </button>
                  )}

                  {slideModifiedCount > 0 && (
                    <button
                      type="button"
                      onClick={(e) => {
                        e.stopPropagation();
                        toggleAllSlideDiffs(slide);
                      }}
                      className="text-xs font-semibold px-2.5 py-1 rounded-md text-amber-800 bg-amber-100 hover:bg-amber-200 border border-amber-300 flex items-center gap-1.5 transition-colors"
                      title={t('toggleAllSlideDiffsTooltip', 'Přepnout porovnání pro všechny upravené texty na tomto snímku')}
                    >
                      <Sparkles className="w-3 h-3 text-amber-600" />
                      <span>
                        {areAllSlideDiffsOpen(slide)
                          ? t('hideSlideDiffs', 'Skrýt porovnání')
                          : t('showSlideDiffs', 'Porovnat změny')}
                      </span>
                    </button>
                  )}

                  {/* Jump button to slide */}
                  <button
                    type="button"
                    onClick={(e) => {
                      e.stopPropagation();
                      onSlideChange(slide.slideNumber);
                    }}
                    className="text-xs font-semibold px-2.5 py-1 rounded-md text-blue-700 bg-blue-100/80 hover:bg-blue-200/80 border border-blue-300 flex items-center gap-1.5 transition-colors"
                    title={t('goToSlideTooltip', 'Zobrazit tento snímek v náhledu vpravo')}
                  >
                    <ExternalLink className="w-3 h-3" />
                    <span>{t('goToSlide', 'Přejít na slide')}</span>
                  </button>
                </div>
              </div>

              {/* Word Document Slide Body */}
              {!isCollapsed && (
                <div className="p-6 sm:p-8 space-y-6">
                  {slide.shapes.length === 0 ? (
                    <div className="text-center py-6 text-slate-400 text-xs italic">
                      {t(
                        'emptySlideText',
                        'Tento snímek neobsahuje žádné editovatelné textové prvky.'
                      )}
                    </div>
                  ) : (
                    slide.shapes.map((shape) => (
                      <div
                        key={`shape-${shape.id}`}
                        className="space-y-3 pl-2 sm:pl-4 border-l-2 border-slate-200/80 hover:border-slate-400 transition-colors"
                      >
                        {/* Shape Context Badge */}
                        <div className="flex items-center justify-between text-xs text-slate-500 font-medium pb-1">
                          <div className="flex items-center gap-2">
                            <span
                              className={`px-2 py-0.5 rounded text-[11px] font-semibold uppercase tracking-wider ${
                                shape.role === 'title'
                                  ? 'bg-amber-100 text-amber-800 border border-amber-200'
                                  : shape.role === 'subtitle'
                                  ? 'bg-orange-100 text-orange-800 border border-orange-200'
                                  : shape.role === 'table-cell' || shape.isTable
                                  ? 'bg-purple-100 text-purple-800 border border-purple-200'
                                  : 'bg-slate-100 text-slate-700 border border-slate-200'
                              }`}
                            >
                              {shape.role === 'title'
                                ? t('titleShape', 'Nadpis snímku')
                                : shape.role === 'subtitle'
                                ? t('subtitleShape', 'Podnadpis')
                                : shape.role === 'table-cell' || shape.isTable
                                ? t('tableShape', 'Tabulka')
                                : t('bodyShape', 'Textové pole / Odrážky')}
                            </span>
                            <span className="text-slate-400 text-xs truncate">
                              {shape.name}
                            </span>
                          </div>

                          {(shape.role === 'table-cell' || shape.isTable) && onTranslateSlide && (
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                onTranslateSlide(slide.slideNumber);
                              }}
                              className="text-[11px] font-semibold text-purple-700 hover:text-purple-900 bg-purple-50 hover:bg-purple-100 px-2 py-0.5 rounded border border-purple-200 flex items-center gap-1 transition-colors"
                              title="Přeložit všechny buňky v této tabulce"
                            >
                              <Languages className="w-3 h-3 text-purple-600" />
                              <span>Přeložit celou tabulku</span>
                            </button>
                          )}
                        </div>

                        {/* Paragraphs in shape */}
                        {shape.paragraphs.map((paragraph) => {
                          const isSelected = selectedParagraphId === paragraph.id;
                          const showDiff = !!showDiffMap[paragraph.id];

                          return (
                            <div
                              key={paragraph.id}
                              ref={(el) => {
                                itemRefs.current[paragraph.id] = el;
                              }}
                              onClick={() => onSelectParagraph(paragraph.id, slide.slideNumber)}
                              className={`p-4 rounded-xl border transition-all duration-200 ${
                                isSelected
                                  ? 'bg-blue-50/70 border-blue-400 shadow-sm ring-2 ring-blue-400/30'
                                  : paragraph.isModified
                                  ? 'bg-amber-50/30 border-amber-200/90 hover:border-amber-300'
                                  : 'bg-slate-50/50 border-slate-200 hover:border-slate-300 hover:bg-slate-50'
                              }`}
                            >
                              {/* Element Header with Mini-Link and Actions */}
                              <div className="flex flex-wrap items-center justify-between gap-2 mb-2 pb-2 border-b border-slate-200/60">
                                <div className="flex items-center gap-2">
                                  {/* Mini Interactive Link to Slide Position */}
                                  <button
                                    type="button"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      onSelectParagraph(paragraph.id, slide.slideNumber);
                                    }}
                                    className={`text-[11px] font-semibold px-2 py-0.5 rounded flex items-center gap-1.5 transition-colors ${
                                      isSelected
                                        ? 'bg-blue-600 text-white shadow-xs'
                                        : 'bg-blue-100/90 text-blue-700 hover:bg-blue-200 border border-blue-200'
                                    }`}
                                    title={t(
                                      'highlightTooltip',
                                      'Klikněte pro zobrazení a zvýraznění tohoto prvku na snímku'
                                    )}
                                  >
                                    <ExternalLink className="w-3 h-3" />
                                    <span>
                                      Slide #{slide.slideNumber} • {t('position', 'Pozice')}{' '}
                                      {paragraph.paragraphIndex + 1}
                                    </span>
                                  </button>

                                  {paragraph.isModified && (
                                    <span className="text-[10px] font-semibold px-1.5 py-0.5 rounded bg-amber-100 text-amber-800 border border-amber-200 flex items-center gap-1">
                                      <span className="w-1.5 h-1.5 rounded-full bg-amber-500"></span>
                                      {t('modifiedShort', 'Upraveno')}
                                    </span>
                                  )}

                                  {paragraph.bulletLevel ? (
                                    <span className="text-[10px] font-medium text-slate-500">
                                      ({t('bulletLevelText', 'Úroveň odrážky')}{' '}
                                      {paragraph.bulletLevel})
                                    </span>
                                  ) : null}
                                </div>

                                {/* Right toolbar: AI actions & Diff */}
                                <div className="flex items-center gap-1">
                                  {paragraph.isModified && (
                                    <button
                                      type="button"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        toggleDiff(paragraph.id);
                                      }}
                                      className={`text-[11px] px-2 py-0.5 rounded font-medium transition-colors ${
                                        showDiff
                                          ? 'bg-slate-800 text-white'
                                          : 'text-slate-600 hover:bg-slate-200/70'
                                      }`}
                                      title={t('compareTooltip', 'Porovnat původní text se současným')}
                                    >
                                      {showDiff
                                        ? t('hideCompare', 'Skrýt porovnání')
                                        : t('compareOriginal', 'Porovnat s originálem')}
                                    </button>
                                  )}

                                  <button
                                    type="button"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      onSingleTranslate(paragraph);
                                    }}
                                    className="p-1 rounded text-indigo-600 hover:bg-indigo-100/70 hover:text-indigo-800 transition-colors text-[11px] font-medium flex items-center gap-1 px-1.5"
                                    title={t('translateTooltip', 'Přeložit tento odstavec')}
                                  >
                                    <Languages className="w-3 h-3" />
                                    <span className="hidden sm:inline">
                                      {t('translateWithAi', 'Přeložit')}
                                    </span>
                                  </button>

                                  <button
                                    type="button"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      handleFetchMoreAlternatives(paragraph);
                                    }}
                                    disabled={loadingAltsForId === paragraph.id}
                                    className="p-1 rounded text-purple-700 hover:bg-purple-100/70 hover:text-purple-900 transition-colors text-[11px] font-medium flex items-center gap-1 px-1.5 disabled:opacity-50"
                                    title="Získat alternativní styly a varianty překladu přes Gemma 4"
                                  >
                                    {loadingAltsForId === paragraph.id ? (
                                      <Loader2 className="w-3 h-3 animate-spin" />
                                    ) : (
                                      <Sparkles className="w-3 h-3" />
                                    )}
                                    <span className="hidden sm:inline">Varianty</span>
                                  </button>

                                  <button
                                    type="button"
                                    onClick={(e) => {
                                      e.stopPropagation();
                                      onSinglePolish(paragraph);
                                    }}
                                    className="p-1 rounded text-amber-700 hover:bg-amber-100 hover:text-amber-900 transition-colors text-[11px] font-medium flex items-center gap-1 px-1.5"
                                    title={t('polishTooltip', 'Opravit gramatiku a přeformulovat s Gemma 4')}
                                  >
                                    <Sparkles className="w-3 h-3" />
                                    <span className="hidden sm:inline">
                                      {t('polishWithAi', 'Korektura')}
                                    </span>
                                  </button>

                                  {paragraph.isModified && (
                                    <button
                                      type="button"
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        onResetParagraphText(paragraph.id);
                                      }}
                                      className="p-1 rounded text-slate-400 hover:text-rose-600 hover:bg-rose-50 transition-colors"
                                      title={t('revertToOriginal', 'Vrátit na originál')}
                                    >
                                      <RotateCcw className="w-3 h-3" />
                                    </button>
                                  )}
                                </div>
                              </div>

                              {/* Diff Comparison View */}
                              {showDiff && paragraph.isModified && (
                                <div className="mb-3 p-2.5 rounded-lg bg-slate-100/90 border border-slate-200 text-xs space-y-1.5">
                                  <div className="flex items-start gap-2">
                                    <span className="text-[10px] font-bold uppercase tracking-wider text-rose-600 w-14 shrink-0 pt-0.5">
                                      {t('origLabel', 'Původní:')}
                                    </span>
                                    <p className="text-slate-600 line-through decoration-rose-300">
                                      {paragraph.originalText}
                                    </p>
                                  </div>
                                  <div className="flex items-start gap-2">
                                    <span className="text-[10px] font-bold uppercase tracking-wider text-emerald-600 w-14 shrink-0 pt-0.5">
                                      {t('newLabel', 'Nový:')}
                                    </span>
                                    <p className="text-slate-900 font-medium">
                                      {paragraph.currentText}
                                    </p>
                                  </div>
                                </div>
                              )}

                              {/* Editable Text Field */}
                              <div className="relative">
                                <textarea
                                  value={paragraph.currentText}
                                  onChange={(e) => onUpdateParagraphText(paragraph.id, e.target.value)}
                                  onFocus={() => onSelectParagraph(paragraph.id, slide.slideNumber)}
                                  rows={Math.max(2, Math.min(8, Math.ceil(paragraph.currentText.length / 70)))}
                                  className={`w-full text-sm leading-relaxed p-2.5 rounded-lg border bg-white focus:outline-none transition-colors resize-y ${
                                    isSelected
                                      ? 'border-blue-400 focus:ring-2 focus:ring-blue-500/20'
                                      : 'border-slate-300 focus:border-blue-400 focus:ring-2 focus:ring-blue-500/20'
                                  } ${
                                    paragraph.shapeRole === 'title'
                                      ? 'font-bold text-slate-900 text-base'
                                      : paragraph.shapeRole === 'subtitle'
                                      ? 'font-medium text-slate-800'
                                      : 'text-slate-800'
                                  }`}
                                  placeholder={t('inputPlaceholder', 'Zadejte text odstavce...')}
                                />
                              </div>

                              {/* Domain Alternative Translations Chips (e.g. Support services vs Ancillary) */}
                              {(() => {
                                const currentAlts =
                                  dynamicAlternatives[paragraph.id] || paragraph.alternatives;
                                if (!currentAlts || currentAlts.length === 0) return null;
                                return (
                                  <div className="mt-2 p-2.5 rounded-xl bg-indigo-50/80 border border-indigo-200/90 text-xs space-y-1.5 animate-in fade-in duration-150">
                                    <div className="flex items-center justify-between">
                                      <span className="text-[10px] font-bold uppercase tracking-wider text-indigo-900 flex items-center gap-1.5">
                                        <Sparkles className="w-3.5 h-3.5 text-indigo-600 shrink-0" />
                                        <span>Varianty překladu (klikněte pro okamžitou záměnu):</span>
                                      </span>
                                      <button
                                        type="button"
                                        onClick={(e) => {
                                          e.stopPropagation();
                                          handleFetchMoreAlternatives(paragraph);
                                        }}
                                        disabled={loadingAltsForId === paragraph.id}
                                        className="text-[10px] font-semibold text-indigo-700 hover:text-indigo-900 hover:underline flex items-center gap-1 disabled:opacity-50"
                                      >
                                        {loadingAltsForId === paragraph.id ? (
                                          <Loader2 className="w-3 h-3 animate-spin" />
                                        ) : (
                                          <Sparkles className="w-3 h-3" />
                                        )}
                                        <span>
                                          {loadingAltsForId === paragraph.id ? 'Hledám...' : '+ Další'}
                                        </span>
                                      </button>
                                    </div>
                                    <div className="flex flex-wrap gap-1.5 pt-0.5">
                                      {currentAlts.map((alt, aIdx) => (
                                        <button
                                          key={`doc-alt-${paragraph.id}-${aIdx}`}
                                          type="button"
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            onUpdateParagraphText(paragraph.id, alt);
                                          }}
                                          className="px-2.5 py-1 rounded-lg text-xs font-semibold bg-white hover:bg-indigo-600 hover:text-white text-indigo-950 border border-indigo-300 shadow-xs transition-all active:scale-95 flex items-center gap-1"
                                          title={`Nahradit text touto variantou: "${alt}"`}
                                        >
                                          <span>{alt}</span>
                                        </button>
                                      ))}
                                    </div>
                                  </div>
                                );
                              })()}

                              {/* AI Notes / Length Fit Warnings Banner */}
                              {paragraph.isModified &&
                                paragraph.originalText.length > 20 &&
                                paragraph.currentText.length > paragraph.originalText.length * 1.25 && (
                                  <div className="mt-2 p-2 rounded-md bg-amber-50 border border-amber-200 text-xs text-amber-900 flex items-center justify-between gap-2">
                                    <span className="flex items-center gap-1.5 font-medium">
                                      <AlertTriangle className="w-3.5 h-3.5 text-amber-600 shrink-0" />
                                      {t('smartFitWarning', 'Text je o 25% delší než originál')}
                                    </span>
                                    <button
                                      type="button"
                                      onClick={() => onSinglePolish(paragraph)}
                                      className="px-2 py-0.5 rounded bg-amber-600 hover:bg-amber-700 text-white font-semibold text-[11px] shrink-0 shadow-xs"
                                    >
                                      {t('shortenWithAi', '⚡ Zhustit s AI')}
                                    </button>
                                  </div>
                                )}

                              {paragraph.aiNotes && (
                                <div className="mt-2 p-2 rounded-md bg-indigo-50/80 border border-indigo-200/80 text-xs text-indigo-900 flex items-start gap-2">
                                  <Lightbulb className="w-3.5 h-3.5 text-indigo-600 shrink-0 mt-0.5" />
                                  <span className="flex-1">{paragraph.aiNotes}</span>
                                </div>
                              )}

                              {paragraph.verificationSuggestion && (
                                <div className="mt-2 p-2.5 rounded-lg bg-teal-50 border border-teal-200 text-xs text-teal-900 flex items-start justify-between gap-2">
                                  <div className="flex items-start gap-2">
                                    <Sparkles className="w-3.5 h-3.5 text-teal-600 shrink-0 mt-0.5" />
                                    <div>
                                      <p className="font-semibold text-teal-950">
                                        {t('auditorSuggestion', 'Doporučení auditora:')}{' '}
                                        {paragraph.verificationIssue}
                                      </p>
                                      <p className="text-teal-800 mt-0.5 font-medium">
                                        „{paragraph.verificationSuggestion}“
                                      </p>
                                    </div>
                                  </div>
                                  <button
                                    type="button"
                                    onClick={() =>
                                      onUpdateParagraphText(
                                        paragraph.id,
                                        paragraph.verificationSuggestion!
                                      )
                                    }
                                    className="px-2 py-1 rounded bg-teal-600 hover:bg-teal-700 text-white font-semibold text-[11px] shrink-0 transition-colors shadow-xs"
                                  >
                                    {t('applySuggestion', 'Použít návrh')}
                                  </button>
                                </div>
                              )}

                              {/* Hyperlink Section */}
                              {paragraph.hyperlink?.url ? (
                                <div className="mt-2.5 p-2 rounded-lg bg-blue-50/90 border border-blue-200 text-xs flex items-center justify-between gap-2">
                                  <div className="flex items-center gap-1.5 min-w-0">
                                    <LinkIcon className="w-3.5 h-3.5 text-blue-600 shrink-0" />
                                    <span className="font-bold text-blue-950 text-[11px] shrink-0">
                                      {t('linkUnderText', 'Odkaz pod textem')}:
                                    </span>
                                    <a
                                      href={paragraph.hyperlink.url}
                                      target="_blank"
                                      rel="noreferrer"
                                      className="text-blue-700 underline truncate hover:text-blue-900 text-xs font-medium"
                                      title={paragraph.hyperlink.url}
                                    >
                                      {paragraph.hyperlink.url}
                                    </a>
                                  </div>
                                  <div className="flex items-center gap-1 shrink-0">
                                    <a
                                      href={paragraph.hyperlink.url}
                                      target="_blank"
                                      rel="noreferrer"
                                      className="p-1 rounded text-blue-600 hover:bg-blue-100 transition-colors"
                                      title={t("openLink", "Otevřít odkaz")}
                                    >
                                      <ExternalLink className="w-3 h-3" />
                                    </a>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        const newUrl = prompt(
                                          t('editUrlPrompt', 'Upravit URL odkazu:'),
                                          paragraph.hyperlink?.url
                                        );
                                        if (newUrl !== null) {
                                          onUpdateParagraphHyperlink(paragraph.id, newUrl.trim());
                                        }
                                      }}
                                      className="text-[10px] text-blue-700 font-semibold hover:underline px-1"
                                    >
                                      {t('editLink', 'Upravit')}
                                    </button>
                                  </div>
                                </div>
                              ) : (
                                <div className="mt-1 flex justify-end">
                                  <button
                                    type="button"
                                    onClick={() => {
                                      const url = prompt(
                                        t(
                                          'addUrlPrompt',
                                          'Zadejte URL adresu odkazu pod tento text (např. https://example.com):'
                                        )
                                      );
                                      if (url && url.trim()) {
                                        onUpdateParagraphHyperlink(paragraph.id, url.trim());
                                      }
                                    }}
                                    className="text-[10px] text-slate-400 hover:text-blue-600 flex items-center gap-1 transition-colors py-0.5"
                                  >
                                    <LinkIcon className="w-3 h-3" />
                                    <span>{t('addLink', '+ Připojit odkaz k textu')}</span>
                                  </button>
                                </div>
                              )}

                              {/* Footer metrics */}
                              <div className="mt-2 flex items-center justify-between text-[11px] text-slate-400">
                                <span>
                                  {paragraph.currentText.length} {t('chars', 'znaků')} •{' '}
                                  {
                                    paragraph.currentText
                                      .trim()
                                      .split(/\s+/)
                                      .filter(Boolean).length
                                  }{' '}
                                  {t('wordsCount', 'slov')}
                                </span>
                                {paragraph.runs[0]?.fontSizePt && (
                                  <span>
                                    ~{Math.round(paragraph.runs[0].fontSizePt)}pt
                                    {paragraph.runs[0].isBold ? ` • ${t('boldIndicator', 'Tučné')}` : ''}
                                  </span>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    ))
                  )}

                  {/* Speaker Notes */}
                  <div className="mt-6 pt-4 border-t border-slate-200">
                    <div className="flex items-center justify-between text-xs font-semibold text-slate-700 mb-2">
                      <span className="flex items-center gap-1.5">
                        <FileText className="w-3.5 h-3.5 text-slate-400" />
                        {t('speakerNotes', 'Poznámka pro řečníka')}
                      </span>
                      {slide.speakerNotes?.isModified && (
                        <span className="text-[10px] text-amber-700 font-semibold bg-amber-50 px-1.5 py-0.5 rounded border border-amber-200">
                          {t('notesModified', 'Poznámka upravena')}
                        </span>
                      )}
                    </div>
                    <textarea
                      value={slide.speakerNotes?.text || ''}
                      onChange={(e) => onUpdateSpeakerNotes(slide.slideNumber, e.target.value)}
                      placeholder={t(
                        'speakerNotesPlaceholder',
                        'Poznámky k tomuto snímku pro přednášejícího...'
                      )}
                      rows={2}
                      className="w-full text-xs p-2.5 rounded-lg border border-slate-200 bg-slate-50 text-slate-700 focus:bg-white focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-400 transition-colors"
                    />
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
};

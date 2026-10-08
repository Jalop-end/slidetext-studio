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
  Layers,
  Link as LinkIcon,
} from 'lucide-react';
import { PptxDocument, PptxParagraph } from '../types/pptx';
import { useLanguage } from '../context/LanguageContext';

interface DocumentEditorProps {
  document: PptxDocument;
  selectedParagraphId: string | null;
  onSelectParagraph: (paragraphId: string, slideNumber: number) => void;
  onUpdateParagraphText: (paragraphId: string, newText: string) => void;
  onResetParagraphText: (paragraphId: string) => void;
  currentSlideNumber: number;
  onSlideChange: (slideNumber: number) => void;
  onSingleTranslate: (paragraph: PptxParagraph) => void;
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
  onSinglePolish,
  onUpdateSpeakerNotes,
  onUpdateParagraphHyperlink,
}) => {
  const { t } = useLanguage();
  const [collapsedSlides, setCollapsedSlides] = useState<Record<number, boolean>>({});
  const [showDiffMap, setShowDiffMap] = useState<Record<string, boolean>>({});
  const itemRefs = useRef<Record<string, HTMLDivElement | null>>({});

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

  return (
    <div className="flex-1 overflow-y-auto bg-slate-100/70 p-4 sm:p-6 lg:p-8 font-sans">
      <div className="max-w-4xl mx-auto space-y-8">
        {/* Document Intro / Document Header */}
        <div className="bg-white rounded-xl p-5 shadow-xs border border-slate-200/80 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-lg bg-blue-50 border border-blue-200 flex items-center justify-center text-blue-600 shrink-0">
              <FileText className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-base font-bold text-slate-900">
                {t('docOverviewTitle', 'Textový přehled prezentace')}
              </h2>
              <p className="text-xs text-slate-500">
                {t(
                  'docOverviewSubtitle',
                  'Texty jsou uspořádány podle pořadí snímků a odstavců. Kliknutím na mini-odkaz se prvek zaměří na snímku.'
                )}
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2 text-xs text-slate-500 bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-200/60">
            <Layers className="w-3.5 h-3.5 text-slate-400" />
            <span>
              {t('total', 'Celkem')} <strong>{document.slides.length}</strong>{' '}
              {t('slidesCount', 'snímků')}
            </span>
          </div>
        </div>

        {/* Slides rendering as Word Pages */}
        {document.slides.map((slide) => {
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
                                  : shape.role === 'table-cell'
                                  ? 'bg-purple-100 text-purple-800 border border-purple-200'
                                  : 'bg-slate-100 text-slate-700 border border-slate-200'
                              }`}
                            >
                              {shape.role === 'title'
                                ? t('titleShape', 'Nadpis snímku')
                                : shape.role === 'subtitle'
                                ? t('subtitleShape', 'Podnadpis')
                                : shape.role === 'table-cell'
                                ? t('tableShape', 'Tabulka')
                                : t('bodyShape', 'Textové pole / Odrážky')}
                            </span>
                            <span className="text-slate-400 text-xs truncate">
                              {shape.name}
                            </span>
                          </div>
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
                                      title="Otevřít odkaz"
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
                                    {paragraph.runs[0].isBold ? ' • Tučné' : ''}
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

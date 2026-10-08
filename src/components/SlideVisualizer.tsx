import React, { useState, useEffect } from 'react';
import {
  ChevronLeft,
  ChevronRight,
  Maximize2,
  Minimize2,
  Tv,
  ExternalLink,
  Layers,
  LayoutGrid,
  RotateCw,
  Eye,
  Sliders,
  ZoomIn,
  ZoomOut,
  Sparkles,
  CheckCircle2,
  Pencil,
  PanelLeftClose,
  PanelLeftOpen,
  X,
  FileCheck,
  Download,
} from 'lucide-react';
import { PptxSlide, PptxParagraph, PptxShape } from '../types/pptx';
import { useLanguage } from '../context/LanguageContext';

interface SlideVisualizerProps {
  slides: PptxSlide[];
  currentSlideNumber: number;
  onSlideChange: (slideNumber: number) => void;
  selectedParagraphId: string | null;
  onSelectParagraph: (paragraphId: string, slideNumber: number) => void;
  onRefreshRender?: () => Promise<void>;
  isRendering?: boolean;
  isGalleryOpen?: boolean;
  onToggleGallery?: (open: boolean) => void;
  onExport?: () => void;
}

export const SlideVisualizer: React.FC<SlideVisualizerProps> = ({
  slides,
  currentSlideNumber,
  onSlideChange,
  selectedParagraphId,
  onSelectParagraph,
  onRefreshRender,
  isRendering = false,
  isGalleryOpen,
  onToggleGallery,
  onExport,
}) => {
  const { t } = useLanguage();
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [isProjectionMode, setIsProjectionMode] = useState(false);
  const [viewMode, setViewMode] = useState<'authentic' | 'overlay' | 'wireframe'>('overlay');
  const [isDeckOpen, setIsDeckOpen] = useState(true);
  const [localGridOpen, setLocalGridOpen] = useState(false);
  const [zoomLevel, setZoomLevel] = useState<number>(100);

  const isModalGridOpen = isGalleryOpen !== undefined ? isGalleryOpen : localGridOpen;
  const setModalGridOpen = (open: boolean) => {
    setLocalGridOpen(open);
    if (onToggleGallery) onToggleGallery(open);
  };

  const currentSlide =
    slides.find((s) => s.slideNumber === currentSlideNumber) || slides[0];

  const prevSlide = () => {
    if (currentSlideNumber > 1) {
      onSlideChange(currentSlideNumber - 1);
    }
  };

  const nextSlide = () => {
    if (currentSlideNumber < slides.length) {
      onSlideChange(currentSlideNumber + 1);
    }
  };

  // Keyboard navigation for projection mode
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (isProjectionMode || isFullscreen) {
        if (e.key === 'ArrowRight' || e.key === 'Space') {
          e.preventDefault();
          nextSlide();
        } else if (e.key === 'ArrowLeft') {
          e.preventDefault();
          prevSlide();
        } else if (e.key === 'Escape') {
          setIsProjectionMode(false);
          setIsFullscreen(false);
          setModalGridOpen(false);
        }
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [currentSlideNumber, isProjectionMode, isFullscreen, slides.length]);

  if (!currentSlide) {
    return (
      <div className="flex-1 flex items-center justify-center p-8 bg-slate-900 text-slate-400">
        Žádný snímek k zobrazení.
      </div>
    );
  }

  // Check if current slide has any modified paragraphs
  const currentSlideHasModifications = currentSlide.shapes.some((s) =>
    s.paragraphs.some((p) => p.isModified)
  );

  return (
    <div
      className={`flex flex-col bg-slate-900 border-l border-slate-800 transition-all ${
        isProjectionMode || isFullscreen
          ? 'fixed inset-0 z-50 bg-slate-950 w-screen h-screen'
          : 'h-full w-full flex-1 overflow-hidden'
      }`}
    >
      {/* Top Visualizer Toolbar */}
      <div className="w-full px-3 sm:px-4 py-2 bg-slate-900/95 border-b border-slate-800 flex items-center justify-between text-xs text-slate-300 shrink-0 gap-2">
        {/* Left Toolbar Section */}
        <div className="flex items-center gap-2">
          {/* Deck Toggle Button */}
          {!isProjectionMode && !isFullscreen && (
            <button
              onClick={() => setIsDeckOpen(!isDeckOpen)}
              className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
              title={isDeckOpen ? 'Skrýt pás snímků' : 'Zobrazit pás snímků'}
            >
              {isDeckOpen ? (
                <PanelLeftClose className="w-4 h-4 text-blue-400" />
              ) : (
                <PanelLeftOpen className="w-4 h-4" />
              )}
            </button>
          )}

          <div className="flex items-center gap-1.5">
            <span className="w-2 h-2 rounded-full bg-orange-500 animate-pulse" />
            <span className="font-bold text-slate-200">
              {t('slide')} #{currentSlide.slideNumber}
            </span>
            <span className="text-slate-500">/ {slides.length}</span>
          </div>

          {/* Grid Overview Modal Trigger */}
          <button
            onClick={() => setModalGridOpen(true)}
            className="hidden sm:inline-flex items-center gap-1.5 px-2 py-1 rounded-md bg-blue-950/60 hover:bg-blue-900/60 text-blue-300 hover:text-white transition-colors border border-blue-800/80 font-medium"
            title="Zobrazit celou prezentaci (všechny snímky najednou s šablonami)"
          >
            <LayoutGrid className="w-3.5 h-3.5 text-blue-400" />
            <span>Celá prezentace</span>
          </button>
        </div>

        {/* Center / Right Toolbar Section: View Modes & Actions */}
        <div className="flex items-center gap-1.5 sm:gap-2">
          {/* View Mode Selector */}
          <div className="flex items-center bg-slate-950/80 p-0.5 rounded-lg border border-slate-800">
            <button
              onClick={() => setViewMode('authentic')}
              className={`px-2 py-1 rounded text-[11px] font-semibold transition-all flex items-center gap-1 ${
                viewMode === 'authentic'
                  ? 'bg-blue-600 text-white shadow-xs'
                  : 'text-slate-400 hover:text-slate-200'
              }`}
              title="Věrný náhled z PowerPointu (originální vzhled a šablona)"
            >
              <Eye className="w-3 h-3" />
              <span className="hidden md:inline">PowerPoint vzhled</span>
            </button>
            <button
              onClick={() => setViewMode('overlay')}
              className={`px-2 py-1 rounded text-[11px] font-semibold transition-all flex items-center gap-1 ${
                viewMode === 'overlay'
                  ? 'bg-blue-600 text-white shadow-xs'
                  : 'text-slate-400 hover:text-slate-200'
              }`}
              title="Interaktivní náhled se zobrazením upravených textů"
            >
              <Sparkles className="w-3 h-3 text-amber-300" />
              <span className="hidden md:inline">S úpravami</span>
            </button>
            <button
              onClick={() => setViewMode('wireframe')}
              className={`px-2 py-1 rounded text-[11px] font-semibold transition-all flex items-center gap-1 ${
                viewMode === 'wireframe'
                  ? 'bg-blue-600 text-white shadow-xs'
                  : 'text-slate-400 hover:text-slate-200'
              }`}
              title="Drátěný model rozvržení prvků"
            >
              <Sliders className="w-3 h-3" />
              <span className="hidden xl:inline">Drátěný model</span>
            </button>
          </div>

          {/* Re-render via PowerPoint COM Button */}
          {onRefreshRender && (
            <button
              onClick={onRefreshRender}
              disabled={isRendering}
              className="px-2.5 py-1 rounded-md text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-200 hover:text-white transition-colors border border-slate-700 flex items-center gap-1.5 disabled:opacity-50"
              title="Vygenerovat čerstvý PowerPoint náhled se všemi textovými úpravami"
            >
              <RotateCw className={`w-3.5 h-3.5 text-emerald-400 ${isRendering ? 'animate-spin' : ''}`} />
              <span className="hidden sm:inline">
                {isRendering ? 'Renderuji...' : 'Aktualizovat PowerPoint'}
              </span>
            </button>
          )}

          {/* Zoom controls */}
          <div className="hidden lg:flex items-center bg-slate-950/80 rounded-lg border border-slate-800 p-0.5">
            <button
              onClick={() => setZoomLevel((z) => Math.max(50, z - 15))}
              className="p-1 rounded text-slate-400 hover:text-white hover:bg-slate-800"
              title="Oddálit"
            >
              <ZoomOut className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => setZoomLevel(100)}
              className="px-1.5 text-[10px] font-mono text-slate-300 hover:text-white"
              title="Obnovit 100%"
            >
              {zoomLevel}%
            </button>
            <button
              onClick={() => setZoomLevel((z) => Math.min(175, z + 15))}
              className="p-1 rounded text-slate-400 hover:text-white hover:bg-slate-800"
              title="Přiblížit"
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Projection / Slideshow Mode Toggle */}
          <button
            onClick={() => setIsProjectionMode(!isProjectionMode)}
            className={`px-2 sm:px-2.5 py-1 rounded-md text-xs font-semibold transition-all flex items-center gap-1.5 ${
              isProjectionMode
                ? 'bg-orange-600 text-white shadow-xs'
                : 'text-orange-400 hover:bg-orange-950/60 bg-orange-950/30 border border-orange-800/60'
            }`}
            title="Spustit režim promítání na celou obrazovku"
          >
            <Tv className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">{isProjectionMode ? 'Ukončit (Esc)' : t('projectionMode')}</span>
          </button>

          {/* Fullscreen toggle */}
          <button
            onClick={() => setIsFullscreen(!isFullscreen)}
            className="p-1 rounded-md text-slate-400 hover:bg-slate-800 hover:text-white transition-colors"
            title={isFullscreen ? 'Zmenšit' : t('fullscreen')}
          >
            {isFullscreen ? <Minimize2 className="w-4 h-4" /> : <Maximize2 className="w-4 h-4" />}
          </button>
        </div>
      </div>

      {/* Main Workspace: Left Deck Sidebar + Central Slide Canvas */}
      <div className="flex-1 w-full h-full min-h-0 min-w-0 flex overflow-hidden relative">
        {/* LEFT DECK NAVIGATOR (Slide Deck Strip) */}
        {!isProjectionMode && !isFullscreen && isDeckOpen && (
          <aside className="w-56 sm:w-64 bg-slate-950/90 border-r border-slate-800 flex flex-col shrink-0 overflow-hidden">
            {/* Deck Header */}
            <div className="p-3 border-b border-slate-800/80 flex items-center justify-between text-xs">
              <div className="flex items-center gap-1.5 font-bold text-slate-200">
                <Layers className="w-4 h-4 text-blue-400" />
                <span>Snímky ({slides.length})</span>
              </div>
              <button
                onClick={() => setModalGridOpen(true)}
                className="text-[11px] text-blue-400 hover:text-blue-300 font-medium"
              >
                Mřížka
              </button>
            </div>

            {/* Scrollable Slide List */}
            <div className="flex-1 overflow-y-auto p-2.5 space-y-2.5 custom-scrollbar">
              {slides.map((s) => {
                const isSelected = s.slideNumber === currentSlideNumber;
                const hasModified = s.shapes.some((sh) =>
                  sh.paragraphs.some((p) => p.isModified)
                );

                return (
                  <div
                    key={`deck-item-${s.slideNumber}`}
                    onClick={() => onSlideChange(s.slideNumber)}
                    className={`group relative rounded-xl p-2 cursor-pointer transition-all border ${
                      isSelected
                        ? 'bg-blue-950/40 border-blue-500 shadow-md ring-1 ring-blue-500/50'
                        : 'bg-slate-900/60 border-slate-800 hover:border-slate-700 hover:bg-slate-800/60'
                    }`}
                  >
                    {/* Thumbnail Viewport (16:9 ratio) */}
                    <div className="relative w-full aspect-[16/9] rounded-lg overflow-hidden bg-slate-800 shadow-xs border border-slate-800/80">
                      {s.renderedImageDataUrl ? (
                        <img
                          src={s.renderedImageDataUrl}
                          alt={`Snímek ${s.slideNumber}`}
                          className="w-full h-full object-cover select-none pointer-events-none"
                          loading="lazy"
                        />
                      ) : (
                        <div
                          className="w-full h-full flex flex-col items-center justify-center p-2 text-center"
                          style={{ backgroundColor: s.backgroundColor || '#ffffff' }}
                        >
                          <span className="text-[10px] font-bold text-slate-800 line-clamp-2">
                            {s.title}
                          </span>
                        </div>
                      )}

                      {/* Number Badge */}
                      <span className="absolute bottom-1 left-1 px-1.5 py-0.5 rounded bg-black/70 backdrop-blur-xs text-[10px] font-mono font-bold text-white shadow-xs">
                        #{s.slideNumber}
                      </span>

                      {/* Modification indicator badge */}
                      {hasModified && (
                        <span className="absolute top-1 right-1 px-1.5 py-0.5 rounded-full bg-emerald-500/90 text-white text-[9px] font-bold shadow-md flex items-center gap-0.5">
                          <CheckCircle2 className="w-2.5 h-2.5" />
                          Upraveno
                        </span>
                      )}
                    </div>

                    {/* Slide Caption */}
                    <div className="mt-1.5 flex items-center justify-between text-[11px]">
                      <span
                        className={`truncate font-semibold ${
                          isSelected ? 'text-blue-300' : 'text-slate-300 group-hover:text-white'
                        }`}
                        title={s.title}
                      >
                        {s.title || `Snímek ${s.slideNumber}`}
                      </span>
                    </div>
                  </div>
                );
              })}
            </div>
          </aside>
        )}

        {/* CENTER SLIDE CANVAS AREA */}
        <div className="flex-1 w-full h-full min-h-0 min-w-0 overflow-auto flex items-center justify-center p-3 sm:p-6 bg-slate-950/90 relative">
          {/* Floating PowerPoint Render Status Banner */}
          {isRendering && (
            <div className="absolute top-4 left-1/2 -translate-x-1/2 z-40 bg-slate-900/95 backdrop-blur-md border border-blue-500/60 px-4 py-2 rounded-full shadow-2xl flex items-center gap-2.5 text-xs text-blue-200 pointer-events-none animate-pulse">
              <RotateCw className="w-4 h-4 text-blue-400 animate-spin" />
              <span className="font-semibold">Načítám věrný náhled z Microsoft PowerPointu (šablona a grafika)...</span>
            </div>
          )}

          {/* Slide Stage Container with Authentic 16:9 Aspect Ratio */}
          <div
            className={`relative w-full max-w-5xl shadow-2xl rounded-xl overflow-hidden select-none transition-transform ${
              isProjectionMode
                ? 'ring-4 ring-orange-500/30 max-w-6xl shadow-orange-950/50'
                : 'ring-1 ring-slate-800'
            }`}
            style={{
              backgroundColor: currentSlide.backgroundColor || '#ffffff',
              aspectRatio: '16 / 9',
              transform: zoomLevel !== 100 ? `scale(${zoomLevel / 100})` : undefined,
              transformOrigin: 'center center',
            }}
          >
            {/* Structural 16:9 spacer ensuring container never collapses */}
            <svg
              viewBox="0 0 1920 1080"
              className="w-full h-auto block pointer-events-none invisible select-none"
              aria-hidden="true"
            />
            {/* 1. True PowerPoint Full HD Slide Image Background */}
            {currentSlide.renderedImageDataUrl ? (
              <img
                src={currentSlide.renderedImageDataUrl}
                alt={`Snímek ${currentSlide.slideNumber} - PowerPoint náhled`}
                className="absolute inset-0 w-full h-full object-contain pointer-events-none select-none z-0"
              />
            ) : currentSlide.backgroundImageDataUrl ? (
              <img
                src={currentSlide.backgroundImageDataUrl}
                alt="Slide background"
                className="absolute inset-0 w-full h-full object-cover pointer-events-none select-none z-0"
              />
            ) : null}

            {/* 2. Embedded images (fallback if PowerPoint COM render is not available) */}
            {!currentSlide.renderedImageDataUrl &&
              currentSlide.images?.map((img) => (
                <img
                  key={img.id}
                  src={img.dataUrl}
                  alt={img.name}
                  className="absolute object-contain pointer-events-none select-none z-10 drop-shadow-sm"
                  style={{
                    left: `${img.bounds.x}%`,
                    top: `${img.bounds.y}%`,
                    width: `${img.bounds.width}%`,
                    height: `${img.bounds.height}%`,
                  }}
                />
              ))}

            {/* 3. Interactive Overlays & Modified Text Replacements */}
            {currentSlide.shapes.map((shape) => {
              const hasSelectedParagraph = shape.paragraphs.some(
                (p) => p.id === selectedParagraphId
              );
              const hasModifiedParagraph = shape.paragraphs.some((p) => p.isModified);

              // In 'authentic' pure view mode: only show selection highlight if a paragraph is actively clicked
              if (viewMode === 'authentic' && !hasSelectedParagraph) {
                return null;
              }

              // In 'overlay' mode: show modified texts and interactive boxes
              // If not modified and no rendered image is available, render standard text so fallback works
              const shouldRenderText =
                viewMode === 'wireframe' ||
                !currentSlide.renderedImageDataUrl ||
                hasModifiedParagraph ||
                hasSelectedParagraph;

              return (
                <div
                  key={shape.id}
                  onClick={() => {
                    if (shape.paragraphs.length > 0) {
                      onSelectParagraph(shape.paragraphs[0].id, currentSlide.slideNumber);
                    }
                  }}
                  className={`absolute cursor-pointer transition-all duration-150 rounded ${
                    hasSelectedParagraph && !isProjectionMode
                      ? 'ring-2 ring-blue-500 bg-blue-500/10 shadow-lg z-30'
                      : hasModifiedParagraph && viewMode === 'overlay'
                      ? 'ring-1 ring-emerald-500/70 bg-white/95 dark:bg-slate-900/95 shadow-md z-20 backdrop-blur-xs'
                      : viewMode === 'wireframe'
                      ? 'border border-dashed border-slate-300 hover:border-blue-400 hover:bg-blue-50/20 z-10'
                      : 'hover:ring-1 hover:ring-blue-300/40 z-10'
                  }`}
                  style={{
                    left: `${shape.bounds.x}%`,
                    top: `${shape.bounds.y}%`,
                    width: `${shape.bounds.width}%`,
                    height: `${shape.bounds.height}%`,
                    backgroundColor:
                      viewMode === 'wireframe'
                        ? shape.backgroundColor || undefined
                        : hasModifiedParagraph
                        ? undefined
                        : undefined,
                    borderColor: viewMode === 'wireframe' ? shape.borderColor : undefined,
                    borderWidth: viewMode === 'wireframe' && shape.borderColor ? '1px' : undefined,
                  }}
                >
                  {/* Selection badge */}
                  {hasSelectedParagraph && !isProjectionMode && (
                    <div className="absolute -top-3.5 left-2 bg-blue-600 text-white font-bold text-[9px] uppercase px-1.5 py-0.5 rounded shadow-md tracking-wider flex items-center gap-1 z-40 pointer-events-none">
                      <span className="w-1.5 h-1.5 rounded-full bg-white animate-ping"></span>
                      Vybráno k editaci
                    </div>
                  )}

                  {/* Modified text indicator badge */}
                  {hasModifiedParagraph && viewMode === 'overlay' && !hasSelectedParagraph && (
                    <div className="absolute -top-3 left-2 bg-emerald-600 text-white font-semibold text-[9px] px-1.5 py-0.2 rounded shadow-xs flex items-center gap-1 z-30 pointer-events-none">
                      <Pencil className="w-2.5 h-2.5" />
                      Upravený text
                    </div>
                  )}

                  {/* Render Paragraphs inside the shape */}
                  {shouldRenderText && (
                    <div className="h-full w-full overflow-hidden flex flex-col justify-start p-1 sm:p-1.5">
                      {shape.paragraphs.map((p) => {
                        const isParaSelected = p.id === selectedParagraphId;

                        return (
                          <div
                            key={p.id}
                            onClick={(e) => {
                              e.stopPropagation();
                              onSelectParagraph(p.id, currentSlide.slideNumber);
                            }}
                            className={`transition-colors rounded px-1 my-0.5 ${
                              isParaSelected && !isProjectionMode
                                ? 'bg-blue-100/90 text-blue-950 font-medium'
                                : p.isModified && viewMode === 'overlay'
                                ? 'text-slate-900 dark:text-slate-100 font-medium'
                                : ''
                            }`}
                            style={{
                              textAlign: p.alignment || 'left',
                              marginLeft: p.bulletLevel ? `${p.bulletLevel * 1.5}rem` : undefined,
                            }}
                          >
                            <p
                              className={`break-words leading-snug ${
                                p.shapeRole === 'title'
                                  ? 'text-slate-900 font-extrabold text-xs sm:text-base md:text-xl tracking-tight'
                                  : p.shapeRole === 'subtitle'
                                  ? 'text-slate-600 font-semibold text-[11px] sm:text-sm md:text-base'
                                  : 'text-slate-700 text-[10px] sm:text-xs md:text-sm'
                              }`}
                              style={{
                                color: p.colorHex || undefined,
                              }}
                            >
                              {p.bulletLevel !== undefined && p.shapeRole === 'body' && (
                                <span className="mr-1.5 text-slate-400 font-bold">•</span>
                              )}
                              {p.currentText}
                              {p.hyperlink?.url && (
                                <span className="inline-flex items-center gap-0.5 ml-1.5 text-blue-600 text-[10px] underline font-medium">
                                  🔗 {p.hyperlink.url}
                                </span>
                              )}
                            </p>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* Slide Navigation Footer */}
      <div className="w-full px-4 py-2.5 bg-slate-900 border-t border-slate-800 flex items-center justify-between gap-4 shrink-0">
        <div className="flex items-center gap-2">
          <button
            onClick={prevSlide}
            disabled={currentSlideNumber <= 1}
            className="p-1.5 rounded-lg text-slate-300 hover:text-white hover:bg-slate-800 disabled:opacity-30 disabled:pointer-events-none transition-colors"
            title={t('prev')}
          >
            <ChevronLeft className="w-5 h-5" />
          </button>

          <span className="text-xs font-semibold text-slate-200">
            {t('slide')} {currentSlide.slideNumber} {t('of')} {slides.length}
          </span>

          <button
            onClick={nextSlide}
            disabled={currentSlideNumber >= slides.length}
            className="p-1.5 rounded-lg text-slate-300 hover:text-white hover:bg-slate-800 disabled:opacity-30 disabled:pointer-events-none transition-colors"
            title={t('next')}
          >
            <ChevronRight className="w-5 h-5" />
          </button>
        </div>

        {/* Slide status & fidelity badge */}
        <div className="hidden sm:flex items-center gap-2 text-xs text-slate-400">
          {currentSlide.renderedImageDataUrl ? (
            <span className="inline-flex items-center gap-1 text-[11px] text-emerald-400 bg-emerald-950/60 px-2 py-0.5 rounded-full border border-emerald-800">
              <FileCheck className="w-3.5 h-3.5" />
              100% PowerPoint věrný náhled (1920×1080)
            </span>
          ) : (
            <span className="text-[11px] text-slate-500">
              Načítání věrného náhledu...
            </span>
          )}
        </div>

        {/* Quick Slide Numbers Strip */}
        <div className="flex items-center gap-1 overflow-x-auto max-w-xs sm:max-w-md py-1">
          {slides.map((s) => {
            const isSelected = s.slideNumber === currentSlideNumber;
            const hasModified = s.shapes.some((sh) => sh.paragraphs.some((p) => p.isModified));

            return (
              <button
                key={`footer-strip-${s.slideNumber}`}
                onClick={() => onSlideChange(s.slideNumber)}
                className={`relative px-2 py-1 rounded-md text-[11px] font-bold transition-all shrink-0 ${
                  isSelected
                    ? 'bg-blue-600 text-white shadow-xs'
                    : 'bg-slate-800 text-slate-400 hover:bg-slate-700 hover:text-white'
                }`}
                title={`Přejít na snímek ${s.slideNumber}`}
              >
                {s.slideNumber}
                {hasModified && (
                  <span className="absolute -top-1 -right-1 w-2 h-2 rounded-full bg-emerald-400 ring-2 ring-slate-900" />
                )}
              </button>
            );
          })}
        </div>
      </div>

      {/* FULL PRESENTATION GALLERY MODAL ("Celá prezentace") */}
      {isModalGridOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-6 bg-slate-950/85 backdrop-blur-sm animate-in fade-in duration-150">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl w-full max-w-6xl max-h-[92vh] flex flex-col overflow-hidden text-slate-100">
            {/* Modal Header */}
            <div className="px-6 py-4 bg-slate-900 border-b border-slate-800 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-xl bg-gradient-to-tr from-blue-600 to-indigo-600 text-white flex items-center justify-center shadow-md shadow-blue-500/20">
                  <LayoutGrid className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-bold text-white text-base">
                    Přehled celé prezentace ({slides.length} snímků)
                  </h3>
                  <p className="text-xs text-slate-400">
                    Prohlédněte si všechny snímky, jejich šablony, grafiku a rozvržení před stažením.
                  </p>
                </div>
              </div>
              <button
                onClick={() => setModalGridOpen(false)}
                className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Rendering Info Banner */}
            {isRendering && (
              <div className="mx-6 mt-4 p-3 rounded-xl bg-blue-950/80 border border-blue-500/50 text-xs text-blue-200 flex items-center gap-2.5 shadow-lg shrink-0">
                <RotateCw className="w-4 h-4 text-blue-400 animate-spin shrink-0" />
                <span>Právě probíhá export věrných náhledů z Microsoft PowerPointu (Full HD 1920x1080). Náhledy se na této stránce automaticky aktualizují.</span>
              </div>
            )}

            {/* Grid of All Slides */}
            <div className="flex-1 overflow-y-auto p-6 grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-5 bg-slate-950/60 custom-scrollbar">
              {slides.map((s) => {
                const isSelected = s.slideNumber === currentSlideNumber;
                const hasModified = s.shapes.some((sh) =>
                  sh.paragraphs.some((p) => p.isModified)
                );

                return (
                  <div
                    key={`grid-modal-slide-${s.slideNumber}`}
                    onClick={() => {
                      onSlideChange(s.slideNumber);
                      setModalGridOpen(false);
                    }}
                    className={`group flex flex-col bg-slate-900 rounded-xl overflow-hidden border cursor-pointer transition-all hover:scale-[1.02] ${
                      isSelected
                        ? 'border-blue-500 shadow-xl shadow-blue-500/10 ring-2 ring-blue-500'
                        : 'border-slate-800 hover:border-slate-700 hover:bg-slate-850'
                    }`}
                  >
                    {/* Thumbnail */}
                    <div className="relative w-full aspect-[16/9] bg-slate-800 overflow-hidden">
                      {s.renderedImageDataUrl ? (
                        <img
                          src={s.renderedImageDataUrl}
                          alt={s.title}
                          className="w-full h-full object-cover select-none pointer-events-none"
                          loading="lazy"
                        />
                      ) : (
                        <div
                          className="w-full h-full flex flex-col items-center justify-center p-3 text-center bg-slate-800/90"
                          style={{ backgroundColor: s.backgroundColor || undefined }}
                        >
                          <div className="w-7 h-7 rounded-lg bg-blue-900/60 flex items-center justify-center mb-1.5 text-blue-300">
                            <RotateCw className="w-4 h-4 animate-spin" />
                          </div>
                          <span className="text-xs font-bold text-slate-200 line-clamp-2">
                            {s.title}
                          </span>
                          <span className="text-[10px] text-slate-400 mt-1">Generuji náhled...</span>
                        </div>
                      )}

                      <span className="absolute bottom-2 left-2 px-2 py-0.5 rounded bg-black/75 backdrop-blur-xs text-xs font-mono font-bold text-white shadow-md">
                        #{s.slideNumber}
                      </span>

                      {hasModified && (
                        <span className="absolute top-2 right-2 px-2 py-0.5 rounded-full bg-emerald-500 text-white text-[10px] font-bold shadow-md flex items-center gap-1">
                          <CheckCircle2 className="w-3 h-3" />
                          Upraveno
                        </span>
                      )}
                    </div>

                    {/* Meta info */}
                    <div className="p-3 flex items-center justify-between">
                      <span className="text-xs font-semibold text-slate-200 group-hover:text-blue-400 truncate">
                        {s.title || `Snímek ${s.slideNumber}`}
                      </span>
                      <button className="text-[11px] text-blue-400 font-medium group-hover:underline">
                        Otevřít
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* Modal Footer */}
            <div className="px-6 py-3.5 bg-slate-900 border-t border-slate-800 flex items-center justify-between text-xs text-slate-400 flex-wrap gap-3">
              <span>Kliknutím na snímek jej otevřete v editoru. Můžete si prohlédnout celou prezentaci nebo ji rovnou stáhnout.</span>
              <div className="flex items-center gap-2">
                {onExport && (
                  <button
                    onClick={() => {
                      setModalGridOpen(false);
                      onExport();
                    }}
                    className="px-3.5 py-1.5 rounded-lg bg-gradient-to-r from-orange-600 to-amber-600 hover:from-orange-700 hover:to-amber-700 text-white font-bold flex items-center gap-1.5 shadow-md shadow-orange-600/20 transition-all"
                    title="Stáhnout upravenou prezentaci (.pptx)"
                  >
                    <Download className="w-3.5 h-3.5" />
                    <span>Stáhnout prezentaci (.pptx)</span>
                  </button>
                )}
                <button
                  onClick={() => {
                    setModalGridOpen(false);
                    setIsProjectionMode(true);
                  }}
                  className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-orange-400 border border-orange-800/60 font-semibold flex items-center gap-1.5 transition-colors"
                  title="Promítat prezentaci na celou obrazovku"
                >
                  <Tv className="w-3.5 h-3.5" />
                  <span>Promítat</span>
                </button>
                <button
                  onClick={() => setModalGridOpen(false)}
                  className="px-4 py-1.5 rounded-lg bg-blue-600 text-white font-semibold hover:bg-blue-500 transition-colors"
                >
                  Zavřít přehled
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

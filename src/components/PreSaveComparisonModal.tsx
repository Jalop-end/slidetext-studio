import React, { useState, useEffect } from 'react';
import {
  X,
  Download,
  ChevronLeft,
  ChevronRight,
  ShieldCheck,
  CheckCircle2,
  AlertTriangle,
  ExternalLink,
  Layers,
  Sparkles,
  Palette,
  Image as ImageIcon,
  Link as LinkIcon,
  Maximize2,
  Check,
  Pencil,
} from 'lucide-react';
import { PptxDocument, PptxSlide } from '../types/pptx';
import { useLanguage } from '../context/LanguageContext';

interface PreSaveComparisonModalProps {
  isOpen: boolean;
  onClose: () => void;
  document: PptxDocument;
  onConfirmExport: (customFileName?: string) => void;
  isExporting: boolean;
}

export const PreSaveComparisonModal: React.FC<PreSaveComparisonModalProps> = ({
  isOpen,
  onClose,
  document,
  onConfirmExport,
  isExporting,
}) => {
  const { t } = useLanguage();
  const [currentSlideIndex, setCurrentSlideIndex] = useState(0);
  const [outputFileName, setOutputFileName] = useState(
    document.fileName.replace(/\.pptx$/i, '') + '_upraveno'
  );

  useEffect(() => {
    setOutputFileName(document.fileName.replace(/\.pptx$/i, '') + '_upraveno');
  }, [document.fileName]);

  if (!isOpen) return null;

  const currentSlide = document.slides[currentSlideIndex] || document.slides[0];

  const prevSlide = () => {
    if (currentSlideIndex > 0) setCurrentSlideIndex(currentSlideIndex - 1);
  };

  const nextSlide = () => {
    if (currentSlideIndex < document.slides.length - 1)
      setCurrentSlideIndex(currentSlideIndex + 1);
  };

  // Collect stats on links, images and modifications
  let totalLinks = 0;
  let totalImages = 0;
  let modifiedParas = 0;

  document.slides.forEach((s) => {
    totalImages += s.images?.length || 0;
    s.shapes.forEach((sh) => {
      sh.paragraphs.forEach((p) => {
        if (p.hyperlink?.url) totalLinks++;
        if (p.isModified) modifiedParas++;
      });
    });
  });

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-2 sm:p-4 bg-slate-950/80 backdrop-blur-sm animate-in fade-in duration-150">
      <div className="bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl w-full max-w-7xl h-[95vh] flex flex-col overflow-hidden text-slate-100">
        {/* Top Header */}
        <div className="px-6 py-3.5 bg-slate-900/95 border-b border-slate-800 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-xl bg-orange-600 text-white flex items-center justify-center shadow-md shadow-orange-600/20">
              <Layers className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="font-bold text-white text-base">
                  {t('preSaveTitle')}
                </h3>
                <span className="text-[11px] font-semibold px-2 py-0.5 rounded-full bg-slate-800 text-slate-300 border border-slate-700">
                  {t('slide')} {currentSlide.slideNumber} {t('of')} {document.slides.length}
                </span>
              </div>
              <p className="text-xs text-slate-400">
                {t('preSaveSubtitle')}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-3">
            {/* Quick Audit Pills */}
            <div className="hidden lg:flex items-center gap-2 bg-slate-800/80 px-3 py-1.5 rounded-lg border border-slate-700/80 text-xs">
              <span className="flex items-center gap-1 text-emerald-400 font-medium">
                <CheckCircle2 className="w-3.5 h-3.5" />
                {modifiedParas} {t('textsModified', 'textů upraveno')}
              </span>
              <span className="text-slate-600">•</span>
              <span className="flex items-center gap-1 text-blue-400 font-medium">
                <LinkIcon className="w-3.5 h-3.5" />
                {totalLinks} {t('linksPreserved', 'odkazů zachováno')}
              </span>
              <span className="text-slate-600">•</span>
              <span className="flex items-center gap-1 text-amber-400 font-medium">
                <ImageIcon className="w-3.5 h-3.5" />
                {totalImages} {t('imagesPositioned', 'obrázků na pozicích')}
              </span>
            </div>

            <button
              onClick={onClose}
              className="p-1.5 rounded-lg text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Dual Window Split View */}
        <div className="flex-1 overflow-auto p-4 sm:p-6 grid grid-cols-1 lg:grid-cols-2 gap-6 bg-slate-950/60">
          {/* WINDOW 1: ORIGINAL SLIDE */}
          <div className="flex flex-col bg-slate-900 rounded-xl border border-slate-800 overflow-hidden shadow-lg">
            <div className="px-4 py-2.5 bg-slate-800/90 border-b border-slate-700/80 flex items-center justify-between text-xs">
              <div className="flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full bg-slate-500" />
                <span className="font-bold text-slate-200">
                  {t('window1Title')}
                </span>
              </div>
              <span className="text-[11px] text-slate-400">{t('originalState')}</span>
            </div>

            <div className="flex-1 flex items-center justify-center p-3 sm:p-4 bg-slate-950/40 overflow-hidden">
              {/* Slide Canvas - Original */}
              <div
                className="relative w-full aspect-[16/9] bg-white rounded-lg shadow-xl overflow-hidden select-none"
                style={{ backgroundColor: currentSlide.backgroundColor || '#ffffff' }}
              >
                {/* Authentic PowerPoint Render Image if available */}
                {(currentSlide.originalRenderedImageDataUrl || currentSlide.renderedImageDataUrl) ? (
                  <img
                    src={currentSlide.originalRenderedImageDataUrl || currentSlide.renderedImageDataUrl}
                    alt={`Původní snímek ${currentSlide.slideNumber} - PowerPoint`}
                    className="absolute inset-0 w-full h-full object-contain pointer-events-none select-none z-0"
                  />
                ) : (
                  <>
                    {/* Embedded Images */}
                    {currentSlide.images?.map((img) => (
                      <img
                        key={`orig-img-${img.id}`}
                        src={img.dataUrl}
                        alt={img.name}
                        className="absolute object-contain pointer-events-none select-none"
                        style={{
                          left: `${img.bounds.x}%`,
                          top: `${img.bounds.y}%`,
                          width: `${img.bounds.width}%`,
                          height: `${img.bounds.height}%`,
                        }}
                      />
                    ))}

                    {/* Shapes with Original Text */}
                    {currentSlide.shapes.map((shape) => (
                      <div
                        key={`orig-sh-${shape.id}`}
                        className="absolute overflow-hidden p-1.5 flex flex-col justify-start rounded"
                        style={{
                          left: `${shape.bounds.x}%`,
                          top: `${shape.bounds.y}%`,
                          width: `${shape.bounds.width}%`,
                          height: `${shape.bounds.height}%`,
                          backgroundColor: shape.backgroundColor || undefined,
                          borderColor: shape.borderColor || undefined,
                          borderWidth: shape.borderColor ? '1px' : undefined,
                          borderStyle: shape.borderColor ? 'solid' : undefined,
                        }}
                      >
                        {shape.paragraphs.map((p) => (
                          <div
                            key={`orig-p-${p.id}`}
                            className="my-0.5"
                            style={{
                              textAlign: p.alignment || 'left',
                              marginLeft: p.bulletLevel ? `${p.bulletLevel * 1.5}rem` : undefined,
                            }}
                          >
                            <p
                              className={`break-words leading-snug ${
                                p.shapeRole === 'title'
                                  ? 'text-slate-900 font-extrabold text-xs sm:text-base md:text-lg tracking-tight'
                                  : p.shapeRole === 'subtitle'
                                  ? 'text-slate-600 font-semibold text-[11px] sm:text-sm'
                                  : 'text-slate-700 text-[10px] sm:text-xs md:text-sm'
                              }`}
                              style={{
                                color: p.colorHex || undefined,
                              }}
                            >
                              {p.bulletLevel !== undefined && p.shapeRole === 'body' && (
                                <span className="mr-1.5 text-slate-400">•</span>
                              )}
                              {p.originalText}
                            </p>
                            {p.hyperlink?.url && (
                              <span className="inline-flex items-center gap-1 text-[9px] text-blue-600 underline font-medium mt-0.5">
                                <LinkIcon className="w-2.5 h-2.5" />
                                {p.hyperlink.url}
                              </span>
                            )}
                          </div>
                        ))}
                      </div>
                    ))}
                  </>
                )}
              </div>
            </div>
          </div>

          {/* WINDOW 2: NEW UPDATED SLIDE */}
          <div className="flex flex-col bg-slate-900 rounded-xl border border-blue-900/60 ring-2 ring-blue-500/30 overflow-hidden shadow-lg">
            <div className="px-4 py-2.5 bg-blue-950/70 border-b border-blue-900/80 flex items-center justify-between text-xs">
              <div className="flex items-center gap-2">
                <span className="w-2.5 h-2.5 rounded-full bg-blue-500 animate-pulse" />
                <span className="font-bold text-blue-200">
                  {t('window2Title')}
                </span>
              </div>
              <span className="text-[11px] font-semibold text-emerald-400 bg-emerald-950/80 px-2 py-0.5 rounded border border-emerald-800">
                OK
              </span>
            </div>

            <div className="flex-1 flex items-center justify-center p-3 sm:p-4 bg-slate-950/40 overflow-hidden">
              {/* Slide Canvas - Updated */}
              <div
                className="relative w-full aspect-[16/9] bg-white rounded-lg shadow-xl overflow-hidden select-none"
                style={{ backgroundColor: currentSlide.backgroundColor || '#ffffff' }}
              >
                {/* 1. Authentic PowerPoint Canvas Background */}
                {currentSlide.renderedImageDataUrl ? (
                  <img
                    src={currentSlide.renderedImageDataUrl}
                    alt={`Nový snímek ${currentSlide.slideNumber} - PowerPoint`}
                    className="absolute inset-0 w-full h-full object-contain pointer-events-none select-none z-0"
                  />
                ) : (
                  currentSlide.images?.map((img) => (
                    <img
                      key={`new-img-${img.id}`}
                      src={img.dataUrl}
                      alt={img.name}
                      className="absolute object-contain pointer-events-none select-none"
                      style={{
                        left: `${img.bounds.x}%`,
                        top: `${img.bounds.y}%`,
                        width: `${img.bounds.width}%`,
                        height: `${img.bounds.height}%`,
                      }}
                    />
                  ))
                )}

                {/* 2. Shapes with Current Text / Modifications */}
                {currentSlide.shapes.map((shape) => {
                  const hasModified = shape.paragraphs.some((p) => p.isModified);

                  // If we have authentic rendered image and no modifications in this shape, let the image shine
                  if (currentSlide.renderedImageDataUrl && !hasModified) {
                    return null;
                  }

                  return (
                    <div
                      key={`new-sh-${shape.id}`}
                      className={`absolute overflow-hidden p-1.5 flex flex-col justify-start rounded transition-all ${
                        hasModified ? 'ring-1 ring-emerald-500/80 bg-white/95 dark:bg-slate-900/95 shadow-md z-10' : ''
                      }`}
                      style={{
                        left: `${shape.bounds.x}%`,
                        top: `${shape.bounds.y}%`,
                        width: `${shape.bounds.width}%`,
                        height: `${shape.bounds.height}%`,
                        backgroundColor: hasModified ? undefined : shape.backgroundColor || undefined,
                        borderColor: shape.borderColor || undefined,
                        borderWidth: shape.borderColor ? '1px' : undefined,
                        borderStyle: shape.borderColor ? 'solid' : undefined,
                      }}
                    >
                      {hasModified && (
                        <div className="text-[9px] font-bold text-emerald-600 dark:text-emerald-400 mb-0.5 flex items-center gap-1">
                          <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                          Upraveno
                        </div>
                      )}
                      {shape.paragraphs.map((p) => (
                        <div
                          key={`new-p-${p.id}`}
                          className={`my-0.5 rounded px-1 ${
                            p.isModified ? 'text-slate-900 dark:text-slate-100 font-medium' : ''
                          }`}
                          style={{
                            textAlign: p.alignment || 'left',
                            marginLeft: p.bulletLevel ? `${p.bulletLevel * 1.5}rem` : undefined,
                          }}
                        >
                          <p
                            className={`break-words leading-snug ${
                              p.shapeRole === 'title'
                                ? 'text-slate-900 font-extrabold text-xs sm:text-base md:text-lg tracking-tight'
                                : p.shapeRole === 'subtitle'
                                ? 'text-slate-600 font-semibold text-[11px] sm:text-sm'
                                : 'text-slate-700 text-[10px] sm:text-xs md:text-sm'
                            }`}
                            style={{
                              color: p.colorHex || undefined,
                            }}
                          >
                            {p.bulletLevel !== undefined && p.shapeRole === 'body' && (
                              <span className="mr-1.5 text-slate-400">•</span>
                            )}
                            {p.currentText}
                          </p>
                          {p.hyperlink?.url && (
                            <span className="inline-flex items-center gap-1 text-[9px] text-blue-600 underline font-semibold mt-0.5">
                              <LinkIcon className="w-2.5 h-2.5" />
                              {p.hyperlink.url}
                            </span>
                          )}
                        </div>
                      ))}
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        </div>

        {/* Bottom Slide Switcher & Confirmation Action */}
        <div className="px-6 py-3.5 bg-slate-900 border-t border-slate-800 flex flex-col sm:flex-row items-center justify-between gap-4">
          {/* Synchronized Pagination Controls */}
          <div className="flex items-center gap-3">
            <button
              onClick={prevSlide}
              disabled={currentSlideIndex <= 0}
              className="p-1.5 rounded-lg text-slate-300 hover:text-white hover:bg-slate-800 disabled:opacity-30 disabled:pointer-events-none transition-colors"
            >
              <ChevronLeft className="w-5 h-5" />
            </button>

            <span className="text-xs font-semibold text-slate-200">
              {t('slide')} {currentSlide.slideNumber} {t('of')} {document.slides.length}
            </span>

            <button
              onClick={nextSlide}
              disabled={currentSlideIndex >= document.slides.length - 1}
              className="p-1.5 rounded-lg text-slate-300 hover:text-white hover:bg-slate-800 disabled:opacity-30 disabled:pointer-events-none transition-colors"
            >
              <ChevronRight className="w-5 h-5" />
            </button>

            <div className="flex items-center gap-1.5 ml-2 overflow-x-auto max-w-xs sm:max-w-md py-1">
              {document.slides.map((s, idx) => {
                const isSelected = idx === currentSlideIndex;
                const hasMod = s.shapes.some((sh) => sh.paragraphs.some((p) => p.isModified));
                return (
                  <button
                    key={`presave-btn-${s.slideNumber}`}
                    onClick={() => setCurrentSlideIndex(idx)}
                    className={`relative px-2.5 py-1 rounded-lg text-xs font-bold transition-all flex items-center gap-1 shrink-0 ${
                      isSelected
                        ? 'bg-blue-600 text-white shadow-md'
                        : 'bg-slate-800 text-slate-400 hover:bg-slate-700 hover:text-white'
                    }`}
                  >
                    <span>#{s.slideNumber}</span>
                    {hasMod && (
                      <span className="w-1.5 h-1.5 rounded-full bg-emerald-400" />
                    )}
                  </button>
                );
              })}
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-3">
            {/* Custom File Name Input */}
            <div className="flex items-center gap-1.5 bg-slate-800/90 px-3 py-1.5 rounded-xl border border-slate-700">
              <span className="text-xs text-slate-300 font-semibold flex items-center gap-1 shrink-0">
                <Pencil className="w-3 h-3 text-orange-400" />
                {t('saveAs')}:
              </span>
              <input
                type="text"
                value={outputFileName}
                onChange={(e) => setOutputFileName(e.target.value)}
                className="bg-slate-950 text-xs font-bold text-white px-2 py-1 rounded-lg border border-slate-700 focus:outline-none focus:border-orange-500 w-36 sm:w-56"
                placeholder={t('presentationNamePlaceholder', 'Název prezentace')}
              />
              <span className="text-xs text-slate-400 font-mono">.pptx</span>
            </div>

            <button
              type="button"
              onClick={onClose}
              className="px-3.5 py-2 rounded-xl text-xs font-semibold text-slate-400 hover:text-white hover:bg-slate-800 transition-colors"
            >
              {t('backToEditing')}
            </button>

            <button
              type="button"
              onClick={() => {
                onConfirmExport(outputFileName);
                onClose();
              }}
              disabled={isExporting}
              className="px-5 py-2.5 rounded-xl bg-gradient-to-r from-orange-600 to-amber-600 hover:from-orange-700 hover:to-amber-700 text-white font-bold text-xs flex items-center gap-2 shadow-lg shadow-orange-600/25 transition-all active:scale-95 disabled:opacity-50"
            >
              <Download className="w-4 h-4" />
              <span>{t('confirmAndDownload')}</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

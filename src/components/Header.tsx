import React, { useState, useEffect } from 'react';
import {
  FileText,
  Download,
  Languages,
  Sparkles,
  ShieldCheck,
  CheckCircle2,
  Search,
  Table,
  Undo2,
  Redo2,
  Pencil,
  Check,
  Globe,
  ChevronDown,
  Cpu,
  LayoutGrid,
} from 'lucide-react';
import { PptxDocument } from '../types/pptx';
import { useLanguage } from '../context/LanguageContext';

interface HeaderProps {
  document: PptxDocument;
  onExport: () => void;
  onOpenTranslate: () => void;
  onOpenPolish: () => void;
  onOpenVerify: () => void;
  onOpenFindReplace: () => void;
  onOpenBilingualExport: () => void;
  onOpenGallery?: () => void;
  onResetUpload: () => void;
  isExporting: boolean;
  activeView: 'split' | 'document' | 'slide';
  onChangeView: (view: 'split' | 'document' | 'slide') => void;
  canUndo?: boolean;
  canRedo?: boolean;
  onUndo?: () => void;
  onRedo?: () => void;
  onRenameDocument?: (newName: string) => void;
}

export const Header: React.FC<HeaderProps> = ({
  document,
  onExport,
  onOpenTranslate,
  onOpenPolish,
  onOpenVerify,
  onOpenFindReplace,
  onOpenBilingualExport,
  onOpenGallery,
  onResetUpload,
  isExporting,
  activeView,
  onChangeView,
  canUndo = false,
  canRedo = false,
  onUndo,
  onRedo,
  onRenameDocument,
}) => {
  const { language, setLanguage, t, availableLanguages } = useLanguage();
  const [isEditingName, setIsEditingName] = useState(false);
  const [tempName, setTempName] = useState(document.fileName.replace(/\.pptx$/i, ''));
  const [isLangMenuOpen, setIsLangMenuOpen] = useState(false);

  useEffect(() => {
    setTempName(document.fileName.replace(/\.pptx$/i, ''));
  }, [document.fileName]);

  const handleSaveName = () => {
    if (tempName.trim() && onRenameDocument) {
      const finalName = tempName.trim().endsWith('.pptx') ? tempName.trim() : `${tempName.trim()}.pptx`;
      onRenameDocument(finalName);
    }
    setIsEditingName(false);
  };

  const modifiedCount = document.slides.reduce(
    (acc, slide) =>
      acc +
      slide.shapes.reduce(
        (sAcc, shape) => sAcc + shape.paragraphs.filter((p) => p.isModified).length,
        0
      ),
    0
  );

  const currentLangObj = availableLanguages.find((l) => l.code === language) || availableLanguages[0];

  return (
    <header className="sticky top-0 z-40 bg-white border-b border-slate-200 shadow-xs">
      {/* Top Tier: Title, Badges, Language Selector & Download Button */}
      <div className="border-b border-slate-100 px-3 sm:px-6">
        <div className="flex items-center justify-between min-h-[3.5rem] py-2 gap-2 sm:gap-4 flex-wrap sm:flex-nowrap">
          {/* Left: Logo & Document Title (Constrained with truncate to prevent overlapping) */}
          <div className="flex items-center gap-2 sm:gap-3 min-w-0 max-w-full sm:max-w-[65%]">
            <div className="w-8 h-8 rounded-lg bg-gradient-to-tr from-amber-600 to-orange-500 flex items-center justify-center text-white shadow-xs shrink-0">
              <FileText className="w-4 h-4" />
            </div>

            <div className="flex items-center gap-2 min-w-0 flex-wrap sm:flex-nowrap">
              {isEditingName ? (
                <div className="flex items-center gap-1 bg-slate-50 p-0.5 rounded border border-blue-400 max-w-xs">
                  <input
                    type="text"
                    value={tempName}
                    onChange={(e) => setTempName(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') handleSaveName();
                      if (e.key === 'Escape') setIsEditingName(false);
                    }}
                    onBlur={handleSaveName}
                    autoFocus
                    className="text-xs font-bold text-slate-900 bg-white px-2 py-0.5 rounded focus:outline-none w-36 sm:w-52"
                    placeholder="Název prezentace"
                  />
                  <button
                    type="button"
                    onClick={handleSaveName}
                    className="p-1 rounded bg-blue-600 text-white hover:bg-blue-700 transition-colors shrink-0"
                    title={t('saveName')}
                  >
                    <Check className="w-3.5 h-3.5" />
                  </button>
                </div>
              ) : (
                <div className="flex items-center gap-1.5 group min-w-0">
                  <h1
                    onClick={() => setIsEditingName(true)}
                    className="font-bold text-slate-900 text-xs sm:text-sm md:text-base max-w-[140px] xs:max-w-[200px] sm:max-w-[280px] md:max-w-[360px] truncate cursor-pointer hover:text-blue-600 transition-colors"
                    title={t('renameTooltip')}
                  >
                    {document.fileName}
                  </h1>
                  <button
                    type="button"
                    onClick={() => setIsEditingName(true)}
                    className="opacity-60 group-hover:opacity-100 p-1 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded transition-all shrink-0"
                    title={t('renameTooltip')}
                  >
                    <Pencil className="w-3 h-3" />
                  </button>
                </div>
              )}

              {/* Status Badges */}
              <div className="hidden xs:flex items-center gap-1.5 shrink-0">
                <span className="text-[11px] px-2 py-0.5 rounded-full bg-slate-100 text-slate-600 font-medium">
                  {document.slides.length} {t('slidesCount')}
                </span>

                {modifiedCount > 0 ? (
                  <span className="text-[11px] px-2 py-0.5 rounded-full bg-amber-50 text-amber-700 font-medium border border-amber-200/80 flex items-center gap-1 animate-pulse">
                    <span className="w-1.5 h-1.5 rounded-full bg-amber-500"></span>
                    {modifiedCount} {t('modified')}
                  </span>
                ) : (
                  <span className="hidden md:inline-flex text-[11px] px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 font-medium border border-emerald-200/80 items-center gap-1">
                    <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                    {t('originalState')}
                  </span>
                )}

                <span
                  className="hidden lg:inline-flex text-[11px] px-2 py-0.5 rounded-full bg-purple-50 text-purple-700 font-medium border border-purple-200/80 items-center gap-1"
                  title="Lokální model Gemma 4 běží přes Ollama (bez nutnosti Gemini API klíče)"
                >
                  <Cpu className="w-3 h-3 text-purple-600" />
                  Gemma 4
                </span>
              </div>
            </div>
          </div>

          {/* Right: UI Language Selector & Export Button */}
          <div className="flex items-center gap-2 shrink-0 ml-auto">
            {/* App UI Language Dropdown */}
            <div className="relative">
              <button
                type="button"
                onClick={() => setIsLangMenuOpen(!isLangMenuOpen)}
                className="px-2 sm:px-2.5 py-1.5 text-xs font-semibold rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-200 flex items-center gap-1.5 transition-colors"
                title={t('appLanguage')}
              >
                <span className="text-sm">{currentLangObj.flag}</span>
                <span className="hidden sm:inline font-medium">{currentLangObj.name}</span>
                <ChevronDown className="w-3 h-3 text-slate-500" />
              </button>

              {isLangMenuOpen && (
                <>
                  <div
                    className="fixed inset-0 z-40"
                    onClick={() => setIsLangMenuOpen(false)}
                  />
                  <div className="absolute right-0 mt-1.5 w-48 bg-white rounded-xl shadow-2xl border border-slate-200 py-1.5 z-50 max-h-80 overflow-y-auto">
                    <div className="px-3 py-1 text-[10px] font-bold uppercase text-slate-400 tracking-wider">
                      {t('appLanguage')}
                    </div>
                    {availableLanguages.map((lang) => (
                      <button
                        key={lang.code}
                        type="button"
                        onClick={() => {
                          setLanguage(lang.code);
                          setIsLangMenuOpen(false);
                        }}
                        className={`w-full px-3 py-1.5 text-xs text-left flex items-center justify-between hover:bg-slate-50 transition-colors ${
                          language === lang.code ? 'font-bold text-blue-600 bg-blue-50/50' : 'text-slate-700'
                        }`}
                      >
                        <span className="flex items-center gap-2">
                          <span className="text-base">{lang.flag}</span>
                          <span>{lang.name}</span>
                        </span>
                        {language === lang.code && <Check className="w-3.5 h-3.5 text-blue-600" />}
                      </button>
                    ))}
                  </div>
                </>
              )}
            </div>

            {/* Save & Download PPTX Button */}
            <button
              onClick={onExport}
              disabled={isExporting}
              className="px-3 sm:px-4 py-1.5 bg-gradient-to-r from-orange-600 to-amber-600 hover:from-orange-700 hover:to-amber-700 text-white font-bold text-xs rounded-lg flex items-center gap-1.5 shadow-sm shadow-orange-600/20 active:scale-95 transition-all disabled:opacity-50 shrink-0"
            >
              <Download className="w-3.5 h-3.5" />
              <span className="hidden md:inline">{t('saveAndDownload')} (.pptx)</span>
              <span className="md:hidden">{t('saveAndDownload')}</span>
            </button>
          </div>
        </div>
      </div>

      {/* Lower Tier: View Switcher, Undo/Redo & Tools (Responsive, non-overlapping) */}
      <div className="px-3 sm:px-6 py-1.5 bg-slate-50/80 border-t border-slate-100 overflow-x-auto">
        <div className="flex items-center justify-between gap-3 min-w-max">
          {/* Left: Undo / Redo & View Switcher */}
          <div className="flex items-center gap-2 sm:gap-3">
            {/* Undo / Redo */}
            {(onUndo || onRedo) && (
              <div className="flex items-center gap-0.5 bg-white p-0.5 rounded-lg border border-slate-200 shadow-2xs">
                <button
                  type="button"
                  onClick={onUndo}
                  disabled={!canUndo}
                  className="p-1 sm:p-1.5 rounded-md text-slate-600 hover:text-slate-900 hover:bg-slate-100 disabled:opacity-30 disabled:pointer-events-none transition-colors"
                  title={`${t('undo')} (Ctrl+Z)`}
                >
                  <Undo2 className="w-3.5 h-3.5" />
                </button>
                <button
                  type="button"
                  onClick={onRedo}
                  disabled={!canRedo}
                  className="p-1 sm:p-1.5 rounded-md text-slate-600 hover:text-slate-900 hover:bg-slate-100 disabled:opacity-30 disabled:pointer-events-none transition-colors"
                  title={`${t('redo')} (Ctrl+Y)`}
                >
                  <Redo2 className="w-3.5 h-3.5" />
                </button>
              </div>
            )}

            {/* View Mode Switcher: "Rozdělený", "Textový přehled", "Snímek" */}
            <div className="flex items-center bg-slate-200/80 p-0.5 rounded-lg text-xs font-medium text-slate-600">
              <button
                onClick={() => onChangeView('split')}
                className={`px-2.5 py-1 rounded-md transition-all ${
                  activeView === 'split'
                    ? 'bg-white text-slate-900 shadow-xs font-semibold'
                    : 'hover:text-slate-900'
                }`}
                title={t('splitView')}
              >
                {t('splitView')}
              </button>
              <button
                onClick={() => onChangeView('document')}
                className={`px-2.5 py-1 rounded-md transition-all ${
                  activeView === 'document'
                    ? 'bg-white text-slate-900 shadow-xs font-semibold'
                    : 'hover:text-slate-900'
                }`}
                title={t('documentOnly')}
              >
                {t('documentOnly')}
              </button>
              <button
                onClick={() => onChangeView('slide')}
                className={`px-2.5 py-1 rounded-md transition-all ${
                  activeView === 'slide'
                    ? 'bg-white text-slate-900 shadow-xs font-semibold'
                    : 'hover:text-slate-900'
                }`}
                title={t('slideOnly')}
              >
                {t('slideOnly')}
              </button>
            </div>
          </div>

          {/* Right: AI Tools & Action Buttons */}
          <div className="flex items-center gap-1.5">
            {onOpenGallery && (
              <button
                onClick={onOpenGallery}
                className="px-2.5 py-1 text-xs font-semibold rounded-md text-blue-700 bg-blue-50 hover:bg-blue-100 border border-blue-200 transition-colors flex items-center gap-1.5 shadow-2xs"
                title="Zobrazit přehled všech snímků celé prezentace (všechny šablony a styly)"
              >
                <LayoutGrid className="w-3.5 h-3.5" />
                <span>Celá prezentace</span>
              </button>
            )}

            <button
              onClick={onOpenTranslate}
              className="px-2.5 py-1 text-xs font-semibold rounded-md text-indigo-700 bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 transition-colors flex items-center gap-1.5 shadow-2xs"
            >
              <Languages className="w-3.5 h-3.5" />
              <span>{t('aiTranslator')}</span>
            </button>

            <button
              onClick={onOpenPolish}
              className="px-2.5 py-1 text-xs font-semibold rounded-md text-amber-800 bg-amber-50 hover:bg-amber-100 border border-amber-200 transition-colors flex items-center gap-1.5 shadow-2xs"
            >
              <Sparkles className="w-3.5 h-3.5" />
              <span>{t('aiPolish')}</span>
            </button>

            <button
              onClick={onOpenVerify}
              className="px-2.5 py-1 text-xs font-semibold rounded-md text-teal-800 bg-teal-50 hover:bg-teal-100 border border-teal-200 transition-colors flex items-center gap-1.5 shadow-2xs"
            >
              <ShieldCheck className="w-3.5 h-3.5" />
              <span>{t('contextAudit')}</span>
            </button>

            <button
              onClick={onOpenFindReplace}
              className="p-1 text-xs font-semibold rounded-md text-slate-700 bg-white hover:bg-slate-100 border border-slate-200 transition-colors flex items-center gap-1 shadow-2xs px-2"
              title={t('findReplace')}
            >
              <Search className="w-3.5 h-3.5 text-slate-500" />
              <span className="hidden sm:inline">{t('findReplace')}</span>
            </button>

            <button
              onClick={onOpenBilingualExport}
              className="p-1 text-xs font-semibold rounded-md text-slate-700 bg-white hover:bg-slate-100 border border-slate-200 transition-colors flex items-center gap-1 shadow-2xs px-2"
              title={t('bilingualTable')}
            >
              <Table className="w-3.5 h-3.5 text-slate-500" />
              <span className="hidden sm:inline">{t('bilingualTable')}</span>
            </button>
          </div>
        </div>
      </div>
    </header>
  );
};

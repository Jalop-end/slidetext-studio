import React, { useState, useRef } from 'react';
import {
  UploadCloud,
  FileText,
  Sparkles,
  ShieldCheck,
  Languages,
  CheckCircle2,
  FileCheck,
  Layers,
  ArrowRight,
  Loader2,
  Globe,
  ChevronDown,
  Check,
} from 'lucide-react';
import { useLanguage } from '../context/LanguageContext';

interface UploadDropzoneProps {
  onFileUpload: (file: File) => void;
  onLoadSample: (sampleType: 'czech-strategy' | 'english-pitch' | 'word-strategy') => void;
  isLoading: boolean;
  loadingMessage?: string;
}

export const UploadDropzone: React.FC<UploadDropzoneProps> = ({
  onFileUpload,
  onLoadSample,
  isLoading,
  loadingMessage,
}) => {
  const { language, setLanguage, t, availableLanguages } = useLanguage();
  const [isDragOver, setIsDragOver] = useState(false);
  const [isLangMenuOpen, setIsLangMenuOpen] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const currentLangObj = availableLanguages.find((l) => l.code === language) || availableLanguages[0];

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      const name = files[0].name.toLowerCase();
      if (name.endsWith('.pptx') || name.endsWith('.docx')) {
        onFileUpload(files[0]);
      }
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      onFileUpload(files[0]);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center p-4 sm:p-6 lg:p-8 relative">
      {/* Top Language Switcher Bar */}
      <div className="absolute top-4 right-4 sm:top-6 sm:right-6 z-20">
        <div className="relative">
          <button
            type="button"
            onClick={() => setIsLangMenuOpen(!isLangMenuOpen)}
            className="px-3 py-1.5 text-xs font-semibold rounded-xl bg-white hover:bg-slate-50 text-slate-700 border border-slate-200 shadow-xs flex items-center gap-2 transition-colors"
          >
            <span className="text-base">{currentLangObj.flag}</span>
            <span className="font-bold">{currentLangObj.name}</span>
            <ChevronDown className="w-3.5 h-3.5 text-slate-400" />
          </button>

          {isLangMenuOpen && (
            <>
              <div
                className="fixed inset-0 z-30"
                onClick={() => setIsLangMenuOpen(false)}
              />
              <div className="absolute right-0 mt-1.5 w-48 bg-white rounded-xl shadow-xl border border-slate-200 py-1.5 z-40 max-h-72 overflow-y-auto">
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
                    className={`w-full px-3 py-1.5 text-xs text-left flex items-center justify-between hover:bg-slate-50 transition-colors ${language === lang.code ? 'font-bold text-blue-600 bg-blue-50/50' : 'text-slate-700'
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
      </div>

      <div className="max-w-3xl w-full space-y-8">
        {/* Hero Title */}
        <div className="text-center space-y-3">
          {/*
          <div className="inline-flex items-center gap-2 px-3 py-1.5 rounded-full bg-orange-100/80 text-orange-800 text-xs font-bold border border-orange-200 shadow-xs mb-2">
            <Sparkles className="w-3.5 h-3.5 text-orange-600" />
            <span>PowerPoint • AI • 100% Layout</span>
          </div>
          */}

          <h1 className="text-3xl sm:text-4xl font-extrabold text-slate-900 tracking-tight">
            {t('dropzoneTitle')}
          </h1>
          <p className="text-sm sm:text-base text-slate-600 max-w-xl mx-auto leading-relaxed">
            {t('dropzoneSubtitle')}
          </p>
        </div>

        {/* Upload Card */}
        <div
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
          onDrop={handleDrop}
          onClick={() => fileInputRef.current?.click()}
          className={`relative border-2 border-dashed rounded-3xl p-8 sm:p-12 text-center cursor-pointer transition-all duration-200 bg-white ${isDragOver
            ? 'border-orange-500 bg-orange-50/40 ring-4 ring-orange-500/10 scale-[1.01]'
            : 'border-slate-300 hover:border-orange-400 hover:bg-slate-50/60 shadow-lg shadow-slate-200/50'
            }`}
        >
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            accept=".pptx,.docx,application/vnd.openxmlformats-officedocument.presentationml.presentation,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            className="hidden"
          />

          {isLoading ? (
            <div className="py-8 flex flex-col items-center justify-center space-y-4">
              <div className="w-16 h-16 rounded-2xl bg-orange-50 border border-orange-200 flex items-center justify-center text-orange-600">
                <Loader2 className="w-8 h-8 animate-spin" />
              </div>
              <p className="font-bold text-slate-900 text-base">
                {loadingMessage || t('loadingPresentation', 'Zpracovávám dokument...')}
              </p>
              <p className="text-xs text-slate-500">
                {t('loadingPresentationDesc', 'Rozbaluji strukturu a načítám textové bloky se 100% zachováním stylů...')}
              </p>
            </div>
          ) : (
            <div className="space-y-4">
              <div className="flex items-center justify-center gap-3">
                <div className="w-14 h-14 rounded-2xl bg-gradient-to-tr from-orange-500 to-amber-500 text-white flex items-center justify-center shadow-lg shadow-orange-500/25" title="PowerPoint (.pptx)">
                  <UploadCloud className="w-7 h-7" />
                </div>
                <div className="w-14 h-14 rounded-2xl bg-gradient-to-tr from-blue-600 to-indigo-600 text-white flex items-center justify-center shadow-lg shadow-blue-500/25" title="Microsoft Word (.docx)">
                  <FileText className="w-7 h-7" />
                </div>
              </div>

              <div>
                <p className="font-bold text-slate-900 text-base sm:text-lg">
                  {t('dragDropOrClick', 'Přetáhněte sem prezentaci (.pptx) nebo dokument Word (.docx)')}
                </p>
                <p className="text-xs sm:text-sm text-slate-500 mt-1">
                  {t('supportedFormatsDesc', 'Podporovány jsou soubory Microsoft PowerPoint (.pptx) i Microsoft Word (.docx)')}
                </p>
              </div>

              <div className="pt-2 flex items-center justify-center gap-2">
                <span className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl bg-slate-900 hover:bg-slate-800 text-white font-bold text-xs shadow-md transition-all">
                  <FileText className="w-4 h-4" />
                  {t('chooseFromComputer', 'Vybrat soubor z počítače (.pptx / .docx)')}
                </span>
              </div>
            </div>
          )}
        </div>

        {/* Quick Samples Section */}
        <div className="bg-white rounded-2xl p-6 border border-slate-200 shadow-sm space-y-4">
          <div className="flex items-center justify-between">
            <div>
              <h3 className="font-bold text-slate-900 text-sm">
                {t('noFileOnHand', 'Nemáte zrovna po ruce svůj soubor?')}
              </h3>
              <p className="text-xs text-slate-500">
                {t('trySampleDesc', 'Vyzkoušejte aplikaci okamžitě s jednou z našich připravených vzorových ukázek:')}
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <button
              type="button"
              onClick={() => onLoadSample('czech-strategy')}
              disabled={isLoading}
              className="p-3.5 rounded-xl border border-slate-200 hover:border-orange-300 hover:bg-orange-50/40 text-left transition-all flex items-start justify-between group disabled:opacity-50"
            >
              <div>
                <div className="flex items-center gap-1.5 font-bold text-xs text-slate-900 group-hover:text-orange-700">
                  <span>{t('sampleCzech', '🇨🇿 Prezentace: Strategie inovací (.pptx)')}</span>
                </div>
                <p className="text-[11px] text-slate-500 mt-1">
                  {t('sampleCzechDesc', '4 snímky, odrážky, podnadpisy, vhodné pro překlad do AJ/NJ')}
                </p>
              </div>
              <ArrowRight className="w-4 h-4 text-slate-400 group-hover:text-orange-600 shrink-0 mt-0.5 transition-transform group-hover:translate-x-0.5" />
            </button>

            <button
              type="button"
              onClick={() => onLoadSample('english-pitch')}
              disabled={isLoading}
              className="p-3.5 rounded-xl border border-slate-200 hover:border-orange-300 hover:bg-orange-50/40 text-left transition-all flex items-start justify-between group disabled:opacity-50"
            >
              <div>
                <div className="flex items-center gap-1.5 font-bold text-xs text-slate-900 group-hover:text-orange-700">
                  <span>{t('sampleEnglish', '🇬🇧 Pitch Deck 2026 (.pptx)')}</span>
                </div>
                <p className="text-[11px] text-slate-500 mt-1">
                  {t('sampleEnglishDesc', 'Anglický pitch deck, perfektní pro test lokalizace do češtiny')}
                </p>
              </div>
              <ArrowRight className="w-4 h-4 text-slate-400 group-hover:text-orange-600 shrink-0 mt-0.5 transition-transform group-hover:translate-x-0.5" />
            </button>

            <button
              type="button"
              onClick={() => onLoadSample('word-strategy')}
              disabled={isLoading}
              className="p-3.5 rounded-xl border border-blue-200 hover:border-blue-400 hover:bg-blue-50/50 text-left transition-all flex items-start justify-between group disabled:opacity-50"
            >
              <div>
                <div className="flex items-center gap-1.5 font-bold text-xs text-blue-900 group-hover:text-blue-700">
                  <span>📄 Výroční zpráva & Plán (.docx)</span>
                </div>
                <p className="text-[11px] text-slate-500 mt-1">
                  Word dokument: nadpisy, formátované odstavce, odrážky a strukturovaná tabulka
                </p>
              </div>
              <ArrowRight className="w-4 h-4 text-blue-400 group-hover:text-blue-600 shrink-0 mt-0.5 transition-transform group-hover:translate-x-0.5" />
            </button>
          </div>
        </div>

        {/* Feature Highlights Grid */}
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 pt-2">
          <div className="p-4 rounded-xl bg-white border border-slate-200/80 shadow-xs space-y-1.5">
            <div className="w-8 h-8 rounded-lg bg-orange-100 text-orange-700 flex items-center justify-center mb-2">
              <FileCheck className="w-4 h-4" />
            </div>
            <h4 className="font-bold text-slate-900 text-xs">
              {t('featureCard1Title', '100% zachování formátování')}
            </h4>
            <p className="text-[11px] text-slate-500 leading-relaxed">
              {t('featureCard1Desc', 'Mění se pouze text. Rozložení objektů, barvy, fonty a pozice zůstávají nedotčeny.')}
            </p>
          </div>

          <div className="p-4 rounded-xl bg-white border border-slate-200/80 shadow-xs space-y-1.5">
            <div className="w-8 h-8 rounded-lg bg-blue-100 text-blue-700 flex items-center justify-center mb-2">
              <Layers className="w-4 h-4" />
            </div>
            <h4 className="font-bold text-slate-900 text-xs">
              {t('featureCard2Title', 'Word přehled & mini-odkazy')}
            </h4>
            <p className="text-[11px] text-slate-500 leading-relaxed">
              {t('featureCard2Desc', 'Text je přehledně vytažen do stránek s mini-odkazy pro přesné zobrazení na snímku.')}
            </p>
          </div>

          <div className="p-4 rounded-xl bg-white border border-slate-200/80 shadow-xs space-y-1.5">
            <div className="w-8 h-8 rounded-lg bg-teal-100 text-teal-700 flex items-center justify-center mb-2">
              <ShieldCheck className="w-4 h-4" />
            </div>
            <h4 className="font-bold text-slate-900 text-xs">
              {t('featureCard3Title', 'AI kontextový auditor')}
            </h4>
            <p className="text-[11px] text-slate-500 leading-relaxed">
              {t('featureCard3Desc', 'Kontroluje návaznost slov, jednotu pojmů i délku textu, aby se slide nezdeformoval.')}
            </p>
          </div>
        </div>
      </div>
    </div>
  );
};

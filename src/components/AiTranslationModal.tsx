import React, { useEffect, useRef, useState } from 'react';
import {
  X,
  Languages,
  Check,
  AlertTriangle,
  Loader2,
  Key,
  Cpu,
  Globe2,
  Save,
  CheckCircle2,
  Sparkles,
} from 'lucide-react';
import { PptxDocument, PptxParagraph, TranslationItem } from '../types/pptx';
import {
  requestAiTranslation,
  TranslationResultItem,
  fetchServerStatus,
  ServerStatusResponse,
} from '../services/aiService';
import { useLanguage } from '../context/LanguageContext';

interface AiTranslationModalProps {
  isOpen: boolean;
  onClose: () => void;
  document: PptxDocument;
  currentSlideNumber: number;
  initialParagraph?: PptxParagraph | null;
  onApplyTranslations: (results: TranslationResultItem[]) => void;
}

const SUPPORTED_LANGUAGES = [
  { code: 'en', name: 'Angličtina (English)', flag: '🇬🇧' },
  { code: 'cs', name: 'Čeština', flag: '🇨🇿' },
  { code: 'de', name: 'Němčina (Deutsch)', flag: '🇩🇪' },
  { code: 'sk', name: 'Slovenčina', flag: '🇸🇰' },
  { code: 'pl', name: 'Polština (Polski)', flag: '🇵🇱' },
  { code: 'es', name: 'Španělština (Español)', flag: '🇪🇸' },
  { code: 'fr', name: 'Francouzština (Français)', flag: '🇫🇷' },
  { code: 'it', name: 'Italština (Italiano)', flag: '🇮🇹' },
  { code: 'uk', name: 'Ukrajinština (Українська)', flag: '🇺🇦' },
  { code: 'zh', name: 'Čínština (中文)', flag: '🇨🇳' },
  { code: 'ja', name: 'Japonština (日本語)', flag: '🇯🇵' },
];

export const AiTranslationModal: React.FC<AiTranslationModalProps> = ({
  isOpen,
  onClose,
  document,
  currentSlideNumber,
  initialParagraph,
  onApplyTranslations,
}) => {
  const { t } = useLanguage();
  const [targetLanguage, setTargetLanguage] = useState<string>('Angličtina (English)');
  const [scope, setScope] = useState<'all' | 'current' | 'single'>(
    initialParagraph ? 'single' : 'all'
  );
  // Default is 100% free fast translator
  const [provider, setProvider] = useState<'free' | 'gemma' | 'custom-key'>('free');
  const [apiKey, setApiKey] = useState<string>(() => {
    try {
      return localStorage.getItem('app_translator_key') || '';
    } catch {
      return '';
    }
  });
  const [isSavedKey, setIsSavedKey] = useState(false);
  const [serverStatus, setServerStatus] = useState<ServerStatusResponse | null>(null);

  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [warning, setWarning] = useState<string | null>(null);
  const [previewResults, setPreviewResults] = useState<TranslationResultItem[] | null>(null);
  const abortRef = useRef<AbortController | null>(null);

  // Fetch server status on mount
  useEffect(() => {
    fetchServerStatus().then((status) => {
      if (status) {
        setServerStatus(status);
        if (status.translator?.hasEnvKey) {
          setProvider('custom-key');
        } else if (status.gemma?.available) {
          // If local Gemma is detected and active, offer Gemma
          // keep 'free' as fast default or let user choose
        }
      }
    });
  }, []);

  // Cancel running request when modal closes / unmounts
  useEffect(() => {
    return () => abortRef.current?.abort();
  }, []);

  if (!isOpen) return null;

  const handleClose = () => {
    abortRef.current?.abort();
    onClose();
  };

  const handleSaveApiKey = () => {
    try {
      localStorage.setItem('app_translator_key', apiKey.trim());
      setIsSavedKey(true);
      setTimeout(() => setIsSavedKey(false), 2500);
    } catch {}
  };

  // Collect items to translate based on scope
  const getItemsToTranslate = (): TranslationItem[] => {
    if (initialParagraph && scope === 'single') {
      return [
        {
          id: initialParagraph.id,
          slideNumber: initialParagraph.slideNumber,
          role: initialParagraph.shapeRole,
          originalText: initialParagraph.currentText,
        },
      ];
    }

    const items: TranslationItem[] = [];
    document.slides.forEach((slide) => {
      if (scope === 'current' && slide.slideNumber !== currentSlideNumber) return;

      slide.shapes.forEach((shape) => {
        shape.paragraphs.forEach((p) => {
          if (p.currentText.trim()) {
            items.push({
              id: p.id,
              slideNumber: slide.slideNumber,
              role: p.shapeRole,
              originalText: p.currentText,
            });
          }
        });
      });

      if (slide.speakerNotes?.text?.trim()) {
        items.push({
          id: `slide-${slide.slideNumber}-notes`,
          slideNumber: slide.slideNumber,
          role: 'footer',
          originalText: slide.speakerNotes.text,
        });
      }
    });

    return items;
  };

  const handleStartTranslation = async () => {
    setIsLoading(true);
    setError(null);
    setWarning(null);
    setPreviewResults(null);

    const items = getItemsToTranslate();
    if (items.length === 0) {
      setError(t('noTextsToTranslate', 'Nebyly nalezeny žádné texty k překladu.'));
      setIsLoading(false);
      return;
    }

    // Only validate key if user explicitly chose "custom-key" and has no server key
    if (provider === 'custom-key' && !apiKey.trim() && !serverStatus?.translator?.hasEnvKey) {
      setError(t('customKeyLabel', 'Zadejte prosím svůj API klíč, nebo zvolte výchozí bezplatný překlad.'));
      setIsLoading(false);
      return;
    }

    abortRef.current?.abort();
    const controller = new AbortController();
    abortRef.current = controller;

    try {
      const backendProvider =
        provider === 'gemma' ? 'gemma' : provider === 'custom-key' ? 'google' : 'auto';

      const response = await requestAiTranslation(items, targetLanguage, {
        apiKey: apiKey.trim() || undefined,
        provider: backendProvider,
        externalSignal: controller.signal,
      });

      if (controller.signal.aborted) return;
      setPreviewResults(response.results);
      if (response.warning) setWarning(response.warning);
    } catch (err: any) {
      if (controller.signal.aborted) return;
      setError(err.message || 'Nepodařilo se provést překlad.');
    } finally {
      if (abortRef.current === controller) abortRef.current = null;
      setIsLoading(false);
    }
  };

  const handleApply = () => {
    if (previewResults) {
      onApplyTranslations(previewResults);
      onClose();
    }
  };

  const totalItemsCount = getItemsToTranslate().length;
  const hasServerKey = !!serverStatus?.translator?.hasEnvKey;
  const isGemmaOnline = !!serverStatus?.gemma?.available;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-150">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-2xl w-full max-h-[92vh] flex flex-col overflow-hidden">
        {/* Modal Header */}
        <div className="px-6 py-4 bg-gradient-to-r from-blue-50 to-indigo-50 border-b border-blue-100 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-blue-600 text-white flex items-center justify-center shadow-md shadow-blue-600/20">
              <Languages className="w-5 h-5" />
            </div>
            <div>
              <h3 className="font-bold text-slate-900 text-base">
                {t('translationModalTitle', 'Překladač prezentace (100% Zdarma, 1:1 layout)')}
              </h3>
              <p className="text-xs text-slate-600">
                {t('translationModalSubtitle', 'Bleskový bezplatný překlad textů se zachováním rozvržení a odrážek')}
              </p>
            </div>
          </div>
          <button
            onClick={handleClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-200/50 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-6 overflow-y-auto space-y-5 flex-1">
          {error && (
            <div className="p-3.5 rounded-xl bg-rose-50 border border-rose-200 text-rose-800 text-xs flex items-center gap-2">
              <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
              <span>{error}</span>
            </div>
          )}

          {warning && (
            <div className="p-3.5 rounded-xl bg-amber-50 border border-amber-200 text-amber-900 text-xs flex items-center gap-2">
              <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
              <span>{warning}</span>
            </div>
          )}

          {!previewResults ? (
            <div className="space-y-5">
              {/* Provider Selection */}
              <div>
                <label className="block text-xs font-bold uppercase tracking-wider text-slate-700 mb-2">
                  {t('chooseTranslationMethod', 'Zvolte způsob překladu')}
                </label>
                <div className="grid grid-cols-3 gap-2">
                  <button
                    type="button"
                    onClick={() => setProvider('free')}
                    className={`p-2.5 rounded-xl border text-left transition-all ${
                      provider === 'free'
                        ? 'bg-emerald-50 border-emerald-500 ring-2 ring-emerald-500/20 shadow-xs'
                        : 'border-slate-200 hover:border-slate-300 bg-white'
                    }`}
                  >
                    <div className="flex items-center gap-1.5 font-bold text-xs text-slate-900">
                      <Globe2 className="w-3.5 h-3.5 text-emerald-600" />
                      <span>{t('providerFree', 'Rychlý překladač (Zdarma)')}</span>
                    </div>
                    <p className="text-[10px] text-emerald-700 font-semibold mt-0.5">
                      {t('providerFreeDesc', 'Bez API klíče • Okamžitý překlad')}
                    </p>
                  </button>

                  <button
                    type="button"
                    onClick={() => setProvider('gemma')}
                    className={`p-2.5 rounded-xl border text-left transition-all relative ${
                      provider === 'gemma'
                        ? 'bg-purple-50 border-purple-500 ring-2 ring-purple-500/20 shadow-xs'
                        : 'border-slate-200 hover:border-slate-300 bg-white'
                    }`}
                  >
                    <div className="flex items-center gap-1.5 font-bold text-xs text-slate-900">
                      <Cpu className="w-3.5 h-3.5 text-purple-600" />
                      <span>{t('providerGemma', 'Gemma 4 (Místní AI)')}</span>
                    </div>
                    <p className="text-[10px] text-purple-700 font-semibold mt-0.5">
                      {isGemmaOnline ? 'Ollama • Aktivní' : t('providerGemmaDesc', 'Ollama • Offline AI bez limitů')}
                    </p>
                  </button>

                  <button
                    type="button"
                    onClick={() => setProvider('custom-key')}
                    className={`p-2.5 rounded-xl border text-left transition-all ${
                      provider === 'custom-key'
                        ? 'bg-blue-50 border-blue-500 ring-2 ring-blue-500/20 shadow-xs'
                        : 'border-slate-200 hover:border-slate-300 bg-white'
                    }`}
                  >
                    <div className="flex items-center gap-1.5 font-bold text-xs text-slate-900">
                      <Key className="w-3.5 h-3.5 text-blue-600" />
                      <span>{t('providerKey', 'Vlastní API klíč')}</span>
                    </div>
                    <p className="text-[10px] text-slate-500 mt-0.5">{t('providerKeyDesc', 'Google Cloud / DeepL')}</p>
                  </button>
                </div>
              </div>

              {/* Status info banner based on selected provider */}
              {provider === 'free' && (
                <div className="p-3 rounded-xl bg-emerald-50 border border-emerald-200 text-xs text-emerald-800 flex items-center gap-2">
                  <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                  <span>
                    {t('freeBanner', '100% zdarma: Překlad běží bez nutnosti placeného API klíče nebo platební karty.')}
                  </span>
                </div>
              )}

              {provider === 'gemma' && (
                <div className="p-3 rounded-xl bg-purple-50 border border-purple-200 text-xs text-purple-800 flex items-center gap-2">
                  <Sparkles className="w-4 h-4 text-purple-600 shrink-0" />
                  <span>
                    {t('gemmaBanner', 'Lokální AI: Překlad provádí váš model Gemma 4 na vašem počítači přes Ollama (zcela zdarma a offline).')}
                  </span>
                </div>
              )}

              {/* API Key Configuration (shown only when custom-key selected) */}
              {provider === 'custom-key' && (
                <div className="p-3.5 rounded-xl bg-slate-50 border border-slate-200 space-y-2">
                  <div className="flex items-center justify-between">
                    <label className="text-xs font-bold text-slate-700 flex items-center gap-1.5">
                      <Key className="w-3.5 h-3.5 text-slate-500" />
                      <span>{t('customKeyLabel', 'Vlastní placený API klíč (Google Cloud nebo DeepL)')}</span>
                    </label>
                    {hasServerKey && !apiKey && (
                      <span className="text-[10px] font-semibold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-md border border-emerald-200 flex items-center gap-1">
                        <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                        Nastaveno v .env.local
                      </span>
                    )}
                  </div>

                  <div className="flex gap-2">
                    <input
                      type="password"
                      value={apiKey}
                      onChange={(e) => setApiKey(e.target.value)}
                      placeholder={t('apiKeyPlaceholder', 'Vložte Google Cloud nebo DeepL API klíč...')}
                      className="flex-1 px-3 py-1.5 rounded-lg border border-slate-300 text-xs text-slate-800 bg-white focus:outline-none focus:ring-2 focus:ring-blue-500/30 focus:border-blue-500"
                    />
                    <button
                      type="button"
                      onClick={handleSaveApiKey}
                      disabled={!apiKey.trim()}
                      className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold flex items-center gap-1 transition-colors disabled:opacity-40"
                      title={t('keySaved', 'Uložit klíč')}
                    >
                      <Save className="w-3 h-3" />
                      <span>{isSavedKey ? t('keySaved', 'Uloženo!') : t('saveKey', 'Uložit')}</span>
                    </button>
                  </div>
                </div>
              )}

              {/* Target Language Selection */}
              <div>
                <label className="block text-xs font-bold uppercase tracking-wider text-slate-700 mb-2">
                  {t('targetLangLabel', 'Cílový jazyk překladu')}
                </label>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                  {SUPPORTED_LANGUAGES.map((lang) => (
                    <button
                      key={lang.code}
                      type="button"
                      onClick={() => setTargetLanguage(lang.name)}
                      className={`px-3 py-2 rounded-xl text-xs font-semibold flex items-center gap-2 border transition-all text-left ${
                        targetLanguage === lang.name
                          ? 'bg-blue-50 border-blue-500 text-blue-950 ring-2 ring-blue-500/20 shadow-xs'
                          : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50 hover:border-slate-300'
                      }`}
                    >
                      <span className="text-base">{lang.flag}</span>
                      <span className="truncate">{lang.name.split(' (')[0]}</span>
                    </button>
                  ))}
                </div>
              </div>

              {/* Translation Scope */}
              <div>
                <label className="block text-xs font-bold uppercase tracking-wider text-slate-700 mb-2">
                  {t('scopeLabel', 'Rozsah překladu')}
                </label>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                  <button
                    type="button"
                    onClick={() => setScope('all')}
                    className={`p-3 rounded-xl border text-left transition-all ${
                      scope === 'all'
                        ? 'bg-blue-50/80 border-blue-500 ring-2 ring-blue-500/20'
                        : 'border-slate-200 hover:border-slate-300 bg-white'
                    }`}
                  >
                    <p className="text-xs font-bold text-slate-900">{t('scopeAll', 'Celá prezentace')}</p>
                    <p className="text-[11px] text-slate-500 mt-0.5">
                      {document.slides.length} {t('slidesCount', 'snímků')}
                    </p>
                  </button>

                  <button
                    type="button"
                    onClick={() => setScope('current')}
                    className={`p-3 rounded-xl border text-left transition-all ${
                      scope === 'current'
                        ? 'bg-blue-50/80 border-blue-500 ring-2 ring-blue-500/20'
                        : 'border-slate-200 hover:border-slate-300 bg-white'
                    }`}
                  >
                    <p className="text-xs font-bold text-slate-900">{t('scopeCurrent', 'Pouze aktuální snímek')}</p>
                    <p className="text-[11px] text-slate-500 mt-0.5">
                      #{currentSlideNumber}
                    </p>
                  </button>

                  {initialParagraph && (
                    <button
                      type="button"
                      onClick={() => setScope('single')}
                      className={`p-3 rounded-xl border text-left transition-all ${
                        scope === 'single'
                          ? 'bg-blue-50/80 border-blue-500 ring-2 ring-blue-500/20'
                          : 'border-slate-200 hover:border-slate-300 bg-white'
                      }`}
                    >
                      <p className="text-xs font-bold text-slate-900">{t('scopeSingle', 'Pouze vybraný odstavec')}</p>
                      <p className="text-[11px] text-slate-500 mt-0.5">
                        1 {t('itemsCountLabel', 'položka')}
                      </p>
                    </button>
                  )}
                </div>
              </div>
            </div>
          ) : (
            /* Results Preview */
            <div className="space-y-4">
              <div className="flex items-center justify-between pb-2 border-b border-slate-200">
                <p className="text-xs font-bold text-slate-900">
                  {t('previewTitle', 'Náhled přeložených textů')}: {previewResults.length} {t('itemsCountLabel', 'položek')}
                </p>
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setPreviewResults(null)}
                    className="text-[11px] text-blue-600 hover:text-blue-800 font-semibold underline"
                  >
                    {t('backToSettings', 'Zpět k nastavení')}
                  </button>
                  <span className="text-xs font-semibold text-blue-700 bg-blue-50 px-2.5 py-1 rounded-md border border-blue-200">
                    {targetLanguage}
                  </span>
                </div>
              </div>

              <div className="space-y-2.5 max-h-72 overflow-y-auto pr-1">
                {previewResults.map((item, idx) => (
                  <div
                    key={item.id}
                    className="p-3 rounded-xl border border-slate-200 bg-slate-50/80 text-xs space-y-1.5"
                  >
                    <div className="flex items-center justify-between text-slate-400 font-semibold text-[10px]">
                      <span>#{idx + 1}</span>
                      {item.notes && <span className="text-emerald-700 font-medium">{item.notes}</span>}
                    </div>
                    <p className="text-slate-900 font-medium whitespace-pre-wrap">{item.translatedText}</p>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
          <button
            type="button"
            onClick={handleClose}
            className="px-4 py-2 text-xs font-semibold text-slate-600 hover:text-slate-900 transition-colors"
          >
            {t('cancelBtn', 'Zrušit')}
          </button>

          {!previewResults ? (
            <button
              type="button"
              onClick={handleStartTranslation}
              disabled={isLoading || totalItemsCount === 0}
              className="px-5 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs flex items-center gap-2 shadow-md shadow-emerald-600/20 transition-all disabled:opacity-50"
            >
              {isLoading ? (
                <>
                  <Loader2 className="w-4 h-4 animate-spin" />
                  <span>{t('translatingProgress', 'Překládám texty...')} ({totalItemsCount})</span>
                </>
              ) : (
                <>
                  <Languages className="w-4 h-4" />
                  <span>{t('startTranslationBtn', 'Spustit překlad')} ({totalItemsCount})</span>
                </>
              )}
            </button>
          ) : (
            <button
              type="button"
              onClick={handleApply}
              className="px-5 py-2.5 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs flex items-center gap-2 shadow-md shadow-blue-600/20 transition-all"
            >
              <Check className="w-4 h-4" />
              <span>{t('applyTranslationsBtn', 'Použít přeložené texty do prezentace')}</span>
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

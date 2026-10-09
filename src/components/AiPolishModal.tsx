import React, { useState } from 'react';
import {
  X,
  Sparkles,
  Check,
  AlertTriangle,
  Loader2,
  Sliders,
  Type,
  FileCheck2,
} from 'lucide-react';
import { PptxDocument, PptxParagraph } from '../types/pptx';
import { requestAiPolish, PolishResultItem } from '../services/aiService';
import { useLanguage } from '../context/LanguageContext';

interface AiPolishModalProps {
  isOpen: boolean;
  onClose: () => void;
  document: PptxDocument;
  currentSlideNumber: number;
  initialParagraph?: PptxParagraph | null;
  onApplyPolished: (results: PolishResultItem[]) => void;
}

export const AiPolishModal: React.FC<AiPolishModalProps> = ({
  isOpen,
  onClose,
  document,
  currentSlideNumber,
  initialParagraph,
  onApplyPolished,
}) => {
  const { t } = useLanguage();
  const [mode, setMode] = useState<'grammar' | 'concise' | 'professional' | 'engaging'>('grammar');
  const [scope, setScope] = useState<'all' | 'current' | 'single'>(
    initialParagraph ? 'single' : 'current'
  );
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [previewResults, setPreviewResults] = useState<PolishResultItem[] | null>(null);

  if (!isOpen) return null;

  const getItemsToPolish = (): { id: string; text: string }[] => {
    if (initialParagraph && scope === 'single') {
      return [{ id: initialParagraph.id, text: initialParagraph.currentText }];
    }

    const items: { id: string; text: string }[] = [];
    document.slides.forEach((slide) => {
      if (scope === 'current' && slide.slideNumber !== currentSlideNumber) return;

      slide.shapes.forEach((shape) => {
        shape.paragraphs.forEach((p) => {
          if (p.currentText.trim()) {
            items.push({ id: p.id, text: p.currentText });
          }
        });
      });
    });

    return items;
  };

  const handleStartPolish = async () => {
    setIsLoading(true);
    setError(null);
    setPreviewResults(null);

    const items = getItemsToPolish();
    if (items.length === 0) {
      setError(t('noTextsToPolish', 'Nebyly nalezeny žádné texty ke korektuře.'));
      setIsLoading(false);
      return;
    }

    try {
      const results = await requestAiPolish(items, mode, 'čeština');
      setPreviewResults(results);
    } catch (err: any) {
      setError(err.message || t('polishFailed', 'Nepodařilo se provést korekturu.'));
    } finally {
      setIsLoading(false);
    }
  };

  const handleApply = () => {
    if (previewResults) {
      onApplyPolished(previewResults);
      onClose();
    }
  };

  const itemsCount = getItemsToPolish().length;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-150">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-2xl w-full max-h-[90vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="px-6 py-4 bg-gradient-to-r from-amber-50 to-orange-50 border-b border-amber-100 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-amber-600 text-white flex items-center justify-center shadow-md shadow-amber-600/20">
              <Sparkles className="w-5 h-5" />
            </div>
            <div>
              <h3 className="font-bold text-slate-900 text-base">
                {t('polishModalTitle', 'AI Korektura a stylistika (Gemma 4)')}
              </h3>
              <p className="text-xs text-slate-600">
                {t('polishModalSubtitle', 'Vylepšení textů, gramatiky a tónu se zachováním grafiky')}
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-200/50 transition-colors"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Content */}
        <div className="p-6 overflow-y-auto space-y-5 flex-1">
          {error && (
            <div className="p-3.5 rounded-xl bg-rose-50 border border-rose-200 text-rose-800 text-xs flex items-center gap-2">
              <AlertTriangle className="w-4 h-4 text-rose-600 shrink-0" />
              <span>{error}</span>
            </div>
          )}

          {!previewResults ? (
            <div className="space-y-5">
              {/* Polish Mode */}
              <div>
                <label className="block text-xs font-bold uppercase tracking-wider text-slate-700 mb-2">
                  {t('polishModalTitle', 'Zvolte typ úpravy')}
                </label>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                  {[
                    {
                      id: 'grammar',
                      title: t('polishModeGrammar', 'Gramatika a pravopis'),
                      desc: t('polishDescGrammar', 'Opraví pravopis, interpunkci a chyby'),
                    },
                    {
                      id: 'concise',
                      title: t('polishModeConcise', 'Stručně a úderně'),
                      desc: t('polishDescConcise', 'Zkrátí text do úderných odrážek'),
                    },
                    {
                      id: 'professional',
                      title: t('polishModeProfessional', 'Manažerský styl'),
                      desc: t('polishDescProfessional', 'Uhladí formulace do formálního stylu'),
                    },
                    {
                      id: 'engaging',
                      title: t('polishModeEngaging', 'Poutavý a moderní'),
                      desc: t('polishDescEngaging', 'Oživí text dynamickými výrazy'),
                    },
                  ].map((m) => (
                    <button
                      key={m.id}
                      type="button"
                      onClick={() => setMode(m.id as any)}
                      className={`p-3.5 rounded-xl border text-left transition-all ${
                        mode === m.id
                          ? 'bg-amber-50/80 border-amber-500 ring-2 ring-amber-500/20'
                          : 'bg-white border-slate-200 hover:border-slate-300'
                      }`}
                    >
                      <p className="text-xs font-bold text-slate-900">{m.title}</p>
                      <p className="text-[11px] text-slate-500 mt-1 leading-relaxed">{m.desc}</p>
                    </button>
                  ))}
                </div>
              </div>

              {/* Scope */}
              <div>
                <label className="block text-xs font-bold uppercase tracking-wider text-slate-700 mb-2">
                  {t('scopeLabel', 'Rozsah korektury')}
                </label>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                  <button
                    type="button"
                    onClick={() => setScope('current')}
                    className={`p-3 rounded-xl border text-left text-xs transition-all ${
                      scope === 'current'
                        ? 'bg-amber-50/80 border-amber-400 font-semibold text-amber-950'
                        : 'border-slate-200 hover:bg-slate-50 text-slate-700'
                    }`}
                  >
                    #{currentSlideNumber}
                  </button>

                  <button
                    type="button"
                    onClick={() => setScope('all')}
                    className={`p-3 rounded-xl border text-left text-xs transition-all ${
                      scope === 'all'
                        ? 'bg-amber-50/80 border-amber-400 font-semibold text-amber-950'
                        : 'border-slate-200 hover:bg-slate-50 text-slate-700'
                    }`}
                  >
                    {t('scopeAll', 'Celá prezentace')} ({document.slides.length})
                  </button>

                  {initialParagraph && (
                    <button
                      type="button"
                      onClick={() => setScope('single')}
                      className={`p-3 rounded-xl border text-left text-xs transition-all ${
                        scope === 'single'
                          ? 'bg-amber-50/80 border-amber-400 font-semibold text-amber-950'
                          : 'border-slate-200 hover:bg-slate-50 text-slate-700'
                      }`}
                    >
                      {t('scopeSingle', 'Pouze vybraný text')}
                    </button>
                  )}
                </div>
              </div>
            </div>
          ) : (
            /* Results preview */
            <div className="space-y-4">
              <div className="flex items-center justify-between pb-2 border-b border-slate-200">
                <p className="text-xs font-bold text-slate-900">
                  {previewResults.length} {t('itemsCountLabel', 'položek')}
                </p>
                <span className="text-xs text-amber-800 bg-amber-50 px-2 py-0.5 rounded-md border border-amber-200 font-medium">
                  {mode}
                </span>
              </div>

              <div className="space-y-3 max-h-72 overflow-y-auto pr-1">
                {previewResults.map((item) => (
                  <div
                    key={item.id}
                    className="p-3 rounded-xl border border-slate-200 bg-slate-50 text-xs space-y-1.5"
                  >
                    <p className="text-slate-900 font-medium whitespace-pre-wrap">{item.polishedText}</p>
                    {item.changesSummary && (
                      <p className="text-slate-500 text-[11px] italic">
                        {item.changesSummary}
                      </p>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 text-xs font-semibold text-slate-600 hover:text-slate-900 transition-colors"
          >
            {t('cancelBtn', 'Zrušit')}
          </button>

          {!previewResults ? (
            <button
              type="button"
              onClick={handleStartPolish}
              disabled={isLoading || itemsCount === 0}
              className="px-5 py-2.5 rounded-xl bg-amber-600 hover:bg-amber-700 text-white font-bold text-xs flex items-center gap-2 shadow-md shadow-amber-600/20 transition-all disabled:opacity-50"
            >
              {isLoading ? (
                <>
                  <Loader2 className="w-4 h-4 animate-spin" />
                  <span>({itemsCount})...</span>
                </>
              ) : (
                <>
                  <Sparkles className="w-4 h-4" />
                  <span>{t('startPolishBtn', 'Spustit korekturu')} ({itemsCount})</span>
                </>
              )}
            </button>
          ) : (
            <button
              type="button"
              onClick={handleApply}
              className="px-5 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs flex items-center gap-2 shadow-md shadow-emerald-600/20 transition-all"
            >
              <Check className="w-4 h-4" />
              <span>{t('applyPolishBtn', 'Použít úpravy do prezentace')}</span>
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

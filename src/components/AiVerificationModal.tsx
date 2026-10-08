import React, { useState } from 'react';
import {
  X,
  ShieldCheck,
  AlertTriangle,
  CheckCircle2,
  HelpCircle,
  Loader2,
  Sparkles,
  ArrowRight,
  Check,
} from 'lucide-react';
import { PptxDocument, VerificationItem } from '../types/pptx';
import {
  requestAiVerification,
  VerificationResponse,
  VerificationResultItem,
} from '../services/aiService';
import { useLanguage } from '../context/LanguageContext';

interface AiVerificationModalProps {
  isOpen: boolean;
  onClose: () => void;
  document: PptxDocument;
  onApplyCorrection: (paragraphId: string, correctedText: string) => void;
}

export const AiVerificationModal: React.FC<AiVerificationModalProps> = ({
  isOpen,
  onClose,
  document,
  onApplyCorrection,
}) => {
  const { t } = useLanguage();
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [verificationData, setVerificationData] = useState<VerificationResponse | null>(null);

  if (!isOpen) return null;

  const handleStartAudit = async () => {
    setIsLoading(true);
    setError(null);

    const items: VerificationItem[] = [];
    document.slides.forEach((slide) => {
      slide.shapes.forEach((shape) => {
        shape.paragraphs.forEach((p) => {
          if (p.currentText.trim()) {
            items.push({
              id: p.id,
              slideNumber: slide.slideNumber,
              originalText: p.originalText,
              currentText: p.currentText,
            });
          }
        });
      });
    });

    if (items.length === 0) {
      setError('Nebyly nalezeny žádné texty k ověření.');
      setIsLoading(false);
      return;
    }

    try {
      const result = await requestAiVerification(items, 'aktuální jazyk prezentace');
      setVerificationData(result);
    } catch (err: any) {
      setError(err.message || 'Nepodařilo se dokončit kontextové ověření.');
    } finally {
      setIsLoading(false);
    }
  };

  const applySingle = (item: VerificationResultItem) => {
    if (item.suggestion) {
      onApplyCorrection(item.id, item.suggestion);
      // Mark as ok in UI
      setVerificationData((prev) => {
        if (!prev) return null;
        return {
          ...prev,
          verifiedItems: prev.verifiedItems.map((vi) =>
            vi.id === item.id ? { ...vi, status: 'ok', issue: 'Oprava aplikována' } : vi
          ),
        };
      });
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-150">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-3xl w-full max-h-[90vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="px-6 py-4 bg-gradient-to-r from-teal-50 to-emerald-50 border-b border-teal-100 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-teal-600 text-white flex items-center justify-center shadow-md shadow-teal-600/20">
              <ShieldCheck className="w-5 h-5" />
            </div>
            <div>
              <h3 className="font-bold text-slate-900 text-base">
                {t('verificationModalTitle', 'Kontextový Audit a Kontrola Pozic (Gemma 4)')}
              </h3>
              <p className="text-xs text-slate-600">
                {t('verificationModalSubtitle', 'Kontrola rozvržení snímků, soudržnosti a přetékání textových polí')}
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

          {!verificationData ? (
            <div className="space-y-4">
              <div className="p-4 rounded-xl bg-slate-50 border border-slate-200 text-xs text-slate-600 space-y-3">
                <p className="font-semibold text-slate-900 text-sm">
                  Co přesně tento audit zkontroluje?
                </p>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-slate-700">
                  <div className="p-3 rounded-lg bg-white border border-slate-200 space-y-1">
                    <p className="font-bold text-teal-800 flex items-center gap-1.5">
                      <CheckCircle2 className="w-3.5 h-3.5 text-teal-600" />
                      Návaznost a logika slov
                    </p>
                    <p className="text-[11px] text-slate-500">
                      Ověří, zda na sebe výrazy přirozeně navazují a dávají přesný smysl v kontextu tématu prezentace.
                    </p>
                  </div>

                  <div className="p-3 rounded-lg bg-white border border-slate-200 space-y-1">
                    <p className="font-bold text-teal-800 flex items-center gap-1.5">
                      <CheckCircle2 className="w-3.5 h-3.5 text-teal-600" />
                      Jednotná terminologie
                    </p>
                    <p className="text-[11px] text-slate-500">
                      Zkontroluje, zda stejné pojmy nebo odborné fráze nejsou na různých slidech přeloženy odlišně.
                    </p>
                  </div>

                  <div className="p-3 rounded-lg bg-white border border-slate-200 space-y-1">
                    <p className="font-bold text-teal-800 flex items-center gap-1.5">
                      <CheckCircle2 className="w-3.5 h-3.5 text-teal-600" />
                      Neporušený význam originálu
                    </p>
                    <p className="text-[11px] text-slate-500">
                      Detekuje případný významový posun nebo zavádějící vyznění po překladu či manuální úpravě.
                    </p>
                  </div>

                  <div className="p-3 rounded-lg bg-white border border-slate-200 space-y-1">
                    <p className="font-bold text-teal-800 flex items-center gap-1.5">
                      <CheckCircle2 className="w-3.5 h-3.5 text-teal-600" />
                      Doporučení 1-klikem
                    </p>
                    <p className="text-[11px] text-slate-500">
                      Při nalezení nesrovnalosti AI vygeneruje opravenou variantu, kterou můžete okamžitě aplikovat.
                    </p>
                  </div>
                </div>
              </div>

              <div className="text-center pt-2">
                <button
                  type="button"
                  onClick={handleStartAudit}
                  disabled={isLoading}
                  className="px-6 py-3 rounded-xl bg-teal-600 hover:bg-teal-700 text-white font-bold text-sm shadow-md shadow-teal-600/20 inline-flex items-center gap-2 transition-all disabled:opacity-50"
                >
                  {isLoading ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" />
                      <span>{t('startAuditBtn', 'Spustit audit prezentace')}...</span>
                    </>
                  ) : (
                    <>
                      <ShieldCheck className="w-4 h-4" />
                      <span>{t('startAuditBtn', 'Spustit audit prezentace')}</span>
                    </>
                  )}
                </button>
              </div>
            </div>
          ) : (
            <div className="space-y-4">
              {/* Overall Assessment Card */}
              <div className="p-4 rounded-xl bg-teal-50/80 border border-teal-200 text-xs">
                <p className="font-bold text-teal-950 text-sm">
                  {t('contextAudit', 'Kontextový Audit')}
                </p>
                <p className="text-teal-900 mt-1 leading-relaxed">
                  {verificationData.overallAssessment}
                </p>
              </div>

              {/* Items List */}
              <div className="space-y-3 max-h-80 overflow-y-auto pr-1">
                {verificationData.verifiedItems.map((item, idx) => {
                  const isOk = item.status === 'ok';
                  const isWarning = item.status === 'warning';

                  return (
                    <div
                      key={item.id}
                      className={`p-3.5 rounded-xl border text-xs space-y-2 transition-all ${
                        isOk
                          ? 'bg-slate-50 border-slate-200'
                          : isWarning
                          ? 'bg-rose-50/60 border-rose-200'
                          : 'bg-amber-50/60 border-amber-200'
                      }`}
                    >
                      <div className="flex items-center justify-between">
                        <span className="font-semibold text-slate-600 text-[11px]">
                          #{idx + 1}
                        </span>
                        <span
                          className={`px-2 py-0.5 rounded text-[10px] font-bold uppercase tracking-wider ${
                            isOk
                              ? 'bg-emerald-100 text-emerald-800'
                              : isWarning
                              ? 'bg-rose-100 text-rose-800'
                              : 'bg-amber-100 text-amber-800'
                          }`}
                        >
                          {isOk ? t('auditStatusOk', 'V pořádku') : isWarning ? t('auditStatusWarning', 'Varování') : t('auditStatusSuggestion', 'Doporučení')}
                        </span>
                      </div>

                      {item.issue && (
                        <p className="text-slate-800 font-medium">
                          {item.issue}
                        </p>
                      )}

                      {item.suggestion && (
                        <div className="p-2.5 rounded-lg bg-white border border-slate-200 flex items-center justify-between gap-3">
                          <div className="min-w-0">
                            <span className="text-[10px] uppercase font-bold text-slate-400 block">
                              {t('auditorSuggestion', 'Doporučení auditora:')}
                            </span>
                            <span className="text-slate-900 font-semibold text-xs">
                              „{item.suggestion}“
                            </span>
                          </div>
                          <button
                            type="button"
                            onClick={() => applySingle(item)}
                            className="px-2.5 py-1.5 rounded-lg bg-teal-600 hover:bg-teal-700 text-white font-bold text-[11px] shrink-0 transition-colors shadow-xs"
                          >
                            {t('applySuggestion', 'Použít návrh')}
                          </button>
                        </div>
                      )}
                    </div>
                  );
                })}
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
            {t('cancelBtn', 'Zavřít')}
          </button>

          {verificationData && (
            <button
              type="button"
              onClick={handleStartAudit}
              disabled={isLoading}
              className="px-4 py-2 rounded-xl text-teal-700 hover:bg-teal-100/70 border border-teal-200 text-xs font-bold transition-colors"
            >
              {t('startAuditBtn', 'Spustit znovu')}
            </button>
          )}
        </div>
      </div>
    </div>
  );
};

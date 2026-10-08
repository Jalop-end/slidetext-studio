import React, { useState } from 'react';
import { X, Search, Replace, Check, AlertCircle, ArrowRight } from 'lucide-react';
import { PptxDocument } from '../types/pptx';
import { useLanguage } from '../context/LanguageContext';

interface FindReplaceModalProps {
  isOpen: boolean;
  onClose: () => void;
  document: PptxDocument;
  onReplaceAll: (findText: string, replaceText: string, matchCase: boolean) => number;
}

export const FindReplaceModal: React.FC<FindReplaceModalProps> = ({
  isOpen,
  onClose,
  document,
  onReplaceAll,
}) => {
  const { t } = useLanguage();
  const [findText, setFindText] = useState('');
  const [replaceText, setReplaceText] = useState('');
  const [matchCase, setMatchCase] = useState(false);
  const [resultCount, setResultCount] = useState<number | null>(null);

  if (!isOpen) return null;

  // Calculate occurrences
  const countOccurrences = () => {
    if (!findText.trim()) return 0;
    let count = 0;
    document.slides.forEach((slide) => {
      slide.shapes.forEach((shape) => {
        shape.paragraphs.forEach((p) => {
          const target = matchCase ? p.currentText : p.currentText.toLowerCase();
          const query = matchCase ? findText : findText.toLowerCase();
          if (query) {
            let pos = target.indexOf(query);
            while (pos !== -1) {
              count++;
              pos = target.indexOf(query, pos + query.length);
            }
          }
        });
      });
    });
    return count;
  };

  const occurrences = countOccurrences();

  const handleExecute = () => {
    if (!findText.trim()) return;
    const replaced = onReplaceAll(findText, replaceText, matchCase);
    setResultCount(replaced);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-150">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-md w-full overflow-hidden">
        {/* Header */}
        <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-blue-100 text-blue-700 flex items-center justify-center">
              <Replace className="w-4 h-4" />
            </div>
            <div>
              <h3 className="font-bold text-slate-900 text-sm">{t('findReplaceTitle', 'Najít a Nahradit v celé prezentaci')}</h3>
              <p className="text-[11px] text-slate-500">{t('findReplace', 'Najít & Nahradit')}</p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded text-slate-400 hover:text-slate-700 hover:bg-slate-200 transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Form */}
        <div className="p-6 space-y-4">
          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1">
              {t('findLabel', 'Hledat')}
            </label>
            <div className="relative">
              <input
                type="text"
                value={findText}
                onChange={(e) => {
                  setFindText(e.target.value);
                  setResultCount(null);
                }}
                placeholder="..."
                className="w-full text-xs p-2.5 pl-8 rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500"
              />
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-3" />
            </div>
            {findText && (
              <p className="text-[11px] text-slate-500 mt-1">
                ({occurrences})
              </p>
            )}
          </div>

          <div>
            <label className="block text-xs font-semibold text-slate-700 mb-1">
              {t('replaceLabel', 'Nahradit za')}
            </label>
            <div className="relative">
              <input
                type="text"
                value={replaceText}
                onChange={(e) => {
                  setReplaceText(e.target.value);
                  setResultCount(null);
                }}
                placeholder="..."
                className="w-full text-xs p-2.5 pl-8 rounded-lg border border-slate-300 focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500"
              />
              <Replace className="w-3.5 h-3.5 text-slate-400 absolute left-2.5 top-3" />
            </div>
          </div>

          <div className="flex items-center gap-2 pt-1">
            <input
              type="checkbox"
              id="matchCase"
              checked={matchCase}
              onChange={(e) => setMatchCase(e.target.checked)}
              className="rounded border-slate-300 text-blue-600 focus:ring-blue-500"
            />
            <label htmlFor="matchCase" className="text-xs text-slate-600 select-none">
              {t('matchCase', 'Rozlišovat velikost písmen')}
            </label>
          </div>

          {resultCount !== null && (
            <div className="p-3 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs flex items-center gap-2">
              <Check className="w-4 h-4 text-emerald-600 shrink-0" />
              <span>{resultCount} {t('replacedCount', 'nahrazeno výskytů')}</span>
            </div>
          )}
        </div>

        {/* Footer */}
        <div className="px-6 py-3.5 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
          <button
            type="button"
            onClick={onClose}
            className="px-3 py-1.5 text-xs font-semibold text-slate-600 hover:text-slate-900 transition-colors"
          >
            {t('cancelBtn', 'Zavřít')}
          </button>

          <button
            type="button"
            onClick={handleExecute}
            disabled={!findText.trim() || occurrences === 0}
            className="px-4 py-2 rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs flex items-center gap-1.5 shadow-sm transition-all disabled:opacity-50"
          >
            <Replace className="w-3.5 h-3.5" />
            <span>{t('replaceAllBtn', 'Nahradit vše')} ({occurrences})</span>
          </button>
        </div>
      </div>
    </div>
  );
};

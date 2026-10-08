import React, { useState } from 'react';
import { X, Table, Download, Copy, Check, FileText } from 'lucide-react';
import { PptxDocument } from '../types/pptx';
import { exportBilingualReport } from '../services/pptxParser';
import { useLanguage } from '../context/LanguageContext';

interface BilingualExportModalProps {
  isOpen: boolean;
  onClose: () => void;
  document: PptxDocument;
}

export const BilingualExportModal: React.FC<BilingualExportModalProps> = ({
  isOpen,
  onClose,
  document,
}) => {
  const { t } = useLanguage();
  const [format, setFormat] = useState<'html' | 'csv' | 'markdown'>('html');
  const [copied, setCopied] = useState(false);

  if (!isOpen) return null;

  const content = exportBilingualReport(document, format);

  const handleCopy = () => {
    navigator.clipboard.writeText(content);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleDownload = () => {
    const ext = format === 'html' ? 'html' : format === 'csv' ? 'csv' : 'md';
    const mime =
      format === 'html'
        ? 'text/html'
        : format === 'csv'
        ? 'text/csv;charset=utf-8;'
        : 'text/markdown';

    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = window.document.createElement('a');
    a.href = url;
    a.download = `${document.fileName.replace(/\.pptx$/i, '')}_srovnani.${ext}`;
    window.document.body.appendChild(a);
    a.click();
    window.document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-xs animate-in fade-in duration-150">
      <div className="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-2xl w-full max-h-[90vh] flex flex-col overflow-hidden">
        {/* Header */}
        <div className="px-6 py-4 bg-slate-50 border-b border-slate-200 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded-lg bg-teal-100 text-teal-700 flex items-center justify-center">
              <Table className="w-4 h-4" />
            </div>
            <div>
              <h3 className="font-bold text-slate-900 text-sm">
                {t('bilingualTitle', 'Export bilingvní tabulky (Původní vs. Nový)')}
              </h3>
              <p className="text-[11px] text-slate-500">
                {t('bilingualSubtitle', 'Přehledné srovnání textů po snímcích pro revizi a překladatelské QA')}
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded text-slate-400 hover:text-slate-700 hover:bg-slate-200 transition-colors"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* Format Selector */}
        <div className="p-6 space-y-4 overflow-y-auto flex-1">
          <div className="flex items-center gap-2">
            {[
              { id: 'html', label: 'HTML Report (Tabulka)' },
              { id: 'csv', label: 'CSV (Excel / Tabulky)' },
              { id: 'markdown', label: 'Markdown (.md)' },
            ].map((f) => (
              <button
                key={f.id}
                type="button"
                onClick={() => setFormat(f.id as any)}
                className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ${
                  format === f.id
                    ? 'bg-teal-600 text-white shadow-xs'
                    : 'bg-slate-100 text-slate-600 hover:bg-slate-200'
                }`}
              >
                {f.label}
              </button>
            ))}
          </div>

          <div className="relative">
            <textarea
              readOnly
              value={content}
              rows={12}
              className="w-full text-xs font-mono p-3 rounded-xl border border-slate-200 bg-slate-50 text-slate-700 focus:outline-none resize-none"
            />
          </div>
        </div>

        {/* Footer */}
        <div className="px-6 py-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between">
          <button
            type="button"
            onClick={onClose}
            className="px-3 py-1.5 text-xs font-semibold text-slate-600 hover:text-slate-900 transition-colors"
          >
            {t('cancelBtn', 'Zavřít')}
          </button>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={handleCopy}
              className="px-3.5 py-2 rounded-xl bg-slate-200 hover:bg-slate-300 text-slate-800 font-semibold text-xs flex items-center gap-1.5 transition-colors"
            >
              {copied ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
              <span>{copied ? 'OK' : t('copyTable', 'Kopírovat do schránky')}</span>
            </button>

            <button
              type="button"
              onClick={handleDownload}
              className="px-4 py-2 rounded-xl bg-teal-600 hover:bg-teal-700 text-white font-bold text-xs flex items-center gap-1.5 shadow-sm transition-colors"
            >
              <Download className="w-3.5 h-3.5" />
              <span>{t('downloadCsv', 'Stáhnout')} (.{format})</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

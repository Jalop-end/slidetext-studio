/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState } from 'react';
import { PptxDocument, PptxParagraph } from './types/pptx';
import { parsePptxFile, updatePptxAndExport, renderPptxSlidesViaBackend } from './services/pptxParser';
import { parseDocxFile, updateDocxAndExport, renderDocxPagesViaBackend } from './services/docxParser';
import { generateSamplePptx } from './services/samplePptxGenerator';
import { generateSampleDocx } from './services/sampleDocxGenerator';
import { Header } from './components/Header';
import { DocumentEditor } from './components/DocumentEditor';
import { SlideVisualizer } from './components/SlideVisualizer';
import { UploadDropzone } from './components/UploadDropzone';
import { AiTranslationModal } from './components/AiTranslationModal';
import { AiPolishModal } from './components/AiPolishModal';
import { AiVerificationModal } from './components/AiVerificationModal';
import { FindReplaceModal } from './components/FindReplaceModal';
import { BilingualExportModal } from './components/BilingualExportModal';
import { PreSaveComparisonModal } from './components/PreSaveComparisonModal';
import { TranslationResultItem, PolishResultItem } from './services/aiService';
import { CheckCircle2, Download, AlertCircle } from 'lucide-react';
import { useLanguage } from './context/LanguageContext';

export default function App() {
  const { t } = useLanguage();
  const [document, setDocument] = useState<PptxDocument | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [loadingMessage, setLoadingMessage] = useState<string>('');
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [currentSlideNumber, setCurrentSlideNumber] = useState<number>(1);
  const [selectedParagraphId, setSelectedParagraphId] = useState<string | null>(null);
  const [activeView, setActiveView] = useState<'split' | 'document' | 'slide'>('split');
  const [isRenderingPptx, setIsRenderingPptx] = useState<boolean>(false);

  // Modals state
  const [isTranslateOpen, setIsTranslateOpen] = useState(false);
  const [isPolishOpen, setIsPolishOpen] = useState(false);
  const [isVerifyOpen, setIsVerifyOpen] = useState(false);
  const [isFindReplaceOpen, setIsFindReplaceOpen] = useState(false);
  const [isBilingualExportOpen, setIsBilingualExportOpen] = useState(false);
  const [isPreSaveOpen, setIsPreSaveOpen] = useState(false);
  const [isGalleryOpen, setIsGalleryOpen] = useState(false);
  const [activeSingleParagraph, setActiveSingleParagraph] = useState<PptxParagraph | null>(null);
  const [translateScope, setTranslateScope] = useState<'all' | 'current' | 'single'>('all');

  // History stack for Undo / Redo
  const [history, setHistory] = useState<PptxDocument[]>([]);
  const [historyIndex, setHistoryIndex] = useState<number>(-1);

  const handleUndo = () => {
    if (historyIndex > 0) {
      const nextIdx = historyIndex - 1;
      setHistoryIndex(nextIdx);
      setDocument(history[nextIdx]);
      showToast(t('stepUndone', 'Krok vrácen zpět.'));
    }
  };

  const handleRedo = () => {
    if (historyIndex < history.length - 1) {
      const nextIdx = historyIndex + 1;
      setHistoryIndex(nextIdx);
      setDocument(history[nextIdx]);
      showToast(t('stepRedone', 'Krok obnoven.'));
    }
  };

  // Push new state to history
  const commitDocumentChange = (newDoc: PptxDocument) => {
    setHistory((prev) => {
      const nextSlice = prev.slice(0, historyIndex + 1);
      // Keep up to 30 steps
      const newStack = [...nextSlice, newDoc];
      return newStack.length > 30 ? newStack.slice(newStack.length - 30) : newStack;
    });
    setHistoryIndex((prev) => prev + 1);
    setDocument(newDoc);
  };

  // Success toast
  const [toastMessage, setToastMessage] = useState<string | null>(null);

  const showToast = (msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage(null);
    }, 4000);
  };

  // Upload handlers
  const handleFileUpload = async (file: File) => {
    setIsLoading(true);
    const isWord = file.name.toLowerCase().endsWith('.docx');
    setLoadingMessage(
      isWord
        ? `Načítám Word dokument... (${file.name})`
        : `${t('loadingPresentation', 'Načítám prezentaci...')} (${file.name})`
    );

    try {
      const arrayBuf = await file.arrayBuffer();

      if (isWord) {
        const parsed = await parseDocxFile(arrayBuf, file.name);
        parsed.originalBuffer = arrayBuf;
        setDocument(parsed);
        setHistory([parsed]);
        setHistoryIndex(0);
        setCurrentSlideNumber(1);
        setSelectedParagraphId(null);
        setIsLoading(false);
        showToast(`Dokument Word načten: ${parsed.slides.length} stránek. Načítám věrný náhled...`);

        // Asynchronously render Word page previews via backend
        setIsRenderingPptx(true);
        try {
          const res = await renderDocxPagesViaBackend(arrayBuf);
          if (res.success && res.pageImages.length > 0) {
            setDocument((prev) => {
              if (!prev) return null;
              const updatedSlides = prev.slides.map((s, idx) => ({
                ...s,
                renderedImageDataUrl: res.pageImages[idx] || s.renderedImageDataUrl,
                originalRenderedImageDataUrl: res.pageImages[idx] || s.originalRenderedImageDataUrl,
              }));
              const updatedDoc = { ...prev, slides: updatedSlides };
              setHistory([updatedDoc]);
              return updatedDoc;
            });
            showToast(`Word náhled připraven (${res.pageImages.length} stránek).`);
          }
        } catch (renderErr) {
          console.warn('Word background render error:', renderErr);
        } finally {
          setIsRenderingPptx(false);
        }
        return;
      }

      // PPTX parsing
      const parsed = await parsePptxFile(arrayBuf, file.name);
      parsed.originalBuffer = arrayBuf;
      setDocument(parsed);
      setHistory([parsed]);
      setHistoryIndex(0);
      setCurrentSlideNumber(1);
      setSelectedParagraphId(null);
      setIsLoading(false);
      showToast(`Prezentace načtena: ${parsed.slides.length} snímků. Načítám věrný náhled...`);

      // Asynchronously render high-fidelity PowerPoint images via backend
      setIsRenderingPptx(true);
      try {
        const renderedImages = await renderPptxSlidesViaBackend(arrayBuf);
        if (renderedImages.length > 0) {
          setDocument((prev) => {
            if (!prev) return null;
            const updatedSlides = prev.slides.map((s, idx) => ({
              ...s,
              renderedImageDataUrl: renderedImages[idx] || s.renderedImageDataUrl,
              originalRenderedImageDataUrl: renderedImages[idx] || s.originalRenderedImageDataUrl,
            }));
            const updatedDoc = { ...prev, slides: updatedSlides };
            setHistory([updatedDoc]);
            return updatedDoc;
          });
          showToast(`PowerPoint náhled připraven (${renderedImages.length} snímků v 1080p).`);
        }
      } catch (renderErr) {
        console.warn('PowerPoint background render error:', renderErr);
      } finally {
        setIsRenderingPptx(false);
      }
    } catch (error: any) {
      console.error('Chyba při zpracování souboru:', error);
      alert('Nepodařilo se načíst soubor: ' + (error.message || 'Ujistěte se, že jde o platný soubor PPTX nebo DOCX.'));
      setIsLoading(false);
      setLoadingMessage('');
    }
  };

  const handleLoadSample = async (sampleType: 'czech-strategy' | 'english-pitch' | 'word-strategy') => {
    setIsLoading(true);

    if (sampleType === 'word-strategy') {
      setLoadingMessage('Generuji vzorový Word dokument...');
      try {
        const sampleBlob = await generateSampleDocx();
        const arrayBuf = await sampleBlob.arrayBuffer();
        const fileName = 'Vyrocni_zprava_2026.docx';
        const parsed = await parseDocxFile(arrayBuf, fileName);
        parsed.originalBuffer = arrayBuf;
        setDocument(parsed);
        setHistory([parsed]);
        setHistoryIndex(0);
        setCurrentSlideNumber(1);
        setSelectedParagraphId(null);
        setIsLoading(false);
        showToast(`Vzorový Word dokument načten (${parsed.slides.length} stránek).`);

        setIsRenderingPptx(true);
        try {
          const res = await renderDocxPagesViaBackend(arrayBuf);
          if (res.success && res.pageImages.length > 0) {
            setDocument((prev) => {
              if (!prev) return null;
              const updatedSlides = prev.slides.map((s, idx) => ({
                ...s,
                renderedImageDataUrl: res.pageImages[idx] || s.renderedImageDataUrl,
                originalRenderedImageDataUrl: res.pageImages[idx] || s.originalRenderedImageDataUrl,
              }));
              const updatedDoc = { ...prev, slides: updatedSlides };
              setHistory([updatedDoc]);
              return updatedDoc;
            });
            showToast(`Věrný Word náhled připraven (${res.pageImages.length} stránek).`);
          }
        } catch (renderErr) {
          console.warn('Word sample background render error:', renderErr);
        } finally {
          setIsRenderingPptx(false);
        }
      } catch (err: any) {
        console.error('Chyba při vzoru Word:', err);
        alert('Nepodařilo se vygenerovat vzorový Word dokument.');
        setIsLoading(false);
        setLoadingMessage('');
      }
      return;
    }

    setLoadingMessage(t('loadingPresentation', 'Načítám vzorovou prezentaci...'));
    try {
      const sample = await generateSamplePptx(sampleType);
      const arrayBuf = await sample.blob.arrayBuffer();
      const parsed = await parsePptxFile(arrayBuf, sample.fileName);
      parsed.originalBuffer = arrayBuf;
      setDocument(parsed);
      setHistory([parsed]);
      setHistoryIndex(0);
      setCurrentSlideNumber(1);
      setSelectedParagraphId(null);
      setIsLoading(false);
      showToast(`Vzorová prezentace načtena (${parsed.slides.length} snímků).`);

      // Asynchronously render high-fidelity PowerPoint images via backend
      setIsRenderingPptx(true);
      try {
        const renderedImages = await renderPptxSlidesViaBackend(arrayBuf);
        if (renderedImages.length > 0) {
          setDocument((prev) => {
            if (!prev) return null;
            const updatedSlides = prev.slides.map((s, idx) => ({
              ...s,
              renderedImageDataUrl: renderedImages[idx] || s.renderedImageDataUrl,
              originalRenderedImageDataUrl: renderedImages[idx] || s.originalRenderedImageDataUrl,
            }));
            const updatedDoc = { ...prev, slides: updatedSlides };
            setHistory([updatedDoc]);
            return updatedDoc;
          });
          showToast(`Věrný PowerPoint náhled připraven pro všech ${renderedImages.length} snímků.`);
        }
      } catch (renderErr) {
        console.warn('Sample background render error:', renderErr);
      } finally {
        setIsRenderingPptx(false);
      }
    } catch (error: any) {
      console.error('Chyba při generování vzoru:', error);
      alert('Nepodařilo se načíst vzorovou prezentaci.');
      setIsLoading(false);
      setLoadingMessage('');
    }
  };

  // Re-render modified presentation or document via backend
  const handleRefreshRender = async () => {
    if (!document) return;
    const isWord = document.docType === 'docx';
    setIsRenderingPptx(true);
    showToast(
      isWord
        ? 'Generuji aktuální Word náhled se všemi provedenými změnami...'
        : t('reRenderStarted', 'Generuji aktuální PowerPoint náhled se všemi provedenými změnami...')
    );

    try {
      if (isWord) {
        const updatedBlob = await updateDocxAndExport(document.rawZip, document);
        const res = await renderDocxPagesViaBackend(updatedBlob);
        if (res.success && res.pageImages.length > 0) {
          setDocument((prev) => {
            if (!prev) return null;
            const updatedSlides = prev.slides.map((s, idx) => ({
              ...s,
              renderedImageDataUrl: res.pageImages[idx] || s.renderedImageDataUrl,
            }));
            return { ...prev, slides: updatedSlides };
          });
          showToast('Word náhled byl úspěšně zaktualizován.');
        } else {
          showToast('Používá se interaktivní náhled dokumentu.');
        }
      } else {
        const updatedBlob = await updatePptxAndExport(document.rawZip, document.slides);
        const renderedImages = await renderPptxSlidesViaBackend(updatedBlob);
        if (renderedImages.length > 0) {
          setDocument((prev) => {
            if (!prev) return null;
            const updatedSlides = prev.slides.map((s, idx) => ({
              ...s,
              renderedImageDataUrl: renderedImages[idx] || s.renderedImageDataUrl,
            }));
            return { ...prev, slides: updatedSlides };
          });
          showToast(t('reRenderSuccess', 'PowerPoint náhled byl úspěšně zaktualizován.'));
        } else {
          showToast('Náhled je aktuální.');
        }
      }
    } catch (err: any) {
      console.warn('Re-render error:', err);
      showToast('Aktualizace náhledu se nezdařila, používá se interaktivní zobrazení.');
    } finally {
      setIsRenderingPptx(false);
    }
  };

  // Document state mutation helpers
  const handleUpdateParagraphText = (paragraphId: string, newText: string) => {
    if (!document) return;

    setDocument((prev) => {
      if (!prev) return null;

      let modCount = 0;
      let wordsCount = 0;

      const updatedSlides = prev.slides.map((slide) => {
        const updatedShapes = slide.shapes.map((shape) => {
          const updatedParagraphs = shape.paragraphs.map((p) => {
            if (p.id === paragraphId) {
              const isModified = newText !== p.originalText;
              if (isModified) modCount++;
              wordsCount += newText.trim().split(/\s+/).filter(Boolean).length;
              return {
                ...p,
                currentText: newText,
                isModified,
              };
            } else {
              if (p.isModified) modCount++;
              wordsCount += p.currentText.trim().split(/\s+/).filter(Boolean).length;
              return p;
            }
          });
          return { ...shape, paragraphs: updatedParagraphs };
        });
        return { ...slide, shapes: updatedShapes };
      });

      return {
        ...prev,
        slides: updatedSlides,
        modifiedCount: modCount,
        totalWords: wordsCount,
      };
    });
  };

  const handleResetParagraphText = (paragraphId: string) => {
    if (!document) return;

    setDocument((prev) => {
      if (!prev) return null;

      const updatedSlides = prev.slides.map((slide) => {
        const updatedShapes = slide.shapes.map((shape) => {
          const updatedParagraphs = shape.paragraphs.map((p) => {
            if (p.id === paragraphId) {
              return {
                ...p,
                currentText: p.originalText,
                isModified: false,
                aiNotes: undefined,
                lengthFitWarning: undefined,
              };
            }
            return p;
          });
          return { ...shape, paragraphs: updatedParagraphs };
        });
        return { ...slide, shapes: updatedShapes };
      });

      return {
        ...prev,
        slides: updatedSlides,
      };
    });
  };

  // Two-way interactive sync: paragraph click / selection
  const handleSelectParagraph = (paragraphId: string, slideNumber: number) => {
    setSelectedParagraphId(paragraphId);
    if (slideNumber !== currentSlideNumber) {
      setCurrentSlideNumber(slideNumber);
    }
  };

  // Export to .pptx or .docx preserving original visual layout & format
  const handleExportDocument = async (customFileName?: string) => {
    if (!document) return;
    setIsExporting(true);
    const isWord = document.docType === 'docx';
    const ext = isWord ? '.docx' : '.pptx';

    try {
      const outputBlob = isWord
        ? await updateDocxAndExport(document.rawZip, document)
        : await updatePptxAndExport(document.rawZip, document.slides);

      // Determine final download file name
      let downloadName = customFileName?.trim();
      if (!downloadName) {
        const originalName = document.fileName.replace(/\.(pptx|docx)$/i, '');
        downloadName = `${originalName}_upraveno${ext}`;
      } else if (!downloadName.toLowerCase().endsWith(ext)) {
        downloadName += ext;
      }

      const downloadUrl = URL.createObjectURL(outputBlob);
      const a = window.document.createElement('a');
      a.href = downloadUrl;
      a.download = downloadName;
      window.document.body.appendChild(a);
      a.click();
      window.document.body.removeChild(a);
      URL.revokeObjectURL(downloadUrl);

      showToast(`${isWord ? 'Dokument Word' : 'Prezentace'} úspěšně uložena jako ${downloadName}`);
    } catch (err: any) {
      console.error('Chyba při exportu:', err);
      alert('Nepodařilo se vyexportovat soubor: ' + (err.message || 'Neznámá chyba'));
    } finally {
      setIsExporting(false);
    }
  };

  const handleRenameDocument = (newName: string) => {
    if (!document || !newName.trim()) return;
    const isWord = document.docType === 'docx';
    const ext = isWord ? '.docx' : '.pptx';
    let clean = newName.trim().replace(/\.(pptx|docx)$/i, '');
    const finalName = `${clean}${ext}`;

    const updated = {
      ...document,
      fileName: finalName,
    };
    commitDocumentChange(updated);
    showToast(`Dokument přejmenován na: ${finalName}`);
  };

  // AI batch applications
  const handleApplyTranslations = (results: TranslationResultItem[]) => {
    if (!document) return;

    const resultMap = new Map<string, TranslationResultItem>();
    results.forEach((r) => resultMap.set(r.id, r));

    let modCount = 0;
    const updatedSlides = document.slides.map((slide) => {
      const updatedShapes = slide.shapes.map((shape) => {
        const updatedParagraphs = shape.paragraphs.map((p) => {
          if (resultMap.has(p.id)) {
            const res = resultMap.get(p.id)!;
            const textChanged = res.translatedText.trim() !== p.originalText.trim();
            modCount++;
            return {
              ...p,
              currentText: res.translatedText,
              isModified: true, // Always mark true so every translated element is tracked and marked edited
              aiNotes: res.notes || (textChanged ? 'Přeloženo' : 'Přeloženo (beze změny)'),
              lengthFitWarning: res.lengthFitWarning,
              alternatives: res.alternatives && res.alternatives.length > 0 ? res.alternatives : p.alternatives,
            };
          }
          if (p.isModified) modCount++;
          return p;
        });
        return { ...shape, paragraphs: updatedParagraphs };
      });

      // Also check speaker notes if translated
      let updatedNotes = slide.speakerNotes;
      const notesId = `slide-${slide.slideNumber}-notes`;
      if (resultMap.has(notesId) && slide.speakerNotes) {
        const res = resultMap.get(notesId)!;
        modCount++;
        updatedNotes = {
          ...slide.speakerNotes,
          text: res.translatedText,
          isModified: true,
        };
      }

      return { ...slide, shapes: updatedShapes, speakerNotes: updatedNotes };
    });

    const newDoc = {
      ...document,
      slides: updatedSlides,
      modifiedCount: modCount,
    };
    commitDocumentChange(newDoc);
    showToast(t('translationsApplied', 'Překlady byly úspěšně aplikovány do prezentace!'));
  };

  const handleApplyPolished = (results: PolishResultItem[]) => {
    if (!document) return;

    const resultMap = new Map<string, PolishResultItem>();
    results.forEach((r) => resultMap.set(r.id, r));

    setDocument((prev) => {
      if (!prev) return null;

      let modCount = 0;
      const updatedSlides = prev.slides.map((slide) => {
        const updatedShapes = slide.shapes.map((shape) => {
          const updatedParagraphs = shape.paragraphs.map((p) => {
            if (resultMap.has(p.id)) {
              const res = resultMap.get(p.id)!;
              modCount++;
              return {
                ...p,
                currentText: res.polishedText,
                isModified: res.polishedText !== p.originalText,
                aiNotes: res.changesSummary,
              };
            }
            if (p.isModified) modCount++;
            return p;
          });
          return { ...shape, paragraphs: updatedParagraphs };
        });
        return { ...slide, shapes: updatedShapes };
      });

      return {
        ...prev,
        slides: updatedSlides,
        modifiedCount: modCount,
      };
    });

    showToast(`Korektura úspěšně aplikována na ${results.length} textů.`);
  };

  const handleApplyCorrection = (paragraphId: string, correctedText: string) => {
    handleUpdateParagraphText(paragraphId, correctedText);
    showToast('Doporučená oprava byla aplikována.');
  };

  const handleUpdateSpeakerNotes = (slideNumber: number, newNotes: string) => {
    if (!document) return;
    setDocument((prev) => {
      if (!prev) return null;
      const updatedSlides = prev.slides.map((slide) => {
        if (slide.slideNumber === slideNumber) {
          return {
            ...slide,
            speakerNotes: {
              text: newNotes,
              originalText: slide.speakerNotes?.originalText || '',
              isModified: true,
              xmlPath: slide.speakerNotes?.xmlPath,
            },
          };
        }
        return slide;
      });
      return { ...prev, slides: updatedSlides };
    });
    showToast(`Poznámka ke snímku #${slideNumber} uložena.`);
  };

  const handleReplaceAll = (findText: string, replaceText: string, matchCase: boolean): number => {
    if (!document || !findText) return 0;
    let replaceCount = 0;
    let modCount = 0;

    const updatedSlides = document.slides.map((slide) => {
      const updatedShapes = slide.shapes.map((shape) => {
        const updatedParagraphs = shape.paragraphs.map((p) => {
          let updatedText = p.currentText;
          if (matchCase) {
            const parts = updatedText.split(findText);
            if (parts.length > 1) {
              replaceCount += parts.length - 1;
              updatedText = parts.join(replaceText);
            }
          } else {
            const regex = new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
            const matches = updatedText.match(regex);
            if (matches) {
              replaceCount += matches.length;
              updatedText = updatedText.replace(regex, replaceText);
            }
          }

          const isModified = updatedText !== p.originalText;
          if (isModified) modCount++;
          return {
            ...p,
            currentText: updatedText,
            isModified,
          };
        });
        return { ...shape, paragraphs: updatedParagraphs };
      });
      return { ...slide, shapes: updatedShapes };
    });

    setDocument((prev) => {
      if (!prev) return null;
      return {
        ...prev,
        slides: updatedSlides,
        modifiedCount: modCount,
      };
    });

    showToast(`Nahrazeno celkem ${replaceCount} výskytů.`);
    return replaceCount;
  };

  const handleUpdateParagraphHyperlink = (paragraphId: string, newUrl: string) => {
    if (!document) return;
    setDocument((prev) => {
      if (!prev) return null;
      let modCount = 0;
      const updatedSlides = prev.slides.map((slide) => {
        const updatedShapes = slide.shapes.map((shape) => {
          const updatedParagraphs = shape.paragraphs.map((p) => {
            if (p.id === paragraphId) {
              modCount++;
              return {
                ...p,
                hyperlink: newUrl.trim() ? { url: newUrl.trim() } : undefined,
                isModified: true,
              };
            }
            if (p.isModified) modCount++;
            return p;
          });
          return { ...shape, paragraphs: updatedParagraphs };
        });
        return { ...slide, shapes: updatedShapes };
      });
      return { ...prev, slides: updatedSlides, modifiedCount: modCount };
    });
    showToast(newUrl.trim() ? 'Odkaz k textu byl úspěšně uložen.' : 'Odkaz byl odebrán.');
  };

  // Single-paragraph & slide-level quick triggers
  const handleSingleTranslate = (paragraph: PptxParagraph) => {
    setActiveSingleParagraph(paragraph);
    setTranslateScope('single');
    setIsTranslateOpen(true);
  };

  const handleTranslateSlide = (slideNum: number) => {
    setCurrentSlideNumber(slideNum);
    setActiveSingleParagraph(null);
    setTranslateScope('current');
    setIsTranslateOpen(true);
  };

  const handleSinglePolish = (paragraph: PptxParagraph) => {
    setActiveSingleParagraph(paragraph);
    setIsPolishOpen(true);
  };

  // Keyboard shortcuts for Undo and Redo
  React.useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
        if (e.shiftKey) {
          e.preventDefault();
          handleRedo();
        } else {
          e.preventDefault();
          handleUndo();
        }
      } else if ((e.ctrlKey || e.metaKey) && e.key === 'y') {
        e.preventDefault();
        handleRedo();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [historyIndex, history]);

  // If no document is loaded, show upload screen
  if (!document) {
    return (
      <UploadDropzone
        onFileUpload={handleFileUpload}
        onLoadSample={handleLoadSample}
        isLoading={isLoading}
        loadingMessage={loadingMessage}
      />
    );
  }

  return (
    <div className="min-h-screen flex flex-col bg-slate-100 text-slate-900 font-sans">
      {/* Toast notification */}
      {toastMessage && (
        <div className="fixed bottom-6 right-6 z-50 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-2xl flex items-center gap-2.5 text-xs font-semibold animate-in slide-in-from-bottom-5 duration-200 border border-slate-700">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{toastMessage}</span>
        </div>
      )}

      {/* Main App Header */}
      <Header
        document={document}
        onExport={() => setIsPreSaveOpen(true)}
        canUndo={historyIndex > 0}
        canRedo={historyIndex < history.length - 1}
        onUndo={handleUndo}
        onRedo={handleRedo}
        onRenameDocument={handleRenameDocument}
        onOpenTranslate={() => {
          setActiveSingleParagraph(null);
          setTranslateScope('all');
          setIsTranslateOpen(true);
        }}
        onOpenPolish={() => {
          setActiveSingleParagraph(null);
          setIsPolishOpen(true);
        }}
        onOpenVerify={() => setIsVerifyOpen(true)}
        onOpenFindReplace={() => setIsFindReplaceOpen(true)}
        onOpenBilingualExport={() => setIsBilingualExportOpen(true)}
        onOpenGallery={() => setIsGalleryOpen(true)}
        onResetUpload={() => setDocument(null)}
        isExporting={isExporting}
        activeView={activeView}
        onChangeView={setActiveView}
      />

      {/* Main Workspace */}
      <main className="flex-1 flex overflow-hidden">
        {/* Left Pane: Word-Like Document Editor */}
        {(activeView === 'split' || activeView === 'document') && (
          <div
            className={`flex flex-col h-full min-h-0 min-w-0 overflow-hidden transition-all duration-200 ${
              activeView === 'split' ? 'w-full lg:w-1/2' : 'w-full'
            }`}
          >
            <DocumentEditor
              document={document}
              selectedParagraphId={selectedParagraphId}
              onSelectParagraph={handleSelectParagraph}
              onUpdateParagraphText={handleUpdateParagraphText}
              onResetParagraphText={handleResetParagraphText}
              currentSlideNumber={currentSlideNumber}
              onSlideChange={setCurrentSlideNumber}
              onSingleTranslate={handleSingleTranslate}
              onTranslateSlide={handleTranslateSlide}
              onSinglePolish={handleSinglePolish}
              onUpdateSpeakerNotes={handleUpdateSpeakerNotes}
              onUpdateParagraphHyperlink={handleUpdateParagraphHyperlink}
            />
          </div>
        )}

        {/* Right Pane: Interactive Slide Visualizer */}
        {(activeView === 'split' || activeView === 'slide') && (
          <div
            className={`flex flex-col h-full min-h-0 min-w-0 overflow-hidden transition-all duration-200 ${
              activeView === 'split' ? 'hidden lg:flex lg:w-1/2' : 'w-full'
            }`}
          >
            <SlideVisualizer
              docType={document.docType}
              slides={document.slides}
              currentSlideNumber={currentSlideNumber}
              onSlideChange={setCurrentSlideNumber}
              selectedParagraphId={selectedParagraphId}
              onSelectParagraph={handleSelectParagraph}
              onRefreshRender={handleRefreshRender}
              isRendering={isRenderingPptx}
              isGalleryOpen={isGalleryOpen}
              onToggleGallery={setIsGalleryOpen}
              onExport={() => setIsPreSaveOpen(true)}
            />
          </div>
        )}
      </main>

      {/* AI Modals – mounted only when open (keyed) so their internal state resets on every opening */}
      {isTranslateOpen && (
        <AiTranslationModal
          key={`translate-${translateScope}-${activeSingleParagraph?.id ?? currentSlideNumber}`}
          isOpen={isTranslateOpen}
          onClose={() => {
            setIsTranslateOpen(false);
            setActiveSingleParagraph(null);
          }}
          document={document}
          currentSlideNumber={currentSlideNumber}
          initialParagraph={activeSingleParagraph}
          initialScope={translateScope}
          onApplyTranslations={handleApplyTranslations}
        />
      )}

      {isPolishOpen && (
        <AiPolishModal
          key={`polish-${activeSingleParagraph?.id ?? 'all'}`}
          isOpen={isPolishOpen}
          onClose={() => {
            setIsPolishOpen(false);
            setActiveSingleParagraph(null);
          }}
          document={document}
          currentSlideNumber={currentSlideNumber}
          initialParagraph={activeSingleParagraph}
          onApplyPolished={handleApplyPolished}
        />
      )}

      <AiVerificationModal
        isOpen={isVerifyOpen}
        onClose={() => setIsVerifyOpen(false)}
        document={document}
        onApplyCorrection={handleApplyCorrection}
      />

      <FindReplaceModal
        isOpen={isFindReplaceOpen}
        onClose={() => setIsFindReplaceOpen(false)}
        document={document}
        onReplaceAll={handleReplaceAll}
      />

      <BilingualExportModal
        isOpen={isBilingualExportOpen}
        onClose={() => setIsBilingualExportOpen(false)}
        document={document}
      />

      <PreSaveComparisonModal
        isOpen={isPreSaveOpen}
        onClose={() => setIsPreSaveOpen(false)}
        document={document}
        onConfirmExport={handleExportDocument}
        isExporting={isExporting}
      />
    </div>
  );
}

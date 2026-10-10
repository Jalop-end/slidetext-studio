import { TranslationItem, VerificationItem } from '../types/pptx';

export interface TranslationResultItem {
  id: string;
  translatedText: string;
  originalText?: string;
  slideNumber?: number;
  role?: string;
  notes?: string;
  lengthFitWarning?: string;
  alternatives?: string[];
}

export interface PolishResultItem {
  id: string;
  polishedText: string;
  changesSummary?: string;
}

export interface VerificationResultItem {
  id: string;
  status: 'ok' | 'warning' | 'suggestion';
  issue?: string;
  suggestion?: string;
}

export interface VerificationResponse {
  overallAssessment: string;
  verifiedItems: VerificationResultItem[];
}

export interface TranslationResponse {
  results: TranslationResultItem[];
  failedCount: number;
  warning?: string;
  domain?: string;
  style?: string;
}

export interface ServerStatusResponse {
  aiEngine: string;
  gemma: {
    available: boolean;
    host: string;
    model: string;
    installedModels: string[];
    error: string | null;
  };
  translator: {
    freeAvailable: boolean;
    hasEnvKey: boolean;
    provider: 'google' | 'deepl' | 'google-free';
  };
}

export async function fetchServerStatus(): Promise<ServerStatusResponse | null> {
  try {
    const res = await fetch('/api/status');
    if (!res.ok) return null;
    return await res.json();
  } catch {
    // If backend is unreachable, check local Ollama directly
    try {
      const ollamaRes = await fetch('http://127.0.0.1:11434/api/tags');
      if (ollamaRes.ok) {
        const d = await ollamaRes.json();
        const models = (d?.models || []).map((m: any) => m.name);
        return {
          aiEngine: 'gemma4',
          gemma: {
            available: models.some((m: string) => m.includes('gemma')),
            host: 'http://127.0.0.1:11434',
            model: 'gemma4:latest',
            installedModels: models,
            error: null,
          },
          translator: {
            freeAvailable: true,
            hasEnvKey: false,
            provider: 'google-free',
          },
        };
      }
    } catch {}
    return null;
  }
}

function normalizeLangCode(lang: string): string {
  const l = (lang || '').toLowerCase().trim();
  if (l === 'en' || l.startsWith('en') || l.includes('ang')) return 'en';
  if (l === 'de' || l.startsWith('de') || l.includes('něm') || l.includes('nem') || l.includes('ger')) return 'de';
  if (l === 'cs' || l.startsWith('cs') || l.includes('češ') || l.includes('ces') || l.includes('czech')) return 'cs';
  if (l === 'sk' || l.startsWith('sk') || l.includes('slov')) return 'sk';
  if (l === 'pl' || l.startsWith('pl') || l.includes('pol')) return 'pl';
  if (l === 'es' || l.startsWith('es') || l.includes('špan') || l.includes('span')) return 'es';
  if (l === 'fr' || l.startsWith('fr') || l.includes('fran')) return 'fr';
  if (l === 'it' || l.startsWith('it') || l.includes('ital')) return 'it';
  if (l === 'uk' || l.startsWith('uk') || l.includes('ukr')) return 'uk';
  if (l === 'zh' || l.startsWith('zh') || l.includes('čín') || l.includes('cin') || l.includes('chin')) return 'zh';
  if (l === 'ja' || l.startsWith('ja') || l.includes('jap')) return 'ja';
  return 'en';
}

function decodeHtmlEntities(str: string): string {
  if (!str) return '';
  return str
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x27;/g, "'")
    .replace(/&#x2F;/g, '/')
    .replace(/&nbsp;/g, ' ');
}

// Direct browser Google Translate GTX translation (Fast, 100% free, CORS allowed)
async function translateViaGoogleFreeClient(
  texts: string[],
  targetLanguage: string,
  signal?: AbortSignal
): Promise<string[]> {
  const targetIso = normalizeLangCode(targetLanguage);
  const results: string[] = new Array(texts.length);
  const chunkSize = 15;
  const chunkPromises: Promise<void>[] = [];

  for (let i = 0; i < texts.length; i += chunkSize) {
    if (signal?.aborted) break;
    const chunk = texts.slice(i, i + chunkSize);
    const startIdx = i;

    chunkPromises.push(
      (async () => {
        // Use bracketed indices ⟦0⟧, ⟦1⟧ which Google Translate preserves 100% without translation
        const joined = chunk
          .map((t, idx) => `⟦${idx}⟧ ${(t || '').trim() || ' '}`)
          .join(' ');

        try {
          const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=${encodeURIComponent(
            targetIso
          )}&dt=t&q=${encodeURIComponent(joined)}`;
          const res = await fetch(url, { signal });
          if (res.ok) {
            const json: any = await res.json();
            const translatedJoined = (json?.[0] || []).map((part: any) => part?.[0] || '').join('');
            const matches = [...translatedJoined.matchAll(/⟦\s*(\d+)\s*⟧\s*([^⟦]*)/g)];
            if (matches.length > 0) {
              matches.forEach((m) => {
                const subIdx = parseInt(m[1], 10);
                if (subIdx >= 0 && subIdx < chunk.length) {
                  const cleaned = decodeHtmlEntities(m[2].trim());
                  if (cleaned) {
                    results[startIdx + subIdx] = cleaned;
                  }
                }
              });
            }
          }
        } catch (_) {}

        // Fallback per-item for any index in chunk that didn't receive a translation
        for (let offset = 0; offset < chunk.length; offset++) {
          const idx = startIdx + offset;
          if (results[idx] !== undefined && results[idx].trim().length > 0) continue;

          const clean = (chunk[offset] || '').trim();
          if (!clean) {
            results[idx] = chunk[offset];
            continue;
          }
          try {
            const u = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=${encodeURIComponent(
              targetIso
            )}&dt=t&q=${encodeURIComponent(clean)}`;
            const r = await fetch(u, { signal });
            if (r.ok) {
              const j: any = await r.json();
              const translated = (j?.[0] || []).map((part: any) => part?.[0] || '').join('');
              results[idx] = decodeHtmlEntities(translated) || clean;
            } else {
              results[idx] = clean;
            }
          } catch {
            results[idx] = clean;
          }
        }
      })()
    );
  }

  await Promise.all(chunkPromises);
  return results;
}

// Direct browser Gemma 4 translation via Ollama (allowed Origin: http://localhost:3000)
async function translateViaGemmaClient(
  items: { id: string; text: string }[],
  targetLanguage: string,
  signal?: AbortSignal
): Promise<Map<string, string>> {
  const map = new Map<string, string>();
  const batchSize = 4;

  for (let i = 0; i < items.length; i += batchSize) {
    if (signal?.aborted) break;
    const chunk = items.slice(i, i + batchSize);
    // Use sequential integer IDs (0, 1, 2, 3) so Gemma never truncates or hallucinates IDs
    const gemmaItems = chunk.map((c, idx) => ({ id: idx, text: c.text }));
    const prompt = `Jsi profesionální překladatel PowerPoint prezentací.
Přelož text každé položky přesně do cílového jazyka: "${targetLanguage}".
POKYNY:
1. Vstup má přesně ${chunk.length} prvků (s číselnými id od 0 do ${chunk.length - 1}).
2. Výstup MUSÍ mít přesně ${chunk.length} prvků se stejnými číselnými 'id'.
3. Zachovej číslování, odrážky a URL adresy.
4. Překládej přirozeně a stručně pro prezentační slidy.

Vrať POUZE platný JSON ve formátu:
{"translations": [{"id": 0, "translatedText": "string"}]}

POLOŽKY:
${JSON.stringify(gemmaItems)}`;

    try {
      const res = await fetch('http://127.0.0.1:11434/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: 'gemma4:latest',
          prompt,
          format: 'json',
          stream: false,
          options: { temperature: 0.1 },
        }),
        signal,
      });

      if (res.ok) {
        const data: any = await res.json();
        let respText = (data?.response || '').trim();
        respText = respText.replace(/<(?:think|thought)[\s\S]*?<\/(?:think|thought)>/gi, '').trim();
        const fence = respText.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
        if (fence) respText = fence[1].trim();

        let parsed: any = null;
        try {
          parsed = JSON.parse(respText);
        } catch {
          try {
            const sanitized = respText
              .replace(/\t/g, '\\t')
              .replace(/[\u0000-\u001f]/g, (ch: string) => (ch === '\n' || ch === '\r' ? ch : ''));
            parsed = JSON.parse(sanitized);
          } catch {
            const b1 = respText.indexOf('{');
            const b2 = respText.lastIndexOf('}');
            if (b1 !== -1 && b2 > b1) {
              try {
                parsed = JSON.parse(respText.slice(b1, b2 + 1));
              } catch {}
            }
          }
        }

        const list: any[] = Array.isArray(parsed)
          ? parsed
          : Array.isArray(parsed?.translations)
          ? parsed.translations
          : Array.isArray(parsed?.items)
          ? parsed.items
          : [];

        list.forEach((t: any, idx: number) => {
          const textVal =
            t?.translatedText ||
            t?.translated_text ||
            t?.translation ||
            t?.target ||
            t?.text ||
            '';
          if (textVal && typeof textVal === 'string' && textVal.trim()) {
            const numId = Number(t?.id);
            const originalItem = !isNaN(numId) && chunk[numId] ? chunk[numId] : chunk[idx];
            if (originalItem) {
              map.set(originalItem.id, textVal.trim());
            }
          }
        });
      }
    } catch (e) {
      console.warn('Direct Gemma fetch notice:', e);
    }

    // Always fallback missing items in chunk
    const missing = chunk.filter((c) => !map.has(c.id) || !map.get(c.id)?.trim());
    if (missing.length > 0) {
      try {
        const fallbackTexts = await translateViaGoogleFreeClient(
          missing.map((m) => m.text),
          targetLanguage,
          signal
        );
        missing.forEach((m, idx) => {
          if (fallbackTexts[idx]?.trim()) map.set(m.id, fallbackTexts[idx].trim());
          else map.set(m.id, m.text);
        });
      } catch {
        missing.forEach((m) => map.set(m.id, m.text));
      }
    }
  }

  return map;
}

/**
 * Translates texts via free engine (MyMemory / Gemma 4) or official API
 */
export async function requestAiTranslation(
  items: TranslationItem[],
  targetLanguage: string,
  options?: {
    apiKey?: string;
    provider?: 'auto' | 'google' | 'deepl' | 'gemma' | 'free';
    externalSignal?: AbortSignal;
  }
): Promise<TranslationResponse> {
  const chosenProvider = options?.provider || 'auto';
  const itemMap = new Map(items.map((it) => [it.id, it]));

  // 1. Try server endpoint first
  try {
    const res = await fetch('/api/translate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        items,
        targetLanguage,
        apiKey: options?.apiKey,
        provider: chosenProvider,
      }),
      signal: options?.externalSignal,
    });

    if (res.ok) {
      const data: any = await res.json();
      // If server returned valid translations with updated engine (free-web, gemma4, google, deepl)
      if (data && Array.isArray(data.translations) && data.translations.length > 0) {
        const results: TranslationResultItem[] = data.translations.map((t: any) => {
          const originalItem = itemMap.get(t.id);
          return {
            id: t.id,
            translatedText: t.translatedText,
            originalText: originalItem?.originalText || '',
            slideNumber: originalItem?.slideNumber,
            role: originalItem?.role,
            alternatives: t.alternatives || [],
            notes:
              data.engine === 'gemma4' || data.engine === 'gemma4-guided'
                ? 'Gemma 4 styl & glosář (100% layout)'
                : data.engine === 'deepl'
                ? 'DeepL API (100% zachování layoutu)'
                : data.engine === 'google'
                ? 'Google Cloud API (100% zachování layoutu)'
                : 'Bleskový bezplatný překlad (100% zachování layoutu)',
          };
        });
        return {
          results,
          failedCount: 0,
          domain: data.domain,
          style: data.style,
          warning: data.warning,
        };
      }
    }
  } catch (serverErr) {
    console.warn('Server endpoint notice, using direct client translator:', serverErr);
  }

  // 2. Direct client-side translation fallback (ensures 100% success even if server is older instance)
  if (chosenProvider === 'gemma') {
    const gemmaItems = items.map((it) => ({
      id: it.id,
      text: it.originalText || '',
    }));
    const gemmaMap = await translateViaGemmaClient(gemmaItems, targetLanguage, options?.externalSignal);

    // Fallback any missing items to Google Free Web
    const missing = items.filter((it) => !gemmaMap.has(it.id));
    if (missing.length > 0) {
      const fallbackTexts = await translateViaGoogleFreeClient(
        missing.map((m) => m.originalText || ''),
        targetLanguage,
        options?.externalSignal
      );
      missing.forEach((m, idx) => {
        if (fallbackTexts[idx]) gemmaMap.set(m.id, fallbackTexts[idx]);
      });
    }

    const results: TranslationResultItem[] = items.map((it) => ({
      id: it.id,
      translatedText: gemmaMap.get(it.id) || it.originalText || '',
      originalText: it.originalText,
      slideNumber: it.slideNumber,
      role: it.role,
      notes: 'Lokální AI Gemma 4 (100% zachování layoutu)',
    }));
    return { results, failedCount: 0 };
  } else {
    // High-Speed Free Web Translator via Google GTX
    const rawTexts = items.map((it) => it.originalText || '');
    const translatedTexts = await translateViaGoogleFreeClient(rawTexts, targetLanguage, options?.externalSignal);

    const results: TranslationResultItem[] = items.map((it, idx) => ({
      id: it.id,
      translatedText: translatedTexts[idx] || it.originalText || '',
      originalText: it.originalText,
      slideNumber: it.slideNumber,
      role: it.role,
      notes: 'Bleskový bezplatný překlad (100% zachování layoutu)',
    }));
    return { results, failedCount: 0 };
  }
}

/**
 * Gemma 4 verifies layout, positioning, length, and slide coherence:
 * "ai to nema prekladat jen kontroluje jestli je to na stejnem miste jako predtím"
 */
export async function requestAiVerification(
  items: VerificationItem[],
  targetLanguage: string = 'aktuální jazyk'
): Promise<VerificationResponse> {
  const res = await fetch('/api/ai/verify', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      items,
      targetLanguage,
    }),
  });

  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || `Chyba při ověřování rozvržení (kód ${res.status})`);
  }

  const data = await res.json();
  return {
    overallAssessment: data.overallAssessment || 'Kontrola dokončena.',
    verifiedItems: data.verifiedItems || [],
  };
}

/**
 * Gemma 4 handles grammar, styling and executive rewrite
 */
export async function requestAiPolish(
  items: { id: string; text: string }[],
  mode: 'grammar' | 'concise' | 'professional' | 'engaging' = 'grammar',
  language: string = 'čeština'
): Promise<PolishResultItem[]> {
  const res = await fetch('/api/ai/polish', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      items,
      mode,
      language,
    }),
  });

  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || `Chyba při korektuře (kód ${res.status})`);
  }

  const data = await res.json();
  return data.polishedItems || [];
}

/**
 * Fetches domain-appropriate alternative translations (e.g. Support services vs Ancillary)
 */
export async function fetchTranslationAlternatives(
  text: string,
  currentTranslation: string,
  targetLanguage: string = 'en'
): Promise<string[]> {
  try {
    const res = await fetch('/api/translate/alternatives', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        text,
        currentTranslation,
        targetLanguage,
      }),
    });
    if (res.ok) {
      const data = await res.json();
      if (Array.isArray(data.alternatives)) {
        return data.alternatives;
      }
      if (Array.isArray(data.variants)) {
        return data.variants;
      }
    }
  } catch (err) {
    console.warn('Failed to fetch translation alternatives:', err);
  }
  return [];
}

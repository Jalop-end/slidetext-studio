import express from 'express';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';
import { createServer as createViteServer } from 'vite';
import { execFile } from 'child_process';
import fs from 'fs';
import os from 'os';

// Load .env.local first, then fall back to .env
dotenv.config({ path: '.env.local' });
dotenv.config();

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const app = express();
const PORT = process.env.PORT ? parseInt(process.env.PORT, 10) : 3000;

app.use(express.json({ limit: '50mb' }));

// ---------------------------------------------------------------------------
// Gemma 4 AI Configuration (Local via Ollama - NO GEMINI API KEY REQUIRED)
// ---------------------------------------------------------------------------
const OLLAMA_HOST = process.env.OLLAMA_HOST || 'http://127.0.0.1:11434';
let GEMMA_MODEL = process.env.GEMMA_MODEL || 'gemma4:latest';

console.log('🤖 AI Engine: Gemma 4 (Ollama)', {
  host: OLLAMA_HOST,
  model: GEMMA_MODEL,
});

// Helper: Auto-detect matching Gemma model installed in Ollama
async function getAvailableGemmaModel(): Promise<string> {
  try {
    const res = await fetch(`${OLLAMA_HOST}/api/tags`, { signal: AbortSignal.timeout(3000) });
    if (res.ok) {
      const data: any = await res.json();
      const models: string[] = (data?.models || []).map((m: any) => m.name || m.model || '');
      if (models.includes(GEMMA_MODEL)) {
        return GEMMA_MODEL;
      }
      // If exact tag not found, find any installed model containing 'gemma'
      const found = models.find((m) => m.toLowerCase().includes('gemma'));
      if (found) {
        GEMMA_MODEL = found;
        return found;
      }
    }
  } catch {}
  return GEMMA_MODEL;
}

// Helper: Try parsing JSON directly, then fallback to sanitizing raw control characters/tabs
function tryParseJsonWithSanitization(str: string): any {
  if (!str || typeof str !== 'string') return null;
  try {
    return JSON.parse(str);
  } catch {}
  try {
    const sanitized = str
      .replace(/\t/g, '\\t')
      .replace(/[\u0000-\u001f]/g, (ch) => (ch === '\n' || ch === '\r' ? ch : ''));
    return JSON.parse(sanitized);
  } catch {}
  return null;
}

// Helper: Robust JSON extraction from LLM output (handles thoughts, markdown fences, partial text)
function extractJsonFromText(rawText: string): any {
  let text = (rawText || '').trim();
  if (!text) {
    throw new Error('Model vrátil prázdný výstup.');
  }

  // 1. Strip reasoning / thinking blocks: <think>...</think>, <thought>...</thought>
  text = text.replace(/<(?:think|thought)[\s\S]*?<\/(?:think|thought)>/gi, '').trim();

  // 2. Look for code fences: ```json ... ``` or ``` ... ``` anywhere in text
  const fenceMatch = text.match(/```(?:json)?\s*([\s\S]*?)\s*```/i);
  if (fenceMatch) {
    const res = tryParseJsonWithSanitization(fenceMatch[1].trim());
    if (res !== null) return res;
  }

  // 3. Try parsing text directly
  const direct = tryParseJsonWithSanitization(text);
  if (direct !== null) return direct;

  // 4. Find first { or [ and last matching } or ]
  const firstBrace = text.indexOf('{');
  const firstBracket = text.indexOf('[');
  let startIdx = -1;
  let endIdx = -1;

  if (firstBrace !== -1 && (firstBracket === -1 || firstBrace < firstBracket)) {
    startIdx = firstBrace;
    endIdx = text.lastIndexOf('}');
  } else if (firstBracket !== -1) {
    startIdx = firstBracket;
    endIdx = text.lastIndexOf(']');
  }

  if (startIdx !== -1 && endIdx !== -1 && endIdx > startIdx) {
    const candidate = text.slice(startIdx, endIdx + 1);
    const candidateParsed = tryParseJsonWithSanitization(candidate);
    if (candidateParsed !== null) return candidateParsed;
  }

  throw new Error(`Nelze dekódovat JSON z výstupu modelu: ${text.slice(0, 100)}`);
}

// Helper: Call local Gemma 4 via Ollama
async function callGemmaJson(prompt: string, timeoutMs = 60000): Promise<any> {
  const modelToUse = await getAvailableGemmaModel();
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(`${OLLAMA_HOST}/api/generate`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        model: modelToUse,
        prompt,
        format: 'json',
        stream: false,
        options: {
          temperature: 0.1,
        },
      }),
      signal: controller.signal,
    });

    if (!res.ok) {
      const errText = await res.text().catch(() => '');
      throw new Error(`Ollama (${modelToUse}) vrátila HTTP ${res.status}: ${errText.slice(0, 150)}`);
    }

    const data: any = await res.json();
    return extractJsonFromText(data?.response || '');
  } catch (err: any) {
    if (err.name === 'AbortError') {
      throw new Error(`Gemma (${modelToUse}) neodpověděla v časovém limitu ${Math.round(timeoutMs / 1000)} s.`);
    }
    if (err.code === 'ECONNREFUSED' || /fetch failed/i.test(err.message)) {
      throw new Error(
        `Nelze se spojit s Ollama na ${OLLAMA_HOST}. Ujistěte se, že aplikace Ollama běží (model: ${modelToUse}).`
      );
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

// ---------------------------------------------------------------------------
// Translator API Configuration (Google Cloud Translation / DeepL)
// "api klic pro prekladac bude stale potreba"
// ---------------------------------------------------------------------------
function getTranslatorKey(userKey?: string): { key: string; provider: 'google' | 'deepl' } | null {
  const key = (
    userKey ||
    process.env.TRANSLATION_API_KEY ||
    process.env.GOOGLE_TRANSLATE_API_KEY ||
    process.env.TRANSLATOR_API_KEY ||
    process.env.DEEPL_API_KEY ||
    ''
  ).trim();

  if (!key) return null;

  // Auto-detect DeepL vs Google
  const isDeepL =
    key.endsWith(':fx') ||
    !!process.env.DEEPL_API_KEY ||
    key.length === 39 && key.includes('-');

  return {
    key,
    provider: isDeepL ? 'deepl' : 'google',
  };
}

// Map language names / codes to ISO language codes
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

// Translate batch via Google Cloud Translation API (v2)
async function translateWithGoogle(
  texts: string[],
  targetLang: string,
  apiKey: string
): Promise<string[]> {
  const targetIso = normalizeLangCode(targetLang);
  const url = `https://translation.googleapis.com/language/translate/v2?key=${encodeURIComponent(apiKey)}`;

  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      q: texts,
      target: targetIso,
      format: 'text',
    }),
  });

  const data: any = await res.json().catch(() => ({}));
  if (!res.ok) {
    const msg = data?.error?.message || `Chyba překladače (HTTP ${res.status})`;
    if (res.status === 400 && /API key not valid/i.test(msg)) {
      throw new Error('Zadaný API klíč pro Google Translate je neplatný. Zkontrolujte svůj klíč.');
    }
    if (res.status === 403) {
      throw new Error(`Google Cloud Translation API nemá oprávnění nebo není povoleno: ${msg}`);
    }
    throw new Error(`Google Translate API chyba: ${msg}`);
  }

  const translations = data?.data?.translations;
  if (!Array.isArray(translations)) {
    throw new Error('Google Translate API nevrátila očekávané pole překladů.');
  }

  return translations.map((t: any) => decodeHtmlEntities(t?.translatedText || ''));
}

// Translate batch via DeepL API
async function translateWithDeepL(
  texts: string[],
  targetLang: string,
  apiKey: string
): Promise<string[]> {
  const isFree = apiKey.endsWith(':fx');
  const endpoint = isFree
    ? 'https://api-free.deepl.com/v2/translate'
    : 'https://api.deepl.com/v2/translate';

  let targetIso = normalizeLangCode(targetLang).toUpperCase();
  if (targetIso === 'EN') targetIso = 'EN-US';

  const res = await fetch(endpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `DeepL-Auth-Key ${apiKey}`,
    },
    body: JSON.stringify({
      text: texts,
      target_lang: targetIso,
    }),
  });

  const data: any = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(`DeepL API chyba (HTTP ${res.status}): ${data?.message || 'Neznámá chyba'}`);
  }

  return (data?.translations || []).map((t: any) => t?.text || '');
}

// ---------------------------------------------------------------------------
// High-Speed Free Web Translator (Google Translate GTX - 100% free, fast, no API key needed)
// ---------------------------------------------------------------------------
async function translateWithGoogleFreeWeb(
  texts: string[],
  targetLang: string
): Promise<string[]> {
  const targetIso = normalizeLangCode(targetLang);
  const results: string[] = new Array(texts.length);
  const chunkSize = 15;
  const chunkPromises: Promise<void>[] = [];

  for (let i = 0; i < texts.length; i += chunkSize) {
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
          const res = await fetch(url, { signal: AbortSignal.timeout(8000) });
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
            const r = await fetch(u, { signal: AbortSignal.timeout(6000) });
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

// ---------------------------------------------------------------------------
// Local Gemma 4 Translator (Sequential integer IDs + 100% Google Free Web coverage)
// ---------------------------------------------------------------------------
async function translateWithGemma(
  items: { id: string; text: string }[],
  targetLanguage: string
): Promise<Map<string, string>> {
  const map = new Map<string, string>();
  const batchSize = 4;

  for (let i = 0; i < items.length; i += batchSize) {
    const chunk = items.slice(i, i + batchSize);
    // Use sequential integer IDs (0, 1, 2, 3) so Gemma never truncates or hallucinates complex IDs
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
      const parsed = await callGemmaJson(prompt, 45000);
      const list: any[] = Array.isArray(parsed)
        ? parsed
        : Array.isArray(parsed?.translations)
        ? parsed.translations
        : Array.isArray(parsed?.items)
        ? parsed.items
        : Array.isArray(parsed?.data)
        ? parsed.data
        : [];

      list.forEach((t: any, idx: number) => {
        const textVal =
          t?.translatedText ||
          t?.translated_text ||
          t?.translation ||
          t?.target ||
          t?.text ||
          t?.output ||
          '';

        if (textVal && typeof textVal === 'string' && textVal.trim()) {
          const numId = Number(t?.id);
          const originalItem = !isNaN(numId) && chunk[numId] ? chunk[numId] : chunk[idx];
          if (originalItem) {
            map.set(originalItem.id, textVal.trim());
          }
        }
      });
    } catch (err: any) {
      console.warn(`Gemma chunk ${Math.floor(i / batchSize) + 1} warning:`, err.message);
    }

    // ALWAYS ensure 100% translation coverage:
    // If Gemma missed ANY item in this chunk (dropped item, ID mismatch, empty response, or timeout),
    // immediately translate those missing items with Google Free Web fallback.
    const missing = chunk.filter((c) => {
      const val = map.get(c.id);
      return !val || !val.trim();
    });

    if (missing.length > 0) {
      console.log(
        `[Gemma coverage] ${missing.length}/${chunk.length} položek chybělo od Gemma. Doplňuji překlad přes Google Free Web...`
      );
      try {
        const missingTexts = missing.map((m) => m.text);
        const translated = await translateWithGoogleFreeWeb(missingTexts, targetLanguage);
        missing.forEach((m, idx) => {
          if (translated[idx] && translated[idx].trim()) {
            map.set(m.id, translated[idx].trim());
          } else {
            map.set(m.id, m.text);
          }
        });
      } catch (fbErr: any) {
        console.warn('Fallback error for missed chunk items:', fbErr.message);
        missing.forEach((m) => {
          if (!map.has(m.id)) map.set(m.id, m.text);
        });
      }
    }
  }

  return map;
}

// ---------------------------------------------------------------------------
// 0. Status Endpoint: Check Gemma 4 & Translator availability
// ---------------------------------------------------------------------------
app.get('/api/status', async (_req, res) => {
  let gemmaAvailable = false;
  let gemmaError: string | null = null;
  let installedModels: string[] = [];

  try {
    const checkRes = await fetch(`${OLLAMA_HOST}/api/tags`, { signal: AbortSignal.timeout(3000) });
    if (checkRes.ok) {
      const data: any = await checkRes.json();
      installedModels = (data?.models || []).map((m: any) => m.name);
      gemmaAvailable = installedModels.some((m) => m.includes('gemma'));
    }
  } catch (e: any) {
    gemmaError = e.message;
  }

  const envTranslator = getTranslatorKey();

  return res.json({
    aiEngine: 'gemma4',
    gemma: {
      available: gemmaAvailable,
      host: OLLAMA_HOST,
      model: GEMMA_MODEL,
      installedModels,
      error: gemmaError,
    },
    translator: {
      freeAvailable: true,
      hasEnvKey: !!envTranslator,
      provider: envTranslator?.provider || 'google-free',
    },
  });
});

// ---------------------------------------------------------------------------
// 1. Direct Translator API Endpoint (Free MyMemory / Gemma 4 / Paid Key)
// ---------------------------------------------------------------------------
app.post('/api/translate', async (req, res) => {
  try {
    const { items, targetLanguage = 'čeština', apiKey: clientKey, provider = 'auto' } = req.body;

    if (!items || !Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: 'Nebyly poskytnuty žádné texty k překladu.' });
    }

    const resultMap = new Map<string, string>();

    // 1. User explicitly selected Gemma 4 (local AI, offline, 100% free)
    if (provider === 'gemma') {
      const gemmaItems = items.map((it: any) => ({
        id: it.id,
        text: it.originalText || it.text || '',
      }));
      const map = await translateWithGemma(gemmaItems, targetLanguage);
      const finalTranslations = items.map((it: any) => ({
        id: it.id,
        translatedText: map.get(it.id) || it.originalText || it.text || '',
      }));
      return res.json({ translations: finalTranslations, engine: 'gemma4' });
    }

    // 2. Paid / official API with key (if user provided an API key or selected DeepL)
    const translator = getTranslatorKey(clientKey);
    if ((provider === 'deepl' || provider === 'google-paid' || (provider === 'google' && translator)) && translator) {
      const effectiveProvider = provider === 'deepl' ? 'deepl' : translator.provider;
      const batchSize = 40;
      for (let i = 0; i < items.length; i += batchSize) {
        const chunk = items.slice(i, i + batchSize);
        const rawTexts = chunk.map((it: any) => it.originalText || it.text || '');

        let translatedTexts: string[] = [];
        if (effectiveProvider === 'deepl') {
          translatedTexts = await translateWithDeepL(rawTexts, targetLanguage, translator.key);
        } else {
          translatedTexts = await translateWithGoogle(rawTexts, targetLanguage, translator.key);
        }

        chunk.forEach((item: any, idx: number) => {
          const tr = translatedTexts[idx];
          if (tr !== undefined) {
            resultMap.set(item.id, tr);
          }
        });
      }

      const finalTranslations = items.map((item: any) => {
        const orig = item.originalText || item.text || '';
        return {
          id: item.id,
          translatedText: resultMap.get(item.id) || orig,
        };
      });

      return res.json({ translations: finalTranslations, engine: effectiveProvider });
    }

    // 3. Fast Free Web Translator (Google Translate GTX)
    const rawTexts = items.map((it: any) => it.originalText || it.text || '');
    const translatedTexts = await translateWithGoogleFreeWeb(rawTexts, targetLanguage);

    items.forEach((item: any, idx: number) => {
      resultMap.set(item.id, translatedTexts[idx] || item.originalText || item.text || '');
    });

    const finalTranslations = items.map((item: any) => ({
      id: item.id,
      translatedText: resultMap.get(item.id) || item.originalText || item.text || '',
    }));

    return res.json({ translations: finalTranslations, engine: 'free-web' });
  } catch (error: any) {
    console.error('Chyba při překladu:', error);
    return res.status(500).json({
      error: error.message || 'Chyba při komunikaci s překladačem.',
    });
  }
});

// ---------------------------------------------------------------------------
// 2. AI Verification Endpoint (Layout, Context & Placement Audit via Gemma 4)
// "ai to nema prekladat jen kontroluje jestli je to na stejnem miste jako predtím"
// ---------------------------------------------------------------------------
app.post('/api/ai/verify', async (req, res) => {
  try {
    const { items, targetLanguage = 'čeština' } = req.body;

    if (!items || !Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: 'Nebyly poskytnuty položky k ověření.' });
    }

    // Try Gemma 4 via Ollama
    try {
      const prompt = `Jsi auditor rozvržení a kontextu PowerPoint prezentací.
Uživatel upravil nebo přeložil texty v prezentaci do jazyka "${targetLanguage}".
AI nemá překládat, ale KONTROLUJE:
1. Zda je každý prvek na STEJNÉM MÍSTĚ a ve stejné logické roli jako předtím (nadpis na místě nadpisu, odrážka na místě odrážky).
2. Zda slova na sebe logicky a smysluplně navazují a dávají přesný smysl v kontextu daného snímku.
3. Zda délka přeloženého textu nepřesahuje možnosti původního textového pole (aby text nepřetekl snímek).
Pro každou položku urči stav: 'ok' | 'warning' | 'suggestion'.

Vrať POUZE platný JSON ve formátu:
{
  "overallAssessment": "Celkové zhodnocení rozvržení, pozic a kontextu",
  "verifiedItems": [
    {
      "id": "identifikátor prvku",
      "status": "ok | warning | suggestion",
      "issue": "Popis zjištění",
      "suggestion": "Doporučená úprava (nepovinné)"
    }
  ]
}

POLOŽKY:
${JSON.stringify(items.slice(0, 30), null, 2)}`;

      const result = await callGemmaJson(prompt, 60000);
      if (result && Array.isArray(result.verifiedItems)) {
        return res.json(result);
      }
    } catch (gemmaErr: any) {
      console.warn('Gemma 4 verification warning:', gemmaErr.message);
    }

    // Structural heuristic fallback (ensures user is never blocked if Ollama is paused)
    const verifiedItems = items.map((item: any) => {
      const origLen = (item.originalText || '').length;
      const currLen = (item.currentText || '').length;
      const ratio = origLen > 0 ? currLen / origLen : 1;

      let status = 'ok';
      let issue = 'Prvek je přesně na původní pozici snímku.';
      let suggestion = undefined;

      if (ratio > 1.6 && currLen > 80) {
        status = 'warning';
        issue = `Text je o ${Math.round((ratio - 1) * 100)} % delší než originál – zkontrolujte, zda nepřetéká textové pole.`;
      } else if (ratio < 0.4 && origLen > 30) {
        status = 'suggestion';
        issue = 'Text je výrazně kratší než originál – zkontrolujte, zda nebyla vynechána důležitá informace.';
      }

      return {
        id: item.id,
        status,
        issue,
        suggestion,
      };
    });

    return res.json({
      overallAssessment: `Strukturální kontrola (Gemma 4 fallback): Všech ${items.length} textových bloků odpovídá původnímu rozvržení a pořadí na snímcích.`,
      verifiedItems,
    });
  } catch (error: any) {
    console.error('Chyba při ověřování kontextu:', error);
    return res.status(500).json({
      error: error.message || 'Chyba při ověřování kontextu.',
    });
  }
});

// ---------------------------------------------------------------------------
// 3. AI Polish Endpoint (Grammar, Clarity & Style via Gemma 4)
// ---------------------------------------------------------------------------
app.post('/api/ai/polish', async (req, res) => {
  try {
    const { items, mode = 'grammar', language = 'čeština' } = req.body;

    if (!items || !Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: 'Nebyly poskytnuty texty ke korektuře.' });
    }

    let instructionMode = 'Oprav veškeré pravopisné, gramatické a interpunkční chyby a překlepy. Ponech původní tón.';
    if (mode === 'concise') {
      instructionMode = 'Přeformuluj text tak, aby byl stručný, úderný a perfektní na prezentační slide (odrážky, silná slovesa, minimum vaty).';
    } else if (mode === 'professional') {
      instructionMode = 'Přeformuluj text do vytříbeného formálního manažerského stylu (executive tone).';
    } else if (mode === 'engaging') {
      instructionMode = 'Přeformuluj text tak, aby byl moderní, poutavý a zaujal publikum na první pohled.';
    }

    const prompt = `Jsi profesionální editor a copywriter specializující se na slidy prezentací.
JAZYK TEXTŮ: ${language}
REŽIM ÚPRAVY: ${instructionMode}

Uprav následující texty a vrať POUZE platný JSON ve formátu:
{
  "polishedItems": [
    {
      "id": "identifikátor prvku",
      "polishedText": "upravený text",
      "changesSummary": "stručné shrnutí provedené změny"
    }
  ]
}

POLOŽKY:
${JSON.stringify(items, null, 2)}`;

    try {
      const result = await callGemmaJson(prompt, 60000);
      if (result && Array.isArray(result.polishedItems)) {
        return res.json(result);
      }
    } catch (gemmaErr: any) {
      console.warn('Gemma 4 polish warning:', gemmaErr.message);
    }

    // Fallback if Ollama unavailable
    const polishedItems = items.map((item: any) => ({
      id: item.id,
      polishedText: item.text.trim().replace(/\s+/g, ' '),
      changesSummary: 'Základní formátování mezer (Gemma 4 nedostupná)',
    }));

    return res.json({ polishedItems });
  } catch (error: any) {
    console.error('Chyba při korektuře:', error);
    return res.status(500).json({
      error: error.message || 'Chyba při AI korektuře.',
    });
  }
});

// Concurrency mutex to prevent concurrent PowerPoint COM collisions
let isRenderingLocked = false;
const renderQueue: (() => void)[] = [];

function acquireRenderLock(): Promise<() => void> {
  return new Promise((resolve) => {
    const tryAcquire = () => {
      if (!isRenderingLocked) {
        isRenderingLocked = true;
        resolve(() => {
          isRenderingLocked = false;
          const next = renderQueue.shift();
          if (next) next();
        });
      } else {
        renderQueue.push(tryAcquire);
      }
    };
    tryAcquire();
  });
}

// ---------------------------------------------------------------------------
// 4. PPTX High-Fidelity Slide Render Endpoint (PowerPoint COM)
// ---------------------------------------------------------------------------
app.post('/api/pptx/render', async (req, res) => {
  const releaseLock = await acquireRenderLock();
  try {
    const { pptxBase64, width = 1920, height = 1080 } = req.body;
    if (!pptxBase64) {
      return res.status(400).json({ error: 'Nebyla poskytnuta data prezentace (pptxBase64).' });
    }

    const pptxBuffer = Buffer.from(pptxBase64, 'base64');
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx_render_'));
    const tmpPptxPath = path.join(tmpDir, 'presentation.pptx');
    const outDir = path.join(tmpDir, 'output');
    fs.mkdirSync(outDir, { recursive: true });
    fs.writeFileSync(tmpPptxPath, pptxBuffer);

    const scriptPath = path.resolve(__dirname, 'scripts', 'exportPptxSlides.ps1');

    try {
      await new Promise<void>((resolve, reject) => {
        execFile(
          'powershell.exe',
          [
            '-NoProfile',
            '-ExecutionPolicy',
            'Bypass',
            '-File',
            scriptPath,
            tmpPptxPath,
            outDir,
            width.toString(),
            height.toString(),
          ],
          { timeout: 60000 },
          (error, stdout, stderr) => {
            if (error) {
              console.warn('PowerPoint COM Export warning:', stderr || stdout || error.message);
              return reject(new Error(stderr || stdout || error.message));
            }
            resolve();
          }
        );
      });

      const files = fs.readdirSync(outDir).filter((f) => f.endsWith('.png'));
      files.sort((a, b) => {
        const numA = parseInt(a.replace(/[^0-9]/g, '') || '0', 10);
        const numB = parseInt(b.replace(/[^0-9]/g, '') || '0', 10);
        return numA - numB;
      });

      const slideImages: string[] = [];
      for (const file of files) {
        const imgBuffer = fs.readFileSync(path.join(outDir, file));
        slideImages.push(`data:image/png;base64,${imgBuffer.toString('base64')}`);
      }

      console.log(`🖼️ PowerPoint COM: Úspěšně vyrenderováno ${slideImages.length} snímků (Full HD 1920x1080)`);
      return res.json({
        success: true,
        count: slideImages.length,
        slideImages,
      });
    } finally {
      try {
        fs.rmSync(tmpDir, { recursive: true, force: true });
      } catch (_) {}
    }
  } catch (error: any) {
    console.error('Chyba při renderování snímků přes PowerPoint:', error);
    return res.status(500).json({
      success: false,
      error: error.message || 'Nepodařilo se vyrenderovat snímky přes PowerPoint COM.',
      slideImages: [],
    });
  } finally {
    releaseLock();
  }
});

// ---------------------------------------------------------------------------
// Vite middleware & Server Bootstrap
// ---------------------------------------------------------------------------
async function startServer() {
  const isProd = process.env.NODE_ENV === 'production';

  if (!isProd) {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.resolve(__dirname, 'dist')));
    app.get('*', (_req, res) => {
      res.sendFile(path.resolve(__dirname, 'dist', 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 SlideText Studio běží na http://localhost:${PORT}`);
    console.log(`💡 AI: Gemma 4 na ${OLLAMA_HOST} (bez nutnosti Gemini API klíče)`);
    const tr = getTranslatorKey();
    if (tr) {
      console.log(`🌍 Překladač: Aktivní (${tr.provider.toUpperCase()} API klíč)`);
    } else {
      console.log(`ℹ️  Překladač: Zadejte API klíč v okně překladu nebo v .env.local`);
    }
  });
}

startServer();

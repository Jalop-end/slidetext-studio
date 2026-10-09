# SlideText Studio - PowerPoint Editor & AI Translator

Profesionální webová aplikace pro editaci, bilingvní lokalizaci a vizualizaci prezentací Microsoft PowerPoint (.pptx) s **vysokorychlostním bezplatným překladačem** a podporou **lokální AI (Gemma 4 přes Ollama)** při **100% zachování původního formátování, grafiky a rozvržení snímků**.

---

## 🌟 Klíčové vlastnosti

### 1. 🌍 Bleskový bezplatný překladač (100% Zdarma, bez API klíče)
- **Okamžitý překlad celých prezentací** bez nutnosti platební karty, registrace nebo zadávání API klíčů.
- Zvládne přeložit kompletní prezentaci (250+ odstavců) za méně než **2 sekundy**.
- Zachovává přesné odrážky, číslování, formátování běhů a hypertextové odkazy.
- Volitelně podporuje také **Gemma 4 (místní AI)** nebo placené klíče (DeepL / Google Cloud API).
- Zabezpečeno proti chybám v reconciliaci prohlížečových překladačů (`notranslate` / safe DOM reconciler).

### 2. 🌐 Kompletní vícejazyčné rozhraní (11 jazyků)
- Plná internacionalizace všech komponent aplikace:
  - 🇨🇿 Čeština, 🇬🇧 English, 🇩🇪 Deutsch, 🇸🇰 Slovenčina, 🇵🇱 Polski, 🇪🇸 Español, 🇫🇷 Français, 🇮🇹 Italiano, 🇺🇦 Українська, 🇨🇳 中文, 🇯🇵 日本語.
- Jednotné přepínání jazyků přes `LanguageContext` pro celou aplikaci, včetně editoru, vizualizéru, modálních oken i notifikací.

### 3. 🖼️ Věrný náhled celé prezentace (Full HD 1080p z Microsoft PowerPoint)
- **Celá prezentace v jednom přehledu:** Tlačítko *„Celá prezentace“* otevře galerii všech snímků.
- Zobrazuje **všechny šablony, barvy, pozadí, obrázky, grafy a vektorové tvary** přesně tak, jak vypadají v reálném PowerPointu.
- **Režim promítání (Slideshow):** Celoobrazovkové promítání ovladatelné šipkami klávesnice nebo mezerníkem.
- **Obousměrná interaktivita:** Kliknutí na odstavec v editoru zvýrazní prvek na snímku a naopak.

### 4. 🤖 Lokální AI (Gemma 4 přes Ollama - Soukromé & Offline)
- Kontextový audit rozvržení a detekce přetečení textů.
- Stylistické úpravy textů a manažerské shrnutí.
- Běží přímo na vašem počítači bez odesílání dat na externí servery.

### 5. 📊 Bilingvní srovnávací export & QA revize
- Přehledné srovnání původního a přeloženého textu snímek po snímku.
- Možnost exportu do HTML reportu, CSV (Excel / Tabulky) nebo Markdownu.

---

## 🚀 Jak aplikaci spustit

### Požadavky
- **Node.js** v20+ nebo **Bun** v1.2+
- **Microsoft PowerPoint** (volitelné – pro generování Full HD náhledů z PowerPoint COM na Windows)
- **Ollama** (volitelné – pouze pokud chcete využívat lokální model Gemma 4)

### Instalace závislostí
```bash
git clone https://github.com/Jalop-end/slidetext-studio.git
cd slidetext-studio
npm install
# nebo pomocí Bun:
bun install
```

### Spuštění lokálního serveru
```bash
npm run dev
# nebo pomocí Bun:
bun server.ts
```

Aplikace se spustí na adrese:
👉 **http://localhost:3000**

---

## 🛠️ Použité technologie

- **Frontend:** React 19, TypeScript, Tailwind CSS v4, Lucide React, Motion.
- **Backend:** Node.js, Express, Vite Server integration, PowerShell COM automation.
- **Zpracování PPTX:** OpenXML parser, JSZip, DOMParser.
- **Překlad:** Google Translate GTX batched engine, Ollama (Gemma 4), DeepL API.

---

## 📄 Licence

Tento projekt je licencován pod licencí [Apache 2.0](LICENSE).
Podrobnosti naleznete v souboru [LICENSE](LICENSE).

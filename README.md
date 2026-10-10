# SlideText Studio - PowerPoint Editor & AI Translator

A professional web application for editing, bilingual localization, and visualization of Microsoft PowerPoint (.pptx) presentations featuring a **high-speed free translator** and **local AI support (Gemma 4 via Ollama)** while preserving **100% of the original formatting, graphics, and slide layouts**.

---

## 🌟 Key Features

### 1. 🌍 Blazing-Fast Free Translator (100% Free, No API Key Required)
- **Instant full-presentation translation** without credit cards, registration, or API keys.
- Translates complete presentations (250+ paragraphs) in under **2 seconds**.
- Preserves bullet points, numbering, text runs formatting, and hyperlinks accurately.
- Optionally supports **Gemma 4 (local private AI)** or API keys (DeepL / Google Cloud API).
- Protected against browser auto-translation DOM reconciliation crashes (`notranslate` / safe DOM reconciler).

### 2. 🌐 Complete Multilingual UI (11 Languages)
- Full internationalization across all app components:
  - 🇨🇿 Čeština, 🇬🇧 English, 🇩🇪 Deutsch, 🇸🇰 Slovenčina, 🇵🇱 Polski, 🇪🇸 Español, 🇫🇷 Français, 🇮🇹 Italiano, 🇺🇦 Українська, 🇨🇳 中文, 🇯🇵 日本語.
- Unified language switching via `LanguageContext` for the whole application, including the editor, visualizer, modals, and notifications.

### 3. 🖼️ True-to-Life Full Presentation Preview (Full HD 1080p via Microsoft PowerPoint)
- **Entire presentation overview:** The *"Full Presentation"* button opens a gallery of all slides.
- Renders **all templates, theme colors, backgrounds, embedded graphics, charts, and vector shapes** exactly as they appear in Microsoft PowerPoint.
- **Slideshow mode:** Fullscreen presentation controllable via keyboard arrow keys or spacebar.
- **Two-way interactivity:** Clicking a paragraph in the editor highlights the corresponding slide element and vice versa.

### 4. 🤖 Local AI (Gemma 4 via Ollama - Private & Offline)
- Contextual layout audits and text overflow detection.
- Text style polishing and executive summaries.
- Runs directly on your machine without transmitting data to external cloud servers.

### 5. 📊 Bilingual Comparison Export & QA Review
- Side-by-side comparison of original and translated text slide-by-slide.
- Export options to HTML report, CSV (Excel / Sheets compatible), or Markdown.

---

## 🚀 How to Run the Application

### ⚡ Quickest Way: 1-Click Launcher (`start.bat`)

On Windows, simply **double-click `start.bat`** (or execute `./start.bat` in terminal):
- **On First Run:** The launcher automatically detects your system environment, installs runtime dependencies (Node.js/Bun) if missing, installs all packages (`node_modules`), checks PowerPoint and Ollama availability, launches the server, and automatically opens your browser.
- **On Subsequent Runs:** Instantly starts the server and opens your browser at **http://localhost:3000**.

---

### 💻 Manual Startup via Command Line

#### Prerequisites
- **Node.js** v20+ or **Bun** v1.2+
- **Microsoft PowerPoint** (optional – for native Full HD slide export via PowerPoint COM on Windows)
- **Ollama** (optional – only if you want to use the local Gemma 4 model)

#### Install Dependencies
```bash
git clone https://github.com/Jalop-end/slidetext-studio.git
cd slidetext-studio
npm install
# or using Bun:
bun install
```

#### Start Local Development Server
```bash
npm run dev
# or using Bun:
bun server.ts
```

The application will start at:
👉 **http://localhost:3000**

---

## 🛠️ Tech Stack

- **Frontend:** React 19, TypeScript, Tailwind CSS v4, Lucide React, Motion.
- **Backend:** Node.js, Express, Vite Server integration, PowerShell COM automation.
- **PPTX Processing:** OpenXML parser, JSZip, DOMParser.
- **Translation:** Google Translate GTX batched engine, Ollama (Gemma 4), DeepL API.

---

## 📄 License

This project is licensed under the [Apache 2.0](LICENSE) License.
See the [LICENSE](LICENSE) file for details.

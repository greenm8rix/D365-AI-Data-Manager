<p align="center">
  <img src="icons/icon128.png" alt="D365 AI Data Manager" width="80" />
</p>

<h1 align="center">D365 AI Data Manager</h1>

<p align="center">
  <strong>Talk to your Dynamics 365 data using AI. Browse, filter, join, and export — right from Chrome.</strong>
</p>

<p align="center">
  <a href="https://github.com/greenm8rix/D365-AI-Data-Manager/blob/main/LICENSE"><img src="https://img.shields.io/badge/license-MIT-blue.svg" alt="License" /></a>
  <img src="https://img.shields.io/badge/manifest-v3-green.svg" alt="Manifest V3" />
  <img src="https://img.shields.io/badge/version-2.2.0-purple.svg" alt="Version" />
  <img src="https://img.shields.io/badge/build-no%20bundler-orange.svg" alt="No Bundler" />
</p>

<p align="center">
  <a href="#-features">Features</a> &bull;
  <a href="#-install">Install</a> &bull;
  <a href="#-ai-setup">AI Setup</a> &bull;
  <a href="#-developer-guide">Developer Guide</a> &bull;
  <a href="#-contributing">Contributing</a>
</p>

---

## What Is This?

Think **"AX 2012 Table Browser"** — but in the browser, with AI, joins, and Power Platform exports.

D365 AI Data Manager is a Chrome extension that connects directly to your Dynamics 365 Finance & Operations OData API. It piggybacks on your existing D365 session — no extra login, no middleware, no backend. Open the extension, pick an entity, and start exploring.

Ask the AI to filter, sort, join, highlight, or export your data using plain English. Or do it manually with the full-featured data grid.

<!--
  SCREENSHOTS: Add screenshots to a `docs/` folder and uncomment these:

  ![Data Grid](docs/screenshots/grid.png)
  ![AI Assistant](docs/screenshots/ai-chat.png)
  ![Power Tools](docs/screenshots/power-tools.png)
-->

---

## Features

### Data Browsing
- Full data grid with **pagination**, **sorting**, and **column reordering**
- **Advanced filters** — equals, contains, greater than, date ranges, null checks, and more
- **Quick search** across all visible columns
- **Column visibility** toggle — show/hide fields, reorder, resize
- **Card view** for a different perspective on your data
- **Dark mode**

### Cross-Entity Joins
- **$expand** OData navigation properties (1:N and N:1 relationships)
- **Manual joins** between any two entities on matching fields
- Inner join filtering — see only rows that match

### AI Assistant
- **Natural language queries** — "show me all vendors with overdue invoices"
- AI generates and executes JavaScript to drive the extension (filter, sort, export, highlight)
- **Multi-provider support:**

  | Provider | Notes |
  |----------|-------|
  | **Ollama** | Free, fully local, private. Recommended for enterprise. |
  | **Google Gemini** | Free tier available |
  | **OpenAI** | GPT-4o, GPT-4, etc. |
  | **Anthropic** | Claude models |
  | **OpenRouter** | Access 100+ models through one API |
  | **inferenc.es** | Built-in analysis platform |
  | **Custom endpoint** | Any OpenAI-compatible API |

- **Auto-execute** mode (opt-in) — AI runs code automatically without confirmation
- AI auto-declutters columns after joins to keep the view clean

### Export & Power Platform
- **CSV**, **Excel (.xlsx)**, **JSON**, **SQL INSERT** statements
- **Power BI** — `.pbids` connection files and M Query
- **Power Automate** — Flow definitions and connector settings
- **Power Apps** — Power Fx `ClearCollect()` formulas
- Export full dataset or selected rows only
- All exports support joined/expanded data

### Other
- **Favorites** and **recent entities** for quick access
- **Saved queries** — store and reuse filter+sort combos
- **Data profiling** — quick stats on column values
- **Keyboard shortcuts** for power users
- **Cross-company** queries enabled by default

---

## Install

### From Chrome Web Store

> Coming soon — or install as an unpacked extension below.

### As Unpacked Extension (Developer)

1. **Clone the repo**
   ```bash
   git clone https://github.com/greenm8rix/D365-AI-Data-Manager.git
   ```

2. **Open Chrome** and navigate to `chrome://extensions/`

3. **Enable Developer Mode** (toggle in the top-right corner)

4. **Click "Load unpacked"** and select the cloned folder

5. **Navigate to your D365 F&O environment** (e.g., `https://yourenv.operations.dynamics.com`)

6. **Click the extension icon** in the toolbar — you're in!

> No build step needed. No `npm install`. It's plain JavaScript loaded directly by Chrome.

---

## AI Setup

The AI assistant works with any supported provider. Here's how to get started with **Ollama** (free, local, private):

### Ollama (Recommended for Enterprise)

Ollama runs LLMs entirely on your machine — your D365 data never leaves your computer.

1. **Install Ollama** from [ollama.com](https://ollama.com)

2. **Pull a model** (pick one):
   ```bash
   # Recommended for most machines (8GB+ RAM)
   ollama pull llama3.1

   # Lighter option (4GB RAM)
   ollama pull llama3.2

   # Best quality (16GB+ RAM)
   ollama pull qwen2.5:14b
   ```

3. **Start Ollama** — it runs on `http://localhost:11434` by default

4. **In the extension**, click the **AI Settings** button (gear icon in the AI panel)

5. Select **Ollama** as the provider. The model list auto-populates.

6. **Start chatting** — ask "show me all customers" or "filter where amount > 1000"

### Other Providers

For cloud providers (Gemini, OpenAI, Anthropic, OpenRouter):

1. Get an API key from the provider
2. Open **AI Settings** in the extension
3. Select the provider and paste your key
4. Pick a model from the dropdown

> API keys are stored locally in `chrome.storage.local` — they never leave your browser.

---

## Architecture

```
d365-table-browser/
├── manifest.json              # Chrome MV3 config
├── background/
│   └── service-worker.js      # All HTTP requests route through here (CORS bypass)
├── popup/
│   ├── popup.html             # Extension popup — entity search + favorites
│   ├── popup.js
│   └── popup.css
├── browser/
│   ├── browser.html           # Main data browser page
│   ├── browser.js             # Core grid logic (~4200 lines)
│   ├── browser.css            # All styles
│   ├── ai-assistant.js        # Agentic AI chat panel
│   ├── ai-settings.js         # AI provider config + auth
│   ├── ai-analyze.js          # AI data analysis via inferenc.es
│   └── power-tools.js         # Power BI / Automate / Apps exports
├── shared/
│   ├── odata-client.js        # OData queries + $metadata parsing
│   ├── storage.js             # chrome.storage.local wrapper
│   └── svg-icons.js           # Inline SVG icons
├── icons/                     # Extension icons (16/48/128px + SVG)
└── privacy-policy.html        # Privacy policy
```

**Key patterns:**
- All HTTP requests go through the **service worker** via `chrome.runtime.sendMessage()` — never use `fetch()` directly in browser/popup pages
- Scripts load in dependency order as plain `<script>` tags — no bundler, no modules
- AI-generated code runs via `safeExecute()` — statement-by-statement eval with error handling
- The `_aiDeferLoadData` flag batches multiple AI steps into a single `loadData()` call

---

## Developer Guide

### Prerequisites
- Chrome or Chromium-based browser
- Access to a D365 Finance & Operations environment
- (Optional) [Ollama](https://ollama.com) for local AI

### Making Changes

1. Edit any `.js`, `.html`, or `.css` file
2. Go to `chrome://extensions/`
3. Click the **reload** button on the extension card
4. Reopen the popup or refresh the browser page

That's it. No compilation, no hot reload needed.

### Key Globals (browser page)

These are available on `window` and used by the AI assistant:

| Global | Type | Description |
|--------|------|-------------|
| `currentEntity` | string | Active entity name |
| `data` | array | Current page of records |
| `entitySchema` | object | Field names, types, and metadata |
| `visibleColumns` | array | Currently displayed columns |
| `filterConfig` | array | Active filters |
| `sortConfig` | object | `{ field, direction }` |
| `odataClient` | ODataClient | OData query interface |

### Key Functions (browser page)

| Function | What it does |
|----------|-------------|
| `loadData()` | Fetch data with current filters/sort/page |
| `addFilter(field, op, value)` | Add a filter and reload |
| `sortByColumn(field, dir)` | Sort and reload |
| `setVisibleColumns(cols)` | Show only these columns |
| `highlightCells(field, op, value, color)` | Conditional cell highlighting |
| `exportData(format)` | Export as csv/excel/json/sql |
| `joinEntity(target, sourceKey, targetKey)` | Manual cross-entity join |

### Service Worker Message Protocol

All browser/popup pages communicate with the service worker via messages:

```javascript
// Example: Fetch OData
chrome.runtime.sendMessage({
  action: 'odataFetch',
  url: 'https://myenv.operations.dynamics.com/data/Customers?$top=10',
  options: { method: 'GET' }
}, response => {
  if (response.success) console.log(response.data);
});
```

Actions: `odataFetch`, `aiApiCall`, `abortAiCall`, `fetchModels`, `getEnvironment`, `inferencesLogin`, `inferencesRefresh`, `inferencesSendOtp`, `inferencesSignUp`, `inferencesVerifyOtp`, `inferencesUpload`

---

## Contributing

Contributions are welcome! This project is intentionally simple — no build tools, no framework dependencies.

### How to Contribute

1. **Fork** this repository
2. **Clone** your fork
   ```bash
   git clone https://github.com/YOUR_USERNAME/D365-AI-Data-Manager.git
   ```
3. **Create a branch** for your feature/fix
   ```bash
   git checkout -b feat/my-feature
   ```
4. **Make your changes** — test by loading as an unpacked extension
5. **Commit** with a clear message
   ```bash
   git commit -m "feat: add support for batch delete"
   ```
6. **Push** and open a **Pull Request**

### Guidelines

- **No build tools.** Keep it as plain JS/HTML/CSS.
- **No external dependencies.** Everything runs standalone in Chrome.
- **All HTTP requests** must go through the service worker — never `fetch()` directly.
- **Test on a real D365 environment** if possible.
- **Follow the existing code style** — check surrounding code before writing.
- Use **conventional commits**: `feat:`, `fix:`, `refactor:`, `docs:`, `chore:`

### Ideas for Contributions

- New export formats
- Additional AI provider integrations
- Accessibility improvements
- Localization / i18n
- Performance optimizations for large datasets
- New filter operators

---

## Privacy

Your data stays in your browser. The extension:

- Authenticates using your **existing D365 session cookies** — no credentials stored
- Stores AI API keys in **`chrome.storage.local`** (local to your browser)
- Makes direct API calls to D365 and AI providers — **no middleman server**
- With Ollama, everything runs **100% locally**

See [privacy-policy.html](privacy-policy.html) for the full privacy policy.

---

## License

[MIT](LICENSE) — use it, fork it, ship it.

---

<p align="center">
  Built for the D365 F&O community by <a href="https://github.com/greenm8rix">greenm8rix</a>
</p>

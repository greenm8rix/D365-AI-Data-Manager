# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

A Chrome Extension (Manifest V3) that browses Dynamics 365 Finance & Operations Data Entities via the OData API. Think "AX 2012 Table Browser" but in the browser. It supports filtering, sorting, data profiling, export (CSV/Excel/JSON/SQL), Power Platform integration, and AI-powered query assistance.

## Development

No build system, bundler, or transpiler. All files are plain JavaScript loaded directly by Chrome.

**To develop/test:** Load as an unpacked extension in `chrome://extensions/` (enable Developer Mode). After code changes, click the reload button on the extension card, then reopen the popup or refresh the browser page.

**To test on a live environment:** Navigate to a D365 F&O instance (e.g., `*.operations.dynamics.com`) first, then click the extension icon. The extension piggybacks on the user's existing D365 session cookies for authentication.

**Ignored directory:** `video/` is a Remotion project for generating promo videos — not part of the extension.

## Architecture

### Extension Entry Points
- **`manifest.json`** - Manifest V3 config. Defines permissions, host patterns, and entry points.
- **`background/service-worker.js`** - Background service worker. ALL network requests (OData + AI APIs) are routed here via `chrome.runtime.sendMessage()` to bypass CORS. Handles: `odataFetch`, `aiApiCall`, `fetchModels`, `inferences*` auth flows, and `abortAiCall`.
- **`popup/popup.html`** - Extension popup. Shows entity search with favorites/recent/categorized lists. Selecting an entity opens the browser page in a new tab.
- **`browser/browser.html`** - Main data browser page. Full-page grid with all features (filters, exports, AI, Power Platform tools).

### Shared Modules (`shared/`)
- **`odata-client.js`** - `ODataClient` class (global `odataClient`). Environment detection, `$metadata` XML parsing, entity querying with OData parameters, schema inference, metadata caching (1hr TTL in `chrome.storage.local`).
- **`storage.js`** - `StorageManager` singleton. Wraps `chrome.storage.local` for favorites, recent entities, saved queries, query history, settings, column preferences, AI settings, and environment data.
- **`svg-icons.js`** - `SVGIcons` object with inline SVGs for Power Platform buttons.

### Browser Page Modules (`browser/`)
- **`browser.js`** - Core browser logic (~4200 lines). Manages data grid rendering, pagination, filtering, sorting, column visibility, context menus, export, row selection, entity switching, manual joins, related entity $expand, keyboard shortcuts, dark mode, and settings modals. Exposes key globals: `currentEntity`, `data`, `entitySchema`, `visibleColumns`, `filterConfig`, `sortConfig`, `pageSize`, `currentPage`, `activeJoin`.
- **`ai-assistant.js`** - `AIAssistant` object. Agentic LLM chat that takes a DOM snapshot and generates JavaScript to drive the extension (filter, sort, export, highlight). Supports Gemini, OpenAI, Anthropic, OpenRouter, Ollama, and custom endpoints. Floating draggable panel.
- **`ai-settings.js`** - `AISettings` object. Manages AI provider config, API key storage, model fetching/caching, and inferenc.es Supabase auth (email/password, OTP, token refresh).
- **`ai-analyze.js`** - `AIAnalyze` object. Exports current data as CSV, uploads to inferenc.es for AI analysis, shows animated progress overlay.
- **`power-tools.js`** - `PowerTools` + `PowerToolsUI` objects. Generates exports for Power BI (.pbids, M Query), Power Automate (flow definitions, connector settings), and Power Apps (Power Fx ClearCollect formulas). Supports joins.

### Script Loading Order (browser.html)
Scripts are loaded as plain `<script>` tags in dependency order: `odata-client.js` -> `storage.js` -> `svg-icons.js` -> `power-tools.js` -> `ai-settings.js` -> `ai-assistant.js` -> `ai-analyze.js` -> `browser.js`. All browser modules share globals via the window scope.

## Key Patterns

- **All HTTP requests go through the service worker.** Never use `fetch()` directly in popup or browser pages. Always `chrome.runtime.sendMessage({ action: 'odataFetch', url, options })`.
- **Service worker message actions:** `odataFetch`, `getEnvironment`, `aiApiCall`, `abortAiCall`, `fetchModels`, `inferencesLogin`, `inferencesRefresh`, `inferencesSendOtp`, `inferencesSignUp`, `inferencesVerifyOtp`, `inferencesUpload`. All handlers return `true` from the listener to keep the async channel open.
- **AI API calls use an allowlist** (`AI_ALLOWED_HOSTS` in service-worker.js). Custom endpoints bypass the allowlist with `skipAllowlist: true` but must use HTTPS (localhost exempted) and cannot target internal network ranges.
- **OData queries default to `cross-company=true`** so results aren't limited to the user's default legal entity.
- **Entity metadata comes from the `$metadata` XML endpoint** and is parsed client-side with DOMParser. Entity schemas are validated against a real `$top=1` probe query to ensure only actual fields are used in `$select`.
- **Modals and panels** are defined inline in `browser.html` and toggled with the `hidden` CSS class.
- **The AI assistant generates JavaScript code** that gets executed via `safeExecute()` (statement-by-statement eval) to manipulate the browser page state. It calls existing globals like `addFilter()`, `loadData()`, `sortByColumn()`, `highlightCells()`, `exportData()`, `setVisibleColumns()`, `joinEntity()`, etc.
- **AI batch control:** When the AI runs multi-step code, `window._aiDeferLoadData = true` prevents each step from triggering `loadData()`. A single `loadData()` fires at the end. Check `_aiDeferLoadData` / `_aiLoadDataNeeded` when adding data-mutating functions.
- **OData path is configurable** via settings (`odataClient.odataPath`, default `/data/`). Always use `odataClient.odataPath` rather than hardcoding `/data/`.

# currentDate
Today's date is 2026-02-17.

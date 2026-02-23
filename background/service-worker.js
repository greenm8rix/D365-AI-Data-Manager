/**
 * D365 AI Data Manager - Background Service Worker
 * Handles cross-origin API requests with proper cookie access
 */

// Active AI request AbortController — allows cancellation from content script
let activeAIAbort = null;

// Listen for messages from popup and browser pages
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'abortAiCall') {
    if (activeAIAbort) {
      activeAIAbort.abort();
      activeAIAbort = null;
    }
    sendResponse({ ok: true });
    return false;
  }

  if (request.action === 'odataFetch') {
    handleODataFetch(request)
      .then(sendResponse)
      .catch(error => sendResponse({ error: error.message }));
    return true; // Keep channel open for async response
  }

  if (request.action === 'getEnvironment') {
    getStoredEnvironment()
      .then(sendResponse)
      .catch(error => sendResponse({ error: error.message }));
    return true;
  }

  if (request.action === 'aiApiCall') {
    handleAIApiCall(request)
      .then(sendResponse)
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true;
  }

  if (request.action === 'inferencesUpload') {
    handleInferencesUpload(request)
      .then(sendResponse)
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true;
  }

  if (request.action === 'fetchModels') {
    handleFetchModels(request)
      .then(sendResponse)
      .catch(error => sendResponse({ success: false, error: error.message }));
    return true;
  }

});

async function handleODataFetch(request) {
  const { url, options = {} } = request;

  // Validate URL — only send session cookies to D365-like domains
  try {
    const urlObj = new URL(url);
    const host = urlObj.hostname.toLowerCase();
    const isD365 = host.endsWith('.dynamics.com') || host.endsWith('.dynamics.us') ||
                   host.endsWith('.dynamics.cn') || host.endsWith('.cloudax.dynamics.com');
    // Also allow manually-configured environments stored in chrome.storage
    let isManual = false;
    if (!isD365) {
      const stored = await chrome.storage.local.get(['currentEnvironment']);
      const env = stored.currentEnvironment;
      if (env && env.baseUrl) {
        try { isManual = new URL(env.baseUrl).hostname.toLowerCase() === host; } catch {}
      }
    }
    if (!isD365 && !isManual) {
      return { success: false, error: 'Blocked: OData URL does not match a D365 environment' };
    }
  } catch (e) {
    return { success: false, error: 'Invalid OData URL' };
  }

  try {
    const response = await fetch(url, {
      ...options,
      credentials: 'include',
      headers: {
        'Accept': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
        'Content-Type': 'application/json',
        'Prefer': 'odata.include-annotations="*"',
        ...options.headers
      }
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP ${response.status}: ${errorText.substring(0, 500)}`);
    }

    // Check content type
    const contentType = response.headers.get('content-type');

    if (contentType && contentType.includes('application/json')) {
      const data = await response.json();
      return { success: true, data };
    } else if (contentType && contentType.includes('application/xml')) {
      const text = await response.text();
      return { success: true, data: text, isXml: true };
    } else {
      const text = await response.text();
      return { success: true, data: text, isText: true };
    }
  } catch (error) {
    console.error('Background fetch error:', error);
    return { success: false, error: error.message };
  }
}

async function getStoredEnvironment() {
  const result = await chrome.storage.local.get(['currentEnvironment']);
  return result.currentEnvironment || null;
}

const AI_ALLOWED_HOSTS = [
  'generativelanguage.googleapis.com',
  'api.openai.com',
  'api.anthropic.com',
  'openrouter.ai',
  'api.inferenc.es',
  'inferenc.es'
];

function isPrivateIP(hostname) {
  // Block internal network ranges for non-localhost custom endpoints
  if (/^(10\.|172\.(1[6-9]|2[0-9]|3[01])\.|192\.168\.)/.test(hostname)) return true;
  if (hostname === '0.0.0.0' || hostname === '[::]') return true;
  return false;
}

async function handleAIApiCall(request) {
  const { url, options = {}, skipAllowlist = false } = request;

  // Validate URL against allowlist (skipped for user-configured custom endpoints)
  try {
    const urlObj = new URL(url);
    if (!skipAllowlist && !AI_ALLOWED_HOSTS.includes(urlObj.hostname)) {
      return { success: false, error: 'Blocked: URL not in AI endpoint allowlist' };
    }
    // For custom endpoints (skipAllowlist), still enforce basic safety:
    // - Allow localhost/127.0.0.1 (Ollama)
    // - Require HTTPS for all other hosts
    // - Block internal network ranges
    if (skipAllowlist) {
      const isLocal = urlObj.hostname === 'localhost' || urlObj.hostname === '127.0.0.1';
      if (!isLocal) {
        if (urlObj.protocol !== 'https:') {
          return { success: false, error: 'Custom endpoints must use HTTPS (except localhost)' };
        }
        if (isPrivateIP(urlObj.hostname)) {
          return { success: false, error: 'Internal network addresses are not allowed' };
        }
      }
    }
  } catch (e) {
    return { success: false, error: 'Invalid URL' };
  }

  console.log('Background: AI API call to', url.replace(/key=[^&]+/, 'key=***'));

  // Create AbortController so content script can cancel in-flight requests
  activeAIAbort = new AbortController();
  const signal = activeAIAbort.signal;

  try {
    const response = await fetch(url, {
      method: options.method || 'POST',
      headers: options.headers || {},
      body: options.body || undefined,
      signal
    });

    activeAIAbort = null;

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`HTTP ${response.status}: ${errorText.substring(0, 500)}`);
    }

    const data = await response.json();
    return { success: true, data };
  } catch (error) {
    activeAIAbort = null;
    if (error.name === 'AbortError') {
      return { success: false, error: 'Aborted by user' };
    }
    console.error('AI API call error:', error);
    return { success: false, error: error.message };
  }
}

// ==================== FETCH MODELS ====================

async function handleFetchModels(request) {
  const { provider, apiKey, ollamaPort } = request;

  try {
    switch (provider) {
      case 'openai': {
        const resp = await fetch('https://api.openai.com/v1/models', {
          headers: { 'Authorization': `Bearer ${apiKey}` }
        });
        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        const data = await resp.json();
        const models = (data.data || [])
          .filter(m => /^(gpt-|o1|o3|o4|chatgpt)/.test(m.id))
          .map(m => ({ value: m.id, label: m.id }))
          .sort((a, b) => a.label.localeCompare(b.label));
        return { success: true, models };
      }

      case 'gemini': {
        const resp = await fetch(`https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`);
        if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
        const data = await resp.json();
        const models = (data.models || [])
          .filter(m => {
            const id = m.name || '';
            // Only include actual Gemini chat models, exclude embedding/tuning/AQA/code etc.
            if (!id.includes('gemini')) return false;
            if (id.includes('embedding') || id.includes('tuning') || id.includes('aqa')) return false;
            return m.supportedGenerationMethods?.includes('generateContent');
          })
          .map(m => ({ value: m.name.replace('models/', ''), label: m.displayName || m.name.replace('models/', '') }))
          .sort((a, b) => a.label.localeCompare(b.label));
        return { success: true, models };
      }

      case 'openrouter': {
        const resp = await fetch('https://openrouter.ai/api/v1/models');
        if (!resp.ok) throw new Error(`OpenRouter HTTP ${resp.status}`);
        const data = await resp.json();
        // OpenRouter returns { data: [...] } array of models
        const allModels = data.data || data.models || [];
        if (allModels.length === 0) throw new Error('OpenRouter returned no models');
        const models = allModels
          .map(m => ({ value: m.id, label: m.name || m.id }))
          .sort((a, b) => a.label.localeCompare(b.label));
        return { success: true, models };
      }

      case 'ollama': {
        const port = ollamaPort || 11434;
        let resp;
        try {
          resp = await fetch(`http://localhost:${port}/api/tags`);
        } catch (e) {
          return { success: false, error: `Cannot reach Ollama at localhost:${port}. Is Ollama running? (ollama serve)` };
        }
        if (!resp.ok) throw new Error(`Ollama HTTP ${resp.status}`);
        const data = await resp.json();
        const models = (data.models || [])
          .map(m => {
            const size = m.size ? ` (${(m.size / 1e9).toFixed(1)}GB)` : '';
            return { value: m.name, label: m.name + size };
          });
        if (models.length === 0) {
          return { success: false, error: 'No models installed. Run: ollama pull llama3.3' };
        }
        return { success: true, models };
      }

      case 'anthropic': {
        const resp = await fetch('https://api.anthropic.com/v1/models', {
          headers: {
            'x-api-key': apiKey,
            'anthropic-version': '2023-06-01'
          }
        });
        if (!resp.ok) throw new Error(`Anthropic HTTP ${resp.status}`);
        const data = await resp.json();
        const models = (data.data || [])
          .map(m => ({ value: m.id, label: m.display_name || m.id }))
          .sort((a, b) => a.label.localeCompare(b.label));
        if (models.length === 0) throw new Error('No models returned');
        return { success: true, models };
      }

      default:
        return { success: false, error: 'Model fetching not supported for this provider' };
    }
  } catch (error) {
    console.error('Fetch models error:', error);
    return { success: false, error: error.message };
  }
}

// ==================== INFERENC.ES UPLOAD ====================

async function handleInferencesUpload(request) {
  const { csvContent, filename, accessToken } = request;
  if (!csvContent || !accessToken) return { success: false, error: 'Missing data or token' };

  try {
    // Step 1: Upload file as FormData
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const formData = new FormData();
    formData.append('files', blob, filename || 'data.csv');

    const uploadResp = await fetch('https://api.inferenc.es/api/upload-raw-file', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${accessToken}`
      },
      body: formData
    });

    if (!uploadResp.ok) {
      const errText = await uploadResp.text();
      return { success: false, error: `Upload failed: ${errText.substring(0, 200)}` };
    }

    const uploadData = await uploadResp.json();
    return { success: true, data: uploadData };
  } catch (error) {
    return { success: false, error: error.message };
  }
}


// When extension is installed or updated
chrome.runtime.onInstalled.addListener(() => {
  console.log('D365 AI Data Manager installed/updated');
});

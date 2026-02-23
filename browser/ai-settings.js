/**
 * D365 AI Data Manager - AI Settings Manager
 * Controls AI feature toggle, API key management, provider selection,
 * and inferenc.es account linking via chrome.identity.launchWebAuthFlow().
 */

const AISettings = {
  settings: null,
  modelCache: {},

  async init() {
    this.settings = await StorageManager.getAISettings();
    this.applyVisibility();
  },

  applyVisibility() {
    const container = document.getElementById('aiButtonsContainer');
    if (!container) return;

    if (this.settings && this.settings.enabled) {
      container.classList.remove('hidden');
      const assistantBtn = document.getElementById('aiAssistantBtn');
      const analyzeBtn = document.getElementById('aiAnalyzeBtn');
      if (assistantBtn) {
        assistantBtn.classList.toggle('hidden', !this.isConfigured());
      }
      if (analyzeBtn) {
        analyzeBtn.classList.toggle('hidden', !this.settings.inferencesLinked && !this.isConfigured());
      }
    } else {
      container.classList.add('hidden');
      document.getElementById('aiPanel')?.classList.add('hidden');
    }
  },

  async populateSettingsUI() {
    if (!this.settings) {
      this.settings = await StorageManager.getAISettings();
    }

    const enabledCheckbox = document.getElementById('settingAIEnabled');
    const detailSection = document.getElementById('aiSettingsDetail');
    const providerSelect = document.getElementById('settingAIProvider');
    const apiKeyInput = document.getElementById('settingAIApiKey');
    const modelSelect = document.getElementById('settingAIModel');
    const endpointInput = document.getElementById('settingAIEndpoint');
    const customGroup = document.getElementById('customEndpointGroup');

    if (enabledCheckbox) {
      enabledCheckbox.checked = this.settings.enabled;
      enabledCheckbox.onchange = () => {
        detailSection?.classList.toggle('hidden', !enabledCheckbox.checked);
      };
    }

    if (detailSection) {
      detailSection.classList.toggle('hidden', !this.settings.enabled);
    }

    const customModelGroup = document.getElementById('customModelGroup');
    const customFormatGroup = document.getElementById('customFormatGroup');
    const modelSelectGroup = document.getElementById('modelSelectGroup');
    const customModelInput = document.getElementById('settingCustomModel');
    const customFormatSelect = document.getElementById('settingCustomFormat');

    const ollamaPortGroup = document.getElementById('ollamaPortGroup');
    const ollamaPortInput = document.getElementById('settingOllamaPort');
    const autoExecuteCheckbox = document.getElementById('settingAutoExecute');
    const customPromptTextarea = document.getElementById('settingCustomPrompt');
    const fetchModelsBtn = document.getElementById('fetchModelsBtn');

    if (providerSelect) {
      providerSelect.value = this.settings.provider || '';
      providerSelect.onchange = async () => {
        const val = providerSelect.value;
        const isCustom = val === 'custom';
        const isOllama = val === 'ollama';
        await this.updateModelOptions(val);
        customGroup?.classList.toggle('hidden', !isCustom);
        customModelGroup?.classList.toggle('hidden', !isCustom);
        customFormatGroup?.classList.toggle('hidden', !isCustom);
        modelSelectGroup?.classList.toggle('hidden', isCustom || !val);
        ollamaPortGroup?.classList.toggle('hidden', !isOllama);
        // Update API key placeholder
        if (apiKeyInput) {
          apiKeyInput.placeholder = isOllama ? '(not required for local)' : 'sk-...';
        }
        // Hide fetch button for custom provider or no selection
        if (fetchModelsBtn) {
          fetchModelsBtn.classList.toggle('hidden', isCustom || !val);
        }
      };
      if (this.settings.provider) {
        await this.updateModelOptions(this.settings.provider);
      }
    }

    const isCustom = this.settings.provider === 'custom';
    const isOllama = this.settings.provider === 'ollama';
    if (apiKeyInput) {
      apiKeyInput.value = this.settings.apiKey || '';
      apiKeyInput.placeholder = isOllama ? '(not required for local)' : 'sk-...';
    }
    if (endpointInput) endpointInput.value = this.settings.customEndpoint || '';
    if (customModelInput) customModelInput.value = this.settings.customModel || '';
    if (customFormatSelect) customFormatSelect.value = this.settings.customFormat || 'openai';
    if (ollamaPortInput) ollamaPortInput.value = this.settings.ollamaPort || 11434;
    if (autoExecuteCheckbox) autoExecuteCheckbox.checked = this.settings.autoExecute === true;
    if (customPromptTextarea) customPromptTextarea.value = this.settings.customPrompt || '';
    customGroup?.classList.toggle('hidden', !isCustom);
    customModelGroup?.classList.toggle('hidden', !isCustom);
    customFormatGroup?.classList.toggle('hidden', !isCustom);
    modelSelectGroup?.classList.toggle('hidden', isCustom || !this.settings.provider);
    ollamaPortGroup?.classList.toggle('hidden', !isOllama);
    if (fetchModelsBtn) fetchModelsBtn.classList.toggle('hidden', isCustom || !this.settings.provider);

    if (modelSelect && this.settings.model) {
      modelSelect.value = this.settings.model;
    }

    this.updateInferencesUI();
  },

  updateInferencesUI() {
    const loginForm = document.getElementById('inferencesLoginForm');
    const loggedInInfo = document.getElementById('inferencesLoggedIn');

    if (this.settings.inferencesLinked && this.settings.inferencesEmail) {
      if (loginForm) loginForm.classList.add('hidden');
      if (loggedInInfo) {
        loggedInInfo.classList.remove('hidden');
        const emailEl = document.getElementById('inferencesEmailDisplay');
        if (emailEl) emailEl.textContent = this.settings.inferencesEmail;
      }
    } else {
      if (loginForm) loginForm.classList.remove('hidden');
      if (loggedInInfo) loggedInInfo.classList.add('hidden');
    }
  },

  // ==================== INFERENC.ES CONNECTION ====================

  async connectInferences() {
    try {
      const redirectUrl = chrome.identity.getRedirectURL();
      const authUrl = `https://inferenc.es/?from=extension&redirect_uri=${encodeURIComponent(redirectUrl)}`;

      const responseUrl = await new Promise((resolve, reject) => {
        chrome.identity.launchWebAuthFlow(
          { url: authUrl, interactive: true },
          (callbackUrl) => {
            if (chrome.runtime.lastError) {
              reject(new Error(chrome.runtime.lastError.message));
            } else {
              resolve(callbackUrl);
            }
          }
        );
      });

      // Parse tokens from callback URL
      const url = new URL(responseUrl);
      const params = new URLSearchParams(url.search);
      const accessToken = params.get('access_token');
      const refreshToken = params.get('refresh_token');
      const email = params.get('email');
      const expiresAt = params.get('expires_at');

      if (!accessToken) throw new Error('No access token received');

      await StorageManager.updateAISettings({
        inferencesLinked: true,
        inferencesSession: accessToken,
        inferencesRefreshToken: refreshToken || null,
        inferencesEmail: email || null,
        inferencesExpiresAt: expiresAt ? parseInt(expiresAt) : null
      });

      this.settings = await StorageManager.getAISettings();
      this.updateInferencesUI();
      this.applyVisibility();
      showToast('inferenc.es linked successfully!');
    } catch (e) {
      // User closed the popup — not an error
      if (e.message?.includes('canceled') || e.message?.includes('closed') || e.message?.includes('user')) {
        return;
      }
      showToast('Login failed: ' + e.message, 'error');
    }
  },

  async logoutInferences() {
    await StorageManager.updateAISettings({
      inferencesLinked: false,
      inferencesSession: null,
      inferencesRefreshToken: null,
      inferencesEmail: null,
      inferencesExpiresAt: null
    });

    this.settings = await StorageManager.getAISettings();
    this.updateInferencesUI();
    this.applyVisibility();
    showToast('inferenc.es unlinked');
  },

  async getValidToken() {
    // Reload settings in case the service worker updated them via token handoff
    this.settings = await StorageManager.getAISettings();
    if (!this.settings.inferencesLinked || !this.settings.inferencesSession) return null;

    // If token is expired, prompt user to reconnect
    if (this.settings.inferencesExpiresAt) {
      const now = Math.floor(Date.now() / 1000);
      if (now >= this.settings.inferencesExpiresAt - 60) {
        return null;
      }
    }

    return this.settings.inferencesSession;
  },

  // ==================== AI PROVIDER SETTINGS ====================

  async updateModelOptions(provider) {
    const modelSelect = document.getElementById('settingAIModel');
    if (!modelSelect) return;

    // Check memory cache first, then storage cache
    let cached = this.modelCache[provider];
    if (!cached || Date.now() - cached.ts >= 3600000) {
      try {
        const stored = await chrome.storage.local.get([`modelCache_${provider}`]);
        const storedCache = stored[`modelCache_${provider}`];
        if (storedCache && storedCache.models && Date.now() - storedCache.ts < 3600000) {
          cached = storedCache;
          this.modelCache[provider] = cached;
        }
      } catch (e) { /* ignore */ }
    }

    if (cached && cached.models && cached.models.length > 0) {
      this.populateModelDropdown(cached.models);
      return;
    }

    // No cached models — show prompt to fetch
    modelSelect.innerHTML = '<option value="">-- Click Fetch to load models --</option>';
  },

  async fetchModels() {
    // Read current UI values (not saved settings — user may not have saved yet)
    const provider = document.getElementById('settingAIProvider')?.value;
    if (!provider || provider === 'custom') return;

    const apiKey = document.getElementById('settingAIApiKey')?.value || '';
    const ollamaPort = parseInt(document.getElementById('settingOllamaPort')?.value) || 11434;

    // Always fetch fresh when user clicks Fetch — skip cache
    const fetchBtn = document.getElementById('fetchModelsBtn');
    if (fetchBtn) { fetchBtn.disabled = true; fetchBtn.textContent = '...'; }

    try {
      const resp = await chrome.runtime.sendMessage({
        action: 'fetchModels',
        provider,
        apiKey,
        ollamaPort
      });

      if (!resp || !resp.success) throw new Error(resp?.error || 'Failed to fetch models');
      if (!resp.models || resp.models.length === 0) throw new Error('No models returned');

      // Cache in memory and persist to storage
      this.modelCache[provider] = { models: resp.models, ts: Date.now() };
      await chrome.storage.local.set({ [`modelCache_${provider}`]: { models: resp.models, ts: Date.now() } });
      this.populateModelDropdown(resp.models);
      showToast(`${resp.models.length} models fetched from ${provider}`);
    } catch (error) {
      showToast('Fetch models failed: ' + error.message, 'error');
    } finally {
      if (fetchBtn) { fetchBtn.disabled = false; fetchBtn.textContent = 'Fetch'; }
    }
  },

  populateModelDropdown(models) {
    const modelSelect = document.getElementById('settingAIModel');
    if (!modelSelect || !models.length) return;

    const currentVal = modelSelect.value;
    const esc = (s) => { const d = document.createElement('div'); d.textContent = s; return d.innerHTML; };
    modelSelect.innerHTML = models.map(m =>
      `<option value="${esc(m.value)}">${esc(m.label)}</option>`
    ).join('');

    // Try to keep current selection
    if (currentVal && [...modelSelect.options].some(o => o.value === currentVal)) {
      modelSelect.value = currentVal;
    }
  },

  // Request optional host permission for custom endpoints only
  async requestCustomEndpointPermission(customEndpoint) {
    if (!customEndpoint) return true;
    try {
      const url = new URL(customEndpoint);
      return await chrome.permissions.request({
        origins: [`${url.protocol}//${url.hostname}/*`]
      });
    } catch (e) {
      console.warn('Permission request failed:', e);
      return false;
    }
  },

  async saveFromUI() {
    const provider = document.getElementById('settingAIProvider')?.value || '';
    const enabled = document.getElementById('settingAIEnabled')?.checked || false;
    const customEndpoint = document.getElementById('settingAIEndpoint')?.value || '';

    // Only request permission for custom endpoints (all others are in host_permissions)
    if (enabled && provider === 'custom' && customEndpoint) {
      const granted = await this.requestCustomEndpointPermission(customEndpoint);
      if (!granted) {
        showToast('Permission denied — need network access to your custom endpoint', 'error');
        return;
      }
    }

    const updates = {
      enabled,
      provider: provider,
      apiKey: document.getElementById('settingAIApiKey')?.value || '',
      model: provider === 'custom'
        ? (document.getElementById('settingCustomModel')?.value || '')
        : (document.getElementById('settingAIModel')?.value || ''),
      customModel: document.getElementById('settingCustomModel')?.value || '',
      customEndpoint: customEndpoint,
      customFormat: document.getElementById('settingCustomFormat')?.value || 'openai',
      ollamaPort: (() => { const p = parseInt(document.getElementById('settingOllamaPort')?.value); return (p > 0 && p <= 65535) ? p : 11434; })(),
      autoExecute: document.getElementById('settingAutoExecute')?.checked !== false,
      customPrompt: document.getElementById('settingCustomPrompt')?.value || ''
    };

    this.settings = await StorageManager.updateAISettings(updates);
    this.applyVisibility();
  },

  getApiEndpoint() {
    if (!this.settings) return '';
    switch (this.settings.provider) {
      case 'gemini':
        return `https://generativelanguage.googleapis.com/v1beta/models/${this.settings.model || 'gemini-2.0-flash'}:generateContent?key=${this.settings.apiKey}`;
      case 'openai':
        return 'https://api.openai.com/v1/chat/completions';
      case 'anthropic':
        return 'https://api.anthropic.com/v1/messages';
      case 'openrouter':
        return 'https://openrouter.ai/api/v1/chat/completions';
      case 'ollama': {
        const port = this.settings.ollamaPort || 11434;
        return `http://localhost:${port}/v1/chat/completions`;
      }
      case 'custom':
        return this.settings.customEndpoint;
      default:
        return '';
    }
  },

  isConfigured() {
    if (!this.settings || !this.settings.enabled || !this.settings.provider) return false;
    // Ollama doesn't need an API key
    if (this.settings.provider === 'ollama') return true;
    return this.settings.apiKey && this.settings.apiKey.length > 0;
  }
};

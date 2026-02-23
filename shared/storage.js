/**
 * Storage Manager for D365 AI Data Manager
 * Handles favorites, recent entities, saved queries, and settings
 */

const StorageManager = {
  /**
   * Get a value from storage
   */
  async get(key) {
    const result = await chrome.storage.local.get([key]);
    return result[key];
  },

  /**
   * Set a value in storage
   */
  async set(key, value) {
    await chrome.storage.local.set({ [key]: value });
  },

  /**
   * Remove a value from storage
   */
  async remove(key) {
    await chrome.storage.local.remove([key]);
  },

  /**
   * Get all storage data
   */
  async getAll() {
    return chrome.storage.local.get(null);
  },

  // ==================== FAVORITES ====================

  /**
   * Get favorite entities
   */
  async getFavorites() {
    return (await this.get('favorites')) || [];
  },

  /**
   * Add entity to favorites
   */
  async addFavorite(entityName) {
    const favorites = await this.getFavorites();
    if (!favorites.includes(entityName)) {
      favorites.unshift(entityName);
      await this.set('favorites', favorites);
    }
  },

  /**
   * Remove entity from favorites
   */
  async removeFavorite(entityName) {
    const favorites = await this.getFavorites();
    const index = favorites.indexOf(entityName);
    if (index > -1) {
      favorites.splice(index, 1);
      await this.set('favorites', favorites);
    }
  },

  /**
   * Check if entity is favorite
   */
  async isFavorite(entityName) {
    const favorites = await this.getFavorites();
    return favorites.includes(entityName);
  },

  // ==================== RECENT ENTITIES ====================

  /**
   * Get recently viewed entities
   */
  async getRecent() {
    return (await this.get('recentEntities')) || [];
  },

  /**
   * Add entity to recent list
   */
  async addRecent(entityName) {
    let recent = await this.getRecent();

    // Remove if already exists
    recent = recent.filter(e => e !== entityName);

    // Add to front
    recent.unshift(entityName);

    // Keep only last 10
    recent = recent.slice(0, 10);

    await this.set('recentEntities', recent);
  },

  // ==================== SAVED QUERIES ====================

  /**
   * Get all saved queries
   */
  async getSavedQueries() {
    return (await this.get('savedQueries')) || [];
  },

  /**
   * Save a query
   */
  async saveQuery(query) {
    const queries = await this.getSavedQueries();
    const newQuery = {
      id: Date.now().toString(),
      name: query.name,
      entityName: query.entityName,
      filter: query.filter,
      orderBy: query.orderBy,
      select: query.select,
      created: Date.now()
    };
    queries.unshift(newQuery);
    await this.set('savedQueries', queries);
    return newQuery;
  },

  /**
   * Delete a saved query
   */
  async deleteQuery(queryId) {
    let queries = await this.getSavedQueries();
    queries = queries.filter(q => q.id !== queryId);
    await this.set('savedQueries', queries);
  },

  // ==================== QUERY HISTORY ====================

  /**
   * Get query history
   */
  async getQueryHistory() {
    return (await this.get('queryHistory')) || [];
  },

  /**
   * Add to query history
   */
  async addToHistory(query) {
    let history = await this.getQueryHistory();

    const historyItem = {
      entityName: query.entityName,
      filter: query.filter,
      timestamp: Date.now()
    };

    // Avoid duplicates
    history = history.filter(h =>
      h.entityName !== historyItem.entityName ||
      h.filter !== historyItem.filter
    );

    history.unshift(historyItem);

    // Keep last 50
    history = history.slice(0, 50);

    await this.set('queryHistory', history);
  },

  // ==================== SETTINGS ====================

  /**
   * Get all settings
   */
  async getSettings() {
    const defaults = {
      theme: 'light',
      pageSize: 100,
      showRowNumbers: true,
      compactMode: false,
      autoRefresh: false,
      refreshInterval: 30,
      dateFormat: 'yyyy-MM-dd',
      numberFormat: 'en-US',
      odataPath: '/data/'
    };

    const stored = (await this.get('settings')) || {};
    return { ...defaults, ...stored };
  },

  /**
   * Update settings
   */
  async updateSettings(updates) {
    const settings = await this.getSettings();
    Object.assign(settings, updates);
    await this.set('settings', settings);
    return settings;
  },

  // ==================== COLUMN PREFERENCES ====================

  /**
   * Get column preferences for an entity
   */
  async getColumnPrefs(entityName) {
    const prefs = (await this.get('columnPrefs')) || {};
    return prefs[entityName] || null;
  },

  /**
   * Save column preferences for an entity
   */
  async saveColumnPrefs(entityName, columns) {
    const prefs = (await this.get('columnPrefs')) || {};
    prefs[entityName] = {
      columns: columns, // Array of { name, visible, width, order }
      updated: Date.now()
    };
    await this.set('columnPrefs', prefs);
  },

  // ==================== ENVIRONMENTS ====================

  /**
   * Get known environments
   */
  async getEnvironments() {
    return (await this.get('environments')) || [];
  },

  /**
   * Add an environment
   */
  async addEnvironment(env) {
    let environments = await this.getEnvironments();

    // Update if exists, otherwise add
    const index = environments.findIndex(e => e.baseUrl === env.baseUrl);
    if (index > -1) {
      environments[index] = { ...environments[index], ...env, lastUsed: Date.now() };
    } else {
      environments.push({ ...env, lastUsed: Date.now() });
    }

    await this.set('environments', environments);
  },

  // ==================== AI SETTINGS ====================

  async getAISettings() {
    const defaults = {
      enabled: false,
      provider: '',           // '', 'gemini', 'openai', 'anthropic', 'custom'
      apiKey: '',
      model: '',
      customEndpoint: '',
      inferencesLinked: false,
      inferencesSession: null,
      inferencesRefreshToken: null,
      inferencesEmail: null,
      inferencesExpiresAt: null
    };
    const stored = (await this.get('aiSettings')) || {};
    return { ...defaults, ...stored };
  },

  async updateAISettings(updates) {
    const settings = await this.getAISettings();
    Object.assign(settings, updates);
    await this.set('aiSettings', settings);
    return settings;
  }
};

/**
 * D365 OData Client
 * Uses session cookies for authentication - no extra login needed!
 */

class ODataClient {
  constructor() {
    this.baseUrl = null;
    this.environment = null;
    this.metadataCache = null;
    this.odataPath = '/data/';
    this._loadOdataPath();
  }

  async _loadOdataPath() {
    try {
      const result = await chrome.storage.local.get(['settings']);
      const settings = result.settings || {};
      let p = settings.odataPath || '/data/';
      // Normalize: ensure leading and trailing slashes
      if (!p.startsWith('/')) p = '/' + p;
      if (!p.endsWith('/')) p = p + '/';
      this.odataPath = p;
    } catch (e) {
      // Fallback to default
    }
  }

  /**
   * Detect D365 environment from the current browser tab
   */
  async detectEnvironment() {
    try {
      // First check if we have a stored environment (for browser page)
      const stored = await chrome.storage.local.get(['currentEnvironment']);
      if (stored.currentEnvironment) {
        this.environment = stored.currentEnvironment;
        this.baseUrl = this.environment.baseUrl;
        console.log('Using stored environment:', this.environment);
        return this.environment;
      }

      // Try to detect from active tab (for popup)
      // host_permissions for *.dynamics.com/* lets us read tab.url for D365 tabs without "tabs" permission
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (!tab || !tab.url) return null;

      const url = new URL(tab.url);

      // D365 F&O patterns:
      // xxx.operations.dynamics.com (prod)
      // xxx.sandbox.operations.dynamics.com (sandbox/uat)
      // xxx-dev.operations.dynamics.com (dev)
      // Government: xxx.operations.dynamics.us, xxx.operations.dynamics.cn, etc.
      if (url.hostname.includes('.operations.dynamics.') ||
          url.hostname.includes('.dynamics.com') ||
          url.hostname.includes('.dynamics.us') ||
          url.hostname.includes('.dynamics.cn')) {

        const hostname = url.hostname;
        const envName = hostname.split('.')[0];

        // Determine environment type
        let envType = 'PROD';
        if (hostname.includes('sandbox')) {
          envType = 'UAT';
        } else if (envName.includes('dev') || envName.includes('-dev')) {
          envType = 'DEV';
        } else if (envName.includes('uat') || envName.includes('test')) {
          envType = 'UAT';
        }
        // Government cloud indicator
        if (hostname.includes('.dynamics.us')) {
          envType = 'GOV';
        }

        this.environment = {
          baseUrl: `${url.protocol}//${hostname}`,
          envName: envName,
          envType: envType,
          hostname: hostname,
          source: 'auto'
        };

        this.baseUrl = this.environment.baseUrl;

        // Store for browser page to use
        await chrome.storage.local.set({ currentEnvironment: this.environment });
        console.log('Detected and stored environment:', this.environment);

        return this.environment;
      }

      // Auto-detect failed — check for a manually configured URL
      const manual = await chrome.storage.local.get(['manualEnvironment']);
      if (manual.manualEnvironment) {
        this.environment = manual.manualEnvironment;
        this.baseUrl = this.environment.baseUrl;
        await chrome.storage.local.set({ currentEnvironment: this.environment });
        console.log('Using manual environment:', this.environment);
        return this.environment;
      }

      return null;
    } catch (error) {
      console.error('Error detecting environment:', error);
      return null;
    }
  }

  /**
   * Set environment URL manually (for gov/custom domains)
   */
  async setManualEnvironment(urlString) {
    try {
      const url = new URL(urlString);
      if (url.protocol !== 'https:' && url.protocol !== 'http:') {
        throw new Error('URL must use https:// or http://');
      }

      const hostname = url.hostname;
      const envName = hostname.split('.')[0];

      // Try to guess env type from hostname
      let envType = 'CUSTOM';
      if (hostname.includes('.dynamics.us')) envType = 'GOV';
      else if (hostname.includes('.dynamics.cn')) envType = 'GOV-CN';
      else if (hostname.includes('sandbox')) envType = 'UAT';
      else if (hostname.includes('dev')) envType = 'DEV';

      this.environment = {
        baseUrl: `${url.protocol}//${hostname}`,
        envName: envName,
        envType: envType,
        hostname: hostname,
        source: 'manual'
      };

      this.baseUrl = this.environment.baseUrl;

      // Store both as current and as manual fallback
      await chrome.storage.local.set({
        currentEnvironment: this.environment,
        manualEnvironment: this.environment
      });

      // Clear metadata cache (different environment = different entities)
      this.metadataCache = null;
      await StorageManager.remove('entityMetadata');

      console.log('Manual environment set:', this.environment);
      return this.environment;
    } catch (error) {
      throw new Error(`Invalid URL: ${error.message}`);
    }
  }

  /**
   * Clear manually configured environment
   */
  async clearManualEnvironment() {
    await chrome.storage.local.remove(['manualEnvironment']);
    // If current env was manual, clear it too
    if (this.environment && this.environment.source === 'manual') {
      this.environment = null;
      this.baseUrl = null;
      await chrome.storage.local.remove(['currentEnvironment']);
    }
  }

  /**
   * Make an authenticated OData request via background service worker
   * This bypasses CORS by using extension permissions
   */
  async fetch(endpoint, options = {}) {
    if (!this.baseUrl) {
      await this.detectEnvironment();
    }

    if (!this.baseUrl) {
      throw new Error('Not connected to a D365 environment. Please navigate to D365 first and reopen the extension.');
    }

    const url = endpoint.startsWith('http')
      ? endpoint
      : `${this.baseUrl}${this.odataPath}${endpoint}`;

    console.log('OData fetch:', url);

    // Route through background service worker to bypass CORS
    const response = await chrome.runtime.sendMessage({
      action: 'odataFetch',
      url: url,
      options: options
    });

    if (!response) {
      throw new Error('No response from background service worker');
    }

    if (!response.success) {
      throw new Error(response.error || 'Unknown error');
    }

    return response.data;
  }

  /**
   * Get list of all data entities
   */
  async getEntities() {
    // First try to get from cache
    const metadata = await this.getMetadata();
    if (metadata && metadata.entities) {
      return metadata.entities;
    }

    if (!this.baseUrl) {
      await this.detectEnvironment();
    }

    if (!this.baseUrl) {
      throw new Error('Not connected to a D365 environment');
    }

    // Fetch from $metadata endpoint via background service worker
    try {
      console.log('Fetching metadata from:', `${this.baseUrl}${this.odataPath}$metadata`);

      const response = await chrome.runtime.sendMessage({
        action: 'odataFetch',
        url: `${this.baseUrl}${this.odataPath}$metadata`,
        options: {
          headers: {
            'Accept': 'application/xml'
          }
        }
      });

      if (!response || !response.success) {
        throw new Error(response?.error || 'Failed to fetch metadata');
      }

      const xmlText = response.data;
      const entities = this.parseMetadataXml(xmlText);

      // Cache the results
      await this.cacheMetadata(entities);

      return entities;
    } catch (error) {
      console.error('Error fetching metadata:', error);
      throw error;
    }
  }

  /**
   * Parse the OData $metadata XML to extract entity information
   */
  parseMetadataXml(xmlText) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlText, 'application/xml');

    const entities = [];
    const entityTypes = doc.querySelectorAll('EntityType');
    const entitySets = doc.querySelectorAll('EntitySet');

    // Build a map of EntitySet names to EntityType names
    const entitySetMap = {};
    const entitySetToType = {};
    entitySets.forEach(set => {
      const name = set.getAttribute('Name');
      const entityType = set.getAttribute('EntityType');
      if (name && entityType) {
        // EntityType format: "Microsoft.Dynamics.DataEntities.CustomerV3"
        const typeName = entityType.split('.').pop();
        entitySetMap[typeName] = name;
        entitySetToType[name] = typeName;
      }
    });

    entityTypes.forEach(entityType => {
      const name = entityType.getAttribute('Name');
      if (!name) return;

      // Get the EntitySet name (the actual URL name)
      const entitySetName = entitySetMap[name] || name;

      // Get properties
      const properties = [];
      const propertyElements = entityType.querySelectorAll('Property');
      propertyElements.forEach(prop => {
        properties.push({
          name: prop.getAttribute('Name'),
          type: prop.getAttribute('Type'),
          nullable: prop.getAttribute('Nullable') !== 'false'
        });
      });

      // Get key properties
      const keys = [];
      const keyElements = entityType.querySelectorAll('Key PropertyRef');
      keyElements.forEach(key => {
        keys.push(key.getAttribute('Name'));
      });

      // Get navigation properties (relationships to other entities)
      const navigationProperties = [];
      const navPropElements = entityType.querySelectorAll('NavigationProperty');
      navPropElements.forEach(navProp => {
        const navName = navProp.getAttribute('Name');
        const navType = navProp.getAttribute('Type');
        if (navName && navType) {
          // Type can be: "Microsoft.Dynamics.DataEntities.SalesOrderLine" or "Collection(Microsoft.Dynamics.DataEntities.SalesOrderLine)"
          const isCollection = navType.startsWith('Collection(');
          const relatedTypeName = navType.replace('Collection(', '').replace(')', '').split('.').pop();
          const relatedEntitySet = entitySetMap[relatedTypeName];

          navigationProperties.push({
            name: navName,
            relatedEntity: relatedEntitySet || relatedTypeName,
            relatedType: relatedTypeName,
            isCollection: isCollection
          });
        }
      });

      entities.push({
        name: entitySetName,
        typeName: name,
        label: this.formatEntityLabel(name),
        properties: properties,
        keys: keys,
        navigationProperties: navigationProperties,
        category: this.categorizeEntity(name)
      });
    });

    // Sort alphabetically
    entities.sort((a, b) => a.name.localeCompare(b.name));

    return entities;
  }

  /**
   * Format entity name into a readable label
   */
  formatEntityLabel(name) {
    // Remove common suffixes
    let label = name
      .replace(/V\d+$/, '')  // Remove version suffix (V3, V2, etc.)
      .replace(/Entity$/, '')
      .replace(/Entities$/, '');

    // Add spaces before capital letters
    label = label.replace(/([A-Z])/g, ' $1').trim();

    return label;
  }

  /**
   * Categorize entity by name pattern
   */
  categorizeEntity(name) {
    const nameLower = name.toLowerCase();

    if (nameLower.includes('customer') || nameLower.includes('sales') ||
        nameLower.includes('quote') || nameLower.includes('order')) {
      return 'Sales';
    }
    if (nameLower.includes('vendor') || nameLower.includes('purchase') ||
        nameLower.includes('supplier')) {
      return 'Purchasing';
    }
    if (nameLower.includes('item') || nameLower.includes('product') ||
        nameLower.includes('inventory') || nameLower.includes('warehouse')) {
      return 'Inventory';
    }
    if (nameLower.includes('ledger') || nameLower.includes('journal') ||
        nameLower.includes('account') || nameLower.includes('fiscal') ||
        nameLower.includes('tax') || nameLower.includes('currency')) {
      return 'Finance';
    }
    if (nameLower.includes('worker') || nameLower.includes('employee') ||
        nameLower.includes('position') || nameLower.includes('payroll')) {
      return 'HR';
    }
    if (nameLower.includes('project') || nameLower.includes('activity')) {
      return 'Project';
    }
    if (nameLower.includes('production') || nameLower.includes('bom') ||
        nameLower.includes('route')) {
      return 'Manufacturing';
    }

    return 'Other';
  }

  /**
   * Get or fetch metadata with caching
   */
  async getMetadata() {
    if (this.metadataCache) {
      return this.metadataCache;
    }

    // Try to load from storage
    const stored = await StorageManager.get('entityMetadata');
    if (stored && stored.baseUrl === this.baseUrl) {
      const age = Date.now() - stored.timestamp;
      // Cache for 1 hour
      if (age < 60 * 60 * 1000) {
        this.metadataCache = stored;
        return stored;
      }
    }

    return null;
  }

  /**
   * Cache metadata to storage
   */
  async cacheMetadata(entities) {
    const metadata = {
      baseUrl: this.baseUrl,
      timestamp: Date.now(),
      entities: entities
    };

    this.metadataCache = metadata;
    await StorageManager.set('entityMetadata', metadata);
  }

  /**
   * Query entity data with OData options
   */
  async queryEntity(entityName, options = {}) {
    const queryParams = [];

    // cross-company - REQUIRED for filtering by dataAreaId or querying across companies
    // D365 F&O specific: Must be set to query non-default company data
    if (options.crossCompany !== false) {
      // Default to true for flexibility - allows filtering by dataAreaId
      queryParams.push('cross-company=true');
    }

    // $select - which columns to return
    if (options.select && options.select.length > 0) {
      queryParams.push(`$select=${options.select.join(',')}`);
    }

    // $filter - where clause
    if (options.filter) {
      queryParams.push(`$filter=${encodeURIComponent(options.filter)}`);
    }

    // $orderby - sort
    if (options.orderBy) {
      const orderStr = options.orderBy.map(o =>
        `${o.field} ${o.direction || 'asc'}`
      ).join(',');
      queryParams.push(`$orderby=${orderStr}`);
    }

    // $top - limit
    if (options.top) {
      queryParams.push(`$top=${options.top}`);
    }

    // $skip - offset
    if (options.skip) {
      queryParams.push(`$skip=${options.skip}`);
    }

    // $count - include total count
    if (options.count) {
      queryParams.push('$count=true');
    }

    // $expand - related entities
    if (options.expand && options.expand.length > 0) {
      queryParams.push(`$expand=${options.expand.join(',')}`);
    }

    const queryString = queryParams.length > 0 ? `?${queryParams.join('&')}` : '';
    const endpoint = `${entityName}${queryString}`;

    const startTime = performance.now();
    const result = await this.fetch(endpoint);
    const queryTime = Math.round(performance.now() - startTime);

    return {
      data: result.value || [],
      count: result['@odata.count'],
      nextLink: result['@odata.nextLink'],
      queryTime: queryTime
    };
  }

  /**
   * Get entity schema (field metadata)
   */
  async getEntitySchema(entityName) {
    // First try cached metadata for type info
    let cachedEntity = null;
    const metadata = await this.getMetadata();
    if (metadata && metadata.entities) {
      cachedEntity = metadata.entities.find(e => e.name === entityName);
    }

    // Validate against a real query to get actual property names
    // (metadata cache can have properties that D365 doesn't accept in $select)
    try {
      const result = await this.queryEntity(entityName, { top: 1 });
      if (result.data && result.data.length > 0) {
        const record = result.data[0];
        const realKeys = new Set(Object.keys(record).filter(key => !key.startsWith('@')));

        if (cachedEntity) {
          // Use cached metadata but filter to only properties that actually exist
          const validProperties = cachedEntity.properties.filter(p => realKeys.has(p.name));
          // Add any real properties missing from cache
          for (const key of realKeys) {
            if (!validProperties.find(p => p.name === key)) {
              validProperties.push({
                name: key,
                type: this.inferType(record[key]),
                nullable: true
              });
            }
          }
          return {
            ...cachedEntity,
            properties: validProperties
          };
        }

        // No cache — build schema from real data
        const properties = [...realKeys].map(key => ({
          name: key,
          type: this.inferType(record[key]),
          nullable: true
        }));

        return {
          name: entityName,
          properties: properties,
          keys: [],
          navigationProperties: []
        };
      }
    } catch (error) {
      console.error('Error probing entity schema:', error);
    }

    // Fallback to cached metadata only if probe failed (e.g. empty entity)
    if (cachedEntity) {
      return cachedEntity;
    }

    return null;
  }

  /**
   * Infer the OData type from a JavaScript value
   */
  inferType(value) {
    if (value === null || value === undefined) return 'Edm.String';
    if (typeof value === 'number') {
      return Number.isInteger(value) ? 'Edm.Int64' : 'Edm.Decimal';
    }
    if (typeof value === 'boolean') return 'Edm.Boolean';
    if (typeof value === 'string') {
      // Check for date patterns
      if (/^\d{4}-\d{2}-\d{2}/.test(value)) {
        return value.includes('T') ? 'Edm.DateTimeOffset' : 'Edm.Date';
      }
      // Check for GUID pattern
      if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value)) {
        return 'Edm.Guid';
      }
      return 'Edm.String';
    }
    return 'Edm.String';
  }
}

// Global instance
const odataClient = new ODataClient();

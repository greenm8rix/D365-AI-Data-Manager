/**
 * D365 AI Data Manager - Main Browser
 * Core data browsing functionality
 */

// ==================== STATE ====================
let currentEntity = null;
let entitySchema = null;
let data = [];
let totalCount = 0;
let selectedRows = new Set();
let visibleColumns = [];
let sortConfig = { field: null, direction: 'asc' };
let filterConfig = [];
let expandConfig = []; // Related entities to expand (JOIN)
let pageSize = 100;
let currentPage = 1;
let queryTime = 0;
let viewMode = 'grid'; // 'grid' or 'cards'
let filteredByJoin = 0; // Track how many rows were filtered by inner join
let highlightConfigs = []; // Stored highlight configs that survive re-renders

// AI batch control — when true, data-changing functions skip loadData() calls
// and safeExecute() calls loadData() once at the end
window._aiDeferLoadData = false;
window._aiLoadDataNeeded = false;

// ==================== DOM ELEMENTS ====================
const entityInput = document.getElementById('entityInput');
const entityDropdown = document.getElementById('entityDropdown');
const envBadge = document.getElementById('envBadge');
const recordCount = document.getElementById('recordCount');
const quickFilter = document.getElementById('quickFilter');
const clearFilterBtn = document.getElementById('clearFilterBtn');
const gridHeader = document.getElementById('gridHeader');
const gridBody = document.getElementById('gridBody');
const loadingOverlay = document.getElementById('loadingOverlay');
const emptyState = document.getElementById('emptyState');
const selectionInfo = document.getElementById('selectionInfo');
const pageInfo = document.getElementById('pageInfo');
const queryTimeEl = document.getElementById('queryTime');
const pageInput = document.getElementById('pageInput');
const totalPagesEl = document.getElementById('totalPages');
const prevPageBtn = document.getElementById('prevPageBtn');
const nextPageBtn = document.getElementById('nextPageBtn');
const activeFiltersDiv = document.getElementById('activeFilters');
const filtersList = document.getElementById('filtersList');

// ==================== ENV BADGE ====================

async function getEnvLabels() {
  const result = await chrome.storage.local.get(['envLabels']);
  return result.envLabels || {};
}

async function saveEnvLabel(hostname, label) {
  const labels = await getEnvLabels();
  const updated = { ...labels, [hostname]: label };
  await chrome.storage.local.set({ envLabels: updated });
}

async function clearEnvLabel(hostname) {
  const labels = await getEnvLabels();
  const updated = { ...labels };
  delete updated[hostname];
  await chrome.storage.local.set({ envLabels: updated });
}

async function applyEnvBadge(env) {
  const labels = await getEnvLabels();
  const customLabel = labels[env.hostname];
  const label = customLabel || env.envType;
  envBadge.textContent = label;
  envBadge.className = `env-badge ${(customLabel || env.envType).toLowerCase()}`;
}

function startBadgeEdit() {
  const env = odataClient.environment;
  if (!env) return;
  const current = envBadge.textContent;
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'env-badge-edit';
  input.value = current;
  input.maxLength = 12;
  envBadge.textContent = '';
  envBadge.appendChild(input);
  input.focus();
  input.select();

  const finish = async () => {
    const val = input.value.trim();
    if (input.parentNode === envBadge) envBadge.removeChild(input);

    if (val && val !== env.envType) {
      await saveEnvLabel(env.hostname, val);
      envBadge.textContent = val;
      envBadge.className = `env-badge ${val.toLowerCase()}`;
    } else {
      await clearEnvLabel(env.hostname);
      envBadge.textContent = env.envType;
      envBadge.className = `env-badge ${env.envType.toLowerCase()}`;
    }
  };

  input.addEventListener('blur', finish);
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') input.blur();
    if (e.key === 'Escape') { input.value = ''; input.blur(); }
  });
}

// ==================== INITIALIZATION ====================
document.addEventListener('DOMContentLoaded', async () => {
  await init();
  setupEventListeners();
  setupKeyboardShortcuts();
});

async function init() {
  try {
    // Load saved settings first
    await loadSavedSettings();

    // Detect environment
    const env = await odataClient.detectEnvironment();
    if (env) {
      await applyEnvBadge(env);
    }

    // Check for entity in URL params
    const params = new URLSearchParams(window.location.search);
    const entityName = params.get('entity');

    if (entityName) {
      await loadEntity(entityName);
    }

    // Load entity list for dropdown
    loadEntityDropdown();

    // Initialize Power Tools icons
    if (typeof PowerToolsUI !== 'undefined') PowerToolsUI.initIcons();

    // Initialize AI settings
    if (typeof AISettings !== 'undefined') await AISettings.init();

    // What's New banner
    const currentVersion = chrome.runtime.getManifest().version;
    const lastVersion = await StorageManager.get('lastSeenVersion');
    if (lastVersion !== currentVersion) {
      showToast('New in v2.0: Power Platform exports, AI Assistant & Analysis!', 'info', 8000);
      await StorageManager.set('lastSeenVersion', currentVersion);
    }

  } catch (error) {
    console.error('Init error:', error);
    showError(error.message);
  }
}

async function loadEntityDropdown() {
  try {
    const entities = await odataClient.getEntities();
    if (!entities) return;

    // Cache for search
    window.allEntities = entities;

  } catch (error) {
    console.error('Error loading entity list:', error);
  }
}

// ==================== ENTITY NAME RESOLUTION ====================
// Uses the same search logic as the entity search bar
function resolveEntityName(entityName) {
  if (!window.allEntities || window.allEntities.length === 0) return entityName;

  // Exact match
  if (window.allEntities.find(e => e.name === entityName)) return entityName;

  // Same search as the entity dropdown: match by name or label (case-insensitive)
  const q = entityName.toLowerCase();
  const matches = window.allEntities.filter(e =>
    e.name.toLowerCase().includes(q) ||
    (e.label && e.label.toLowerCase().includes(q))
  );

  // No match found — return original name (loadEntity will validate and throw)
  if (matches.length === 0) return entityName;

  // Exact case-insensitive name match
  const ciExact = matches.find(e => e.name.toLowerCase() === q);
  if (ciExact) return ciExact.name;

  // Prefer V2/V3 data entities, then shortest name
  matches.sort((a, b) => {
    const aV = /V\d+$/.test(a.name) ? 0 : 1;
    const bV = /V\d+$/.test(b.name) ? 0 : 1;
    if (aV !== bV) return aV - bV;
    return a.name.length - b.name.length;
  });

  console.log(`Auto-corrected entity: "${entityName}" → "${matches[0].name}"`);
  return matches[0].name;
}

// ==================== ENTITY LOADING ====================
async function loadEntity(entityName) {
  entityName = resolveEntityName(entityName);

  // Validate entity exists BEFORE resetting any state
  if (window.allEntities && window.allEntities.length > 0) {
    const exists = window.allEntities.find(e => e.name === entityName);
    if (!exists) {
      const msg = `Entity "${entityName}" does not exist. Use searchEntities() to find the correct name.`;
      showError(msg);
      throw new Error(msg);
    }
  }

  currentEntity = entityName;
  entityInput.value = entityName;
  document.title = `${entityName} - D365 AI Data Manager`;

  // Add to recent
  await StorageManager.addRecent(entityName);

  // Reset state
  selectedRows.clear();
  filterConfig = [];
  expandConfig = [];
  activeJoin = null;
  filteredByJoin = 0;
  visibleColumns = [];
  data = [];
  totalCount = 0;
  highlightConfigs = [];
  sortConfig = { field: null, direction: 'asc' };
  currentPage = 1;
  updateActiveFilters();

  // Reset join button
  const joinBtn = document.getElementById('manualJoinBtn');
  if (joinBtn) joinBtn.innerHTML = '&#128200; Join';

  // Get schema
  showLoading();
  try {
    entitySchema = await odataClient.getEntitySchema(entityName);

    if (!entitySchema || !entitySchema.properties) {
      throw new Error(`Could not load schema for "${entityName}". Entity may not exist or is not an OData entity.`);
    }

    // Safe OData types that D365 can serialize without issues
    const safeODataTypes = ['Edm.String', 'Edm.Int16', 'Edm.Int32', 'Edm.Int64', 'Edm.Decimal', 'Edm.Double',
                            'Edm.Single', 'Edm.Boolean', 'Edm.Byte', 'Edm.Date', 'Edm.DateTime',
                            'Edm.DateTimeOffset', 'Edm.Guid', 'Edm.Binary', 'Edm.Time', 'Edm.Duration'];

    // Default visible columns (first 15, excluding internal fields and problematic types)
    visibleColumns = entitySchema.properties
      .filter(p => {
        if (p.name.startsWith('@') || p.name.startsWith('_')) return false;
        if (!p.type) return true;
        if (safeODataTypes.includes(p.type)) return true;
        if (p.type.startsWith('Microsoft.Dynamics.DataEntities.')) return true;
        console.warn(`Excluding column "${p.name}" from default view due to problematic type: ${p.type}`);
        return false;
      })
      .slice(0, 15)
      .map(p => p.name);

    // Load data
    await loadData();

  } catch (error) {
    hideLoading();
    console.error('Error loading entity:', error);
    showError(`Failed to load ${entityName}: ${error.message}`);
    throw error; // Re-throw so safeExecute can catch and break
  }
}

async function loadData() {
  showLoading();
  window.aiLastError = null;

  try {
    // Separate filters for base entity vs joined entity columns
    const { baseFilters, joinedFilters } = separateFiltersByEntity();

    // If there are filters on joined columns, we need to pre-filter by querying the joined entity first
    let joinedKeyFilter = '';
    const hasJoinedFilters = Object.keys(joinedFilters).length > 0;

    if (hasJoinedFilters) {
      if (!activeJoin) {
        // User has joined column filters but no active join - warn and remove invalid filters
        console.warn('Joined column filters exist but no active join - removing invalid filters');
        showToast('Filters on joined columns require an active join. Clearing invalid filters.', 'warning');
        filterConfig = filterConfig.filter(f => !f.field || !f.field.includes('.'));
        updateActiveFilters();
      } else {
        joinedKeyFilter = await buildJoinedColumnFilter(joinedFilters);
      }
    }

    // Build filter string (only for base entity columns)
    const filterString = buildFilterString(baseFilters, joinedKeyFilter);

    // Build orderby - only use if it's a column from current entity (not joined)
    const orderBy = sortConfig.field && !sortConfig.field.includes('.')
      ? [{ field: sortConfig.field, direction: sortConfig.direction }]
      : [];

    // Safe OData types that D365 can serialize
    const safeODataTypes = ['Edm.String', 'Edm.Int16', 'Edm.Int32', 'Edm.Int64', 'Edm.Decimal', 'Edm.Double',
                            'Edm.Single', 'Edm.Boolean', 'Edm.Byte', 'Edm.Date', 'Edm.DateTime',
                            'Edm.DateTimeOffset', 'Edm.Guid', 'Edm.Binary', 'Edm.Time', 'Edm.Duration'];

    const isSafeType = (type) => {
      if (!type) return true; // Unknown type - assume safe
      if (safeODataTypes.includes(type)) return true;
      if (type.startsWith('Microsoft.Dynamics.DataEntities.')) return true; // D365 enum type
      return false;
    };

    // Filter visible columns to only safe types
    let odataSelectColumns = visibleColumns.filter(col => {
      if (col.includes('.')) return false; // Exclude joined columns
      const colType = getFieldType(col);
      if (isSafeType(colType)) return true;
      console.warn(`Excluding column "${col}" from $select due to problematic type: ${colType}`);
      return false;
    });

    // CRITICAL: If no columns selected or all filtered out, build safe list from schema
    // Never pass undefined for $select - D365 will return ALL columns including problematic ones
    if (odataSelectColumns.length === 0 && entitySchema && entitySchema.properties) {
      console.warn('No safe columns in visibleColumns, building safe list from schema');
      odataSelectColumns = entitySchema.properties
        .filter(p => isSafeType(p.type))
        .slice(0, 20) // Limit to first 20 safe columns
        .map(p => p.name);
      console.log('Safe columns from schema:', odataSelectColumns);
    }

    // Log the query for debugging
    console.log('Query params:', {
      entity: currentEntity,
      select: odataSelectColumns,
      selectCount: odataSelectColumns.length,
      visibleColumnsCount: visibleColumns.length,
      schemaPropertiesCount: entitySchema?.properties?.length || 0,
      filter: filterString,
      orderBy: orderBy,
      top: pageSize,
      skip: (currentPage - 1) * pageSize
    });

    // ALWAYS specify $select to avoid D365 returning problematic columns
    // Only skip $select if we have no schema at all
    const useSelect = odataSelectColumns.length > 0 ? odataSelectColumns : undefined;
    if (!useSelect && entitySchema) {
      console.error('WARNING: No safe columns found! Query may fail.');
    }

    const result = await odataClient.queryEntity(currentEntity, {
      select: useSelect,
      filter: filterString || undefined,
      orderBy: orderBy.length > 0 ? orderBy : undefined,
      expand: expandConfig.length > 0 ? expandConfig : undefined,
      top: pageSize,
      skip: (currentPage - 1) * pageSize,
      count: true
    });

    data = result.data;
    totalCount = result.count || data.length;
    queryTime = result.queryTime;
    filteredByJoin = 0; // Reset join filter count

    // Process expanded data (flatten nested navigation properties)
    if (expandConfig.length > 0) {
      processExpandedData();
    }

    // Re-apply active join if one exists
    if (activeJoin) {
      await reapplyJoin();
    }

    // Apply client-side quick filter for joined columns (server can't filter these)
    applyClientSideQuickFilter();

    renderData();
    updatePagination();
    hideLoading();

  } catch (error) {
    console.error('Error loading data:', error);
    console.error('Filter config:', filterConfig);
    console.error('Generated filter:', buildFilterString());
    hideLoading();

    // Expose a short, actionable error to AI so it can self-correct
    const rawErr = error.message || 'Unknown query error';
    const httpMatch = rawErr.match(/^HTTP (\d+)/);
    const filterStr = buildFilterString();
    if (httpMatch) {
      window.aiLastError = `Query failed (HTTP ${httpMatch[1]}). Filter was: ${filterStr || 'none'}. Clear filters or try a different approach.`;
    } else {
      window.aiLastError = `Query failed: ${rawErr.substring(0, 120)}. Clear filters or try a different approach.`;
    }

    // Show debug panel automatically on error
    document.getElementById('debugPanel').classList.remove('hidden');
    updateDebugInfo();

    // Check if error is related to $expand (serialization issues)
    const errorMsg = error.message || '';
    if (expandConfig.length > 0 && (
        errorMsg.includes('Dynamics.AX.Application') ||
        errorMsg.includes('serialization') ||
        errorMsg.includes('given model does not contain')
    )) {
      showError(`Query failed due to $expand. The expanded entity has fields with unsupported types.\n\n` +
                `Try: Click "Related" → "Clear All" → Use "Join" button instead for reliable cross-entity queries.`);

      // Offer to auto-clear expand
      if (confirm('The $expand failed due to D365 serialization issues. Clear the expand and retry?')) {
        expandConfig = [];
        document.getElementById('relatedEntitiesBtn').innerHTML = '&#128279; Related';
        loadData();
        return;
      }
    } else {
      showError(`Query failed: ${error.message}`);
    }
  }
}

// ==================== EXPANDED DATA PROCESSING ====================
/**
 * Process OData $expand results - flatten nested navigation properties into columns
 * When OData returns expanded data, it comes as nested objects like:
 * { CustomerAccount: "C001", CustomerGroups: { GroupId: "10", Name: "Wholesale" } }
 * We flatten these to: { CustomerAccount: "C001", "CustomerGroups.GroupId": "10", "CustomerGroups.Name": "Wholesale" }
 */
function processExpandedData() {
  if (!data || data.length === 0 || expandConfig.length === 0) return;

  const expandedColumns = new Set();

  // Process each row
  data = data.map(row => {
    const flattenedRow = { ...row };

    // For each expanded navigation property
    expandConfig.forEach(navPropName => {
      const expandedData = row[navPropName];

      if (expandedData) {
        if (Array.isArray(expandedData)) {
          // Collection navigation property (one-to-many)
          // For collections, we'll show a count and allow viewing in side panel
          flattenedRow[`${navPropName}.__count`] = expandedData.length;
          flattenedRow[`${navPropName}.__data`] = expandedData;
          expandedColumns.add(`${navPropName}.__count`);

          // Also flatten first item's fields for preview
          if (expandedData.length > 0) {
            Object.keys(expandedData[0]).forEach(key => {
              if (!key.startsWith('@')) {
                const colName = `${navPropName}.${key}`;
                // Show first item's value with indicator if there are more
                flattenedRow[colName] = expandedData[0][key];
                expandedColumns.add(colName);
              }
            });
          }
        } else if (typeof expandedData === 'object') {
          // Single navigation property (one-to-one / many-to-one)
          Object.keys(expandedData).forEach(key => {
            if (!key.startsWith('@')) {
              const colName = `${navPropName}.${key}`;
              flattenedRow[colName] = expandedData[key];
              expandedColumns.add(colName);
            }
          });
        }
      }

      // Remove the original nested object (keep data clean for display)
      delete flattenedRow[navPropName];
    });

    return flattenedRow;
  });

  // Add expanded columns to visibleColumns if not already there
  expandedColumns.forEach(col => {
    if (!visibleColumns.includes(col)) {
      visibleColumns.push(col);
    }
  });

  console.log('Processed expanded data. New columns:', Array.from(expandedColumns));
}

// ==================== GRID RENDERING ====================
function renderData() {
  if (viewMode === 'cards') {
    renderCards();
  } else {
    renderGrid();
  }
  // Re-apply any active highlights after DOM rebuild
  reapplyHighlights();
}

function renderGrid() {
  // Show grid, hide cards container
  document.getElementById('dataGrid').style.display = 'table';
  const cardsContainer = document.getElementById('cardsContainer');
  if (cardsContainer) cardsContainer.style.display = 'none';

  if (!data || data.length === 0) {
    gridHeader.innerHTML = '';
    gridBody.innerHTML = '';
    emptyState.classList.remove('hidden');
    return;
  }

  emptyState.classList.add('hidden');

  // Get columns from first row if schema not available
  const columns = visibleColumns.length > 0
    ? visibleColumns
    : Object.keys(data[0]).filter(k => !k.startsWith('@'));

  // Render header - group by entity
  let headerHtml = '<tr>';
  headerHtml += '<th class="checkbox-col"><input type="checkbox" id="selectAllCheckbox"></th>';

  columns.forEach(col => {
    const isJoinedCol = col.includes('.');
    const sortClass = sortConfig.field === col
      ? (sortConfig.direction === 'asc' ? 'sorted-asc' : 'sorted-desc')
      : '';
    const joinedClass = isJoinedCol ? 'joined-column-header' : '';

    // For joined columns, show just the field name but with entity indicator
    let displayName;
    if (isJoinedCol) {
      const [entityName, fieldName] = col.split('.');
      displayName = `<span class="joined-entity-tag">${escapeHtml(entityName)}</span>${escapeHtml(formatColumnName(fieldName))}`;
    } else {
      displayName = escapeHtml(formatColumnName(col));
    }

    headerHtml += `<th class="${sortClass} ${joinedClass}" data-column="${escapeHtml(col)}">${displayName}</th>`;
  });

  headerHtml += '</tr>';
  gridHeader.innerHTML = headerHtml;

  // Render body
  let bodyHtml = '';
  data.forEach((row, index) => {
    const isSelected = selectedRows.has(index);
    bodyHtml += `<tr data-index="${index}" class="${isSelected ? 'selected' : ''}">`;
    bodyHtml += `<td class="checkbox-cell"><input type="checkbox" ${isSelected ? 'checked' : ''}></td>`;

    columns.forEach(col => {
      const value = row[col];
      const cellClass = getCellClass(value, col);
      const displayValue = formatCellValue(value, col);
      bodyHtml += `<td class="${cellClass}" title="${escapeHtml(String(value ?? ''))}">${displayValue}</td>`;
    });

    bodyHtml += '</tr>';
  });

  gridBody.innerHTML = bodyHtml;

  // Update info
  updateSelectionInfo();
  updateRecordCount();
  queryTimeEl.textContent = `Query: ${queryTime}ms`;

  // Add event listeners
  setupGridEventListeners();
}

function renderCards() {
  // Hide grid, show cards
  document.getElementById('dataGrid').style.display = 'none';

  // Create or get cards container
  let cardsContainer = document.getElementById('cardsContainer');
  if (!cardsContainer) {
    cardsContainer = document.createElement('div');
    cardsContainer.id = 'cardsContainer';
    cardsContainer.className = 'cards-container';
    document.getElementById('gridContainer').appendChild(cardsContainer);
  }
  cardsContainer.style.display = 'grid';

  if (!data || data.length === 0) {
    cardsContainer.innerHTML = '';
    emptyState.classList.remove('hidden');
    return;
  }

  emptyState.classList.add('hidden');

  // Get columns
  const columns = visibleColumns.length > 0
    ? visibleColumns
    : Object.keys(data[0]).filter(k => !k.startsWith('@'));

  // Find key columns (first few important ones)
  const keyColumns = columns.slice(0, 3);
  const otherColumns = columns.slice(3);

  // Separate columns by entity for cards
  const currentEntityOtherCols = otherColumns.filter(col => !col.includes('.'));
  const joinedOtherCols = otherColumns.filter(col => col.includes('.'));

  // Group joined columns by entity
  const joinedByEntity = {};
  joinedOtherCols.forEach(col => {
    const [entityName] = col.split('.');
    if (!joinedByEntity[entityName]) {
      joinedByEntity[entityName] = [];
    }
    joinedByEntity[entityName].push(col);
  });

  // Render cards
  cardsContainer.innerHTML = data.map((row, index) => {
    const isSelected = selectedRows.has(index);
    const keyValues = keyColumns.map(col => `
      <div class="card-key-field">
        <span class="card-key-value">${formatCellValue(row[col], col)}</span>
      </div>
    `).join('');

    // Current entity fields
    const currentEntityValues = currentEntityOtherCols.map(col => `
      <div class="card-field">
        <span class="card-label">${escapeHtml(formatColumnName(col))}</span>
        <span class="card-value">${formatCellValue(row[col], col)}</span>
      </div>
    `).join('');

    // Joined entity fields - grouped by entity
    let joinedSections = '';
    Object.keys(joinedByEntity).forEach(entityName => {
      const joinedFields = joinedByEntity[entityName].map(col => {
        const fieldName = col.split('.')[1];
        return `
          <div class="card-field joined">
            <span class="card-label">${escapeHtml(formatColumnName(fieldName))}</span>
            <span class="card-value">${formatCellValue(row[col], col)}</span>
          </div>
        `;
      }).join('');

      joinedSections += `
        <div class="card-joined-section">
          <div class="card-joined-header">${escapeHtml(entityName)}</div>
          ${joinedFields}
        </div>
      `;
    });

    return `
      <div class="data-card ${isSelected ? 'selected' : ''}" data-index="${index}">
        <div class="card-header">
          <input type="checkbox" ${isSelected ? 'checked' : ''}>
          <div class="card-title">${keyValues}</div>
        </div>
        <div class="card-body">
          ${currentEntityValues}
          ${joinedSections}
        </div>
      </div>
    `;
  }).join('');

  // Update info
  updateSelectionInfo();
  updateRecordCount();
  queryTimeEl.textContent = `Query: ${queryTime}ms`;

  // Add event listeners for cards
  setupCardEventListeners();
}

function setupCardEventListeners() {
  const cardsContainer = document.getElementById('cardsContainer');
  if (!cardsContainer) return;

  cardsContainer.querySelectorAll('.data-card').forEach(card => {
    card.addEventListener('click', (e) => {
      if (e.target.type !== 'checkbox') {
        toggleRowSelection(parseInt(card.dataset.index));
        renderCards();
      }
    });

    card.querySelector('input[type="checkbox"]').addEventListener('change', () => {
      toggleRowSelection(parseInt(card.dataset.index));
      renderCards();
    });
  });
}

function setViewMode(mode) {
  viewMode = mode;

  // Update button states
  document.querySelectorAll('.btn-view').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.view === mode);
  });

  // Re-render
  renderData();
}

function formatColumnName(name) {
  // Add spaces before capitals and clean up
  return name
    .replace(/([A-Z])/g, ' $1')
    .replace(/^./, str => str.toUpperCase())
    .trim();
}

function formatCellValue(value, column) {
  if (value === null || value === undefined) {
    return '<span class="null">null</span>';
  }

  if (typeof value === 'boolean') {
    return value
      ? '<span class="bool-true">✓ Yes</span>'
      : '<span class="bool-false">✗ No</span>';
  }

  if (typeof value === 'number') {
    // Format with locale, but also add data attribute for sorting
    return `<span data-raw="${escapeHtml(String(value))}">${value.toLocaleString()}</span>`;
  }

  // Check for date strings
  if (typeof value === 'string') {
    if (/^\d{4}-\d{2}-\d{2}T/.test(value)) {
      const date = new Date(value);
      return `<span class="date-display" data-raw="${escapeHtml(value)}">${date.toLocaleDateString()} ${date.toLocaleTimeString()}</span>`;
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
      return `<span class="date-display" data-raw="${escapeHtml(value)}">${new Date(value + 'T00:00:00').toLocaleDateString()}</span>`;
    }
    // Check for GUIDs - show shortened version
    if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(value)) {
      return `<span class="guid-value" title="${escapeHtml(value)}">${escapeHtml(value.substring(0, 8))}...</span>`;
    }
    // Empty string
    if (value === '') {
      return '<span class="empty-string">(empty)</span>';
    }
  }

  return escapeHtml(String(value));
}

function getCellClass(value, column) {
  if (value === null || value === undefined) {
    return 'null-value';
  }
  if (typeof value === 'number') {
    return 'number-value';
  }
  if (typeof value === 'boolean') {
    return 'bool-value';
  }
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}/.test(value)) {
    return 'date-value';
  }
  return '';
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// ==================== FILTERING ====================
// D365 F&O uses wildcard asterisks IN THE VALUE, not standard OData functions!
// Instead of contains(field,'value'), use: field eq '*value*'
// Instead of startswith(field,'value'), use: field eq 'value*'

/**
 * Normalize entity name for consistent comparison.
 * D365 entity names should be treated case-insensitively.
 */
function normalizeEntityName(entityName) {
  if (!entityName) return '';
  return entityName.toLowerCase();
}

/**
 * Separate filter config into base entity filters and joined entity filters.
 * Joined entity filters are grouped by entity name (normalized for case-insensitive matching).
 */
function separateFiltersByEntity() {
  const baseFilters = [];
  const joinedFilters = {}; // { normalizedEntityName: [filters] }

  filterConfig.forEach(filter => {
    if (!filter.field) return;

    if (filter.field.includes('.')) {
      // This is a joined column filter (e.g., "DocumentType.FilePlace")
      const [entityName, fieldName] = filter.field.split('.');
      const normalizedEntity = normalizeEntityName(entityName);
      if (!joinedFilters[normalizedEntity]) {
        joinedFilters[normalizedEntity] = [];
      }
      joinedFilters[normalizedEntity].push({
        ...filter,
        originalField: filter.field,
        field: fieldName, // Use just the field name for the joined entity query
        sourceEntity: entityName // Keep original casing for display
      });
    } else {
      // Base entity filter
      baseFilters.push(filter);
    }
  });

  return { baseFilters, joinedFilters };
}

/**
 * For filters on joined columns, query the joined entity to find matching
 * join key values, then return a filter condition for the main entity.
 * This implements a "semi-join" pattern.
 */
async function buildJoinedColumnFilter(joinedFilters) {
  if (!activeJoin) {
    console.log('buildJoinedColumnFilter: No active join');
    return '';
  }

  const targetEntity = activeJoin.targetEntity;
  const targetField = activeJoin.targetField;
  const currentField = activeJoin.currentField;
  const normalizedTarget = normalizeEntityName(targetEntity);

  console.log('buildJoinedColumnFilter:', {
    activeJoinEntity: targetEntity,
    normalizedTarget: normalizedTarget,
    joinedFilterEntities: Object.keys(joinedFilters),
    hasTargetSchema: !!activeJoin.targetSchema
  });

  // Check if any of our filters apply to the active join's target entity (case-insensitive)
  const filtersForJoinedEntity = joinedFilters[normalizedTarget];
  if (!filtersForJoinedEntity || filtersForJoinedEntity.length === 0) {
    console.log('No filters apply to the joined entity:', targetEntity, '(normalized:', normalizedTarget, ') - available:', Object.keys(joinedFilters));
    return '';
  }

  console.log(`Building semi-join filter for ${targetEntity}:`, filtersForJoinedEntity);

  try {
    // Build filter conditions for the joined entity using the joined entity's schema for type lookups
    const joinedEntityConditions = filtersForJoinedEntity.map(filter => {
      const filterWithFullName = { ...filter, field: filter.originalField || filter.field };
      return buildFilterConditionForJoinedEntity(filter.field, filter.operator, filter.value, activeJoin.targetSchema);
    }).filter(c => c);

    if (joinedEntityConditions.length === 0) return '';

    const joinedFilter = joinedEntityConditions.join(' and ');
    console.log(`Querying ${targetEntity} with filter: ${joinedFilter}`);

    // Query the joined entity to get matching records
    const result = await odataClient.queryEntity(targetEntity, {
      select: [targetField],
      filter: joinedFilter,
      top: 1000 // Limit to prevent too many values
    });

    if (!result.data || result.data.length === 0) {
      console.log('No matching records in joined entity - will return no results');
      return `${currentField} eq '__NO_MATCH__'`;
    }

    // Get unique join key values
    const matchingValues = [...new Set(result.data.map(row => row[targetField]).filter(v => v != null))];
    console.log(`Found ${matchingValues.length} matching values in ${targetEntity}.${targetField}`);

    if (matchingValues.length === 0) {
      return `${currentField} eq '__NO_MATCH__'`;
    }

    // Build IN-style filter: field eq 'val1' or field eq 'val2' ...
    const fieldType = getFieldType(currentField);
    const isNumeric = ['Edm.Int16', 'Edm.Int32', 'Edm.Int64', 'Edm.Decimal', 'Edm.Double'].includes(fieldType);

    const valueConditions = matchingValues.slice(0, 100).map(v => { // Limit to 100 values
      if (isNumeric) {
        return `${currentField} eq ${v}`;
      }
      return `${currentField} eq '${escapeODataString(String(v))}'`;
    });

    const inFilter = `(${valueConditions.join(' or ')})`;
    console.log(`Semi-join filter: ${inFilter.substring(0, 200)}${inFilter.length > 200 ? '...' : ''}`);

    return inFilter;

  } catch (error) {
    console.error('Error building semi-join filter:', error);
    return '';
  }
}

function buildFilterString(baseFilters = null, additionalFilter = '') {
  const conditions = [];
  const filtersToProcess = baseFilters !== null ? baseFilters : filterConfig;

  // Quick filter - search ALL visible string columns using D365 wildcard syntax
  // When there's an active join, skip server-side filtering - we'll do it client-side
  // so we can search joined columns too
  const quickFilterValue = quickFilter.value.trim();
  if (quickFilterValue && entitySchema && entitySchema.properties && !activeJoin) {
    // No join - use server-side filtering (faster for large datasets)
    const stringColumns = entitySchema.properties
      .filter(p => p.type === 'Edm.String' && visibleColumns.includes(p.name) && !p.name.includes('.'))
      .map(p => p.name);

    if (stringColumns.length > 0) {
      const searchValue = escapeODataString(quickFilterValue);

      // Determine the search pattern
      let pattern;
      if (quickFilterValue.includes('*')) {
        // User specified wildcards - use as-is
        pattern = searchValue;
      } else {
        // Default: contains behavior with wildcards on both ends
        pattern = `*${searchValue}*`;
      }

      // Build OR conditions for all string columns (limit to 10 to avoid too complex queries)
      const searchCols = stringColumns.slice(0, 10);
      console.log('Quick filter (server-side) searching columns:', searchCols);

      const searchConditions = searchCols.map(col => `${col} eq '${pattern}'`);
      conditions.push(`(${searchConditions.join(' or ')})`);
    } else {
      console.warn('Quick filter: No string columns found in visible columns');
    }
  } else if (quickFilterValue && activeJoin) {
    // With join - client-side filtering will handle this after join is applied
    console.log('Quick filter will be applied client-side to include joined columns');
  }

  // Advanced filters - build with proper AND/OR logic (only process base entity filters)
  let advancedFilterParts = [];
  filtersToProcess.forEach((filter, index) => {
    // Skip joined column filters - they're handled separately via semi-join
    if (filter.field && filter.field.includes('.')) return;

    const condition = buildFilterCondition(filter);
    if (condition) {
      const logic = filter.logic || (index > 0 ? 'and' : null);
      advancedFilterParts.push({ condition, logic });
    }
  });

  // Build the advanced filter string with proper grouping
  if (advancedFilterParts.length > 0) {
    let advancedFilter = advancedFilterParts[0].condition;
    for (let i = 1; i < advancedFilterParts.length; i++) {
      const part = advancedFilterParts[i];
      advancedFilter += ` ${part.logic} ${part.condition}`;
    }
    conditions.push(`(${advancedFilter})`);
  }

  // Add the semi-join filter for joined columns (if provided)
  if (additionalFilter) {
    conditions.push(additionalFilter);
  }

  return conditions.length > 0 ? conditions.join(' and ') : '';
}

function buildFilterCondition(filter) {
  const { field, operator, value } = filter;

  // Skip if no field selected
  if (!field) return null;

  // For null checks — D365 F&O doesn't support "ne null" syntax
  // Strings: use eq '' / ne '' (D365 treats empty strings and null the same)
  // Others: use eq null / not(field eq null)
  if (operator === 'null' || operator === 'notnull') {
    const fieldType = getFieldType(field);
    const isString = !fieldType || fieldType === 'Edm.String';
    if (operator === 'null') {
      return isString ? `${field} eq ''` : `${field} eq null`;
    } else {
      return isString ? `${field} ne ''` : `not(${field} eq null)`;
    }
  }

  // Skip if no value provided for other operators (but allow empty string for 'eq'/'ne' — filters for blank fields)
  if (value === undefined || value === null) {
    return null;
  }
  if (value === '' && operator !== 'eq' && operator !== 'ne') {
    return null;
  }

  // Get field type from schema to determine how to format value
  const fieldType = getFieldType(field);
  const numericTypes = ['Edm.Int16', 'Edm.Int32', 'Edm.Int64', 'Edm.Decimal', 'Edm.Double', 'Edm.Single', 'Edm.Byte'];
  const dateTypes = ['Edm.Date', 'Edm.DateTime', 'Edm.DateTimeOffset'];

  const isNumericField = numericTypes.includes(fieldType);
  const isBoolField = fieldType === 'Edm.Boolean';
  const isDateField = dateTypes.includes(fieldType);
  const isGuidField = fieldType === 'Edm.Guid';
  // Enum types don't start with "Edm." - they have custom type names like "Microsoft.Dynamics.DataEntities.DocuTypeGroup"
  const isEnumField = fieldType && !fieldType.startsWith('Edm.');

  console.log(`Filter field "${field}" type: ${fieldType}, isNumeric: ${isNumericField}, isEnum: ${isEnumField}`);

  // Format the value based on field type - BE STRICT about types
  let formattedValue;
  if (isBoolField) {
    formattedValue = (value.toLowerCase() === 'true' || value === '1') ? 'true' : 'false';
  } else if (isEnumField) {
    // Enum fields need the fully qualified type name: Microsoft.Dynamics.DataEntities.EnumType'Value'
    formattedValue = `${fieldType}'${escapeODataString(value)}'`;
  } else if (isNumericField) {
    // ONLY numeric if the schema explicitly says it's numeric
    formattedValue = value;
  } else if (isGuidField) {
    // GUIDs can be used with or without quotes in OData - use without quotes
    formattedValue = value;
  } else if (isDateField) {
    // Dates should not have quotes in OData v4
    formattedValue = value;
  } else {
    // Everything else (including Edm.String and unknown types) - needs quotes
    formattedValue = `'${escapeODataString(value)}'`;
  }

  // D365 F&O uses wildcard * in values instead of OData functions!
  switch (operator) {
    case 'eq':
      return `${field} eq ${formattedValue}`;
    case 'ne':
      return `${field} ne ${formattedValue}`;
    case 'contains':
      // D365: Use wildcards in value instead of contains() function
      // Note: Wildcards don't work for enum/numeric/date fields - fall back to equals
      if (isEnumField || isNumericField || isDateField || isBoolField) {
        return `${field} eq ${formattedValue}`;
      }
      return `${field} eq '*${escapeODataString(value)}*'`;
    case 'startswith':
      // D365: Use wildcard at end instead of startswith() function
      if (isEnumField || isNumericField || isDateField || isBoolField) {
        return `${field} eq ${formattedValue}`;
      }
      return `${field} eq '${escapeODataString(value)}*'`;
    case 'endswith':
      // D365: Use wildcard at start instead of endswith() function
      if (isEnumField || isNumericField || isDateField || isBoolField) {
        return `${field} eq ${formattedValue}`;
      }
      return `${field} eq '*${escapeODataString(value)}'`;
    case 'gt':
      return `${field} gt ${formattedValue}`;
    case 'ge':
      return `${field} ge ${formattedValue}`;
    case 'lt':
      return `${field} lt ${formattedValue}`;
    case 'le':
      return `${field} le ${formattedValue}`;
    default:
      return null;
  }
}

/**
 * Build a filter condition for a joined entity using its schema.
 * This is used when filtering on joined columns via semi-join.
 */
function buildFilterConditionForJoinedEntity(field, operator, value, targetSchema) {
  // Skip if no field selected
  if (!field) return null;

  // For null checks, don't need a value
  if (operator === 'null') {
    return `${field} eq null`;
  }
  if (operator === 'notnull') {
    return `${field} ne null`;
  }

  // Skip if no value provided for other operators
  if (value === undefined || value === null || value === '') {
    return null;
  }

  // Get field type from the target entity's schema
  const prop = targetSchema?.properties?.find(p => p.name === field);
  const fieldType = prop?.type || 'Edm.String';

  const numericTypes = ['Edm.Int16', 'Edm.Int32', 'Edm.Int64', 'Edm.Decimal', 'Edm.Double', 'Edm.Single', 'Edm.Byte'];
  const dateTypes = ['Edm.Date', 'Edm.DateTime', 'Edm.DateTimeOffset'];

  const isNumericField = numericTypes.includes(fieldType);
  const isBoolField = fieldType === 'Edm.Boolean';
  const isDateField = dateTypes.includes(fieldType);
  const isGuidField = fieldType === 'Edm.Guid';
  const isEnumField = fieldType && !fieldType.startsWith('Edm.');

  console.log(`Joined filter field "${field}" type: ${fieldType}, isNumeric: ${isNumericField}, isEnum: ${isEnumField}`);

  // Format the value based on field type
  let formattedValue;
  if (isBoolField) {
    formattedValue = (value.toLowerCase() === 'true' || value === '1') ? 'true' : 'false';
  } else if (isEnumField) {
    formattedValue = `${fieldType}'${escapeODataString(value)}'`;
  } else if (isNumericField) {
    formattedValue = value;
  } else if (isGuidField) {
    formattedValue = value;
  } else if (isDateField) {
    formattedValue = value;
  } else {
    formattedValue = `'${escapeODataString(value)}'`;
  }

  switch (operator) {
    case 'eq':
      return `${field} eq ${formattedValue}`;
    case 'ne':
      return `${field} ne ${formattedValue}`;
    case 'contains':
      if (isEnumField || isNumericField || isDateField || isBoolField) {
        return `${field} eq ${formattedValue}`;
      }
      return `${field} eq '*${escapeODataString(value)}*'`;
    case 'startswith':
      if (isEnumField || isNumericField || isDateField || isBoolField) {
        return `${field} eq ${formattedValue}`;
      }
      return `${field} eq '${escapeODataString(value)}*'`;
    case 'endswith':
      if (isEnumField || isNumericField || isDateField || isBoolField) {
        return `${field} eq ${formattedValue}`;
      }
      return `${field} eq '*${escapeODataString(value)}'`;
    case 'gt':
      return `${field} gt ${formattedValue}`;
    case 'ge':
      return `${field} ge ${formattedValue}`;
    case 'lt':
      return `${field} lt ${formattedValue}`;
    case 'le':
      return `${field} le ${formattedValue}`;
    default:
      return null;
  }
}

// Helper to get field type from schema
// Can optionally pass a specific schema (for joined entity lookups)
function getFieldType(fieldName, schema = null) {
  const targetSchema = schema || entitySchema;
  if (!targetSchema || !targetSchema.properties) return 'Edm.String';

  // Handle joined column names (EntityName.FieldName)
  if (fieldName.includes('.') && !schema) {
    const [entityName, actualFieldName] = fieldName.split('.');
    // Look up in active join's target schema if this matches (case-insensitive)
    if (activeJoin &&
        normalizeEntityName(activeJoin.targetEntity) === normalizeEntityName(entityName) &&
        activeJoin.targetSchema) {
      const prop = activeJoin.targetSchema.properties?.find(p => p.name === actualFieldName);
      return prop?.type || 'Edm.String';
    }
    return 'Edm.String';
  }

  const prop = targetSchema.properties.find(p => p.name === fieldName);
  return prop?.type || 'Edm.String';
}

function escapeODataString(str) {
  return str.replace(/'/g, "''");
}

function addFilter(field, operator, value, logic) {
  // Validate joined column filters have a matching active join
  if (field && field.includes('.')) {
    const [entityName] = field.split('.');
    if (!activeJoin) {
      showToast(`Cannot filter on "${field}" - no active join. Use Join button first.`, 'error');
      return;
    }
    if (normalizeEntityName(activeJoin.targetEntity) !== normalizeEntityName(entityName)) {
      showToast(`Cannot filter on "${field}" - joined entity is "${activeJoin.targetEntity}"`, 'error');
      return;
    }
  }

  // Validate field exists in current entity schema (prevents OData 400 errors)
  if (field && !field.includes('.') && entitySchema && entitySchema.properties) {
    const fieldExists = entitySchema.properties.some(p => p.name === field);
    if (!fieldExists) {
      const msg = `Field "${field}" does not exist on entity "${currentEntity}". Available columns: ${entitySchema.properties.slice(0, 10).map(p => p.name).join(', ')}...`;
      showToast(msg, 'error');
      throw new Error(msg);
    }
  }

  // logic: 'and' (default), 'or'. First filter has no logic prefix.
  const filterLogic = filterConfig.length === 0 ? null : (logic || 'and');
  filterConfig.push({ field, operator, value, logic: filterLogic });
  updateActiveFilters();
  currentPage = 1;
  if (window._aiDeferLoadData) {
    window._aiLoadDataNeeded = true;
    return Promise.resolve();
  }
  return loadData();
}

/**
 * Apply client-side quick filter when there's an active join.
 * When joined, server-side filtering is skipped so we can search ALL columns
 * (both base entity and joined entity columns).
 */
function applyClientSideQuickFilter() {
  const searchValue = quickFilter.value.trim().toLowerCase();
  if (!searchValue || !activeJoin || !data || data.length === 0) {
    return; // No search, no join, or no data - nothing to do
  }

  // Convert wildcard pattern to check
  // Remove wildcards for simple contains check
  const searchTerm = searchValue.replace(/\*/g, '');
  if (!searchTerm) return;

  // Get all searchable columns (both base and joined)
  const searchableColumns = visibleColumns.filter(col => !col.startsWith('@') && !col.startsWith('_'));

  console.log('Client-side quick filter searching columns:', searchableColumns.length, 'columns for term:', searchTerm);

  // Filter data - keep rows where ANY column contains the search term
  const beforeCount = data.length;
  data = data.filter(row => {
    for (const col of searchableColumns) {
      const value = row[col];
      if (value != null && String(value).toLowerCase().includes(searchTerm)) {
        return true;
      }
    }
    return false;
  });

  const filtered = beforeCount - data.length;
  if (filtered > 0) {
    console.log(`Client-side quick filter: kept ${data.length} of ${beforeCount} rows matching "${searchTerm}"`);
  }
}

function removeFilter(index) {
  filterConfig.splice(index, 1);
  updateActiveFilters();
  currentPage = 1;
  loadData();
}

function clearAllFilters() {
  filterConfig = [];
  quickFilter.value = '';
  clearFilterBtn.classList.add('hidden');
  updateActiveFilters();
  currentPage = 1;
  if (window._aiDeferLoadData) {
    window._aiLoadDataNeeded = true;
    return Promise.resolve();
  }
  return loadData();
}

function updateActiveFilters() {
  if (filterConfig.length === 0) {
    activeFiltersDiv.classList.add('hidden');
    return;
  }

  activeFiltersDiv.classList.remove('hidden');

  filtersList.innerHTML = filterConfig.map((filter, index) => `
    <div class="filter-chip">
      <span>${escapeHtml(filter.field)} ${escapeHtml(filter.operator)} "${escapeHtml(filter.value || '')}"</span>
      <button class="remove" data-index="${index}">&times;</button>
    </div>
  `).join('');

  // Add remove handlers
  filtersList.querySelectorAll('.remove').forEach(btn => {
    btn.addEventListener('click', () => {
      removeFilter(parseInt(btn.dataset.index));
    });
  });
}

// ==================== SORTING ====================
function sortByColumn(column) {
  // Validate field exists in current entity schema
  if (column && !column.includes('.') && entitySchema && entitySchema.properties) {
    const fieldExists = entitySchema.properties.some(p => p.name === column);
    if (!fieldExists) {
      const msg = `Field "${column}" does not exist on entity "${currentEntity}".`;
      showToast(msg, 'error');
      throw new Error(msg);
    }
  }

  if (sortConfig.field === column) {
    // Toggle direction
    sortConfig.direction = sortConfig.direction === 'asc' ? 'desc' : 'asc';
  } else {
    sortConfig.field = column;
    sortConfig.direction = 'asc';
  }

  currentPage = 1;
  if (window._aiDeferLoadData) {
    window._aiLoadDataNeeded = true;
    return Promise.resolve();
  }
  return loadData();
}

// ==================== PAGINATION ====================
function updatePagination() {
  const totalPages = Math.ceil(totalCount / pageSize) || 1;
  const startRow = (currentPage - 1) * pageSize + 1;
  const displayedRows = data.length;

  if (activeJoin && activeJoin.innerOnly) {
    // Inner join mode - counts are not reliable, show honest info
    pageInfo.textContent = `Showing ${displayedRows} matched rows`;

    // Hide total pages, show just current page
    pageInput.value = currentPage;
    pageInput.disabled = true; // Can't jump to specific page
    totalPagesEl.textContent = '?';
    totalPagesEl.title = 'Total pages unknown with inner join filtering';

    prevPageBtn.disabled = currentPage <= 1;
    // Allow next if we got a full page (might be more data)
    // or if the server says there's more
    nextPageBtn.disabled = (currentPage * pageSize) >= totalCount;
  } else {
    // Normal mode - accurate counts
    const endRow = Math.min(startRow + displayedRows - 1, totalCount);
    pageInfo.textContent = `Showing ${startRow}-${endRow} of ${totalCount.toLocaleString()}`;

    pageInput.value = currentPage;
    pageInput.disabled = false;
    totalPagesEl.textContent = totalPages;
    totalPagesEl.title = '';

    prevPageBtn.disabled = currentPage <= 1;
    nextPageBtn.disabled = currentPage >= totalPages;
  }
}

function goToPage(page) {
  const totalPages = Math.ceil(totalCount / pageSize) || 1;
  page = Math.max(1, Math.min(page, totalPages));

  if (page !== currentPage) {
    currentPage = page;
    if (window._aiDeferLoadData) {
      window._aiLoadDataNeeded = true;
      return Promise.resolve();
    }
    return loadData();
  }
}

// ==================== SELECTION ====================
function toggleRowSelection(index) {
  if (selectedRows.has(index)) {
    selectedRows.delete(index);
  } else {
    selectedRows.add(index);
  }
  updateSelectionInfo();
  renderGrid();
}

function selectAllRows() {
  if (selectedRows.size === data.length) {
    selectedRows.clear();
  } else {
    data.forEach((_, index) => selectedRows.add(index));
  }
  updateSelectionInfo();
  renderGrid();
}

function updateSelectionInfo() {
  selectionInfo.textContent = `${selectedRows.size} selected`;
}

function updateRecordCount() {
  const joinIndicator = document.getElementById('joinIndicator');

  if (activeJoin && activeJoin.innerOnly) {
    // Show displayed count when inner join is filtering rows
    recordCount.textContent = `${data.length} matched (this page)`;
    recordCount.className = 'record-count inner-join-active';
    recordCount.title = `Showing ${data.length} rows with matches in ${activeJoin.targetEntity}. Total matched count unknown - browse pages to find more.`;

    // Show join indicator in footer
    if (joinIndicator) {
      joinIndicator.textContent = `INNER JOIN → ${activeJoin.targetEntity}`;
      joinIndicator.classList.remove('hidden');
    }
  } else if (activeJoin) {
    // Join active but not filtering (left join)
    recordCount.textContent = `${totalCount.toLocaleString()} records`;
    recordCount.className = 'record-count';
    recordCount.title = `Joined with ${activeJoin.targetEntity}`;

    // Show join indicator
    if (joinIndicator) {
      joinIndicator.textContent = `LEFT JOIN → ${activeJoin.targetEntity}`;
      joinIndicator.classList.remove('hidden');
    }
  } else {
    recordCount.textContent = `${totalCount.toLocaleString()} records`;
    recordCount.className = 'record-count';
    recordCount.title = '';

    // Hide join indicator
    if (joinIndicator) {
      joinIndicator.classList.add('hidden');
    }
  }
}

// ==================== EXPORT ====================
function exportData(format, selectionOnly = false) {
  const exportData = selectionOnly
    ? Array.from(selectedRows).map(i => data[i])
    : data;

  if (exportData.length === 0) {
    alert('No data to export');
    return;
  }

  switch (format) {
    case 'csv':
      exportToCsv(exportData);
      break;
    case 'excel':
      exportToExcel(exportData);
      break;
    case 'json':
      exportToJson(exportData);
      break;
    case 'sql':
      exportToSql(exportData);
      break;
  }
}

function exportToCsv(exportData, filename = null) {
  if (!exportData || exportData.length === 0) {
    showToast('No data to export');
    return;
  }

  const columns = Object.keys(exportData[0]).filter(k => !k.startsWith('@') && !k.startsWith('__'));

  let csv = columns.join(',') + '\n';

  exportData.forEach(row => {
    const values = columns.map(col => {
      let value = row[col];
      if (value === null || value === undefined) return '';
      value = String(value);
      // Escape quotes and wrap in quotes if contains comma or quote
      if (value.includes(',') || value.includes('"') || value.includes('\n')) {
        value = '"' + value.replace(/"/g, '""') + '"';
      }
      return value;
    });
    csv += values.join(',') + '\n';
  });

  const exportFilename = filename || currentEntity;
  downloadFile(csv, `${exportFilename}.csv`, 'text/csv');
  console.log(`CSV export complete: ${exportFilename}.csv (${exportData.length} rows, ${columns.length} columns)`);
}

function exportToExcel(exportData) {
  // Simple XLSX generation (CSV with Excel MIME type for basic support)
  // For full XLSX support, would need xlsx library
  const columns = Object.keys(exportData[0]).filter(k => !k.startsWith('@'));

  let tsv = columns.join('\t') + '\n';

  exportData.forEach(row => {
    const values = columns.map(col => {
      let value = row[col];
      if (value === null || value === undefined) return '';
      return String(value).replace(/\t/g, ' ').replace(/\n/g, ' ');
    });
    tsv += values.join('\t') + '\n';
  });

  downloadFile(tsv, `${currentEntity}.xls`, 'application/vnd.ms-excel');
}

function exportToJson(exportData, filename = null) {
  if (!exportData || exportData.length === 0) {
    showToast('No data to export');
    return;
  }

  const cleanData = exportData.map(row => {
    const clean = {};
    Object.keys(row).forEach(key => {
      if (!key.startsWith('@') && !key.startsWith('__')) {
        clean[key] = row[key];
      }
    });
    return clean;
  });

  const json = JSON.stringify(cleanData, null, 2);
  const exportFilename = filename || currentEntity;
  downloadFile(json, `${exportFilename}.json`, 'application/json');
  console.log(`JSON export complete: ${exportFilename}.json (${exportData.length} rows)`);
}

function exportToSql(exportData) {
  const columns = Object.keys(exportData[0]).filter(k => !k.startsWith('@'));

  let sql = `-- SQL INSERT statements for ${currentEntity}\n`;
  sql += `-- Generated: ${new Date().toISOString()}\n\n`;

  exportData.forEach(row => {
    const values = columns.map(col => {
      const value = row[col];
      if (value === null || value === undefined) return 'NULL';
      if (typeof value === 'number') return value;
      if (typeof value === 'boolean') return value ? 1 : 0;
      return `'${String(value).replace(/'/g, "''")}'`;
    });

    sql += `INSERT INTO ${currentEntity} (${columns.join(', ')}) VALUES (${values.join(', ')});\n`;
  });

  downloadFile(sql, `${currentEntity}.sql`, 'text/plain');
}

function downloadFile(content, filename, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Export ALL records from the entity
async function exportAllData(format) {
  if (!currentEntity) {
    alert('No entity selected');
    return;
  }

  // Ask user about column preference
  const exportChoice = confirm(
    `Export ALL records from ${currentEntity}\n\n` +
    `Total records: ~${totalCount.toLocaleString()}\n` +
    `Selected columns: ${visibleColumns.length}\n\n` +
    `Click OK to export selected columns only\n` +
    `Click Cancel to export ALL columns`
  );

  const useSelectedColumns = exportChoice;

  showLoading();
  const loadingText = loadingOverlay.querySelector('span');
  const originalText = loadingText.textContent;

  // Add cancel button
  let cancelled = false;
  const cancelBtn = document.createElement('button');
  cancelBtn.textContent = 'Cancel Export';
  cancelBtn.className = 'btn-secondary';
  cancelBtn.style.marginTop = '12px';
  cancelBtn.onclick = () => { cancelled = true; };
  loadingOverlay.appendChild(cancelBtn);

  try {
    const allData = [];
    const batchSize = 5000; // Fetch in batches
    let skip = 0;
    let hasMore = true;

    // Build filter string (same as current view)
    const filterString = buildFilterString();

    // Columns to select - either user's selection or all
    const odataSelectColumns = useSelectedColumns
      ? visibleColumns.filter(col => !col.includes('.'))
      : undefined; // undefined = all columns

    console.log('Export ALL - using columns:', odataSelectColumns || 'ALL');

    while (hasMore && !cancelled) {
      loadingText.textContent = `Fetching... ${allData.length.toLocaleString()} / ~${totalCount.toLocaleString()} records`;

      const result = await odataClient.queryEntity(currentEntity, {
        select: odataSelectColumns,
        filter: filterString || undefined,
        top: batchSize,
        skip: skip,
        count: false
      });

      if (result.data && result.data.length > 0) {
        allData.push(...result.data);
        skip += result.data.length;
        hasMore = result.data.length === batchSize;
        console.log(`Fetched batch: ${result.data.length} rows, total: ${allData.length}`);
      } else {
        hasMore = false;
      }

      // Check at intervals to let user decide to continue
      if (allData.length > 0 && allData.length % 100000 === 0) {
        const continueExport = confirm(
          `Fetched ${allData.length.toLocaleString()} records so far.\n\n` +
          `Continue fetching more? This may take a while.\n` +
          `Click Cancel to export what we have.`
        );
        if (!continueExport) hasMore = false;
      }
    }

    if (cancelled) {
      showToast('Export cancelled');
      return;
    }

    loadingText.textContent = `Preparing export of ${allData.length.toLocaleString()} records...`;

    // If join is active, apply it to all data
    if (activeJoin) {
      loadingText.textContent = `Applying join to ${allData.length.toLocaleString()} records...`;
      await applyJoinToData(allData);
    }

    // Export based on format
    console.log(`Exporting ${allData.length} records as ${format}...`);

    if (format === 'csv') {
      exportToCsv(allData, `${currentEntity}_ALL`);
    } else if (format === 'json') {
      exportToJson(allData, `${currentEntity}_ALL`);
    }

    showToast(`Exported ${allData.length.toLocaleString()} records!`);
    console.log(`Export complete: ${allData.length} records`);

  } catch (error) {
    console.error('Export all failed:', error);
    alert('Export failed: ' + error.message);
  } finally {
    loadingText.textContent = originalText;
    cancelBtn.remove();
    hideLoading();
  }
}

// Apply join to a dataset (for export all)
async function applyJoinToData(dataToJoin) {
  if (!activeJoin) return;

  // Get unique join values
  const joinValues = [...new Set(dataToJoin.map(row => row[activeJoin.currentField]).filter(v => v != null))];

  if (joinValues.length === 0) return;

  // Fetch target data in batches
  let targetData = [];
  const batchSize = 5000;

  for (let i = 0; i < joinValues.length; i += 100) {
    const batchValues = joinValues.slice(i, i + 100);
    const filterConditions = batchValues.map(v => {
      if (typeof v === 'string') {
        return `${activeJoin.targetField} eq '${escapeODataString(v)}'`;
      }
      return `${activeJoin.targetField} eq ${v}`;
    });

    try {
      const result = await odataClient.queryEntity(activeJoin.targetEntity, {
        select: activeJoin.selectedColumns,
        filter: `(${filterConditions.join(' or ')})`,
        top: batchSize
      });
      targetData.push(...(result.data || []));
    } catch (e) {
      console.error('Error fetching join data batch:', e);
    }
  }

  // Build lookup
  const targetLookup = {};
  targetData.forEach(row => {
    const key = String(row[activeJoin.targetField]);
    if (!targetLookup[key]) targetLookup[key] = [];
    targetLookup[key].push(row);
  });

  // Merge
  for (let i = 0; i < dataToJoin.length; i++) {
    const row = dataToJoin[i];
    const joinKey = String(row[activeJoin.currentField]);
    const targetRows = targetLookup[joinKey] || [];
    const targetRow = targetRows[0] || {};

    activeJoin.selectedColumns.forEach(col => {
      row[`${activeJoin.targetEntity}.${col}`] = targetRow[col];
    });
  }

  // Filter if inner join
  if (activeJoin.innerOnly) {
    const before = dataToJoin.length;
    for (let i = dataToJoin.length - 1; i >= 0; i--) {
      const hasMatch = targetLookup[String(dataToJoin[i][activeJoin.currentField])]?.length > 0;
      if (!hasMatch) {
        dataToJoin.splice(i, 1);
      }
    }
    console.log(`Inner join filter: ${before} → ${dataToJoin.length}`);
  }
}

// ==================== UI HELPERS ====================
function showLoading(text = 'Loading data...') {
  loadingOverlay.classList.remove('hidden');
  loadingOverlay.querySelector('.loading-text').textContent = text;
  // Reset progress
  document.getElementById('progressBar')?.classList.add('hidden');
  document.getElementById('loadingStats')?.classList.add('hidden');
}

function updateLoadingProgress(current, total, stats = '') {
  const progressBar = document.getElementById('progressBar');
  const progressFill = document.getElementById('progressFill');
  const loadingStats = document.getElementById('loadingStats');
  const loadingText = loadingOverlay.querySelector('.loading-text');

  if (progressBar && progressFill) {
    progressBar.classList.remove('hidden');
    const percent = total > 0 ? Math.min((current / total) * 100, 100) : 0;
    progressFill.style.width = `${percent}%`;
  }

  loadingText.textContent = `Loading... ${current.toLocaleString()} / ${total.toLocaleString()}`;

  if (loadingStats && stats) {
    loadingStats.classList.remove('hidden');
    loadingStats.textContent = stats;
  }
}

function hideLoading() {
  loadingOverlay.classList.add('hidden');
  document.getElementById('progressBar')?.classList.add('hidden');
  document.getElementById('loadingStats')?.classList.add('hidden');
}

// Load ALL records into the current view (not just export)
async function loadAllIntoView() {
  if (!currentEntity) {
    alert('No entity selected');
    return;
  }

  if (totalCount > 50000) {
    const confirmed = confirm(
      `This will load ALL ${totalCount.toLocaleString()} records into the browser.\n\n` +
      `⚠️ This may use a lot of memory and slow down your browser.\n\n` +
      `Continue?`
    );
    if (!confirmed) return;
  }

  showLoading('Loading all records...');

  // Add cancel support
  let cancelled = false;
  const cancelBtn = document.createElement('button');
  cancelBtn.textContent = 'Cancel';
  cancelBtn.className = 'btn-secondary';
  cancelBtn.style.marginTop = '12px';
  cancelBtn.onclick = () => { cancelled = true; };
  loadingOverlay.appendChild(cancelBtn);

  try {
    const allData = [];
    const batchSize = 5000;
    let skip = 0;
    let hasMore = true;

    const filterString = buildFilterString();
    const odataSelectColumns = visibleColumns.filter(col => !col.includes('.'));

    const startTime = Date.now();

    while (hasMore && !cancelled) {
      updateLoadingProgress(allData.length, totalCount);

      const result = await odataClient.queryEntity(currentEntity, {
        select: odataSelectColumns.length > 0 ? odataSelectColumns : undefined,
        filter: filterString || undefined,
        top: batchSize,
        skip: skip,
        count: false
      });

      if (result.data && result.data.length > 0) {
        allData.push(...result.data);
        skip += result.data.length;
        hasMore = result.data.length === batchSize && allData.length < totalCount;
      } else {
        hasMore = false;
      }
    }

    if (cancelled) {
      showToast('Load cancelled');
      hideLoading();
      cancelBtn.remove();
      return;
    }

    // Apply join if active
    if (activeJoin) {
      loadingOverlay.querySelector('.loading-text').textContent = 'Applying join...';
      await applyJoinToData(allData);
    }

    // Update data and re-render
    data = allData;
    totalCount = allData.length;
    filteredByJoin = 0;

    const elapsed = ((Date.now() - startTime) / 1000).toFixed(1);
    showToast(`Loaded ${allData.length.toLocaleString()} records in ${elapsed}s`);

    renderData();
    updatePagination();
    updateRecordCount();

  } catch (error) {
    console.error('Load all failed:', error);
    alert('Failed to load all records: ' + error.message);
  } finally {
    cancelBtn.remove();
    hideLoading();
  }
}

function showError(message) {
  hideLoading();
  emptyState.querySelector('.message').textContent = 'Error';
  emptyState.querySelector('.hint').textContent = message;
  emptyState.classList.remove('hidden');
}

// Toast notification with optional type: 'success' (default), 'warning', 'error'
function showToast(message, typeOrDuration = 'success', duration = 3000) {
  const toast = document.getElementById('toast');

  // Handle legacy calls: showToast(message, durationNumber)
  let type = 'success';
  if (typeof typeOrDuration === 'number') {
    duration = typeOrDuration;
  } else if (typeof typeOrDuration === 'string') {
    type = typeOrDuration;
  }

  toast.textContent = message;
  toast.className = 'toast'; // Reset classes
  toast.classList.add(`toast-${type}`);
  toast.classList.remove('hidden');
  setTimeout(() => toast.classList.add('hidden'), duration);
}

// Copy to clipboard with feedback
async function copyToClipboard(text, successMessage = 'Copied!') {
  try {
    await navigator.clipboard.writeText(text);
    showToast(successMessage);
    return true;
  } catch (err) {
    console.error('Copy failed:', err);
    showToast('Copy failed');
    return false;
  }
}

// ==================== CONTEXT MENU ====================
let contextMenuTarget = { row: null, column: null, value: null };

function showContextMenu(e, rowIndex, column, value) {
  e.preventDefault();

  contextMenuTarget = { row: rowIndex, column, value };

  const menu = document.getElementById('contextMenu');
  menu.style.left = `${e.pageX}px`;
  menu.style.top = `${e.pageY}px`;
  menu.classList.remove('hidden');

  // Adjust if menu goes off screen
  const rect = menu.getBoundingClientRect();
  if (rect.right > window.innerWidth) {
    menu.style.left = `${e.pageX - rect.width}px`;
  }
  if (rect.bottom > window.innerHeight) {
    menu.style.top = `${e.pageY - rect.height}px`;
  }
}

function hideContextMenu() {
  document.getElementById('contextMenu').classList.add('hidden');
}

function handleContextMenuAction(action) {
  const { row, column, value } = contextMenuTarget;

  switch (action) {
    case 'copy-cell':
      copyToClipboard(String(value ?? ''), 'Cell value copied!');
      break;

    case 'copy-row':
      if (row !== null && data[row]) {
        copyToClipboard(JSON.stringify(data[row], null, 2), 'Row copied as JSON!');
      }
      break;

    case 'filter-equals':
      if (column && value !== null && value !== undefined) {
        addFilter(column, 'eq', String(value));
        showToast(`Filter added: ${column} = ${value}`);
      }
      break;

    case 'filter-not-equals':
      if (column && value !== null && value !== undefined) {
        addFilter(column, 'ne', String(value));
        showToast(`Filter added: ${column} ≠ ${value}`);
      }
      break;

    case 'show-details':
      if (row !== null) {
        showRowDetails(row);
      }
      break;

    case 'copy-odata-url':
      copyODataUrl();
      break;
  }

  hideContextMenu();
}

function copyODataUrl() {
  const filterString = buildFilterString();
  const odataSelectColumns = visibleColumns.filter(col => !col.includes('.'));

  const queryParams = ['cross-company=true'];
  if (odataSelectColumns.length > 0) {
    queryParams.push(`$select=${odataSelectColumns.join(',')}`);
  }
  if (filterString) {
    queryParams.push(`$filter=${encodeURIComponent(filterString)}`);
  }
  if (sortConfig.field && !sortConfig.field.includes('.')) {
    queryParams.push(`$orderby=${sortConfig.field} ${sortConfig.direction}`);
  }
  queryParams.push(`$top=${pageSize}`);
  queryParams.push('$count=true');

  const fullUrl = `${odataClient.baseUrl}${odataClient.odataPath}${currentEntity}?${queryParams.join('&')}`;
  copyToClipboard(fullUrl, 'OData URL copied!');
}

// ==================== ROW DETAILS PANEL ====================
function showRowDetails(rowIndex) {
  const row = data[rowIndex];
  if (!row) return;

  const panel = document.getElementById('sidePanel');
  const title = document.getElementById('panelTitle');
  const content = document.getElementById('panelContent');

  title.textContent = `Row Details (#${rowIndex + 1})`;

  // Get all columns including joined
  const allColumns = Object.keys(row).filter(k => !k.startsWith('@') && !k.startsWith('__'));

  // Separate by entity
  const mainCols = allColumns.filter(c => !c.includes('.'));
  const joinedCols = allColumns.filter(c => c.includes('.'));

  let html = `<div class="row-details">`;

  // Main entity fields
  html += `<div class="detail-section">
    <div class="detail-section-header">${currentEntity}</div>
    ${mainCols.map(col => {
      const val = row[col];
      const displayVal = formatDetailValue(val);
      const typeClass = getValueTypeClass(val);
      return `
        <div class="detail-field" data-column="${col}" data-value="${escapeHtml(String(val ?? ''))}">
          <span class="detail-label">${col}</span>
          <span class="detail-value ${typeClass}" title="Click to copy">${displayVal}</span>
        </div>
      `;
    }).join('')}
  </div>`;

  // Joined entity fields
  if (joinedCols.length > 0) {
    const joinedByEntity = {};
    joinedCols.forEach(col => {
      const [entity, field] = col.split('.');
      if (!joinedByEntity[entity]) joinedByEntity[entity] = [];
      joinedByEntity[entity].push({ col, field, value: row[col] });
    });

    Object.keys(joinedByEntity).forEach(entity => {
      html += `<div class="detail-section joined">
        <div class="detail-section-header">${entity} <span class="join-tag">JOINED</span></div>
        ${joinedByEntity[entity].map(({ col, field, value }) => {
          const displayVal = formatDetailValue(value);
          const typeClass = getValueTypeClass(value);
          return `
            <div class="detail-field" data-column="${col}" data-value="${escapeHtml(String(value ?? ''))}">
              <span class="detail-label">${field}</span>
              <span class="detail-value ${typeClass}" title="Click to copy">${displayVal}</span>
            </div>
          `;
        }).join('')}
      </div>`;
    });
  }

  html += `</div>`;

  // Action buttons
  html += `
    <div class="detail-actions">
      <button class="btn-secondary" id="copyRowJsonBtn" data-row-index="${rowIndex}">
        📋 Copy as JSON
      </button>
    </div>
  `;

  content.innerHTML = html;

  // Add click-to-copy on values
  content.querySelectorAll('.detail-field').forEach(field => {
    field.addEventListener('click', () => {
      const value = field.dataset.value;
      copyToClipboard(value, 'Value copied!');
    });
  });

  // Add click handler for copy JSON button
  const copyBtn = content.querySelector('#copyRowJsonBtn');
  if (copyBtn) {
    copyBtn.addEventListener('click', () => {
      const idx = parseInt(copyBtn.dataset.rowIndex);
      copyToClipboard(JSON.stringify(data[idx], null, 2), 'Row copied!');
    });
  }

  panel.classList.remove('hidden');
}

function formatDetailValue(value) {
  if (value === null || value === undefined) {
    return '<span class="null-value">null</span>';
  }
  if (typeof value === 'boolean') {
    return value
      ? '<span class="bool-true">✓ true</span>'
      : '<span class="bool-false">✗ false</span>';
  }
  if (typeof value === 'number') {
    return `<span class="number-value">${value.toLocaleString()}</span>`;
  }
  if (typeof value === 'string') {
    if (/^\d{4}-\d{2}-\d{2}T/.test(value)) {
      const date = new Date(value);
      return `<span class="date-value">${date.toLocaleString()}</span>`;
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
      return `<span class="date-value">${new Date(value + 'T00:00:00').toLocaleDateString()}</span>`;
    }
    // Check for URLs
    if (/^https?:\/\//.test(value)) {
      return `<a href="${escapeHtml(value)}" target="_blank" class="link-value">${escapeHtml(value)}</a>`;
    }
  }
  return escapeHtml(String(value));
}

function getValueTypeClass(value) {
  if (value === null || value === undefined) return 'type-null';
  if (typeof value === 'boolean') return 'type-bool';
  if (typeof value === 'number') return 'type-number';
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}/.test(value)) return 'type-date';
  return 'type-string';
}

// ==================== EVENT LISTENERS ====================
function setupEventListeners() {
  // Env badge click-to-edit
  envBadge.addEventListener('click', startBadgeEdit);

  // Entity search
  entityInput.addEventListener('focus', () => showEntityDropdown());
  entityInput.addEventListener('input', (e) => filterEntityDropdown(e.target.value));
  entityInput.addEventListener('blur', () => {
    setTimeout(() => entityDropdown.classList.add('hidden'), 200);
  });

  // Quick filter
  quickFilter.addEventListener('input', () => {
    clearFilterBtn.classList.toggle('hidden', !quickFilter.value);
  });

  quickFilter.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      currentPage = 1;
      loadData();
    }
  });

  clearFilterBtn.addEventListener('click', () => {
    quickFilter.value = '';
    clearFilterBtn.classList.add('hidden');
    currentPage = 1;
    loadData();
  });

  // Refresh
  document.getElementById('refreshBtn').addEventListener('click', () => loadData());

  // Load All button
  document.getElementById('loadAllBtn')?.addEventListener('click', loadAllIntoView);

  // Pagination
  prevPageBtn.addEventListener('click', () => goToPage(currentPage - 1));
  nextPageBtn.addEventListener('click', () => goToPage(currentPage + 1));
  pageInput.addEventListener('change', (e) => goToPage(parseInt(e.target.value)));

  // Export
  document.getElementById('exportBtn').addEventListener('click', (e) => {
    e.stopPropagation();
    document.getElementById('exportMenu').classList.toggle('hidden');
  });

  document.getElementById('exportMenu').querySelectorAll('button').forEach(btn => {
    btn.addEventListener('click', async () => {
      const format = btn.dataset.format;
      document.getElementById('exportMenu').classList.add('hidden');

      if (format.includes('-all')) {
        // Export all records
        await exportAllData(format.replace('-all', ''));
      } else if (format.includes('-selection')) {
        // Export selection only
        exportData(format.replace('-selection', ''), true);
      } else {
        // Export current page
        exportData(format, false);
      }
    });
  });

  // Close dropdown when clicking outside
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.export-dropdown')) {
      document.getElementById('exportMenu').classList.add('hidden');
    }
    if (!e.target.closest('.power-dropdown')) {
      document.getElementById('powerMenu')?.classList.add('hidden');
    }
  });

  // Power Platform dropdown
  document.getElementById('powerBtn')?.addEventListener('click', (e) => {
    e.stopPropagation();
    document.getElementById('powerMenu').classList.toggle('hidden');
    document.getElementById('exportMenu').classList.add('hidden');
  });

  document.getElementById('powerMenu')?.querySelectorAll('button[data-power]').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.dataset.power;
      document.getElementById('powerMenu').classList.add('hidden');
      if (typeof PowerToolsUI !== 'undefined') {
        PowerToolsUI.handleAction(action);
      }
    });
  });

  // AI Assistant button
  document.getElementById('aiAssistantBtn')?.addEventListener('click', () => {
    if (typeof AIAssistant !== 'undefined') {
      AIAssistant.togglePanel();
    }
  });

  // AI Analyze button
  document.getElementById('aiAnalyzeBtn')?.addEventListener('click', () => {
    if (typeof AIAnalyze !== 'undefined') {
      AIAnalyze.analyze();
    }
  });

  // inferenc.es connect/logout
  document.getElementById('inferencesConnectBtn')?.addEventListener('click', () => {
    if (typeof AISettings !== 'undefined') AISettings.connectInferences();
  });
  document.getElementById('inferencesLogoutBtn')?.addEventListener('click', () => {
    if (typeof AISettings !== 'undefined') AISettings.logoutInferences();
  });

  // Fetch models button
  document.getElementById('fetchModelsBtn')?.addEventListener('click', () => {
    if (typeof AISettings !== 'undefined') AISettings.fetchModels();
  });

  // Dark mode
  document.getElementById('darkModeBtn').addEventListener('click', () => {
    document.body.classList.toggle('dark');
    StorageManager.updateSettings({ theme: document.body.classList.contains('dark') ? 'dark' : 'light' });
  });

  // View mode toggle
  document.querySelectorAll('.btn-view').forEach(btn => {
    btn.addEventListener('click', () => {
      setViewMode(btn.dataset.view);
    });
  });

  // Settings
  document.getElementById('settingsBtn').addEventListener('click', showSettingsModal);

  // Clear all filters
  document.getElementById('clearAllFiltersBtn').addEventListener('click', clearAllFilters);

  // Query builder
  document.getElementById('queryBuilderBtn').addEventListener('click', showQueryBuilder);
  document.getElementById('cancelQueryBtn').addEventListener('click', hideQueryBuilder);
  document.getElementById('applyQueryBtn').addEventListener('click', applyQueryBuilder);
  document.getElementById('addConditionBtn').addEventListener('click', () => addQueryCondition());
  document.getElementById('addGroupBtn').addEventListener('click', () => addQueryGroup());

  // Related entities (expand)
  document.getElementById('relatedEntitiesBtn').addEventListener('click', showRelatedEntitiesModal);
  document.getElementById('applyExpandBtn').addEventListener('click', applyExpand);
  document.getElementById('clearExpandBtn').addEventListener('click', clearExpand);

  // Manual join
  document.getElementById('manualJoinBtn').addEventListener('click', showManualJoinModal);
  document.getElementById('cancelJoinBtn').addEventListener('click', () => {
    document.getElementById('manualJoinModal').classList.add('hidden');
  });
  document.getElementById('executeJoinBtn').addEventListener('click', executeJoin);

  // Join target entity search
  const joinTargetInput = document.getElementById('joinTargetEntity');
  const joinEntityDropdown = document.getElementById('joinEntityDropdown');

  joinTargetInput.addEventListener('focus', () => {
    if (window.allEntities) {
      showJoinEntityDropdown(joinTargetInput.value);
    }
  });

  joinTargetInput.addEventListener('input', (e) => {
    showJoinEntityDropdown(e.target.value);
  });

  joinTargetInput.addEventListener('blur', () => {
    setTimeout(() => joinEntityDropdown.classList.add('hidden'), 200);
  });

  // Join field changes
  document.getElementById('joinCurrentField').addEventListener('change', (e) => {
    joinConfig.currentField = e.target.value;
    updateJoinPreview();
  });

  document.getElementById('joinTargetField').addEventListener('change', (e) => {
    joinConfig.targetField = e.target.value;
    updateJoinPreview();
  });

  // Columns modal
  document.getElementById('columnsBtn').addEventListener('click', showColumnsModal);
  document.getElementById('applyColumnsBtn').addEventListener('click', applyColumnsSelection);
  document.getElementById('showAllColumnsBtn').addEventListener('click', () => selectAllColumns(true));
  document.getElementById('hideAllColumnsBtn').addEventListener('click', () => selectAllColumns(false));

  // Column search filter
  document.getElementById('columnSearch').addEventListener('input', (e) => {
    renderColumnsList(e.target.value);
  });

  // Modal close buttons
  document.querySelectorAll('.modal-close').forEach(btn => {
    btn.addEventListener('click', () => {
      btn.closest('.modal').classList.add('hidden');
    });
  });

  // Close modals on backdrop click
  document.querySelectorAll('.modal').forEach(modal => {
    modal.addEventListener('click', (e) => {
      if (e.target === modal) {
        modal.classList.add('hidden');
      }
    });
  });

  // Context menu actions
  document.getElementById('contextMenu')?.querySelectorAll('button').forEach(btn => {
    btn.addEventListener('click', () => {
      handleContextMenuAction(btn.dataset.action);
    });
  });

  // Hide context menu on click elsewhere
  document.addEventListener('click', (e) => {
    if (!e.target.closest('#contextMenu')) {
      hideContextMenu();
    }
  });

  // Hide context menu on scroll
  document.querySelector('.grid-container')?.addEventListener('scroll', hideContextMenu);

  // Close side panel
  document.getElementById('closePanelBtn')?.addEventListener('click', () => {
    document.getElementById('sidePanel').classList.add('hidden');
  });
}

function setupGridEventListeners() {
  // Header click for sorting
  gridHeader.querySelectorAll('th[data-column]').forEach(th => {
    th.addEventListener('click', () => sortByColumn(th.dataset.column));
  });

  // Select all checkbox
  const selectAllCheckbox = document.getElementById('selectAllCheckbox');
  if (selectAllCheckbox) {
    selectAllCheckbox.addEventListener('change', selectAllRows);
  }

  // Row and cell interactions
  gridBody.querySelectorAll('tr').forEach(row => {
    const rowIndex = parseInt(row.dataset.index);

    // Checkbox selection
    row.querySelector('input[type="checkbox"]')?.addEventListener('change', () => {
      toggleRowSelection(rowIndex);
    });

    // Cell interactions
    row.querySelectorAll('td').forEach((cell, cellIndex) => {
      if (cell.classList.contains('checkbox-cell')) return;

      const columns = visibleColumns.length > 0
        ? visibleColumns
        : Object.keys(data[0] || {}).filter(k => !k.startsWith('@'));
      const column = columns[cellIndex - 1]; // -1 for checkbox column
      const value = data[rowIndex]?.[column];

      // Right-click context menu
      cell.addEventListener('contextmenu', (e) => {
        showContextMenu(e, rowIndex, column, value);
      });

      // Double-click to copy
      cell.addEventListener('dblclick', () => {
        copyToClipboard(String(value ?? ''), 'Copied!');
      });

      // Single click to select row
      cell.addEventListener('click', (e) => {
        if (e.detail === 1) { // Single click only
          toggleRowSelection(rowIndex);
        }
      });
    });

    // Double-click row to show details
    row.addEventListener('dblclick', (e) => {
      if (e.target.type !== 'checkbox') {
        showRowDetails(rowIndex);
      }
    });
  });
}

function setupKeyboardShortcuts() {
  document.addEventListener('keydown', (e) => {
    // Ctrl+F - Focus quick filter
    if (e.ctrlKey && e.key === 'f') {
      e.preventDefault();
      quickFilter.focus();
    }

    // Ctrl+R - Refresh
    if (e.ctrlKey && e.key === 'r') {
      e.preventDefault();
      loadData();
    }

    // Ctrl+E - Export
    if (e.ctrlKey && e.key === 'e') {
      e.preventDefault();
      document.getElementById('exportMenu').classList.toggle('hidden');
    }

    // Ctrl+D - Dark mode
    if (e.ctrlKey && e.key === 'd') {
      e.preventDefault();
      document.body.classList.toggle('dark');
    }

    // Ctrl+Shift+D - Toggle debug panel
    if (e.ctrlKey && e.shiftKey && e.key === 'D') {
      e.preventDefault();
      toggleDebugPanel();
    }

    // Ctrl+A - Select all
    if (e.ctrlKey && e.key === 'a' && document.activeElement.tagName !== 'INPUT') {
      e.preventDefault();
      selectAllRows();
    }

    // Escape - Clear selection
    if (e.key === 'Escape') {
      selectedRows.clear();
      updateSelectionInfo();
      renderGrid();
    }
  });
}

// ==================== DEBUG PANEL ====================
function toggleDebugPanel() {
  const panel = document.getElementById('debugPanel');
  panel.classList.toggle('hidden');
  if (!panel.classList.contains('hidden')) {
    updateDebugInfo();
  }
}

function updateDebugInfo() {
  const filterString = buildFilterString();
  document.getElementById('debugFilter').textContent = filterString || '(none)';

  if (currentEntity && odataClient.baseUrl) {
    const queryParams = [];
    // D365 F&O requires cross-company=true for dataAreaId filtering
    queryParams.push('cross-company=true');
    if (visibleColumns.length > 0) {
      queryParams.push(`$select=${visibleColumns.join(',')}`);
    }
    if (filterString) {
      queryParams.push(`$filter=${encodeURIComponent(filterString)}`);
    }
    if (expandConfig.length > 0) {
      queryParams.push(`$expand=${expandConfig.join(',')}`);
    }
    queryParams.push(`$top=${pageSize}`);
    queryParams.push(`$skip=${(currentPage - 1) * pageSize}`);
    queryParams.push('$count=true');

    const fullUrl = `${odataClient.baseUrl}${odataClient.odataPath}${currentEntity}?${queryParams.join('&')}`;
    document.getElementById('debugUrl').textContent = fullUrl;
  }
}

// Setup debug panel close button
document.getElementById('closeDebugBtn')?.addEventListener('click', () => {
  document.getElementById('debugPanel').classList.add('hidden');
});

// ==================== ENTITY DROPDOWN ====================
let entitiesLoading = false;

async function loadEntitiesIfNeeded() {
  if (window.allEntities && window.allEntities.length > 0) return true;
  if (entitiesLoading) return false; // Already loading

  entitiesLoading = true;
  try {
    const entities = await odataClient.getEntities();
    window.allEntities = entities || [];
    entitiesLoading = false;
    return true;
  } catch (error) {
    console.error('Error loading entities:', error);
    entitiesLoading = false;
    return false;
  }
}

async function showEntityDropdown() {
  entityDropdown.classList.remove('hidden');

  // Load entities if not already loaded
  if (!window.allEntities || window.allEntities.length === 0) {
    entityDropdown.innerHTML = '<div class="entity-option"><small>Loading entities...</small></div>';
    const loaded = await loadEntitiesIfNeeded();
    if (!loaded) {
      entityDropdown.innerHTML = '<div class="entity-option"><small>Failed to load entities</small></div>';
      return;
    }
  }

  renderEntityDropdown(entityInput.value);
}

async function filterEntityDropdown(query) {
  // Make sure entities are loaded first
  if (!window.allEntities || window.allEntities.length === 0) {
    entityDropdown.innerHTML = '<div class="entity-option"><small>Loading entities...</small></div>';
    entityDropdown.classList.remove('hidden');
    const loaded = await loadEntitiesIfNeeded();
    if (!loaded) {
      entityDropdown.innerHTML = '<div class="entity-option"><small>Failed to load</small></div>';
      return;
    }
  }

  renderEntityDropdown(query);
}

function renderEntityDropdown(query) {
  if (!window.allEntities || window.allEntities.length === 0) {
    entityDropdown.innerHTML = '<div class="entity-option"><small>No entities available</small></div>';
    return;
  }

  const filtered = query
    ? window.allEntities.filter(e =>
        e.name.toLowerCase().includes(query.toLowerCase()) ||
        (e.label && e.label.toLowerCase().includes(query.toLowerCase()))
      ).slice(0, 20)
    : window.allEntities.slice(0, 20);

  if (filtered.length === 0) {
    entityDropdown.innerHTML = '<div class="entity-option"><small>No entities found matching "' + escapeHtml(query) + '"</small></div>';
    entityDropdown.classList.remove('hidden');
    return;
  }

  entityDropdown.innerHTML = filtered.map(e => `
    <div class="entity-option" data-entity="${escapeHtml(e.name)}">
      <strong>${escapeHtml(e.name)}</strong>
      <small>${escapeHtml(e.label || '')}</small>
    </div>
  `).join('');

  entityDropdown.classList.remove('hidden');

  entityDropdown.querySelectorAll('.entity-option').forEach(opt => {
    // Use mousedown instead of click - mousedown fires before blur
    opt.addEventListener('mousedown', (event) => {
      event.preventDefault();
      event.stopPropagation();
      const entityName = opt.dataset.entity;
      if (entityName) {
        console.log('Selected entity:', entityName);
        entityDropdown.classList.add('hidden');
        loadEntity(entityName);
      }
    });
  });
}

// ==================== QUERY BUILDER ====================
function showQueryBuilder() {
  document.getElementById('queryBuilderModal').classList.remove('hidden');
  renderQueryConditions();
}

function hideQueryBuilder() {
  document.getElementById('queryBuilderModal').classList.add('hidden');
}

function renderQueryConditions() {
  const container = document.getElementById('queryConditions');

  if (filterConfig.length === 0) {
    container.innerHTML = '<div class="empty-hint">Click "Add Condition" to create a filter</div>';
    document.getElementById('filterPreview').textContent = '-';
    return;
  }

  // Build column list, separating by entity when joins are active
  const allColumns = entitySchema?.properties?.map(p => p.name) || [];
  const joinedColumns = visibleColumns.filter(c => c.includes('.'));
  const currentEntityColumns = visibleColumns.filter(c => !c.includes('.') && !c.startsWith('@'));

  // Determine which entities are involved
  const joinedEntities = [...new Set(joinedColumns.map(c => c.split('.')[0]))];
  const hasJoinedData = joinedEntities.length > 0;

  const isNullOp = (op) => op === 'null' || op === 'notnull';
  const isStringOp = (op) => ['contains', 'startswith', 'endswith'].includes(op);

  // Get field type for display
  const getShortType = (fieldName) => {
    const type = getFieldType(fieldName);
    return type ? type.replace('Edm.', '') : 'String';
  };

  // Build column options grouped by entity
  const buildColumnOptions = (selectedField) => {
    if (!hasJoinedData) {
      // No joins - just list all columns
      const cols = entitySchema?.properties || visibleColumns.map(c => ({ name: c }));
      return cols.map(col => {
        const colName = col.name || col;
        const colType = getShortType(colName);
        return `<option value="${escapeHtml(colName)}" ${selectedField === colName ? 'selected' : ''}>${escapeHtml(colName)} (${escapeHtml(colType)})</option>`;
      }).join('');
    }

    // With joins - group by entity using optgroups
    let html = '';

    // Current entity columns
    html += `<optgroup label="📋 ${escapeHtml(currentEntity)}">`;
    currentEntityColumns.forEach(colName => {
      const colType = getShortType(colName);
      html += `<option value="${escapeHtml(colName)}" ${selectedField === colName ? 'selected' : ''}>${escapeHtml(colName)} (${escapeHtml(colType)})</option>`;
    });
    html += '</optgroup>';

    // Joined entity columns
    joinedEntities.forEach(entity => {
      html += `<optgroup label="🔗 ${escapeHtml(entity)} (joined)">`;
      joinedColumns
        .filter(c => c.startsWith(entity + '.'))
        .forEach(colName => {
          const displayName = colName.split('.').slice(1).join('.');
          const colType = getShortType(colName);
          html += `<option value="${escapeHtml(colName)}" ${selectedField === colName ? 'selected' : ''}>${escapeHtml(displayName)} (${escapeHtml(colType)})</option>`;
        });
      html += '</optgroup>';
    });

    return html;
  };

  container.innerHTML = filterConfig.map((filter, index) => {
    const fieldType = getFieldType(filter.field);
    const isStringField = fieldType === 'Edm.String' || !fieldType;
    const logic = filter.logic || (index > 0 ? 'and' : null);

    // Logic separator for non-first conditions
    const logicHtml = logic ? `
      <div class="logic-separator ${logic}">
        <select class="logic-select" data-index="${index}">
          <option value="and" ${logic === 'and' ? 'selected' : ''}>AND</option>
          <option value="or" ${logic === 'or' ? 'selected' : ''}>OR</option>
        </select>
      </div>
    ` : '';

    return `${logicHtml}
    <div class="condition-row" data-index="${index}">
      <select class="field-select">
        ${buildColumnOptions(filter.field)}
      </select>
      <select class="operator-select">
        <option value="eq" ${filter.operator === 'eq' ? 'selected' : ''}>equals</option>
        <option value="ne" ${filter.operator === 'ne' ? 'selected' : ''}>not equals</option>
        <option value="contains" ${filter.operator === 'contains' ? 'selected' : ''} ${!isStringField ? 'disabled' : ''}>contains (text)</option>
        <option value="startswith" ${filter.operator === 'startswith' ? 'selected' : ''} ${!isStringField ? 'disabled' : ''}>starts with (text)</option>
        <option value="endswith" ${filter.operator === 'endswith' ? 'selected' : ''} ${!isStringField ? 'disabled' : ''}>ends with (text)</option>
        <option value="gt" ${filter.operator === 'gt' ? 'selected' : ''}>greater than</option>
        <option value="lt" ${filter.operator === 'lt' ? 'selected' : ''}>less than</option>
        <option value="ge" ${filter.operator === 'ge' ? 'selected' : ''}>greater or equal</option>
        <option value="le" ${filter.operator === 'le' ? 'selected' : ''}>less or equal</option>
        <option value="null" ${filter.operator === 'null' ? 'selected' : ''}>is null</option>
        <option value="notnull" ${filter.operator === 'notnull' ? 'selected' : ''}>is not null</option>
      </select>
      <input type="text" class="value-input" value="${escapeHtml(filter.value || '')}"
             placeholder="${isNullOp(filter.operator) ? 'N/A' : 'Enter value...'}"
             ${isNullOp(filter.operator) ? 'disabled' : ''}>
      <button class="btn-remove" data-index="${index}">&times;</button>
    </div>
  `;
  }).join('');

  // Add event listeners
  container.querySelectorAll('.condition-row').forEach(row => {
    const index = parseInt(row.dataset.index);

    row.querySelector('.field-select').addEventListener('change', (e) => {
      filterConfig[index].field = e.target.value;
      // Check if current operator is valid for new field type
      const fieldType = getFieldType(e.target.value);
      const isStringField = fieldType === 'Edm.String' || !fieldType;
      const isStringOp = ['contains', 'startswith', 'endswith'].includes(filterConfig[index].operator);
      // If current operator is text-only but field is not string, reset to 'eq'
      if (isStringOp && !isStringField) {
        filterConfig[index].operator = 'eq';
      }
      // Re-render to update available operators
      renderQueryConditions();
    });

    row.querySelector('.operator-select').addEventListener('change', (e) => {
      filterConfig[index].operator = e.target.value;
      // Re-render to update value input disabled state
      renderQueryConditions();
    });

    row.querySelector('.value-input').addEventListener('input', (e) => {
      filterConfig[index].value = e.target.value;
      updateFilterPreview();
    });

    row.querySelector('.btn-remove').addEventListener('click', () => {
      filterConfig.splice(index, 1);
      // Update logic of next condition if this was the first one
      if (index === 0 && filterConfig.length > 0) {
        filterConfig[0].logic = null;
      }
      renderQueryConditions();
    });
  });

  // Add logic selector listeners
  container.querySelectorAll('.logic-select').forEach(select => {
    select.addEventListener('change', (e) => {
      const index = parseInt(e.target.dataset.index);
      filterConfig[index].logic = e.target.value;
      updateFilterPreview();
    });
  });

  updateFilterPreview();
}

function addQueryCondition(logic = 'and') {
  const columns = entitySchema?.properties?.map(p => p.name) || visibleColumns;
  filterConfig.push({
    field: columns[0] || '',
    operator: 'contains',
    value: '',
    logic: filterConfig.length === 0 ? null : logic // First condition has no logic prefix
  });
  renderQueryConditions();
}

function addQueryGroup() {
  // Add a new condition with OR logic (starts a new OR group)
  addQueryCondition('or');
}

function updateFilterPreview() {
  const filterString = buildFilterString();
  document.getElementById('filterPreview').textContent = filterString || '-';
}

function applyQueryBuilder() {
  hideQueryBuilder();
  updateActiveFilters();
  currentPage = 1;
  loadData();
}

// ==================== COLUMNS MODAL ====================
let allColumnsForModal = []; // Store all columns for filtering
let joinedColumnsForModal = {}; // Store joined columns grouped by entity

function showColumnsModal() {
  const modal = document.getElementById('columnsModal');
  const searchInput = document.getElementById('columnSearch');

  // Reset search
  searchInput.value = '';

  // Store columns for filtering
  allColumnsForModal = entitySchema?.properties || [];

  // Get joined columns
  const joinedCols = visibleColumns.filter(col => col.includes('.'));

  // Group joined columns by entity
  joinedColumnsForModal = {};
  joinedCols.forEach(col => {
    const [entityName, fieldName] = col.split('.');
    if (!joinedColumnsForModal[entityName]) {
      joinedColumnsForModal[entityName] = [];
    }
    joinedColumnsForModal[entityName].push({ name: col, fieldName });
  });

  renderColumnsList('');
  modal.classList.remove('hidden');

  // Focus search input
  setTimeout(() => searchInput.focus(), 100);
}

function renderColumnsList(searchQuery) {
  const list = document.getElementById('columnsList');
  const query = searchQuery.toLowerCase().trim();

  // Filter current entity columns
  const currentEntityCols = allColumnsForModal.filter(col => {
    if (col.name.includes('.')) return false; // Skip joined cols here
    if (!query) return true;
    return col.name.toLowerCase().includes(query);
  });

  // Filter joined columns
  const filteredJoinedByEntity = {};
  Object.keys(joinedColumnsForModal).forEach(entityName => {
    const filtered = joinedColumnsForModal[entityName].filter(col => {
      if (!query) return true;
      return col.fieldName.toLowerCase().includes(query) ||
             entityName.toLowerCase().includes(query);
    });
    if (filtered.length > 0) {
      filteredJoinedByEntity[entityName] = filtered;
    }
  });

  let html = '';

  // Safe OData types
  const safeODataTypes = ['Edm.String', 'Edm.Int16', 'Edm.Int32', 'Edm.Int64', 'Edm.Decimal', 'Edm.Double',
                          'Edm.Single', 'Edm.Boolean', 'Edm.Byte', 'Edm.Date', 'Edm.DateTime',
                          'Edm.DateTimeOffset', 'Edm.Guid', 'Edm.Binary', 'Edm.Time', 'Edm.Duration'];

  const isProblematicType = (type) => {
    if (!type) return false;
    if (safeODataTypes.includes(type)) return false;
    if (type.startsWith('Microsoft.Dynamics.DataEntities.')) return false;
    return true; // Problematic type like Dynamics.AX.Application.*
  };

  // Current entity columns
  if (currentEntityCols.length > 0) {
    html += `<div class="column-group">
      <div class="column-group-header">${escapeHtml(currentEntity || 'Current Entity')} <span class="column-count">(${currentEntityCols.length})</span></div>
      ${currentEntityCols.map(col => {
        const problematic = isProblematicType(col.type);
        return `
        <div class="column-item ${problematic ? 'problematic' : ''}" data-column="${col.name}" ${problematic ? 'title="This column has a non-standard type and may cause errors"' : ''}>
          <input type="checkbox" id="col_${col.name}" ${visibleColumns.includes(col.name) ? 'checked' : ''} ${problematic ? 'disabled' : ''}>
          <label for="col_${col.name}">${highlightMatch(col.name, query)}${problematic ? ' ⚠️' : ''}</label>
          <span class="type">${col.type?.split('.').pop() || 'String'}</span>
        </div>
      `}).join('')}
    </div>`;
  }

  // Joined entity columns
  Object.keys(filteredJoinedByEntity).forEach(entityName => {
    html += `<div class="column-group joined">
      <div class="column-group-header">${escapeHtml(entityName)} <span class="join-tag">JOINED</span> <span class="column-count">(${filteredJoinedByEntity[entityName].length})</span></div>
      ${filteredJoinedByEntity[entityName].map(col => `
        <div class="column-item joined-column" data-column="${col.name}">
          <input type="checkbox" id="col_${col.name.replace('.', '_')}" ${visibleColumns.includes(col.name) ? 'checked' : ''} data-fullname="${col.name}">
          <label for="col_${col.name.replace('.', '_')}">${highlightMatch(col.fieldName, query)}</label>
          <span class="type">Joined</span>
        </div>
      `).join('')}
    </div>`;
  });

  if (html === '') {
    html = `<div class="empty-hint">No columns found matching "${escapeHtml(searchQuery)}"</div>`;
  }

  list.innerHTML = html;
}

// Highlight matching text in column names
function highlightMatch(text, query) {
  if (!query) return escapeHtml(text);
  const safe = escapeHtml(text);
  const regex = new RegExp(`(${escapeRegexChars(query)})`, 'gi');
  return safe.replace(regex, '<mark>$1</mark>');
}

function escapeRegexChars(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function applyColumnsSelection() {
  const checkboxes = document.querySelectorAll('#columnsList input[type="checkbox"]');
  visibleColumns = Array.from(checkboxes)
    .filter(cb => cb.checked)
    .map(cb => cb.dataset.fullname || cb.id.replace('col_', ''));

  document.getElementById('columnsModal').classList.add('hidden');
  loadData();
}

function selectAllColumns(show) {
  // Only affect currently visible (filtered) checkboxes
  const checkboxes = document.querySelectorAll('#columnsList input[type="checkbox"]');
  checkboxes.forEach(cb => cb.checked = show);
}

// ==================== SETTINGS ====================
async function showSettingsModal() {
  const modal = document.getElementById('settingsModal');
  const settings = await StorageManager.getSettings();

  // Load current settings into form
  document.getElementById('settingPageSize').value = settings.pageSize || 100;
  document.getElementById('settingTheme').value = settings.theme || 'light';
  document.getElementById('settingShowRowNumbers').checked = settings.showRowNumbers !== false;
  document.getElementById('settingCompactMode').checked = settings.compactMode === true;

  // Load OData path
  const odataPathInput = document.getElementById('settingOdataPath');
  if (odataPathInput) {
    odataPathInput.value = settings.odataPath || '/data/';
  }

  // Show environment info
  const env = odataClient.environment;
  const envInfoEl = document.getElementById('settingEnvInfo');
  const manualUrlInput = document.getElementById('settingManualUrl');
  const envLabelInput = document.getElementById('settingEnvLabel');
  if (env) {
    const sourceLabel = env.source === 'manual' ? ' (manual)' : ' (auto-detected)';
    envInfoEl.textContent = `${env.envType}: ${env.hostname}${sourceLabel}`;
    if (env.source === 'manual' && manualUrlInput) {
      manualUrlInput.value = env.baseUrl;
    }
    // Populate badge label input with current custom label
    const labels = await getEnvLabels();
    if (envLabelInput) {
      envLabelInput.value = labels[env.hostname] || '';
      envLabelInput.placeholder = env.envType;
    }
  }

  // Badge label save/reset
  document.getElementById('saveEnvLabelBtn').onclick = async () => {
    const env = odataClient.environment;
    if (!env) { showToast('No environment detected', 'warn'); return; }
    const val = envLabelInput?.value?.trim();
    if (!val) { showToast('Enter a label', 'warn'); return; }
    await saveEnvLabel(env.hostname, val);
    envBadge.textContent = val;
    envBadge.className = `env-badge ${val.toLowerCase()}`;
    showToast(`Badge set to "${val}"`);
  };

  document.getElementById('resetEnvLabelBtn').onclick = async () => {
    const env = odataClient.environment;
    if (!env) return;
    await clearEnvLabel(env.hostname);
    envBadge.textContent = env.envType;
    envBadge.className = `env-badge ${env.envType.toLowerCase()}`;
    if (envLabelInput) { envLabelInput.value = ''; envLabelInput.placeholder = env.envType; }
    showToast('Badge reset to auto-detected');
  };

  // Manual URL buttons
  document.getElementById('setManualUrlBtn').onclick = async () => {
    const url = manualUrlInput?.value?.trim();
    if (!url) { showToast('Enter a URL first', 'warning'); return; }
    try {
      const newEnv = await odataClient.setManualEnvironment(url);
      envInfoEl.textContent = `${newEnv.envType}: ${newEnv.hostname} (manual)`;
      await applyEnvBadge(newEnv);
      showToast('Environment set! Reload entity to use new URL.');
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  document.getElementById('clearManualUrlBtn').onclick = async () => {
    await odataClient.clearManualEnvironment();
    if (manualUrlInput) manualUrlInput.value = '';
    envInfoEl.textContent = '--';
    showToast('Manual URL cleared. Auto-detect will be used.');
  };

  // Populate AI settings
  if (typeof AISettings !== 'undefined') {
    AISettings.populateSettingsUI();
  }

  // Clear cache button
  document.getElementById('clearCacheBtn').onclick = async () => {
    await StorageManager.remove('entityMetadata');
    odataClient.metadataCache = null;
    alert('Entity cache cleared! Entities will be reloaded on next use.');
  };

  // Reset all storage
  document.getElementById('clearStorageBtn').onclick = async () => {
    if (confirm('This will clear all settings, favorites, and cached data. Continue?')) {
      await chrome.storage.local.clear();
      location.reload();
    }
  };

  // Save button
  document.getElementById('saveSettingsBtn').onclick = async () => {
    let odataPathRaw = (document.getElementById('settingOdataPath')?.value || '/data/').trim();
    if (!odataPathRaw.startsWith('/')) odataPathRaw = '/' + odataPathRaw;
    if (!odataPathRaw.endsWith('/')) odataPathRaw = odataPathRaw + '/';
    const newSettings = {
      pageSize: parseInt(document.getElementById('settingPageSize').value),
      theme: document.getElementById('settingTheme').value,
      showRowNumbers: document.getElementById('settingShowRowNumbers').checked,
      compactMode: document.getElementById('settingCompactMode').checked,
      odataPath: odataPathRaw
    };

    await StorageManager.updateSettings(newSettings);

    // Apply OData path to client
    odataClient.odataPath = newSettings.odataPath;

    // Apply settings
    pageSize = newSettings.pageSize;

    if (newSettings.theme === 'dark') {
      document.body.classList.add('dark');
    } else {
      document.body.classList.remove('dark');
    }

    if (newSettings.compactMode) {
      document.body.classList.add('compact');
    } else {
      document.body.classList.remove('compact');
    }

    // Save AI settings
    if (typeof AISettings !== 'undefined') {
      await AISettings.saveFromUI();
    }

    modal.classList.add('hidden');

    // Reload data with new page size
    if (currentEntity) {
      currentPage = 1;
      loadData();
    }
  };

  modal.classList.remove('hidden');
}

// Load settings on init
async function loadSavedSettings() {
  const settings = await StorageManager.getSettings();

  pageSize = settings.pageSize || 100;

  if (settings.theme === 'dark') {
    document.body.classList.add('dark');
  }

  if (settings.compactMode) {
    document.body.classList.add('compact');
  }
}

// ==================== RELATED ENTITIES (EXPAND) ====================
async function showRelatedEntitiesModal() {
  const modal = document.getElementById('relatedEntitiesModal');
  const list = document.getElementById('relatedEntitiesList');
  const expandInfo = document.getElementById('expandInfo');

  let html = '';
  let hasAnyRelations = false;

  // Get navigation properties from current entity
  const currentNavProps = entitySchema?.navigationProperties || [];

  // Get navigation properties from joined entity (if any)
  let joinedNavProps = [];
  let joinedEntityName = null;
  if (activeJoin) {
    joinedEntityName = activeJoin.targetEntity;
    try {
      const joinedSchema = await odataClient.getEntitySchema(joinedEntityName);
      joinedNavProps = joinedSchema?.navigationProperties || [];
    } catch (e) {
      console.error('Error getting joined entity schema:', e);
    }
  }

  // Find common relationships (entities that both have relationships to)
  const currentRelatedEntities = new Set(currentNavProps.map(np => np.relatedEntity));
  const joinedRelatedEntities = new Set(joinedNavProps.map(np => np.relatedEntity));
  const commonRelated = [...currentRelatedEntities].filter(e => joinedRelatedEntities.has(e));

  // Show common relationships section if any
  if (commonRelated.length > 0 && activeJoin) {
    html += `<div class="related-section">
      <div class="related-section-header common">
        <span>🔗 Common Relationships</span>
        <small>Both entities relate to these</small>
      </div>`;

    commonRelated.forEach(commonEntity => {
      const fromCurrent = currentNavProps.find(np => np.relatedEntity === commonEntity);
      const fromJoined = joinedNavProps.find(np => np.relatedEntity === commonEntity);
      html += `
        <div class="related-item common-relation">
          <div class="common-relation-info">
            <div class="common-relation-entity">📊 ${commonEntity}</div>
            <div class="common-relation-paths">
              <span>${currentEntity}.${fromCurrent?.name}</span>
              <span>${joinedEntityName}.${fromJoined?.name}</span>
            </div>
          </div>
        </div>`;
    });
    html += '</div>';
    hasAnyRelations = true;
  }

  // Current entity relations
  if (currentNavProps.length > 0) {
    html += `<div class="related-section">
      <div class="related-section-header">
        <span>${currentEntity}</span>
        <small>${currentNavProps.length} relationships</small>
      </div>`;

    html += currentNavProps.map(navProp => `
      <label class="related-item">
        <input type="checkbox" value="${navProp.name}" data-entity="current" ${expandConfig.includes(navProp.name) ? 'checked' : ''}>
        <div class="related-item-info">
          <div class="related-item-name">${navProp.name}</div>
          <div class="related-item-entity">→ ${navProp.relatedEntity}</div>
        </div>
        <span class="related-item-type ${navProp.isCollection ? 'collection' : ''}">
          ${navProp.isCollection ? 'Many' : 'One'}
        </span>
      </label>
    `).join('');

    html += '</div>';
    hasAnyRelations = true;
  }

  // Joined entity relations (if join is active)
  if (activeJoin && joinedNavProps.length > 0) {
    html += `<div class="related-section joined">
      <div class="related-section-header joined">
        <span>${joinedEntityName}</span>
        <span class="join-tag">JOINED</span>
        <small>${joinedNavProps.length} relationships</small>
      </div>`;

    html += joinedNavProps.map(navProp => `
      <label class="related-item joined-relation">
        <input type="checkbox" value="${joinedEntityName}.${navProp.name}" data-entity="joined" disabled title="Expand on joined entities not supported yet">
        <div class="related-item-info">
          <div class="related-item-name">${navProp.name}</div>
          <div class="related-item-entity">→ ${navProp.relatedEntity}</div>
        </div>
        <span class="related-item-type ${navProp.isCollection ? 'collection' : ''}">
          ${navProp.isCollection ? 'Many' : 'One'}
        </span>
      </label>
    `).join('');

    html += '</div>';
    hasAnyRelations = true;
  }

  if (!hasAnyRelations) {
    html = '<div class="empty-hint">No related entities found.<br><small>These entities have no navigation properties defined in the metadata.</small></div>';
    expandInfo.classList.add('hidden');
  }

  list.innerHTML = html;

  // Add change listeners
  list.querySelectorAll('input[type="checkbox"]:not(:disabled)').forEach(cb => {
    cb.addEventListener('change', updateExpandPreview);
  });

  updateExpandPreview();
  modal.classList.remove('hidden');
}

function updateExpandPreview() {
  const list = document.getElementById('relatedEntitiesList');
  const expandInfo = document.getElementById('expandInfo');
  const expandPreview = document.getElementById('expandPreview');

  const selectedExpands = Array.from(list.querySelectorAll('input[type="checkbox"]:checked'))
    .map(cb => cb.value);

  if (selectedExpands.length > 0) {
    expandInfo.classList.remove('hidden');
    expandPreview.textContent = `$expand=${selectedExpands.join(',')}`;
  } else {
    expandInfo.classList.add('hidden');
    expandPreview.textContent = '-';
  }
}

function applyExpand() {
  const list = document.getElementById('relatedEntitiesList');

  // Get previously expanded navigation properties to clean up their columns
  const oldExpandedNavProps = [...expandConfig];

  expandConfig = Array.from(list.querySelectorAll('input[type="checkbox"]:checked'))
    .map(cb => cb.value);

  console.log('Applied expansions:', expandConfig);

  // Remove old expanded columns from visibleColumns
  if (oldExpandedNavProps.length > 0) {
    visibleColumns = visibleColumns.filter(col => {
      // Keep column if it's not from an old expanded nav property
      const colPrefix = col.split('.')[0];
      return !oldExpandedNavProps.includes(colPrefix);
    });
  }

  document.getElementById('relatedEntitiesModal').classList.add('hidden');

  // Update the Related button to show count
  const btn = document.getElementById('relatedEntitiesBtn');
  if (expandConfig.length > 0) {
    btn.innerHTML = `&#128279; Related <span class="expand-badge">${expandConfig.length}</span>`;
  } else {
    btn.innerHTML = '&#128279; Related';
  }

  // Reload data with expansions
  currentPage = 1;
  loadData();
}

function clearExpand() {
  const list = document.getElementById('relatedEntitiesList');
  list.querySelectorAll('input[type="checkbox"]').forEach(cb => {
    cb.checked = false;
  });
  updateExpandPreview();

  // If there were previous expansions, remove their columns and reload
  if (expandConfig.length > 0) {
    // Remove expanded columns from visibleColumns
    visibleColumns = visibleColumns.filter(col => {
      const colPrefix = col.split('.')[0];
      return !expandConfig.includes(colPrefix);
    });

    expandConfig = [];

    // Update button
    const btn = document.getElementById('relatedEntitiesBtn');
    btn.innerHTML = '&#128279; Related';

    // Close modal and reload
    document.getElementById('relatedEntitiesModal').classList.add('hidden');
    currentPage = 1;
    if (window._aiDeferLoadData) {
      window._aiLoadDataNeeded = true;
      return;
    }
    loadData();
  }
}

// ==================== MANUAL JOIN ====================
let joinConfig = {
  targetEntity: null,
  targetSchema: null,
  currentField: null,
  targetField: null,
  selectedColumns: [],
  innerOnly: false // If true, filter out rows with no match (INNER JOIN behavior)
};

// Active join that persists across data reloads (pagination, refresh, etc.)
let activeJoin = null;

function showManualJoinModal() {
  if (!currentEntity || !entitySchema) {
    alert('Please load an entity first');
    return;
  }

  const modal = document.getElementById('manualJoinModal');

  // Set current entity name
  document.getElementById('joinCurrentEntity').textContent = currentEntity;

  // Populate current entity fields
  const currentFieldSelect = document.getElementById('joinCurrentField');
  currentFieldSelect.innerHTML = '<option value="">Select field...</option>' +
    entitySchema.properties.map(p =>
      `<option value="${p.name}">${p.name} (${(p.type || 'String').replace('Edm.', '')})</option>`
    ).join('');

  // Reset target entity
  document.getElementById('joinTargetEntity').value = '';
  document.getElementById('joinTargetField').innerHTML = '<option value="">Select entity first...</option>';
  document.getElementById('joinTargetField').disabled = true;
  document.getElementById('joinColumnsSection').classList.add('hidden');
  document.getElementById('joinPreview').classList.add('hidden');
  document.getElementById('executeJoinBtn').disabled = true;

  joinConfig = {
    targetEntity: null,
    targetSchema: null,
    currentField: null,
    targetField: null,
    selectedColumns: [],
    innerOnly: false
  };

  // Reset inner join checkbox
  const innerOnlyCheckbox = document.getElementById('joinInnerOnly');
  if (innerOnlyCheckbox) innerOnlyCheckbox.checked = false;

  modal.classList.remove('hidden');
}

async function showJoinEntityDropdown(query) {
  const dropdown = document.getElementById('joinEntityDropdown');
  if (!dropdown) return;

  // Load entities if not already loaded
  if (!window.allEntities || window.allEntities.length === 0) {
    dropdown.innerHTML = '<div class="entity-option"><small>Loading entities...</small></div>';
    dropdown.classList.remove('hidden');
    try {
      const entities = await odataClient.getEntities();
      window.allEntities = entities || [];
    } catch (error) {
      console.error('Error loading entities:', error);
      dropdown.innerHTML = '<div class="entity-option"><small>Failed to load entities</small></div>';
      return;
    }
  }

  const filtered = query
    ? window.allEntities.filter(e =>
        e.name.toLowerCase().includes(query.toLowerCase()) ||
        (e.label && e.label.toLowerCase().includes(query.toLowerCase()))
      ).slice(0, 15)
    : window.allEntities.slice(0, 15);

  if (filtered.length === 0) {
    dropdown.innerHTML = '<div class="entity-option"><small>No entities found</small></div>';
    dropdown.classList.remove('hidden');
    return;
  }

  dropdown.innerHTML = filtered.map(e => `
    <div class="entity-option" data-entity="${escapeHtml(e.name)}">
      <strong>${escapeHtml(e.name)}</strong>
      <small>${escapeHtml(e.label || '')}</small>
    </div>
  `).join('');

  dropdown.classList.remove('hidden');

  dropdown.querySelectorAll('.entity-option').forEach(opt => {
    // Use mousedown instead of click - mousedown fires before blur
    opt.addEventListener('mousedown', async (event) => {
      event.preventDefault();
      event.stopPropagation();
      const entityName = opt.dataset.entity;
      if (entityName) {
        console.log('Join: Selected target entity:', entityName);
        document.getElementById('joinTargetEntity').value = entityName;
        dropdown.classList.add('hidden');
        await loadTargetEntitySchema(entityName);
      }
    });
  });
}

async function loadTargetEntitySchema(entityName) {
  joinConfig.targetEntity = entityName;

  try {
    // Get schema for target entity
    const schema = await odataClient.getEntitySchema(entityName);
    joinConfig.targetSchema = schema;

    // Populate target field select
    const targetFieldSelect = document.getElementById('joinTargetField');
    targetFieldSelect.disabled = false;
    targetFieldSelect.innerHTML = '<option value="">Select field...</option>' +
      schema.properties.map(p =>
        `<option value="${p.name}">${p.name} (${(p.type || 'String').replace('Edm.', '')})</option>`
      ).join('');

    // Show columns selection
    const columnsSection = document.getElementById('joinColumnsSection');
    const columnsList = document.getElementById('joinColumnsList');

    // Filter to only show columns with safe OData types (exclude custom D365 types that cause serialization errors)
    const safeTypes = ['Edm.String', 'Edm.Int16', 'Edm.Int32', 'Edm.Int64', 'Edm.Decimal', 'Edm.Double',
                       'Edm.Single', 'Edm.Boolean', 'Edm.Byte', 'Edm.Date', 'Edm.DateTime',
                       'Edm.DateTimeOffset', 'Edm.Guid', 'Edm.Binary', 'Edm.Time', 'Edm.Duration'];

    const safeColumns = schema.properties.filter(p => {
      // Allow standard Edm types
      if (!p.type || safeTypes.includes(p.type)) return true;
      // Allow enum types (Microsoft.Dynamics.DataEntities.*)
      if (p.type.startsWith('Microsoft.Dynamics.DataEntities.')) return true;
      // Exclude custom types like Dynamics.AX.Application.* which cause serialization errors
      console.log(`Excluding column "${p.name}" with problematic type: ${p.type}`);
      return false;
    });

    // Show first 30 safe columns as selectable
    const displayCols = safeColumns.slice(0, 30);
    columnsList.innerHTML = displayCols.map(p => `
      <label class="join-column-item">
        <input type="checkbox" value="${p.name}" checked>
        ${p.name}
        <span class="col-type">${(p.type || 'String').replace('Edm.', '').replace('Microsoft.Dynamics.DataEntities.', '')}</span>
      </label>
    `).join('');

    columnsSection.classList.remove('hidden');

    // Add change listeners for columns
    columnsList.querySelectorAll('input').forEach(cb => {
      cb.addEventListener('change', updateJoinPreview);
    });

    updateJoinPreview();

  } catch (error) {
    console.error('Error loading target schema:', error);
    alert('Failed to load entity schema: ' + error.message);
  }
}

function updateJoinPreview() {
  const preview = document.getElementById('joinPreview');
  const previewText = document.getElementById('joinPreviewText');
  const executeBtn = document.getElementById('executeJoinBtn');

  // Get selected columns
  const columnsList = document.getElementById('joinColumnsList');
  joinConfig.selectedColumns = Array.from(columnsList?.querySelectorAll('input:checked') || [])
    .map(cb => cb.value);

  if (joinConfig.currentField && joinConfig.targetEntity && joinConfig.targetField && joinConfig.selectedColumns.length > 0) {
    preview.classList.remove('hidden');
    previewText.textContent = `JOIN ${currentEntity}.${joinConfig.currentField} = ${joinConfig.targetEntity}.${joinConfig.targetField}\n` +
      `SELECT ${joinConfig.selectedColumns.slice(0, 5).join(', ')}${joinConfig.selectedColumns.length > 5 ? '...' : ''} FROM ${joinConfig.targetEntity}`;
    executeBtn.disabled = false;
  } else {
    preview.classList.add('hidden');
    executeBtn.disabled = true;
  }
}

async function executeJoin() {
  if (!joinConfig.currentField || !joinConfig.targetEntity || !joinConfig.targetField) {
    alert('Please configure the join first');
    return;
  }

  // Clear any filters on old joined columns (from previous join) since they're no longer valid
  const oldJoinedFilters = filterConfig.filter(f => f.field && f.field.includes('.'));
  if (oldJoinedFilters.length > 0) {
    const clearedFields = oldJoinedFilters.map(f => f.field).join(', ');
    console.log('Clearing old joined column filters:', clearedFields);
    showToast(`Cleared ${oldJoinedFilters.length} filter(s) from previous join: ${clearedFields}`, 'warning');
    filterConfig = filterConfig.filter(f => !f.field || !f.field.includes('.'));
    updateActiveFilters();
  }

  // Get inner join option from checkbox (unless set programmatically by AI joinEntity())
  if (!joinConfig._programmatic) {
    const innerOnlyCheckbox = document.getElementById('joinInnerOnly');
    joinConfig.innerOnly = innerOnlyCheckbox ? innerOnlyCheckbox.checked : false;
  }

  console.log('Executing join with config:', {
    targetEntity: joinConfig.targetEntity,
    targetField: joinConfig.targetField,
    currentField: joinConfig.currentField,
    hasSchema: !!joinConfig.targetSchema,
    schemaProperties: joinConfig.targetSchema?.properties?.length || 0
  });

  showLoading();

  try {
    // Get unique values from current data for the join field
    const joinValues = [...new Set(data.map(row => row[joinConfig.currentField]).filter(v => v != null))];

    if (joinValues.length === 0) {
      hideLoading();
      alert('No values found in the join field');
      return;
    }

    console.log(`Joining on ${joinValues.length} unique values`);

    // Safe OData types
    const safeODataTypes = ['Edm.String', 'Edm.Int16', 'Edm.Int32', 'Edm.Int64', 'Edm.Decimal', 'Edm.Double',
                            'Edm.Single', 'Edm.Boolean', 'Edm.Byte', 'Edm.Date', 'Edm.DateTime',
                            'Edm.DateTimeOffset', 'Edm.Guid', 'Edm.Binary', 'Edm.Time', 'Edm.Duration'];

    // Filter selected columns to only safe types (in case UI filter didn't catch all)
    const safeSelectedColumns = joinConfig.selectedColumns.filter(colName => {
      const colDef = joinConfig.targetSchema?.properties?.find(p => p.name === colName);
      const colType = colDef?.type;
      if (!colType) return true; // Unknown - allow
      if (safeODataTypes.includes(colType)) return true;
      if (colType.startsWith('Microsoft.Dynamics.DataEntities.')) return true;
      console.warn(`Join: Excluding column "${colName}" due to problematic type: ${colType}`);
      return false;
    });

    // Always include the join field
    if (!safeSelectedColumns.includes(joinConfig.targetField)) {
      safeSelectedColumns.unshift(joinConfig.targetField);
    }

    console.log('Join safe columns:', safeSelectedColumns);

    // Build filter for target entity using D365 wildcard/eq syntax
    // For many values, we need to batch or use IN-like logic
    let targetData = [];

    // Get the target field type to handle enums correctly
    const targetFieldType = joinConfig.targetSchema?.properties?.find(p => p.name === joinConfig.targetField)?.type || 'Edm.String';
    const isTargetEnum = targetFieldType && !targetFieldType.startsWith('Edm.');

    if (joinValues.length <= 10) {
      // Small number - use OR conditions
      const filterConditions = joinValues.map(v => {
        if (isTargetEnum) {
          // Enum field - use fully qualified type
          return `${joinConfig.targetField} eq ${targetFieldType}'${escapeODataString(String(v))}'`;
        } else if (typeof v === 'string') {
          return `${joinConfig.targetField} eq '${escapeODataString(v)}'`;
        }
        return `${joinConfig.targetField} eq ${v}`;
      });
      const filter = `(${filterConditions.join(' or ')})`;

      const result = await odataClient.queryEntity(joinConfig.targetEntity, {
        select: safeSelectedColumns,
        filter: filter,
        top: 5000
      });
      targetData = result.data;
    } else {
      // Many values - query all and filter client-side (or batch)
      // For now, query with a reasonable limit
      const result = await odataClient.queryEntity(joinConfig.targetEntity, {
        select: safeSelectedColumns,
        top: 10000
      });
      // Filter client-side
      const joinValueSet = new Set(joinValues.map(v => String(v)));
      targetData = result.data.filter(row =>
        joinValueSet.has(String(row[joinConfig.targetField]))
      );
    }

    console.log(`Retrieved ${targetData.length} records from target entity`);

    // Build lookup map
    const targetLookup = {};
    targetData.forEach(row => {
      const key = String(row[joinConfig.targetField]);
      if (!targetLookup[key]) {
        targetLookup[key] = [];
      }
      targetLookup[key].push(row);
    });

    // Merge data - add target columns to current data
    let joinedData = data.map(row => {
      const joinKey = String(row[joinConfig.currentField]);
      const targetRows = targetLookup[joinKey] || [];

      // Take first matching row (or could aggregate)
      const targetRow = targetRows[0] || {};
      const hasMatch = targetRows.length > 0;

      // Merge, prefixing target columns to avoid conflicts
      const merged = { ...row, __hasJoinMatch: hasMatch };
      safeSelectedColumns.forEach(col => {
        merged[`${joinConfig.targetEntity}.${col}`] = targetRow[col];
      });

      return merged;
    });

    // If innerOnly is true, filter out rows with no match (INNER JOIN)
    filteredByJoin = 0;
    if (joinConfig.innerOnly) {
      const beforeCount = joinedData.length;
      joinedData = joinedData.filter(row => row.__hasJoinMatch);
      const afterCount = joinedData.length;
      filteredByJoin = beforeCount - afterCount;
      console.log(`Inner join: filtered ${filteredByJoin} unmatched rows (${beforeCount} → ${afterCount})`);
    }

    // Remove the temporary flag
    joinedData.forEach(row => delete row.__hasJoinMatch);

    // Store active join for re-application on refresh/pagination (use safe columns)
    activeJoin = {
      targetEntity: joinConfig.targetEntity,
      currentField: joinConfig.currentField,
      targetField: joinConfig.targetField,
      selectedColumns: [...safeSelectedColumns], // Use filtered safe columns
      innerOnly: joinConfig.innerOnly,
      targetSchema: joinConfig.targetSchema // For type lookups when filtering on joined columns
    };

    // Update visible columns to include joined columns (filter out any existing joined cols first)
    const baseColumns = visibleColumns.filter(col => !col.includes('.'));
    const newColumns = safeSelectedColumns.map(col => `${joinConfig.targetEntity}.${col}`);
    visibleColumns = [...baseColumns, ...newColumns];

    // Update data
    data = joinedData;

    // Update the Join button to show active join
    const joinBtn = document.getElementById('manualJoinBtn');
    joinBtn.innerHTML = `&#128200; Join <span class="expand-badge">${joinConfig.targetEntity}</span>`;

    // Close modal and re-render
    document.getElementById('manualJoinModal').classList.add('hidden');
    hideLoading();
    renderData();

    // Show success message
    console.log('Join completed successfully. Active join stored for refresh/pagination.');

  } catch (error) {
    hideLoading();
    console.error('Join failed:', error);
    alert('Join failed: ' + error.message);
  }
}

// Programmatic join — called by AI assistant
// Usage: joinEntity('ReleasedProductsV2', 'ItemNumber', 'ItemNumber', true) — true = inner join (only matched rows)
async function joinEntity(targetEntity, currentField, targetField, innerOnly) {
  if (!currentEntity || !entitySchema) {
    showToast('Load an entity first', 'error');
    return;
  }

  // Auto-correct entity name
  targetEntity = resolveEntityName(targetEntity);

  showLoading();

  try {
    // Fetch REAL columns by querying 1 record (metadata cache can have wrong property names)
    let realColumns = null;
    try {
      const probe = await odataClient.queryEntity(targetEntity, { top: 1 });
      if (probe.data && probe.data.length > 0) {
        realColumns = new Set(Object.keys(probe.data[0]).filter(k => !k.startsWith('@') && !k.startsWith('_')));
      }
    } catch (e) {
      console.warn('Probe query failed for', targetEntity, '- using schema only:', e.message);
    }

    // Fetch target entity schema for type information
    const schema = await odataClient.getEntitySchema(targetEntity);
    if (!schema || !schema.properties) {
      throw new Error(`Could not load schema for ${targetEntity}`);
    }

    // Auto-select safe columns — only include properties that ACTUALLY exist on the entity
    const safeTypes = ['Edm.String', 'Edm.Int16', 'Edm.Int32', 'Edm.Int64', 'Edm.Decimal', 'Edm.Double',
                       'Edm.Single', 'Edm.Boolean', 'Edm.Byte', 'Edm.Date', 'Edm.DateTime',
                       'Edm.DateTimeOffset', 'Edm.Guid'];

    const safeColumns = schema.properties
      .filter(p => {
        if (!(!p.type || safeTypes.includes(p.type) || p.type.startsWith('Microsoft.Dynamics.DataEntities.'))) return false;
        // If we have real columns from probe, only include properties that actually exist
        if (realColumns && !realColumns.has(p.name)) {
          console.warn(`Join: Skipping "${p.name}" — not in actual entity response`);
          return false;
        }
        return true;
      })
      .map(p => p.name);

    // Verify target join field exists
    if (realColumns && !realColumns.has(targetField)) {
      throw new Error(`Field "${targetField}" does not exist on ${targetEntity}. Available: ${[...realColumns].slice(0, 10).join(', ')}...`);
    }

    // Always include target join field
    const selectedColumns = [targetField, ...safeColumns.filter(c => c !== targetField)].slice(0, 20);

    // Set up joinConfig and execute (_programmatic prevents executeJoin from overwriting innerOnly)
    joinConfig = {
      targetEntity,
      targetSchema: schema,
      currentField,
      targetField,
      selectedColumns,
      innerOnly: innerOnly === true,
      _programmatic: true
    };

    await executeJoin();
    showToast(`Joined with ${targetEntity}`);
  } catch (error) {
    hideLoading();
    showToast('Join failed: ' + error.message, 'error');
    console.error('joinEntity failed:', error);
  }
}

// ==================== HIGHLIGHTING ====================

function highlightCells(field, operator, value, color) {
  // Store config so highlights survive re-renders
  highlightConfigs.push({ type: 'cell', field, operator, value, color });
  const count = applyHighlight({ type: 'cell', field, operator, value, color });
  showToast(`Highlighted ${count} cells in ${field}`, 'info');
}

function highlightRows(field, operator, value, color) {
  // Store config so highlights survive re-renders
  highlightConfigs.push({ type: 'row', field, operator, value, color });
  const count = applyHighlight({ type: 'row', field, operator, value, color });
  showToast(`Highlighted ${count} rows`, 'info');
}

function clearHighlights() {
  highlightConfigs = [];
  gridBody.querySelectorAll('.highlight-cell, .highlight-row').forEach(el => {
    el.classList.remove('highlight-cell', 'highlight-row',
      'highlight-red', 'highlight-green', 'highlight-yellow',
      'highlight-blue', 'highlight-orange', 'highlight-purple');
  });
  showToast('Highlights cleared');
}

/**
 * Apply a single highlight config to the current DOM. Returns match count.
 */
function applyHighlight(cfg) {
  const colorMap = {
    red: 'highlight-red', green: 'highlight-green', yellow: 'highlight-yellow',
    blue: 'highlight-blue', orange: 'highlight-orange', purple: 'highlight-purple'
  };
  const cls = colorMap[cfg.color] || colorMap.yellow;

  const columns = visibleColumns.length > 0
    ? visibleColumns
    : Object.keys(data[0] || {}).filter(k => !k.startsWith('@'));

  if (cfg.type === 'cell') {
    const colIdx = columns.indexOf(cfg.field);
    if (colIdx === -1) return 0;

    let count = 0;
    const rows = gridBody.querySelectorAll('tr');
    rows.forEach((tr, i) => {
      if (i >= data.length) return;
      if (matchHighlight(data[i], cfg)) {
        const cell = tr.children[colIdx + 1]; // +1 to skip checkbox cell
        if (cell) { cell.classList.add('highlight-cell', cls); count++; }
      }
    });
    return count;
  } else {
    let count = 0;
    const rows = gridBody.querySelectorAll('tr');
    rows.forEach((tr, i) => {
      if (i >= data.length) return;
      if (matchHighlight(data[i], cfg)) {
        tr.classList.add('highlight-row', cls);
        count++;
      }
    });
    return count;
  }
}

function matchHighlight(row, cfg) {
  const rawVal = row[cfg.field];
  const strVal = rawVal == null ? '' : String(rawVal);
  const numVal = parseFloat(rawVal);
  switch (cfg.operator) {
    case 'eq': return strVal === String(cfg.value);
    case 'ne': return strVal !== String(cfg.value);
    case 'contains': return strVal.toLowerCase().includes(String(cfg.value).toLowerCase());
    case 'gt': return !isNaN(numVal) && numVal > Number(cfg.value);
    case 'lt': return !isNaN(numVal) && numVal < Number(cfg.value);
    case 'ge': return !isNaN(numVal) && numVal >= Number(cfg.value);
    case 'le': return !isNaN(numVal) && numVal <= Number(cfg.value);
    case 'null': return rawVal == null || strVal === '';
    case 'notnull': return rawVal != null && strVal !== '';
    default: return strVal === String(cfg.value);
  }
}

/**
 * Re-apply all stored highlights after a re-render.
 */
function reapplyHighlights() {
  if (highlightConfigs.length === 0) return;
  for (const cfg of highlightConfigs) {
    applyHighlight(cfg);
  }
}

// ==================== DATA ANALYSIS ====================
// Operate on loaded data[] — no additional OData queries

function summarizeData(field) {
  if (!data || data.length === 0) return 'No data loaded';
  const counts = {};
  data.forEach(row => {
    const key = row[field] == null ? '(blank)' : String(row[field]);
    counts[key] = (counts[key] || 0) + 1;
  });
  const sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 30);
  const total = data.length;
  const unique = Object.keys(counts).length;
  const lines = sorted.map(([val, count]) => `  ${val}: ${count} (${(count / total * 100).toFixed(1)}%)`);
  window.aiAnalysisResult = `Summary of "${field}" (${total} rows, ${unique} unique values):\n${lines.join('\n')}`;
  showToast(`Summarized ${field}: ${unique} unique values`);
}

function computeStats(field) {
  if (!data || data.length === 0) return 'No data loaded';
  const nums = data.map(r => parseFloat(r[field])).filter(n => !isNaN(n));
  if (nums.length === 0) { window.aiAnalysisResult = `"${field}" has no numeric values`; return; }
  const sum = nums.reduce((a, b) => a + b, 0);
  const sorted = [...nums].sort((a, b) => a - b);
  const median = sorted.length % 2 === 0
    ? (sorted[sorted.length / 2 - 1] + sorted[sorted.length / 2]) / 2
    : sorted[Math.floor(sorted.length / 2)];
  window.aiAnalysisResult = `Stats for "${field}" (${nums.length} values):\n  Min: ${Math.min(...nums)}\n  Max: ${Math.max(...nums)}\n  Sum: ${sum}\n  Avg: ${(sum / nums.length).toFixed(2)}\n  Median: ${median}`;
  showToast(`Stats computed for ${field}`);
}

function getDistinctValues(field) {
  if (!data || data.length === 0) return 'No data loaded';
  const vals = [...new Set(data.map(r => r[field] == null ? '(blank)' : String(r[field])))].sort();
  window.aiAnalysisResult = `Distinct "${field}" (${vals.length} unique):\n${vals.slice(0, 50).join(', ')}` +
    (vals.length > 50 ? `\n...and ${vals.length - 50} more` : '');
  showToast(`${vals.length} distinct values for ${field}`);
}

function crossTab(field1, field2) {
  if (!data || data.length === 0) return 'No data loaded';
  const combos = {};
  data.forEach(row => {
    const k = `${row[field1] == null ? '(blank)' : row[field1]} × ${row[field2] == null ? '(blank)' : row[field2]}`;
    combos[k] = (combos[k] || 0) + 1;
  });
  const sorted = Object.entries(combos).sort((a, b) => b[1] - a[1]).slice(0, 30);
  window.aiAnalysisResult = `Cross-tab "${field1}" × "${field2}" (top 30):\n` +
    sorted.map(([k, c]) => `  ${k}: ${c}`).join('\n');
  showToast(`Cross-tab: ${field1} × ${field2}`);
}

async function compareEntities(targetEntity) {
  if (!currentEntity || !data || data.length === 0) {
    window.aiAnalysisResult = 'Load an entity with data first';
    showToast('No data to compare', 'error');
    return;
  }

  // Auto-correct entity name
  targetEntity = resolveEntityName(targetEntity);

  showToast(`Comparing with ${targetEntity}...`, 'info');

  try {
    // Get current entity columns and sample values (from already-loaded data)
    const currentCols = Object.keys(data[0])
      .filter(k => !k.startsWith('@') && !k.startsWith('_') && !k.includes('.'));

    // Get target entity sample data (1 query, no $select to get real columns)
    const targetResult = await odataClient.queryEntity(targetEntity, { top: 100 });
    if (!targetResult.data || targetResult.data.length === 0) {
      window.aiAnalysisResult = `${targetEntity} has no data to compare with.`;
      showToast(`${targetEntity} is empty`, 'warning');
      return;
    }

    const targetCols = Object.keys(targetResult.data[0])
      .filter(k => !k.startsWith('@') && !k.startsWith('_'));

    // 1. Exact column name matches
    const nameMatches = currentCols.filter(c => targetCols.includes(c));

    // 2. Value overlap analysis — find columns that share actual data values
    // Build distinct value sets from current entity (already loaded, no extra query)
    const currentDistinct = {};
    for (const col of currentCols) {
      const vals = new Set();
      for (let i = 0; i < Math.min(data.length, 500); i++) {
        const v = data[i][col];
        if (v != null && v !== '' && v !== 'null') vals.add(String(v));
      }
      // Skip columns with too few or too many distinct values (unlikely join keys)
      if (vals.size >= 2 && vals.size <= 5000) {
        currentDistinct[col] = vals;
      }
    }

    // Build distinct value sets from target entity
    const targetDistinct = {};
    for (const col of targetCols) {
      const vals = new Set();
      for (const row of targetResult.data) {
        const v = row[col];
        if (v != null && v !== '' && v !== 'null') vals.add(String(v));
      }
      if (vals.size >= 2 && vals.size <= 5000) {
        targetDistinct[col] = vals;
      }
    }

    // Find column pairs with value overlap
    const valueMatches = [];
    for (const cCol of Object.keys(currentDistinct)) {
      for (const tCol of Object.keys(targetDistinct)) {
        let overlap = 0;
        for (const v of currentDistinct[cCol]) {
          if (targetDistinct[tCol].has(v)) overlap++;
          if (overlap > 50) break; // Enough to confirm match
        }
        if (overlap >= 2) {
          const smaller = Math.min(currentDistinct[cCol].size, targetDistinct[tCol].size);
          const overlapPct = Math.round(overlap / smaller * 100);
          if (overlapPct >= 10) {
            valueMatches.push({
              currentCol: cCol,
              targetCol: tCol,
              overlap,
              overlapPct,
              sampleValues: [...currentDistinct[cCol]].filter(v => targetDistinct[tCol].has(v)).slice(0, 3)
            });
          }
        }
      }
    }

    valueMatches.sort((a, b) => b.overlapPct - a.overlapPct);

    // Build result
    let result = `=== Entity Comparison: ${currentEntity} ↔ ${targetEntity} ===\n\n`;

    result += `EXACT NAME MATCHES (${nameMatches.length}):\n`;
    if (nameMatches.length > 0) {
      nameMatches.forEach(col => result += `  • ${col}\n`);
    } else {
      result += '  (none)\n';
    }

    result += `\nJOINABLE COLUMNS (shared values):\n`;
    if (valueMatches.length > 0) {
      valueMatches.slice(0, 15).forEach(m => {
        const tag = m.currentCol === m.targetCol ? ' [SAME NAME]' : '';
        result += `  • ${currentEntity}.${m.currentCol} ↔ ${targetEntity}.${m.targetCol}: ${m.overlap} shared values (${m.overlapPct}% overlap)${tag}`;
        result += ` — e.g. "${m.sampleValues.join('", "')}"\n`;
      });
    } else {
      result += '  (no value overlap found in sampled data)\n';
    }

    result += `\nBEST JOIN CANDIDATES:\n`;
    const bestJoins = valueMatches.filter(m => m.overlapPct >= 30).slice(0, 5);
    if (bestJoins.length > 0) {
      bestJoins.forEach(m => {
        result += `  joinEntity('${targetEntity}', '${m.currentCol}', '${m.targetCol}', true)\n`;
      });
    } else if (nameMatches.length > 0) {
      nameMatches.slice(0, 3).forEach(col => {
        result += `  joinEntity('${targetEntity}', '${col}', '${col}', true)  — same name, check if values match\n`;
      });
    } else {
      result += '  No strong candidates found. Entities may not be directly related.\n';
    }

    result += `\n${currentEntity}: ${currentCols.length} columns, ${data.length} rows loaded`;
    result += `\n${targetEntity}: ${targetCols.length} columns, ${targetResult.data.length} rows sampled`;

    window.aiAnalysisResult = result;
    showToast(`Compared ${currentEntity} ↔ ${targetEntity}`);

  } catch (error) {
    window.aiAnalysisResult = `Error comparing: ${error.message}`;
    showToast('Compare failed: ' + error.message, 'error');
  }
}

function setPageSize(size) {
  const capped = Math.min(Math.max(1, size), 100000);
  pageSize = capped;
  const pageSizeEl = document.getElementById('settingPageSize');
  if (pageSizeEl) pageSizeEl.value = capped;
  currentPage = 1;
  if (window._aiDeferLoadData) {
    window._aiLoadDataNeeded = true;
    return Promise.resolve();
  }
  return loadData();
}

function setVisibleColumns(columns) {
  if (!Array.isArray(columns) || columns.length === 0) return;
  visibleColumns = columns.filter(c =>
    entitySchema?.properties?.some(p => p.name === c) || c.includes('.')
  );
  if (window._aiDeferLoadData) {
    window._aiLoadDataNeeded = true;
    return Promise.resolve();
  }
  return loadData();
}

function searchEntities(query) {
  if (!query || !window.allEntities || window.allEntities.length === 0) {
    window.aiAnalysisResult = 'Entity list not loaded yet. Try loading an entity first.';
    return;
  }
  const q = query.toLowerCase();
  const all = window.allEntities
    .filter(e => {
      const n = (e.name || '').toLowerCase();
      const l = (e.label || '').toLowerCase();
      return n.includes(q) || l.includes(q);
    });

  // Sort: V2/V3 data entities first (primary D365 entities), then by name length (shorter = more relevant)
  all.sort((a, b) => {
    const aIsData = /V\d+$/.test(a.name) ? 0 : 1;
    const bIsData = /V\d+$/.test(b.name) ? 0 : 1;
    if (aIsData !== bIsData) return aIsData - bIsData;
    return a.name.length - b.name.length;
  });

  const top = all.slice(0, 30);
  const total = all.length;
  const lines = top.map(e => {
    const label = e.label && e.label !== e.name ? ` (${e.label})` : '';
    const cat = e.category ? ` [${e.category}]` : '';
    return `  ${e.name}${label}${cat}`;
  });
  if (lines.length > 0) {
    window.aiAnalysisResult = `Found ${total} entities matching "${query}" (showing top ${lines.length}, V2/V3 data entities listed first):\n${lines.join('\n')}\n\nIMPORTANT: Use EXACTLY one of these names in loadEntity(). Do NOT modify or guess names.`;
  } else {
    window.aiAnalysisResult = `No entities matching "${query}" out of ${window.allEntities.length} total entities. Try a SHORTER or DIFFERENT keyword (e.g. instead of "formulalines" try "formula", instead of "productionbom" try "bom"). Do NOT call loadEntity() — the entity does not exist.`;
  }
}

function getRelatedEntities() {
  const navProps = entitySchema?.navigationProperties || [];
  if (navProps.length === 0) {
    window.aiAnalysisResult = 'No navigation properties found for this entity.';
    return;
  }
  const lines = navProps.map(np =>
    `- ${np.name} → ${np.relatedEntity} (${np.isCollection ? 'collection' : 'single'})`
  );
  const active = expandConfig.length > 0 ? `\nCurrently expanded: ${expandConfig.join(', ')}` : '';
  window.aiAnalysisResult = `Related entities for ${currentEntity} (${navProps.length} relationships):\n${lines.join('\n')}${active}\n\nUse expandEntity('navPropertyName') to expand, or clearExpand() to clear.`;
}

function expandEntity(navPropertyNames) {
  if (!entitySchema) {
    throw new Error('No entity loaded. Load an entity first.');
  }
  const navProps = entitySchema.navigationProperties || [];
  const names = Array.isArray(navPropertyNames) ? navPropertyNames : [navPropertyNames];
  for (const name of names) {
    if (!navProps.some(np => np.name === name)) {
      const available = navProps.slice(0, 10).map(np => np.name).join(', ');
      const msg = `Navigation property "${name}" not found on ${currentEntity}. Call getRelatedEntities() first to see available properties. Available: ${available}${navProps.length > 10 ? '...' : ''}`;
      showToast(msg, 'error');
      throw new Error(msg);
    }
    if (!expandConfig.includes(name)) {
      expandConfig.push(name);
    }
  }
  const btn = document.getElementById('relatedEntitiesBtn');
  if (btn && expandConfig.length > 0) {
    btn.innerHTML = `&#128279; Related <span class="expand-badge">${expandConfig.length}</span>`;
  }
  currentPage = 1;
  if (window._aiDeferLoadData) {
    window._aiLoadDataNeeded = true;
    return Promise.resolve();
  }
  return loadData();
}

// Re-apply an active join after data reload (pagination, refresh, etc.)
async function reapplyJoin() {
  // Capture join config locally to prevent race conditions during async operations
  const join = activeJoin;
  if (!join) return;

  console.log('Re-applying active join:', join);

  try {
    // Get unique values from current data for the join field
    const joinValues = [...new Set(data.map(row => row[join.currentField]).filter(v => v != null))];

    if (joinValues.length === 0) {
      console.warn(`No values found for join field "${join.currentField}" in ${data.length} rows. Available columns: ${data.length > 0 ? Object.keys(data[0]).slice(0, 10).join(', ') : 'none'}`);
      showToast(`Join warning: "${join.currentField}" has no values in current data`, 'warning');
      return;
    }

    // Fetch target entity data
    let targetData = [];

    if (joinValues.length <= 10) {
      // Small number - use OR conditions
      const filterConditions = joinValues.map(v => {
        if (typeof v === 'string') {
          return `${join.targetField} eq '${escapeODataString(v)}'`;
        }
        return `${join.targetField} eq ${v}`;
      });
      const filter = `(${filterConditions.join(' or ')})`;

      const result = await odataClient.queryEntity(join.targetEntity, {
        select: join.selectedColumns,
        filter: filter,
        top: 5000
      });
      targetData = result.data;
    } else {
      // Many values - query with limit and filter client-side
      const result = await odataClient.queryEntity(join.targetEntity, {
        select: join.selectedColumns,
        top: 10000
      });
      const joinValueSet = new Set(joinValues.map(v => String(v)));
      targetData = result.data.filter(row =>
        joinValueSet.has(String(row[join.targetField]))
      );
    }

    // Build lookup map
    const targetLookup = {};
    targetData.forEach(row => {
      const key = String(row[join.targetField]);
      if (!targetLookup[key]) {
        targetLookup[key] = [];
      }
      targetLookup[key].push(row);
    });

    // Merge data
    let mergedData = data.map(row => {
      const joinKey = String(row[join.currentField]);
      const targetRows = targetLookup[joinKey] || [];
      const targetRow = targetRows[0] || {};
      const hasMatch = targetRows.length > 0;

      const merged = { ...row, __hasJoinMatch: hasMatch };
      join.selectedColumns.forEach(col => {
        merged[`${join.targetEntity}.${col}`] = targetRow[col];
      });

      return merged;
    });

    // If innerOnly is true, filter out rows with no match
    filteredByJoin = 0;
    if (join.innerOnly) {
      const beforeCount = mergedData.length;
      mergedData = mergedData.filter(row => row.__hasJoinMatch);
      const afterCount = mergedData.length;
      filteredByJoin = beforeCount - afterCount;
      console.log(`Inner join re-apply: filtered ${filteredByJoin} unmatched rows`);
    }

    // Remove the temporary flag and update data
    mergedData.forEach(row => delete row.__hasJoinMatch);
    data = mergedData;

    console.log('Join re-applied successfully');

  } catch (error) {
    console.error('Error re-applying join:', error);
    // Don't fail the whole load, just log the error
  }
}

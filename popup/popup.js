/**
 * D365 AI Data Manager - Popup
 * Entity search and selection
 */

// State
let allEntities = [];
let favorites = [];
let recent = [];
let searchQuery = '';

// DOM Elements
const entitySearch = document.getElementById('entitySearch');
const entityList = document.getElementById('entityList');
const envBadge = document.getElementById('envBadge');
const statusEl = document.getElementById('status');
const errorPanel = document.getElementById('errorPanel');
const errorMessage = document.getElementById('errorMessage');
const refreshBtn = document.getElementById('refreshBtn');
const retryBtn = document.getElementById('retryBtn');

// Initialize
document.addEventListener('DOMContentLoaded', async () => {
  await init();
});

async function init() {
  try {
    // Detect environment
    let env = await odataClient.detectEnvironment();

    if (!env) {
      // Show manual URL input as fallback
      showManualUrlInput();
      return;
    }

    // Update environment badge
    envBadge.textContent = env.envType;
    envBadge.className = `env-badge ${env.envType.toLowerCase()}`;

    // Save environment
    await StorageManager.addEnvironment(env);

    // Load favorites and recent
    favorites = await StorageManager.getFavorites();
    recent = await StorageManager.getRecent();

    // Load entities
    await loadEntities();

  } catch (error) {
    console.error('Init error:', error);
    showError(error.message);
  }
}

async function loadEntities(forceRefresh = false) {
  showLoading();

  try {
    // Try to get from cache first
    if (!forceRefresh) {
      const cached = await odataClient.getMetadata();
      if (cached && cached.entities) {
        allEntities = cached.entities;
        renderEntities();
        statusEl.textContent = `${allEntities.length} entities (cached)`;
        return;
      }
    }

    // Fetch fresh
    statusEl.textContent = 'Fetching entities...';
    allEntities = await odataClient.getEntities();

    // Cache results
    await odataClient.cacheMetadata(allEntities);

    renderEntities();
    statusEl.textContent = `${allEntities.length} entities`;

  } catch (error) {
    console.error('Error loading entities:', error);
    showError(`Failed to load entities: ${error.message}`);
  }
}

function renderEntities() {
  const query = searchQuery.toLowerCase().trim();

  // Filter entities
  let filtered = allEntities;
  if (query) {
    filtered = allEntities.filter(e =>
      e.name.toLowerCase().includes(query) ||
      e.label.toLowerCase().includes(query)
    );
  }

  // Group entities
  const favoriteEntities = filtered.filter(e => favorites.includes(e.name));
  const recentEntities = filtered.filter(e =>
    recent.includes(e.name) && !favorites.includes(e.name)
  );

  // Group by category for all entities
  const categorized = {};
  filtered.forEach(e => {
    if (!categorized[e.category]) {
      categorized[e.category] = [];
    }
    categorized[e.category].push(e);
  });

  let html = '';

  // Favorites section
  if (favoriteEntities.length > 0 && !query) {
    html += renderSection('Favorites', favoriteEntities, 'star');
  }

  // Recent section
  if (recentEntities.length > 0 && !query) {
    html += renderSection('Recent', recentEntities.slice(0, 5), 'clock');
  }

  // All entities (or search results)
  if (query) {
    // Show flat list for search
    html += renderSection(`Results (${filtered.length})`, filtered, 'search');
  } else {
    // Show categorized
    const categoryOrder = ['Sales', 'Purchasing', 'Inventory', 'Finance', 'HR', 'Project', 'Manufacturing', 'Other'];
    categoryOrder.forEach(cat => {
      if (categorized[cat] && categorized[cat].length > 0) {
        html += renderSection(cat, categorized[cat], 'folder');
      }
    });
  }

  if (html === '') {
    html = `
      <div class="empty-state">
        <div class="icon">&#128269;</div>
        <span class="title">No entities found</span>
        <span class="subtitle">Try a different search term</span>
      </div>
    `;
  }

  entityList.innerHTML = html;

  // Add event listeners
  entityList.querySelectorAll('.entity-item').forEach(item => {
    item.addEventListener('click', (e) => {
      if (!e.target.closest('.btn-action')) {
        openEntity(item.dataset.entity);
      }
    });
  });

  entityList.querySelectorAll('.btn-action.favorite').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      toggleFavorite(btn.dataset.entity);
    });
  });

  entityList.querySelectorAll('.section-header').forEach(header => {
    header.addEventListener('click', () => {
      const section = header.closest('.entity-section');
      section.classList.toggle('collapsed');
    });
  });
}

function renderSection(title, entities, icon) {
  const iconMap = {
    'star': '&#11088;',
    'clock': '&#128337;',
    'folder': '&#128193;',
    'search': '&#128269;'
  };

  const entitiesHtml = entities.map(e => {
    const isFav = favorites.includes(e.name);
    const highlighted = searchQuery
      ? highlightMatch(e.name, searchQuery)
      : e.name;

    return `
      <div class="entity-item" data-entity="${escapeHtml(e.name)}">
        <div class="entity-icon">&#128203;</div>
        <div class="entity-info">
          <div class="entity-name">${highlighted}</div>
          <div class="entity-label">${escapeHtml(e.label)}</div>
        </div>
        <div class="entity-actions">
          <button class="btn-action favorite ${isFav ? 'active' : ''}"
                  data-entity="${escapeHtml(e.name)}"
                  title="${isFav ? 'Remove from favorites' : 'Add to favorites'}">
            ${isFav ? '&#9733;' : '&#9734;'}
          </button>
        </div>
      </div>
    `;
  }).join('');

  return `
    <div class="entity-section">
      <div class="section-header">
        <span class="section-icon">${iconMap[icon] || iconMap.folder}</span>
        <span>${title}</span>
        <span class="section-count">${entities.length}</span>
      </div>
      <div class="section-content">
        ${entitiesHtml}
      </div>
    </div>
  `;
}

function highlightMatch(text, query) {
  if (!query) return escapeHtml(text);
  const safe = escapeHtml(text);
  const regex = new RegExp(`(${escapeRegex(query)})`, 'gi');
  return safe.replace(regex, '<span class="highlight">$1</span>');
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

async function openEntity(entityName) {
  // Add to recent
  await StorageManager.addRecent(entityName);

  // Store the entity to browse
  await StorageManager.set('browseEntity', entityName);

  // Open browser in new tab
  chrome.tabs.create({
    url: chrome.runtime.getURL(`browser/browser.html?entity=${encodeURIComponent(entityName)}`)
  });

  // Close popup
  window.close();
}

async function toggleFavorite(entityName) {
  const isFav = favorites.includes(entityName);

  if (isFav) {
    await StorageManager.removeFavorite(entityName);
    favorites = favorites.filter(f => f !== entityName);
  } else {
    await StorageManager.addFavorite(entityName);
    favorites.unshift(entityName);
  }

  renderEntities();
}

function showLoading() {
  entityList.innerHTML = `
    <div class="loading">
      <div class="spinner"></div>
      <span>Loading entities...</span>
    </div>
  `;
  hideError();
}

function showError(message) {
  errorMessage.textContent = message;
  errorPanel.classList.remove('hidden');

  entityList.innerHTML = `
    <div class="empty-state">
      <div class="icon">&#9888;</div>
      <span class="title">Connection Error</span>
      <span class="subtitle">${escapeHtml(message)}</span>
    </div>
  `;
}

function hideError() {
  errorPanel.classList.add('hidden');
}

// Event Listeners
entitySearch.addEventListener('input', (e) => {
  searchQuery = e.target.value;
  renderEntities();
});

entitySearch.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    // Open first result
    const firstEntity = entityList.querySelector('.entity-item');
    if (firstEntity) {
      openEntity(firstEntity.dataset.entity);
    }
  }
});

refreshBtn.addEventListener('click', () => {
  loadEntities(true);
});

retryBtn.addEventListener('click', () => {
  init();
});

function showManualUrlInput() {
  entityList.innerHTML = `
    <div class="empty-state" style="padding:16px">
      <div class="icon">&#9888;</div>
      <span class="title">D365 environment not detected</span>
      <span class="subtitle" style="margin-bottom:12px">Navigate to D365 first, or enter the URL manually for government/custom domains.</span>
      <input type="text" id="manualUrlInput" class="search-input" placeholder="https://your-org.operations.dynamics.us" style="margin-bottom:8px;font-size:12px">
      <div style="display:flex;gap:8px;width:100%">
        <button id="manualUrlBtn" class="btn-small" style="flex:1;padding:6px;background:#0078d4;color:#fff;border:none;border-radius:4px;cursor:pointer">Connect</button>
        <button id="manualRetryBtn" class="btn-small" style="flex:1;padding:6px;background:#f3f3f3;border:1px solid #d1d1d1;border-radius:4px;cursor:pointer">Retry auto-detect</button>
      </div>
    </div>
  `;

  document.getElementById('manualUrlBtn')?.addEventListener('click', async () => {
    const url = document.getElementById('manualUrlInput')?.value?.trim();
    if (!url) return;
    try {
      const env = await odataClient.setManualEnvironment(url);
      envBadge.textContent = env.envType;
      envBadge.className = `env-badge ${env.envType.toLowerCase()}`;
      await StorageManager.addEnvironment(env);
      await loadEntities();
    } catch (err) {
      showError(err.message);
    }
  });

  document.getElementById('manualRetryBtn')?.addEventListener('click', () => init());

  // Allow Enter key in the URL input
  document.getElementById('manualUrlInput')?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') document.getElementById('manualUrlBtn')?.click();
  });
}

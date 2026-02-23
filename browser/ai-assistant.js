/**
 * D365 AI Data Manager - AI Query Assistant
 * Agentic LLM that reads DOM state and generates JS to drive the extension.
 * Supports Gemini, OpenAI, Anthropic, and custom endpoints.
 * Floating draggable chat window.
 */

const AIAssistant = {
  messages: [],
  isProcessing: false,
  aborted: false,

  // ==================== PANEL ====================

  togglePanel() {
    const panel = document.getElementById('aiPanel');
    if (!panel) return;

    if (!AISettings.isConfigured()) {
      showToast('Configure AI settings first (Settings > AI Features)');
      showSettingsModal();
      return;
    }

    panel.classList.toggle('hidden');
    if (!panel.classList.contains('hidden')) {
      document.getElementById('aiInput')?.focus();
    }
  },

  // ==================== DRAG LOGIC ====================

  initDrag() {
    const panel = document.getElementById('aiPanel');
    const header = document.getElementById('aiPanelHeader');
    if (!panel || !header) return;

    let isDragging = false;
    let offsetX = 0;
    let offsetY = 0;

    header.addEventListener('mousedown', (e) => {
      if (e.target.tagName === 'BUTTON') return;
      isDragging = true;
      const rect = panel.getBoundingClientRect();
      offsetX = e.clientX - rect.left;
      offsetY = e.clientY - rect.top;
      panel.style.transition = 'none';
    });

    document.addEventListener('mousemove', (e) => {
      if (!isDragging) return;
      e.preventDefault();
      const x = Math.max(0, Math.min(e.clientX - offsetX, window.innerWidth - 100));
      const y = Math.max(0, Math.min(e.clientY - offsetY, window.innerHeight - 100));
      panel.style.left = x + 'px';
      panel.style.top = y + 'px';
      panel.style.right = 'auto';
      panel.style.bottom = 'auto';
    });

    document.addEventListener('mouseup', () => {
      isDragging = false;
      if (panel) panel.style.transition = '';
    });
  },

  // ==================== DOM SNAPSHOT ====================

  snapshot() {
    const headers = [...document.querySelectorAll('#gridHeader th[data-column]')]
      .map(th => th.dataset.column);

    // Use actual loaded data[] instead of just DOM rows — AI sees what's really loaded
    const rows = (data || []).slice(0, 50).map(row =>
      headers.map(h => {
        const v = row[h];
        return v === null || v === undefined ? '' : String(v);
      })
    );

    const total = totalCount || 0;
    const pages = pageSize > 0 ? Math.ceil(total / pageSize) : 1;

    return {
      entity: currentEntity || null,
      totalCount: total,
      displayedRows: data ? data.length : 0,
      totalPages: pages,
      headers: headers,
      sampleRows: rows,
      visibleColumns: visibleColumns || [],
      filters: filterConfig || [],
      sort: sortConfig || {},
      pageSize: pageSize,
      currentPage: currentPage,
      hasJoin: typeof activeJoin !== 'undefined' && activeJoin && activeJoin.targetEntity ? true : false,
      joinInfo: typeof activeJoin !== 'undefined' && activeJoin ? activeJoin : null
    };
  },

  // ==================== SYSTEM PROMPT ====================

  buildSystemPrompt(snap) {
    return `You are an AI assistant for the D365 AI Data Manager extension. You help users query, filter, join, and analyze D365 Finance & Operations data entities.

KEY CAPABILITIES: You can switch entities, filter, sort, JOIN entities together, compare entities to find join fields, expand related entities, highlight cells/rows, and run analytics (summarize, stats, distinct, crossTab). JOINING IS A CORE FEATURE — when users ask about data across two entities, use compareEntities() then joinEntity().

CURRENT STATE:
- Entity: ${snap.entity || 'none loaded'}
- Total records: ${snap.totalCount}
- Displayed rows: ${snap.displayedRows}
- Page: ${snap.currentPage} of ${snap.totalPages} (page size: ${snap.pageSize})
- Columns: ${snap.headers.join(', ') || 'none'}
- Active filters: ${snap.filters.length > 0 ? JSON.stringify(snap.filters) : 'none'}
- Sort: ${snap.sort.field ? snap.sort.field + ' ' + snap.sort.direction : 'none'}
- Join: ${snap.hasJoin && snap.joinInfo ? `YES — ${snap.joinInfo.targetEntity} on ${snap.joinInfo.currentField}=${snap.joinInfo.targetField}${snap.joinInfo.innerOnly ? ' (inner)' : ' (left)'}` : 'none (use compareEntities + joinEntity to join)'}

SAMPLE DATA (first ${snap.sampleRows.length} rows):
${snap.headers.join(' | ')}
${snap.sampleRows.map(r => r.join(' | ')).join('\n')}

AVAILABLE ACTIONS (use \`\`\`js code blocks):
- loadEntity(entityName) — switch to ANY entity. IMPORTANT: Only pass entity names that came from searchEntities() results or that you see in the current state. loadEntity will FAIL and throw an error if the entity doesn't exist. NEVER guess or invent entity names.
- addFilter(field, operator, value, logic) — operators: eq, ne, contains, startswith, endswith, gt, lt, ge, le, null, notnull. logic: 'and' (default) or 'or'. CRITICAL: When filtering the SAME field for MULTIPLE values (e.g. "show PersonnelNumber X, Y, Z"), you MUST use 'or' logic: addFilter('Field', 'eq', 'val1'); addFilter('Field', 'eq', 'val2', 'or'); addFilter('Field', 'eq', 'val3', 'or'). Using AND for same-field eq returns NOTHING because a field cannot equal two values at once. Only the FIRST filter for a field uses 'and', all subsequent same-field filters MUST use 'or'. Null handling: addFilter('Field', 'null') for IS NULL, addFilter('Field', 'notnull') for IS NOT NULL. Do NOT pass a value for null/notnull operators.
- clearAllFilters() — remove all filters
- sortByColumn(field) — toggle sort
- exportData(format) — csv, excel, json, sql
- joinEntity(targetEntity, currentField, targetField, innerOnly) — join current entity with another. innerOnly=true to ONLY show matched rows (inner join), false/omit for left join (keep all). Example: joinEntity('ReleasedProductsV2', 'ItemNumber', 'ItemNumber', true) — only shows PO lines that exist in released products.
- goToPage(pageNumber) — navigate pages
- summarizeData(field) — group by field, count occurrences, top 30. For "which X has the most Y".
- computeStats(field) — min, max, sum, avg, median for numeric fields.
- getDistinctValues(field) — list unique values of a field.
- crossTab(field1, field2) — cross-tabulate two fields.
- setPageSize(size) — change rows per page (max 5000) and reload. Load more data before analyzing.
- highlightCells(field, operator, value, color) — highlight specific cells matching a condition. Colors: red, green, yellow, blue, orange, purple. Example: highlightCells('Amount', 'gt', '1000', 'red') highlights amounts over 1000 in red.
- highlightRows(field, operator, value, color) — highlight entire rows matching a condition. Same operators and colors as highlightCells.
- clearHighlights() — remove all highlights.
- compareEntities(targetEntity) — BEFORE joining, compare current entity with a target entity. Shows: exact column name matches, columns with overlapping values, and best join candidates with ready-to-use joinEntity() calls. ALWAYS use this before joinEntity() when you're unsure which fields to join on.
- setVisibleColumns(['col1', 'col2', ...]) — show only specific columns. Works with base columns AND joined/expanded columns (e.g. 'JoinedEntity.FieldName'). After a join or expand, use this to show only the relevant columns from both entities. PROACTIVELY call this after joins/expands to declutter the grid.
- searchEntities(keyword) — search the entity list by keyword. Returns matching entity names. You MUST call this BEFORE loadEntity() — it is your ONLY way to discover valid entity names. Entity names may be singular or plural (e.g. CustomersV3 vs CompanyInfoEntity). After getting results, call loadEntity() with an EXACT name from the search results.
- getRelatedEntities() — list all navigation properties (OData $expand relationships) for the current entity. Shows related entity names, whether they're collections or single references. ALWAYS call this in its own code block — results come back asynchronously.
- expandEntity(navPropertyName) — expand a related entity using OData $expand (server-side JOIN). Pass the EXACT navigation property name from getRelatedEntities() results. Can pass a string or array of strings. After expansion, new columns appear as "NavPropertyName.FieldName" (e.g. if you expand "FormulaLines", columns appear as "FormulaLines.ItemNumber", "FormulaLines.Quantity", etc.).
- clearExpand() — remove all $expand expansions.

D365 DATA MODEL RULES:
- Line entities (e.g. PurchaseOrderLinesV2, SalesOrderLines) only contain rows WHERE lines exist. A PO with zero lines has NO rows in PurchaseOrderLinesV2.
- To find POs with no lines, you must query the HEADER entity (PurchaseOrderHeadersV2) and compare.
- You CAN switch to different entities using loadEntity(). Do this freely when the user's question requires a different entity.
- Common header-line pairs: PurchaseOrderHeadersV2/PurchaseOrderLinesV2, SalesOrderHeadersV2/SalesOrderLinesV2, VendorInvoiceHeadersV2/VendorInvoiceLines
- KNOWN ENTITIES you can loadEntity() directly WITHOUT searching: CustomersV3, VendorsV2, ReleasedProductsV2, PurchaseOrderHeadersV2, PurchaseOrderLinesV2, SalesOrderHeadersV2, SalesOrderLinesV2, VendorInvoiceHeadersV2, VendorInvoiceLines, InventOnHandV2, GeneralJournalAccountEntryV2, MainAccountsV2, WorkersV2, LegalEntities, OperatingUnits, WareHousesV2. For ANY OTHER entity, use searchEntities() first.
- Entity names are PascalCase, usually ending in V2 or V3 (e.g. CustomersV3, VendorsV2, PurchaseOrderHeadersV2). These V2/V3 entities are the PRIMARY data entities — always prefer them over internal table names (like CustTable, VendTable).
- KEY PRODUCT FIELDS: ReleasedProductsV2 has DefaultOrderType (production/purchase/transfer/none), ProductType (Item/Service), ItemModelGroupId, BOMType. To distinguish manufactured vs purchased items, filter DefaultOrderType.

RULES:
1. Be concise. Explain briefly what you're doing, then act.
2. Use \`\`\`js code blocks for actions. They get auto-executed. IMPORTANT: Put each logical step in its OWN SEPARATE code block. Do NOT chain multiple actions in one block. Each block runs sequentially, and the grid refreshes between blocks so the user sees each step happen. Example — do THIS:
\`\`\`js
clearAllFilters()
\`\`\`
\`\`\`js
addFilter('Status', 'eq', 'Open')
\`\`\`
\`\`\`js
sortByColumn('Amount')
\`\`\`
NOT this: clearAllFilters(); addFilter('Status', 'eq', 'Open'); sortByColumn('Amount') all in one block.
Exception: Multiple addFilter() calls for the same logical filter group (e.g. OR conditions on the same field) CAN go in one block.
3. You CAN switch entities, add filters, sort, highlight — do it. Act decisively.
4. After entity switch or join, you'll get updated state automatically.
5. NEVER call fetch() directly. Use only the functions above.
6. ENTITY SEARCH PROTOCOL — follow this EXACTLY:
   a) FIRST call searchEntities('keyword') in its own code block. Do NOT combine with loadEntity.
   b) WAIT for the search results. Read the entity names returned.
   c) PICK THE BEST MATCH YOURSELF and call loadEntity() with it. NEVER ask the user "which entity do you want?" or "which is most relevant?" — YOU are the data expert, YOU decide. Prefer V2/V3 entities (e.g. CustomersV3 over CustTable). If multiple candidates exist, pick the one most relevant to the user's question and go. If it turns out wrong, you can always switch.
   d) If search returns 0 results, try 2-3 different shorter keywords (e.g. 'formula' → 'bom' → 'recipe').
   e) If ALL searches return 0 results, tell the user: "No matching entities found. This data may not exist in this environment."
   f) NEVER call loadEntity() with a name you made up or guessed. The entity list is finite — if searchEntities doesn't find it, it doesn't exist.
7. For "how many" or "which X has the most Y", use summarizeData(). For numeric analysis, use computeStats().
8. If only 100 rows are loaded but you need more for accurate analysis, call setPageSize(1000) first.
9. NEVER use exportData() to investigate or analyze — you CANNOT see exported files. You already get sample data in every state update. The user knows how to export themselves — only export when they explicitly ask.
10. To show ONLY rows that match a join (e.g. "hide POs not in released products"), use innerOnly=true: joinEntity('Entity', 'field', 'field', true). Do NOT try to filter on joined columns to achieve this.
11. BEFORE joining, ALWAYS call compareEntities(targetEntity) first to discover which columns have matching values. Never guess join fields — compare first, then use the best candidate from the results.
12. To show/hide columns, use setVisibleColumns(['col1', 'col2']). Pass exact column names from the Columns list. After joins, include joined columns as 'Entity.Field' (e.g. ['OrderNumber', 'Amount', 'ReleasedProducts.ItemName']). After expands, include expanded columns as 'NavProperty.Field'. ALWAYS call setVisibleColumns after a join or expand to show only relevant columns — don't leave the user staring at 50+ columns.
13. SHOW YOUR EVIDENCE: When you make a claim about the data (e.g. "there are 5 overdue invoices", "these vendors have no activity", "Amount exceeds threshold"), you MUST highlight the relevant rows or cells on screen so the user can SEE what you're referring to. Use highlightRows() to mark entire rows you're talking about, or highlightCells() to mark specific values. Always call clearHighlights() first to remove previous highlights. Use color to convey meaning: red for problems/issues/overdue, green for good/matches, yellow for warnings/attention, blue for informational. The user is looking at the screen — if you claim something, it must be visible and highlighted.
14. RELATED ENTITIES PROTOCOL — to get data from related/child entities:
   a) FIRST call getRelatedEntities() in its own code block to see what navigation properties exist.
   b) Read the results — they list nav property names, related entity names, and whether they're collections or single references.
   c) THEN call expandEntity('NavPropertyName') with an EXACT name from the results.
   d) After expansion, related data columns appear as "NavPropertyName.FieldName" in the grid. You can filter and highlight these columns.
   e) This is a SERVER-SIDE join ($expand) — much more efficient than client-side joinEntity() for parent-child relationships.
   f) Use joinEntity() for cross-entity joins (e.g. PO lines + Products). Use expandEntity() for built-in relationships (e.g. header → lines, entity → related lookup).
15. NO LOOPS — Do NOT repeat the same failed approach. If an action fails or returns no useful data:
   a) Try a DIFFERENT approach (different entity, different field, expand instead of join, etc.)
   b) NEVER call the same function with the same arguments twice.
   c) If you've tried 3+ approaches without success, tell the user what you tried and what didn't work — do NOT keep trying.
   d) When you see "[Step X of Y]" in feedback, plan accordingly. If you're past step 5, wrap up and present whatever you've found.${AISettings.settings?.customPrompt ? `\n\nUSER CUSTOM INSTRUCTIONS:\n${AISettings.settings.customPrompt}` : ''}`;
  },

  // ==================== MESSAGING ====================

  addMessage(role, content) {
    this.messages.push({ role, content });
    this.renderMessages();
  },

  renderMessages() {
    const container = document.getElementById('aiMessages');
    if (!container) return;

    let html = '<div class="ai-message system">Ask me about your data. I can filter, sort, join, analyze, highlight, and export.</div>';

    this.messages.forEach(msg => {
      const cls = msg.role === 'user' ? 'user' : 'assistant';
      html += `<div class="ai-message ${cls}">${this.formatContent(msg.content)}</div>`;
    });

    container.innerHTML = html;
    container.scrollTop = container.scrollHeight;
  },

  escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  },

  formatContent(text) {
    let safe = this.escapeHtml(text);
    // Replace code blocks with real execution results if available
    let blockIdx = 0;
    const results = this._lastBlockResults || [];
    safe = safe.replace(/```(?:js|javascript)?\s*\n[\s\S]*?```/g, () => {
      const r = results[blockIdx++];
      if (!r) return '<span class="ai-action-tag">Action executed</span>';
      // Extract function names from code for display
      const fns = [...r.code.matchAll(/^(\w+)\s*\(/gm)].map(m => m[1]);
      const label = fns.length > 0 ? fns.join(', ') : 'code';
      if (r.errors.length > 0) {
        return `<span class="ai-action-tag ai-action-error" title="${this.escapeHtml(r.errors.join('; '))}">${label} — failed</span>`;
      }
      return `<span class="ai-action-tag">${label} — done</span>`;
    });
    this._lastBlockResults = null;
    return safe
      .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
      .replace(/`(.*?)`/g, '<code>$1</code>')
      .replace(/\n/g, '<br>');
  },

  // ==================== SEND MESSAGE ====================

  async sendMessage() {
    const input = document.getElementById('aiInput');
    if (!input) return;

    const userMessage = input.value.trim();
    if (!userMessage || this.isProcessing) return;

    input.value = '';
    this.addMessage('user', userMessage);
    this.isProcessing = true;
    this.aborted = false;
    this.toggleStopButton(true);
    this.showThinking();

    try {
      await this.processMultiStep(userMessage);
    } catch (error) {
      console.error('AI Assistant error:', error);
      if (!this.aborted) {
        this.addMessage('assistant', 'Error: ' + error.message);
      }
    } finally {
      this.isProcessing = false;
      this.toggleStopButton(false);
      this.hideThinking();
    }
  },

  toggleStopButton(showStop) {
    const sendBtn = document.getElementById('aiSendBtn');
    const stopBtn = document.getElementById('aiStopBtn');
    if (sendBtn) sendBtn.classList.toggle('hidden', showStop);
    if (stopBtn) stopBtn.classList.toggle('hidden', !showStop);
  },

  abort() {
    this.aborted = true;
    // Tell service worker to abort the in-flight HTTP request
    chrome.runtime.sendMessage({ action: 'abortAiCall' }).catch(() => {});
    this.hideThinking();
    this.addMessage('assistant', 'Stopped by user.');
  },

  showThinking() {
    const container = document.getElementById('aiMessages');
    if (!container) return;
    const el = document.createElement('div');
    el.className = 'ai-message assistant thinking';
    el.id = 'aiThinking';
    el.textContent = 'Thinking...';
    container.appendChild(el);
    container.scrollTop = container.scrollHeight;
  },

  hideThinking() {
    document.getElementById('aiThinking')?.remove();
  },

  // ==================== MULTI-STEP PROCESSING ====================

  async processMultiStep(userMessage) {
    const snap = this.snapshot();
    const maxSteps = 10;

    const llmMessages = [
      { role: 'system', content: this.buildSystemPrompt(snap) },
      ...this.messages.slice(-16)
    ];

    // Track action history to detect loops
    const actionHistory = [];

    for (let step = 0; step < maxSteps; step++) {
      if (this.aborted) break;

      // Trim intermediate messages if context is growing too large within this turn
      if (llmMessages.length > 25) {
        const sysMsg = llmMessages[0];
        llmMessages.splice(0, llmMessages.length, sysMsg, ...llmMessages.slice(-12));
      }

      const response = await this.callLLM(llmMessages);
      if (!response || this.aborted) break;

      const jsBlocks = [...response.matchAll(/```(?:js|javascript)?\s*\n([\s\S]*?)```/g)];
      let triggeredNavigation = false;
      let triggeredAnalysis = false;

      const dataChanging = ['loadEntity', 'joinEntity', 'executeJoin', 'loadData', 'addFilter', 'clearAllFilters', 'sortByColumn', 'goToPage', 'setPageSize', 'setVisibleColumns', 'expandEntity', 'clearExpand'];
      const analysisFns = ['summarizeData', 'computeStats', 'getDistinctValues', 'crossTab', 'compareEntities', 'searchEntities', 'getRelatedEntities'];

      window.aiAnalysisResult = null;

      // Confirmation dialog if auto-execute is off
      let blocksToRun = jsBlocks;
      if (jsBlocks.length > 0 && AISettings.settings && AISettings.settings.autoExecute === false) {
        const decision = await this.showConfirmationDialog(jsBlocks);
        if (decision === 'skip') {
          this.addMessage('assistant', response);
          break;
        }
        if (Array.isArray(decision)) {
          blocksToRun = decision;
        }
        // 'run' means run all
      }

      // Track what actions are being taken this step
      const stepActions = blocksToRun.map(b => b[1].trim()).join(';');
      actionHistory.push(stepActions);

      const allErrors = [];
      const blockResults = []; // Track per-block results for display
      for (let bi = 0; bi < blocksToRun.length; bi++) {
        const block = blocksToRun[bi];
        if (this.aborted) break;
        try {
          const result = await this.safeExecute(block[1]);
          blockResults.push({ code: block[1].trim(), executed: result.executed, errors: result.errors });
          if (result.errors.length > 0) allErrors.push(...result.errors);
          if (dataChanging.some(fn => block[1].includes(fn))) triggeredNavigation = true;
          if (analysisFns.some(fn => block[1].includes(fn))) triggeredAnalysis = true;
        } catch (e) {
          console.warn('AI script error:', e.message);
          blockResults.push({ code: block[1].trim(), executed: 0, errors: [e.message] });
          allErrors.push(e.message);
        }
        // Let browser paint between blocks so user sees each step
        if (bi < blocksToRun.length - 1) {
          await new Promise(r => setTimeout(r, 150));
        }
      }
      // Store block results so formatContent can use real execution info
      this._lastBlockResults = blockResults;

      if (this.aborted) break;

      const errorFeedback = allErrors.length > 0
        ? `\n\nERRORS from your actions:\n${allErrors.map(e => `- ${e}`).join('\n')}\nFix these errors — try a DIFFERENT approach (different entity, different field, different method).`
        : '';

      // Collect analysis results (searchEntities, compareEntities, etc.)
      const analysisOutput = window.aiAnalysisResult ? `\n\nAnalysis result:\n${window.aiAnalysisResult}` : '';
      window.aiAnalysisResult = null;

      // Step counter for the AI to see
      const stepsRemaining = maxSteps - step - 1;
      const stepTag = `[Step ${step + 1} of ${maxSteps}${stepsRemaining <= 3 ? ` — ${stepsRemaining} steps left, WRAP UP soon` : ''}]`;

      // Detect repeated actions (loop detection)
      let loopWarning = '';
      if (actionHistory.length >= 3) {
        const last3 = actionHistory.slice(-3);
        // Check if the same action pattern repeats
        if (last3[0] === last3[1] && last3[1] === last3[2]) {
          loopWarning = '\n\nWARNING: You are repeating the same action 3 times. STOP this approach. Either try something completely different or tell the user what you found so far.';
        }
        // Check for similar patterns (same function names even with different args)
        const fnPattern = (s) => [...s.matchAll(/(\w+)\s*\(/g)].map(m => m[1]).join(',');
        const patterns = last3.map(fnPattern);
        if (patterns[0] === patterns[1] && patterns[1] === patterns[2] && !loopWarning) {
          loopWarning = '\n\nWARNING: You keep calling the same functions repeatedly. Try a different approach or present your findings to the user.';
        }
      }

      if (triggeredNavigation) {
        this.addMessage('assistant', response);
        // Wait for the grid to render with new data before taking snapshot
        await this.waitForTableLoad();
        // Let user see the result before next LLM call
        await new Promise(r => setTimeout(r, 400));
        const newSnap = this.snapshot();

        // Rebuild system prompt with fresh state so LLM has accurate context
        llmMessages[0] = { role: 'system', content: this.buildSystemPrompt(newSnap) };

        llmMessages.push({ role: 'assistant', content: response });
        llmMessages.push({
          role: 'user',
          content: `${stepTag} [Data updated. Entity: ${newSnap.entity}, ${newSnap.totalCount} total records, showing ${newSnap.displayedRows} rows, page ${newSnap.currentPage}/${newSnap.totalPages}. Columns: ${newSnap.headers.join(', ')}. Sample:\n${newSnap.headers.join(' | ')}\n${newSnap.sampleRows.slice(0, 10).map(r => r.join(' | ')).join('\n')}${analysisOutput}${errorFeedback}${loopWarning}]\nNow analyze what you see and continue.`
        });
        continue;
      }

      // If there were errors but no navigation triggered, still feed them back so AI can retry
      if (allErrors.length > 0 && !triggeredAnalysis) {
        this.addMessage('assistant', response);
        await new Promise(r => setTimeout(r, 400));
        llmMessages.push({ role: 'assistant', content: response });
        llmMessages.push({
          role: 'user',
          content: `${stepTag} [${analysisOutput ? analysisOutput.trim() + '\n\n' : ''}${errorFeedback.trim()}${loopWarning}]\nFix the issue — try a DIFFERENT approach.`
        });
        continue;
      }

      if (triggeredAnalysis && analysisOutput) {
        this.addMessage('assistant', response);
        await new Promise(r => setTimeout(r, 400));
        llmMessages.push({ role: 'assistant', content: response });
        llmMessages.push({
          role: 'user',
          content: `${stepTag} [${analysisOutput.trim()}${errorFeedback}${loopWarning}]\nUse these results to continue answering the user's question. Take action — load an entity, apply filters, etc.`
        });
        continue;
      }

      this.addMessage('assistant', response);
      break;
    }
  },

  showConfirmationDialog(jsBlocks) {
    return new Promise(resolve => {
      const dialog = document.getElementById('aiConfirmDialog');
      const list = document.getElementById('aiConfirmList');
      if (!dialog || !list) { resolve('run'); return; }

      list.innerHTML = jsBlocks.map((block, i) => {
        const code = block[1].trim().substring(0, 200);
        const escaped = this.escapeHtml(code);
        return `<label class="ai-confirm-item">
          <input type="checkbox" checked data-idx="${i}">
          <code>${escaped}</code>
        </label>`;
      }).join('');

      dialog.classList.remove('hidden');

      const cleanup = () => {
        dialog.classList.add('hidden');
        document.getElementById('aiConfirmRunAll')?.removeEventListener('click', onRunAll);
        document.getElementById('aiConfirmRunSelected')?.removeEventListener('click', onRunSelected);
        document.getElementById('aiConfirmSkip')?.removeEventListener('click', onSkip);
      };

      const onRunAll = () => { cleanup(); resolve('run'); };
      const onRunSelected = () => {
        const checked = [...list.querySelectorAll('input[type="checkbox"]:checked')]
          .map(cb => jsBlocks[parseInt(cb.dataset.idx)]);
        cleanup();
        resolve(checked.length > 0 ? checked : 'skip');
      };
      const onSkip = () => { cleanup(); resolve('skip'); };

      document.getElementById('aiConfirmRunAll')?.addEventListener('click', onRunAll);
      document.getElementById('aiConfirmRunSelected')?.addEventListener('click', onRunSelected);
      document.getElementById('aiConfirmSkip')?.addEventListener('click', onSkip);
    });
  },

  // Whitelist of safe functions the AI may call
  SAFE_FUNCTIONS: {
    loadEntity: true,
    addFilter: true,
    clearAllFilters: true,
    sortByColumn: true,
    exportData: true,
    showManualJoinModal: true,
    joinEntity: true,
    goToPage: true,
    loadData: true,
    executeJoin: true,
    highlightCells: true,
    highlightRows: true,
    clearHighlights: true,
    summarizeData: true,
    computeStats: true,
    getDistinctValues: true,
    crossTab: true,
    setPageSize: true,
    compareEntities: true,
    setVisibleColumns: true,
    searchEntities: true,
    getRelatedEntities: true,
    expandEntity: true,
    clearExpand: true
  },

  async safeExecute(code) {
    const statements = code.split(/[;\n]/).map(s => s.trim()).filter(Boolean);
    let executed = 0;
    const errors = [];

    // Defer loadData() — flush between different operation types for sequential visual feedback
    window._aiDeferLoadData = true;
    window._aiLoadDataNeeded = false;
    let prevFnName = null;

    // Helper: flush deferred loadData and yield to browser
    const flushAndYield = async () => {
      if (window._aiLoadDataNeeded) {
        window._aiDeferLoadData = false;
        window._aiLoadDataNeeded = false;
        try { await loadData(); } catch (e) { errors.push('loadData() failed: ' + e.message); }
        if (window.aiLastError) { errors.push(window.aiLastError); window.aiLastError = null; }
        window._aiDeferLoadData = true;
        // Let browser paint so user sees the update
        await new Promise(r => setTimeout(r, 120));
      }
    };

    for (const stmt of statements) {
      // Skip comments
      if (stmt.startsWith('//')) continue;

      const match = stmt.match(/^(\w+)\s*\(([\s\S]*)\)$/);
      if (!match) {
        showToast(`AI: skipped "${stmt.substring(0, 40)}"`, 'warn');
        continue;
      }

      const fnName = match[1];
      const argsStr = match[2].trim();

      if (!this.SAFE_FUNCTIONS[fnName]) {
        showToast(`AI: blocked ${fnName}()`, 'warn');
        errors.push(`${fnName}() is not an allowed function`);
        continue;
      }

      if (typeof window[fnName] !== 'function') {
        showToast(`AI: ${fnName}() not found`, 'error');
        errors.push(`${fnName}() does not exist`);
        continue;
      }

      // Flush deferred loadData when switching to a different operation type
      // This batches consecutive addFilter() calls but flushes between different ops
      if (prevFnName && fnName !== prevFnName) {
        await flushAndYield();
      }

      try {
        window.aiLastError = null;
        // LLMs output single quotes (valid JS) but JSON.parse needs double quotes
        const jsonSafe = argsStr ? argsStr.replace(/'([^']*)'/g, '"$1"') : '';
        const args = jsonSafe ? JSON.parse(`[${jsonSafe}]`) : [];

        // Harden args: limit count, string length, block prototype pollution
        const MAX_ARGS = 10;
        const MAX_STR_LEN = 2000;
        if (args.length > MAX_ARGS) {
          errors.push(`${fnName}(): too many arguments (${args.length})`);
          continue;
        }
        let argBlocked = false;
        for (const arg of args) {
          if (typeof arg === 'string' && arg.length > MAX_STR_LEN) {
            errors.push(`${fnName}(): argument string too long (${arg.length})`);
            argBlocked = true; break;
          }
          if (arg !== null && typeof arg === 'object') {
            const keys = Object.keys(arg);
            if (keys.some(k => k === '__proto__' || k === 'constructor' || k === 'prototype')) {
              errors.push(`${fnName}(): suspicious argument blocked`);
              argBlocked = true; break;
            }
          }
        }
        if (argBlocked) continue;

        const result = window[fnName](...args);
        // Await async functions (loadEntity, joinEntity, loadData, etc.)
        if (result && typeof result.then === 'function') {
          await result;
        }
        // Check if loadData (or other async fn) set an error
        if (window.aiLastError) {
          errors.push(`${fnName}(${argsStr}) → query error: ${window.aiLastError}`);
        }
        executed++;
        prevFnName = fnName;
        // Discovery functions produce results the AI needs to see
        // before it can make informed decisions. Stop the block here
        // so results are fed back first.
        if (fnName === 'searchEntities' || fnName === 'compareEntities' || fnName === 'getRelatedEntities') {
          break;
        }
      } catch (e) {
        showToast(`AI: ${fnName}() failed — ${e.message}`, 'error');
        errors.push(`${fnName}(${argsStr}) threw: ${e.message}`);
        // If loadEntity or joinEntity failed, stop the block — everything after depends on it
        if (fnName === 'loadEntity' || fnName === 'joinEntity') {
          break;
        }
      }
    }

    // Flush any remaining deferred loadData
    window._aiDeferLoadData = false;
    if (window._aiLoadDataNeeded) {
      window._aiLoadDataNeeded = false;
      try {
        await loadData();
      } catch (e) {
        errors.push('loadData() failed: ' + e.message);
      }
      if (window.aiLastError) {
        errors.push(window.aiLastError);
        window.aiLastError = null;
      }
    }

    if (executed > 0) {
      showToast(`AI executed ${executed} action${executed > 1 ? 's' : ''}`);
    }

    return { executed, errors };
  },

  waitForTableLoad() {
    return new Promise(resolve => {
      let resolved = false;
      const observer = new MutationObserver(() => {
        if (!resolved) {
          resolved = true;
          observer.disconnect();
          setTimeout(resolve, 200); // small settle for paint
        }
      });
      const tbody = document.querySelector('#gridBody');
      if (tbody) {
        observer.observe(tbody, { childList: true });
      }
      setTimeout(() => {
        if (!resolved) {
          resolved = true;
          observer.disconnect();
          resolve();
        }
      }, 3000);
    });
  },

  // ==================== LLM API CALLS ====================

  async callLLM(messages) {
    if (this.aborted) return null;
    const settings = AISettings.settings;
    if (!settings || !settings.provider) throw new Error('AI not configured');

    switch (settings.provider) {
      case 'gemini': return this.callGemini(messages, settings);
      case 'openai': return this.callOpenAI(messages, settings);
      case 'anthropic': return this.callAnthropic(messages, settings);
      case 'openrouter': return this.callOpenRouter(messages, settings);
      case 'ollama': return this.callOllama(messages, settings);
      case 'custom': return this.callCustom(messages, settings);
      default: throw new Error('Unknown provider: ' + settings.provider);
    }
  },

  async callGemini(messages, settings) {
    const endpoint = AISettings.getApiEndpoint();

    const systemMsg = messages.find(m => m.role === 'system');
    const chatMessages = messages.filter(m => m.role !== 'system');

    const contents = chatMessages.map(m => ({
      role: m.role === 'assistant' ? 'model' : 'user',
      parts: [{ text: m.content }]
    }));

    const body = {
      contents: contents,
      systemInstruction: systemMsg ? { parts: [{ text: systemMsg.content }] } : undefined,
      generationConfig: { maxOutputTokens: 4096 }
    };

    const resp = await chrome.runtime.sendMessage({
      action: 'aiApiCall',
      url: endpoint,
      options: {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
      }
    });

    if (!resp || !resp.success) throw new Error(resp?.error || 'Gemini API call failed');

    return resp.data?.candidates?.[0]?.content?.parts?.[0]?.text || null;
  },

  async callOpenAI(messages, settings) {
    const isCustom = settings.provider === 'custom';
    const endpoint = isCustom
      ? settings.customEndpoint
      : 'https://api.openai.com/v1/chat/completions';

    const resp = await chrome.runtime.sendMessage({
      action: 'aiApiCall',
      url: endpoint,
      skipAllowlist: isCustom,
      options: {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${settings.apiKey}`
        },
        body: JSON.stringify({
          model: settings.model || 'gpt-4o-mini',
          messages: messages,
          max_tokens: 4096
        })
      }
    });

    if (!resp || !resp.success) throw new Error(resp?.error || 'OpenAI API call failed');

    return resp.data?.choices?.[0]?.message?.content || null;
  },

  async callAnthropic(messages, settings) {
    const systemMsg = messages.find(m => m.role === 'system');
    const chatMessages = messages.filter(m => m.role !== 'system');

    const resp = await chrome.runtime.sendMessage({
      action: 'aiApiCall',
      url: 'https://api.anthropic.com/v1/messages',
      options: {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': settings.apiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true'
        },
        body: JSON.stringify({
          model: settings.model || 'claude-sonnet-4-5-20250929',
          system: systemMsg?.content || '',
          messages: chatMessages.map(m => ({ role: m.role, content: m.content })),
          max_tokens: 4096
        })
      }
    });

    if (!resp || !resp.success) throw new Error(resp?.error || 'Anthropic API call failed');

    const content = resp.data?.content || [];
    return content.filter(b => b.type === 'text').map(b => b.text).join('\n') || null;
  },

  async callOpenRouter(messages, settings) {
    const resp = await chrome.runtime.sendMessage({
      action: 'aiApiCall',
      url: 'https://openrouter.ai/api/v1/chat/completions',
      options: {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${settings.apiKey}`,
          'HTTP-Referer': chrome.runtime.getURL('/'),
          'X-Title': 'D365 AI Data Manager'
        },
        body: JSON.stringify({
          model: settings.model || 'anthropic/claude-sonnet-4',
          messages: messages,
          max_tokens: 4096
        })
      }
    });

    if (!resp || !resp.success) throw new Error(resp?.error || 'OpenRouter API call failed');
    return resp.data?.choices?.[0]?.message?.content || null;
  },

  async callOllama(messages, settings) {
    const port = settings.ollamaPort || 11434;
    const resp = await chrome.runtime.sendMessage({
      action: 'aiApiCall',
      url: `http://localhost:${port}/v1/chat/completions`,
      skipAllowlist: true,
      options: {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: settings.model || 'llama3.3',
          messages: messages
        })
      }
    });

    if (!resp || !resp.success) throw new Error(resp?.error || 'Ollama API call failed');
    return resp.data?.choices?.[0]?.message?.content || null;
  },

  async callCustom(messages, settings) {
    const endpoint = settings.customEndpoint;
    if (!endpoint) throw new Error('Custom endpoint URL not configured');

    const format = settings.customFormat || 'openai';
    const model = settings.customModel || settings.model || '';

    if (format === 'anthropic') {
      // Anthropic-compatible format
      const systemMsg = messages.find(m => m.role === 'system');
      const chatMessages = messages.filter(m => m.role !== 'system');

      const resp = await chrome.runtime.sendMessage({
        action: 'aiApiCall',
        url: endpoint,
        skipAllowlist: true,
        options: {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'x-api-key': settings.apiKey,
            'anthropic-version': '2023-06-01'
          },
          body: JSON.stringify({
            model: model,
            system: systemMsg?.content || '',
            messages: chatMessages.map(m => ({ role: m.role, content: m.content })),
            max_tokens: 4096
          })
        }
      });

      if (!resp || !resp.success) throw new Error(resp?.error || 'Custom API call failed');

      const content = resp.data?.content || [];
      return content.filter(b => b.type === 'text').map(b => b.text).join('\n') || null;
    } else {
      // OpenAI-compatible format (default)
      const resp = await chrome.runtime.sendMessage({
        action: 'aiApiCall',
        url: endpoint,
        skipAllowlist: true,
        options: {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${settings.apiKey}`
          },
          body: JSON.stringify({
            model: model,
            messages: messages,
            max_tokens: 4096
          })
        }
      });

      if (!resp || !resp.success) throw new Error(resp?.error || 'Custom API call failed');

      return resp.data?.choices?.[0]?.message?.content || null;
    }
  }
};

// ==================== EVENT LISTENERS ====================

document.addEventListener('DOMContentLoaded', () => {
  // Init drag
  AIAssistant.initDrag();

  document.getElementById('aiSendBtn')?.addEventListener('click', () => AIAssistant.sendMessage());

  document.getElementById('aiStopBtn')?.addEventListener('click', () => AIAssistant.abort());

  document.getElementById('aiInput')?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      AIAssistant.sendMessage();
    }
  });

  document.getElementById('closeAIPanelBtn')?.addEventListener('click', () => {
    document.getElementById('aiPanel')?.classList.add('hidden');
  });
});

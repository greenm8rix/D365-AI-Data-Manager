/**
 * D365 AI Data Manager - Power Platform Export Tools
 * The value: entity + filters + columns + sort are already configured → one-click to use elsewhere.
 * Relies on globals from browser.js: currentEntity, odataClient, visibleColumns,
 * filterConfig, sortConfig, pageSize, buildFilterString, copyToClipboard, showToast, downloadFile, activeJoin
 */

const PowerTools = {
  // ==================== HELPERS ====================

  getBaseODataUrl() {
    if (!currentEntity || !odataClient.baseUrl) {
      throw new Error('No entity loaded or environment not detected');
    }
    return `${odataClient.baseUrl}${odataClient.odataPath}${currentEntity}`;
  },

  getCurrentFilter() {
    return typeof buildFilterString === 'function' ? buildFilterString() : '';
  },

  getCurrentSelect() {
    const cols = (visibleColumns || []).filter(col => !col.includes('.'));
    return cols.length > 0 ? cols.join(',') : '';
  },

  getCurrentOrderBy() {
    if (sortConfig && sortConfig.field && !sortConfig.field.includes('.')) {
      return `${sortConfig.field} ${sortConfig.direction || 'asc'}`;
    }
    return '';
  },

  buildFullODataUrl(entity, baseUrl) {
    const base = baseUrl || odataClient.baseUrl;
    const ent = entity || currentEntity;
    const url = `${base}${odataClient.odataPath}${ent}`;
    const filter = this.getCurrentFilter();
    const select = this.getCurrentSelect();
    const orderBy = this.getCurrentOrderBy();

    const params = ['cross-company=true'];
    if (select) params.push(`$select=${select}`);
    if (filter) params.push(`$filter=${encodeURIComponent(filter)}`);
    if (orderBy) params.push(`$orderby=${orderBy}`);
    params.push(`$top=${pageSize}`);
    params.push('$count=true');

    return `${url}?${params.join('&')}`;
  },

  _hasJoin() {
    return typeof activeJoin !== 'undefined' && activeJoin && activeJoin.targetEntity;
  },

  // ==================== ODATA URL ====================

  copyODataUrl() {
    const url = this.buildFullODataUrl();
    let text = url;

    if (this._hasJoin()) {
      const join = activeJoin;
      const targetUrl = `${odataClient.baseUrl}${odataClient.odataPath}${join.targetEntity}?cross-company=true&$count=true`;
      text += `\n\nJoined entity: ${targetUrl}\nJoin: ${currentEntity}.${join.currentField} = ${join.targetEntity}.${join.targetField}`;
    }

    copyToClipboard(text, 'OData URL copied!');
  },

  // ==================== POWER AUTOMATE ====================
  // D365 F&O connector in Power Automate takes these parameters directly.
  // User adds "List items" action → D365 Finance & Operations connector → pastes these values.

  generateConnectorSettings() {
    const filter = this.getCurrentFilter();
    const select = this.getCurrentSelect();
    const orderBy = this.getCurrentOrderBy();

    const lines = [];
    lines.push(`Entity name: ${currentEntity}`);
    lines.push(`Cross-company: Yes`);
    if (filter) lines.push(`$filter: ${filter}`);
    if (select) lines.push(`$select: ${select}`);
    if (orderBy) lines.push(`$orderby: ${orderBy}`);
    lines.push(`$top: ${pageSize}`);
    lines.push(`$count: true`);
    lines.push(`\nInstance URL: ${odataClient.baseUrl}`);

    if (this._hasJoin()) {
      const join = activeJoin;
      lines.push(`\n--- Second action for joined entity ---`);
      lines.push(`Entity name: ${join.targetEntity}`);
      lines.push(`Cross-company: Yes`);
      lines.push(`Join: ${currentEntity}.${join.currentField} = ${join.targetEntity}.${join.targetField}`);
    }

    return lines.join('\n');
  },

  copyConnectorSettings() {
    copyToClipboard(this.generateConnectorSettings(), 'D365 connector settings copied!');
  },

  // ==================== POWER BI ====================

  generatePbidsFile() {
    const connectionUrl = this.buildFullODataUrl();

    return JSON.stringify({
      version: "0.1",
      connections: [{
        details: {
          protocol: "odata",
          address: { url: connectionUrl }
        },
        mode: "Import"
      }]
    }, null, 2);
  },

  generateMQuery() {
    const baseUrl = this.getBaseODataUrl();
    const select = this.getCurrentSelect();
    const filter = this.getCurrentFilter();

    const params = ['cross-company=true'];
    if (select) params.push(`$select=${select}`);
    if (filter) params.push(`$filter=${encodeURIComponent(filter)}`);

    const fullUrl = `${baseUrl}?${params.join('&')}`;

    if (this._hasJoin()) {
      return this._generateJoinedMQuery();
    }

    return `let
    Source = OData.Feed(
        "${fullUrl}",
        null,
        [Implementation = "2.0", ODataVersion = 4]
    )
in
    Source`;
  },

  _generateJoinedMQuery() {
    const baseUrl = odataClient.baseUrl;
    const join = activeJoin;

    const mainUrl = `${baseUrl}${odataClient.odataPath}${currentEntity}?cross-company=true`;
    const targetUrl = `${baseUrl}${odataClient.odataPath}${join.targetEntity}?cross-company=true`;

    const joinKind = join.innerOnly ? 'JoinKind.Inner' : 'JoinKind.LeftOuter';
    const targetCols = (join.targetColumns || []).map(c => `"${c}"`).join(', ');

    return `let
    Main = OData.Feed("${mainUrl}", null, [Implementation="2.0", ODataVersion=4]),
    Target = OData.Feed("${targetUrl}", null, [Implementation="2.0", ODataVersion=4]),
    Merged = Table.NestedJoin(Main, {"${join.currentField}"}, Target, {"${join.targetField}"}, "Target_", ${joinKind}),
    Expanded = Table.ExpandTableColumn(Merged, "Target_", {${targetCols}})
in
    Expanded`;
  },

  downloadPbidsFile() {
    const pbids = this.generatePbidsFile();
    downloadFile(pbids, `${currentEntity}.pbids`, 'application/json');
    showToast('.pbids downloaded — open with Power BI Desktop');
  },

  copyMQuery() {
    copyToClipboard(this.generateMQuery(), 'M Query copied — paste into Power Query Advanced Editor');
  },

  // ==================== POWER AUTOMATE FLOW ACTIONS ====================
  // Generates JSON action blocks for Code View paste.
  // Chain: InitializeVariable → HTTP Fetch → SetVariable

  _generateEntityActions(entity, odataUrl) {
    const initName = `Init_arr${entity}`;
    const fetchName = `Fetch_${entity}`;
    const setName = `Set_arr${entity}`;

    return {
      [initName]: {
        type: 'InitializeVariable',
        inputs: {
          variables: [{ name: `arr${entity}`, type: 'array' }]
        },
        runAfter: {}
      },
      [fetchName]: {
        type: 'Http',
        inputs: {
          uri: odataUrl,
          method: 'GET',
          headers: {
            'Authorization': "Bearer @{variables('access_token')}",
            'Accept': 'application/json',
            'Prefer': 'odata.maxpagesize=5000'
          }
        },
        runAfter: { [initName]: ['Succeeded'] }
      },
      [setName]: {
        type: 'SetVariable',
        inputs: {
          name: `arr${entity}`,
          value: `@body('${fetchName}')?['value']`
        },
        runAfter: { [fetchName]: ['Succeeded'] }
      }
    };
  },

  generatePAFlowDefinition() {
    const mainUrl = this.buildFullODataUrl();
    const actions = this._generateEntityActions(currentEntity, mainUrl);

    if (this._hasJoin()) {
      const join = activeJoin;
      const targetUrl = `${odataClient.baseUrl}${odataClient.odataPath}${join.targetEntity}?cross-company=true&$count=true`;
      const targetActions = this._generateEntityActions(join.targetEntity, targetUrl);
      const targetInitKey = `Init_arr${join.targetEntity}`;
      const mainSetKey = `Set_arr${currentEntity}`;
      targetActions[targetInitKey].runAfter = { [mainSetKey]: ['Succeeded'] };
      Object.assign(actions, targetActions);
    }

    return {
      '$schema': 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
      contentVersion: '1.0.0.0',
      parameters: {},
      triggers: {
        manual: {
          type: 'Request',
          kind: 'Button',
          inputs: { schema: { type: 'object', properties: {}, required: [] } }
        }
      },
      actions
    };
  },

  downloadPAFlow() {
    const definition = this.generatePAFlowDefinition();
    const json = JSON.stringify(definition, null, 2);
    downloadFile(json, `D365_${currentEntity}_flow.json`, 'application/json');
    showToast('Flow definition downloaded — import via Power Platform CLI or API');
  },

  copyPAFlowActions() {
    const definition = this.generatePAFlowDefinition();
    const json = JSON.stringify(definition.actions, null, 2);
    copyToClipboard(json, 'Flow actions JSON copied');
  },

  // ==================== POWER FX (POWER APPS) ====================
  // Generates ClearCollect() formula for canvas app OnStart or Button.OnSelect

  _filterToPowerFx(filter) {
    const { field, operator, value } = filter;
    if (!field) return null;

    // Null checks
    if (operator === 'null') return `IsBlank(${field})`;
    if (operator === 'notnull') return `!IsBlank(${field})`;

    // Type-aware value formatting
    const fieldType = typeof getFieldType === 'function' ? getFieldType(field) : null;
    const numericTypes = ['Edm.Int16', 'Edm.Int32', 'Edm.Int64', 'Edm.Decimal', 'Edm.Double', 'Edm.Single', 'Edm.Byte'];
    const isNumeric = numericTypes.includes(fieldType);
    const isBool = fieldType === 'Edm.Boolean';

    const fmtVal = isBool ? (value === 'true' || value === '1' ? 'true' : 'false')
      : isNumeric ? value
      : `"${(value || '').replace(/"/g, '""')}"`;

    switch (operator) {
      case 'eq': return `${field} = ${fmtVal}`;
      case 'ne': return `${field} <> ${fmtVal}`;
      case 'gt': return `${field} > ${fmtVal}`;
      case 'ge': return `${field} >= ${fmtVal}`;
      case 'lt': return `${field} < ${fmtVal}`;
      case 'le': return `${field} <= ${fmtVal}`;
      case 'contains': return `${fmtVal} in ${field}`;
      case 'startswith': return `StartsWith(${field}, ${fmtVal})`;
      case 'endswith': return `EndsWith(${field}, ${fmtVal})`;
      default: return `${field} = ${fmtVal}`;
    }
  },

  generatePowerFxCollection() {
    const entity = currentEntity;

    // Build filter conditions from filterConfig
    const filters = (filterConfig || []).filter(f => f.field && !f.field.includes('.'));
    const pfxParts = [];
    for (let i = 0; i < filters.length; i++) {
      const cond = this._filterToPowerFx(filters[i]);
      if (!cond) continue;
      if (i === 0) {
        pfxParts.push(cond);
      } else {
        const logic = filters[i].logic || 'and';
        if (logic === 'or') {
          // Merge with previous using ||
          const prev = pfxParts.pop();
          pfxParts.push(`${prev} || ${cond}`);
        } else {
          pfxParts.push(cond);
        }
      }
    }

    // Innermost: entity reference or Filter()
    let expr = `'${entity}'`;
    if (pfxParts.length > 0) {
      const indent = '        ';
      expr = `Filter(\n${indent}'${entity}',\n${indent}${pfxParts.join(',\n' + indent)}\n    )`;
    }

    // Wrap with ShowColumns if columns selected
    const cols = (visibleColumns || []).filter(c => !c.includes('.'));
    if (cols.length > 0 && cols.length < 20) {
      const colList = cols.map(c => `"${c}"`).join(', ');
      expr = `ShowColumns(\n        ${expr},\n        ${colList}\n    )`;
    }

    // Wrap with SortByColumns if sort active
    if (sortConfig && sortConfig.field && !sortConfig.field.includes('.')) {
      const dir = sortConfig.direction === 'desc' ? 'SortOrder.Descending' : 'SortOrder.Ascending';
      expr = `SortByColumns(\n        ${expr},\n        "${sortConfig.field}", ${dir}\n    )`;
    }

    // Wrap with FirstN for row limit
    const limit = pageSize || 100;
    if (limit < 100000) {
      expr = `FirstN(\n        ${expr},\n        ${limit}\n    )`;
    }

    let formula = `ClearCollect(\n    D365Data,\n    ${expr}\n);`;

    // If join active, add second collection + merge
    if (this._hasJoin()) {
      const join = activeJoin;
      formula += `\n\nClearCollect(\n    D365Target,\n    '${join.targetEntity}'\n);`;
      formula += `\n\n// Merge: ${currentEntity}.${join.currentField} = ${join.targetEntity}.${join.targetField}`;
      formula += `\nClearCollect(\n    D365Joined,\n    AddColumns(\n        D365Data,\n        "Matched",\n        LookUp(D365Target, ${join.targetField} = D365Data[@${join.currentField}])\n    )\n);`;
    }

    return formula;
  },

  copyPowerFxCollection() {
    const fx = this.generatePowerFxCollection();
    copyToClipboard(fx, 'Power Fx copied — paste into App.OnStart or Button.OnSelect');
  }
};


// ==================== UI HANDLER ====================

const PowerToolsUI = {
  initIcons() {
    if (typeof SVGIcons === 'undefined') return;

    const iconMap = {
      'powerBtnIcon': SVGIcons.powerPlatform,
      'powerAutomateIcon': SVGIcons.powerAutomate,
      'powerAppsIcon': SVGIcons.powerApps,
      'powerBIIcon': SVGIcons.powerBI,
      'aiAssistantIcon': SVGIcons.aiAssistant,
      'aiAnalyzeIcon': SVGIcons.aiAnalyze
    };

    Object.entries(iconMap).forEach(([id, svg]) => {
      const el = document.getElementById(id);
      if (el) el.innerHTML = svg;
    });
  },

  handleAction(action) {
    if (!currentEntity) {
      showToast('Load an entity first');
      return;
    }

    try {
      switch (action) {
        case 'copy-odata-url': PowerTools.copyODataUrl(); break;
        case 'automate-connector': PowerTools.copyConnectorSettings(); break;
        case 'automate-flow-download': PowerTools.downloadPAFlow(); break;
        case 'automate-flow-json': PowerTools.copyPAFlowActions(); break;
        case 'powerapps-fx': PowerTools.copyPowerFxCollection(); break;
        case 'powerbi-pbids': PowerTools.downloadPbidsFile(); break;
        case 'powerbi-mquery': PowerTools.copyMQuery(); break;
        default:
          console.warn('Unknown power action:', action);
      }
    } catch (error) {
      showToast('Error: ' + error.message);
      console.error('Power Tools error:', error);
    }
  }
};

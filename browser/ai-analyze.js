/**
 * D365 AI Data Manager - Analyze with AI (inferenc.es integration)
 * Flow: confirm → export CSV → upload to inferenc.es → trigger analysis → check email
 */

const AIAnalyze = {
  cancelled: false,

  async analyze() {
    if (!currentEntity || !data || data.length === 0) {
      showToast('Load data first');
      return;
    }

    // Check if user has inferenc.es account linked
    const aiSettings = await StorageManager.getAISettings();

    if (!aiSettings.inferencesLinked || !aiSettings.inferencesSession) {
      this.promptLogin();
      return;
    }

    // Ask user to confirm sharing data
    const columns = Object.keys(data[0]).filter(k => !k.startsWith('@') && !k.startsWith('__'));
    const confirmed = confirm(
      `Send ${data.length} rows x ${columns.length} columns from "${currentEntity}" to inferenc.es for AI analysis?\n\n` +
      `Your data will be uploaded securely. You'll receive the analysis results by email.`
    );
    if (!confirmed) return;

    this.cancelled = false;

    // Get a valid (non-expired) token
    const token = await AISettings.getValidToken();
    if (!token) {
      showToast('Session expired. Please reconnect inferenc.es in Settings.');
      return;
    }

    // Start the animated flow
    this.showOverlay();
    this.setStep(1, 'Exporting your data...');

    try {
      // Step 1: Generate CSV
      const csvString = this.generateCsv();

      this.setStepDetail(1, `${currentEntity}.csv (${data.length} rows x ${columns.length} columns)`);
      this.setProgress(1, 100);

      if (this.cancelled) return this.hideOverlay();

      // Step 2: Upload to inferenc.es and trigger analysis
      this.setStep(2, 'Uploading to inferenc.es...');
      this.setProgress(2, 30);

      const filename = `${currentEntity}_${new Date().toISOString().slice(0, 10)}.csv`;

      const uploadResp = await chrome.runtime.sendMessage({
        action: 'inferencesUpload',
        csvContent: csvString,
        filename: filename,
        accessToken: token
      });

      if (this.cancelled) return this.hideOverlay();

      if (!uploadResp || !uploadResp.success) {
        throw new Error(uploadResp?.error || 'Upload failed');
      }

      this.setProgress(2, 100);

      // Step 3: Done — tell user to check email
      this.setStep(3, 'Analysis started!');
      this.setProgress(3, 100);
      this.setStatus('Check your email (or spam folder) for the results.');

      // Auto-hide after a few seconds
      setTimeout(() => {
        this.hideOverlay();
        showToast('Analysis submitted! Check your email for results.');
      }, 3000);

    } catch (error) {
      console.error('Analyze error:', error);
      this.hideOverlay();
      showToast('Upload failed: ' + error.message, 'error');
    }
  },

  promptLogin() {
    // Launch the connect flow directly instead of sending user to settings
    if (typeof AISettings !== 'undefined') {
      AISettings.connectInferences();
    } else {
      showToast('Link your inferenc.es account first: Settings > AI Features');
    }
  },

  generateCsv() {
    const columns = Object.keys(data[0]).filter(k => !k.startsWith('@') && !k.startsWith('__'));
    let csv = columns.map(col => {
      if (col.includes(',') || col.includes('"') || col.includes('\n')) {
        return '"' + col.replace(/"/g, '""') + '"';
      }
      return col;
    }).join(',') + '\n';

    data.forEach(row => {
      const values = columns.map(col => {
        let value = row[col];
        if (value === null || value === undefined) return '';
        value = String(value);
        if (value.includes(',') || value.includes('"') || value.includes('\n')) {
          value = '"' + value.replace(/"/g, '""') + '"';
        }
        return value;
      });
      csv += values.join(',') + '\n';
    });

    return csv;
  },

  // ==================== ANIMATION OVERLAY ====================

  showOverlay() {
    const overlay = document.getElementById('analyzeOverlay');
    if (overlay) {
      overlay.classList.remove('hidden');
      for (let i = 1; i <= 3; i++) {
        const step = document.getElementById(`analyzeStep${i}`);
        if (step) {
          step.classList.remove('active', 'done');
          step.classList.add('pending');
        }
      }
      document.getElementById('analyzeStatus').textContent = '';
    }
  },

  hideOverlay() {
    document.getElementById('analyzeOverlay')?.classList.add('hidden');
  },

  setStep(stepNum, text) {
    for (let i = 1; i < stepNum; i++) {
      const prev = document.getElementById(`analyzeStep${i}`);
      if (prev) {
        prev.classList.remove('active', 'pending');
        prev.classList.add('done');
      }
    }

    const step = document.getElementById(`analyzeStep${stepNum}`);
    if (step) {
      step.classList.remove('pending');
      step.classList.add('active');
    }

    const textEl = step?.querySelector('.step-text');
    if (textEl) textEl.textContent = text;
  },

  setStepDetail(stepNum, detail) {
    const el = document.getElementById(`analyzeDetail${stepNum}`);
    if (el) el.textContent = detail;
  },

  setProgress(stepNum, percent) {
    const fill = document.getElementById(`analyzeProg${stepNum}`);
    if (fill) fill.style.width = `${percent}%`;
  },

  setStatus(text) {
    const el = document.getElementById('analyzeStatus');
    if (el) el.textContent = text;
  }
};

// Cancel button listener
document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('analyzeCancelBtn')?.addEventListener('click', () => {
    AIAnalyze.cancelled = true;
    AIAnalyze.hideOverlay();
    showToast('Analysis cancelled');
  });
});

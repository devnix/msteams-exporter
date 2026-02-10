// Extension state
let exportedData = null;

// DOM elements
const exportBtn = document.getElementById('export-btn');
const exportCompleteBtn = document.getElementById('export-complete-btn');
const debugBtn = document.getElementById('debug-btn');
const downloadBtn = document.getElementById('download-btn');
const statusText = document.getElementById('status-text');
const statusIcon = document.querySelector('.status-icon');
const progressContainer = document.getElementById('progress-container');
const progressFill = document.getElementById('progress-fill');
const progressText = document.getElementById('progress-text');
const resultsContainer = document.getElementById('results');
const messageCount = document.getElementById('message-count');
const errorContainer = document.getElementById('error');
const errorText = document.getElementById('error-text');
const infoComplete = document.getElementById('info-complete');
const versionSpan = document.getElementById('version');

// Set version from manifest
versionSpan.textContent = 'v' + chrome.runtime.getManifest().version;

// Update status
function updateStatus(text, icon = '⚪', type = 'default') {
  statusText.textContent = text;
  statusIcon.textContent = icon;
}

// Show error
function showError(message) {
  errorText.textContent = message;
  errorContainer.classList.remove('hidden');
  updateStatus('Error', '❌', 'error');
}

// Hide error
function hideError() {
  errorContainer.classList.add('hidden');
}

// Show progress
function showProgress(count) {
  progressContainer.classList.remove('hidden');
  progressText.textContent = `${count} messages extracted`;
  // Indeterminate progress animation
  progressFill.style.width = '70%';
}

// Hide progress
function hideProgress() {
  progressContainer.classList.add('hidden');
  progressFill.style.width = '0%';
}

// Show results
function showResults(count) {
  resultsContainer.classList.remove('hidden');
  messageCount.textContent = `${count} messages`;
}

// Check if we are on Teams
async function checkTeamsPage() {
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    if (!tab) {
      updateStatus('No active tab', '⚠️');
      return false;
    }

    const isTeams = tab.url && (
      tab.url.includes('teams.microsoft.com') ||
      tab.url.includes('teams.cloud.microsoft')
    );

    if (isTeams) {
      updateStatus('Connected to Teams', '✅');
      exportBtn.disabled = false;
      exportCompleteBtn.disabled = false;
      debugBtn.disabled = false;
      return true;
    } else {
      updateStatus('Navigate to Teams first', '⚠️');
      exportBtn.disabled = true;
      exportCompleteBtn.disabled = true;
      debugBtn.disabled = true;
      return false;
    }
  } catch (error) {
    console.error('Error checking page:', error);
    updateStatus('Connection error', '❌');
    return false;
  }
}

// Debug DOM
async function debugDOM() {
  hideError();
  updateStatus('Analyzing DOM...', '🔍');

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    const result = await chrome.tabs.sendMessage(tab.id, {
      action: 'debug'
    });

    if (result.success) {
      console.log('Debug info:', result.data);
      updateStatus('Debug complete (see console)', '✅');
      alert(`Debug complete:\n\n${JSON.stringify(result.data, null, 2)}`);
    } else {
      showError(result.error || 'Debug error');
    }
  } catch (error) {
    console.error('Debug error:', error);
    showError(`Error: ${error.message}`);
  }
}

// Export conversation (visible messages only)
async function exportConversation() {
  hideError();
  hideProgress();
  resultsContainer.classList.add('hidden');

  updateStatus('Extracting messages...', '⏳');
  exportBtn.disabled = true;
  exportCompleteBtn.disabled = true;

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    // Send message to content script
    const result = await chrome.tabs.sendMessage(tab.id, {
      action: 'export'
    });

    if (result.success) {
      exportedData = result.data;
      const count = result.messageCount || 0;

      hideProgress();
      updateStatus('Export complete', '✅');
      showResults(count);
      exportBtn.disabled = false;
      exportCompleteBtn.disabled = false;
    } else {
      showError(result.error || 'Unknown export error');
      exportBtn.disabled = false;
      exportCompleteBtn.disabled = false;
    }
  } catch (error) {
    console.error('Export error:', error);
    showError(`Error: ${error.message}`);
    exportBtn.disabled = false;
    exportCompleteBtn.disabled = false;
  }
}

// Export complete conversation (with auto-scroll)
async function exportConversationComplete() {
  hideError();
  hideProgress();
  resultsContainer.classList.add('hidden');

  updateStatus('Starting auto-scroll...', '🔄');
  exportBtn.disabled = true;
  exportCompleteBtn.disabled = true;
  showProgress(0);

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    // Send message to content script
    const result = await chrome.tabs.sendMessage(tab.id, {
      action: 'export-complete'
    });

    if (result.success) {
      exportedData = result.data;
      const count = result.messageCount || 0;
      const iterations = result.iterations || 0;

      hideProgress();
      updateStatus(`Complete export (${iterations} iterations)`, '✅');
      showResults(count);
      exportBtn.disabled = false;
      exportCompleteBtn.disabled = false;
    } else {
      showError(result.error || 'Unknown complete export error');
      exportBtn.disabled = false;
      exportCompleteBtn.disabled = false;
      hideProgress();
    }
  } catch (error) {
    console.error('Complete export error:', error);
    showError(`Error: ${error.message}`);
    exportBtn.disabled = false;
    exportCompleteBtn.disabled = false;
    hideProgress();
  }
}

// Download file
function downloadFile() {
  if (!exportedData) {
    showError('No data to download');
    return;
  }

  try {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    const filename = `teams-export-${timestamp}.txt`;

    // Create blob
    const blob = new Blob([exportedData], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);

    // Create temporary <a> element to download
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';

    document.body.appendChild(a);
    a.click();

    // Cleanup
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 100);

    updateStatus('File downloaded', '💾');
  } catch (error) {
    console.error('Download error:', error);
    showError('Download error: ' + error.message);
  }
}

// Event listeners
exportBtn.addEventListener('click', exportConversation);
exportCompleteBtn.addEventListener('click', exportConversationComplete);
debugBtn.addEventListener('click', debugDOM);
downloadBtn.addEventListener('click', downloadFile);

// Show info when hovering over complete export button
exportCompleteBtn.addEventListener('mouseenter', () => {
  infoComplete.classList.remove('hidden');
});

exportCompleteBtn.addEventListener('mouseleave', () => {
  infoComplete.classList.add('hidden');
});

// Keep tooltip visible when hovering over it
infoComplete.addEventListener('mouseenter', () => {
  infoComplete.classList.remove('hidden');
});

infoComplete.addEventListener('mouseleave', () => {
  infoComplete.classList.add('hidden');
});

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  checkTeamsPage();
});

// Listen for tab changes
chrome.tabs.onActivated.addListener(() => {
  checkTeamsPage();
});

chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.status === 'complete') {
    checkTeamsPage();
  }
});

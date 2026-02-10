// Background service worker

console.log('[Teams Exporter] Service worker started');

// Handle installation
chrome.runtime.onInstalled.addListener((details) => {
  if (details.reason === 'install') {
    console.log('[Teams Exporter] Extension installed');
  } else if (details.reason === 'update') {
    console.log('[Teams Exporter] Extension updated');
  }
});

// Handle messages from content scripts
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  console.log('[Teams Exporter] Message received in background:', request);

  // Background-specific tasks can be handled here
  // For now, popup communicates directly with content script

  return false;
});

// Handle clicks on icon (optional - we already have popup)
chrome.action.onClicked.addListener((tab) => {
  console.log('[Teams Exporter] Icon clicked on tab:', tab.id);
});

// Content script - Runs in Teams context

Logger.info('Content script loaded');

// Extractor state
let isExtracting = false;

/**
 * Handle messages from popup
 */
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  Logger.debug('Message received:', request);

  if (request.action === 'debug') {
    handleDebug(sendResponse);
    return true;  // Keep channel open for async response
  }

  if (request.action === 'export') {
    handleExport(sendResponse);
    return true;  // Keep channel open for async response
  }

  if (request.action === 'export-complete') {
    handleExportComplete(sendResponse);
    return true;  // Keep channel open for async response
  }

  return false;
});

/**
 * Handle debug request
 */
async function handleDebug(sendResponse) {
  try {
    Logger.info('Starting DOM debug...');

    const debugInfo = debugDOM();

    Logger.info('Debug completed:', debugInfo);

    sendResponse({
      success: true,
      data: debugInfo
    });

  } catch (error) {
    Logger.error('Debug error:', error);
    sendResponse({
      success: false,
      error: error.message
    });
  }
}

/**
 * Handle export request
 */
async function handleExport(sendResponse) {
  if (isExtracting) {
    sendResponse({
      success: false,
      error: 'An export is already in progress'
    });
    return;
  }

  try {
    isExtracting = true;
    Logger.info('Starting export...');

    // Extract thread title
    const threadTitle = extractThreadTitle();
    Logger.info('Thread title:', threadTitle);

    // Extract messages
    const messages = extractAllMessages();
    Logger.info(`Extracted ${messages.length} messages`);

    if (messages.length === 0) {
      throw new Error('No messages found. Verify you are in a Teams conversation.');
    }

    // Format to IRC
    const exportData = createExport(messages, threadTitle);
    Logger.info('Export created:', {
      messageCount: exportData.messageCount,
      size: exportData.size
    });

    sendResponse({
      success: true,
      data: exportData.content,
      messageCount: exportData.messageCount,
      threadTitle: threadTitle
    });

  } catch (error) {
    Logger.error('Export error:', error);
    sendResponse({
      success: false,
      error: error.message
    });
  } finally {
    isExtracting = false;
  }
}

/**
 * Handle complete export request (with auto-scroll)
 */
async function handleExportComplete(sendResponse) {
  if (isExtracting) {
    sendResponse({
      success: false,
      error: 'An export is already in progress'
    });
    return;
  }

  try {
    isExtracting = true;
    Logger.info('Starting complete export with auto-scroll...');

    // Export with auto-scroll
    const result = await exportCompleteConversation((progress) => {
      Logger.debug('Progress:', progress);
      // Send progress to popup via port (if connected)
      // For now just log
    });

    if (result.success) {
      sendResponse({
        success: true,
        data: result.data,
        messageCount: result.messageCount,
        iterations: result.iterations
      });
    } else {
      throw new Error(result.error || 'Unknown error in complete export');
    }

  } catch (error) {
    Logger.error('Complete export error:', error);
    sendResponse({
      success: false,
      error: error.message
    });
  } finally {
    isExtracting = false;
  }
}

/**
 * Detect when a new conversation loads
 */
function detectConversationLoad() {
  // Observe DOM changes to detect navigation
  const observer = new MutationObserver(debounce(() => {
    Logger.debug('DOM changed - possible new conversation');
  }, 500));

  observer.observe(document.body, {
    childList: true,
    subtree: true
  });
}

// Initialize
Logger.info('Teams Exporter ready');
detectConversationLoad();

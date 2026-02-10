// Auto-scroll to load all messages from Teams

/**
 * Auto-scroll configuration
 */
const SCROLL_CONFIG = {
  viewport: '[data-tid="channel-replies-viewport"]',
  scrollStep: 500,           // Pixels to scroll each time
  scrollDelay: 800,          // Delay between scrolls (ms)
  loadTimeout: 5000,         // Timeout to wait for new messages
  maxIterations: 200,        // Maximum iterations (to avoid infinite loops)
  topThreshold: 50           // Threshold to consider we reached the top
};

/**
 * Find the scrollable container
 */
function findScrollContainer() {
  // Try first with specific selector
  let container = document.querySelector(SCROLL_CONFIG.viewport);

  if (container) {
    return container;
  }

  // Fallback: find nearest scrollable parent to messages
  const firstMessage = document.querySelector('[data-tid="channel-replies-pane-message"]');
  if (!firstMessage) {
    Logger.error('No messages found to locate scroll container');
    return null;
  }

  let parent = firstMessage.parentElement;
  let depth = 0;

  while (parent && depth < 15) {
    const style = window.getComputedStyle(parent);
    const isScrollable = (
      style.overflowY === 'auto' ||
      style.overflowY === 'scroll' ||
      style.overflow === 'auto' ||
      style.overflow === 'scroll'
    );

    if (isScrollable && parent.scrollHeight > parent.clientHeight) {
      Logger.info(`Scroll container found at depth ${depth}`, parent);
      return parent;
    }

    parent = parent.parentElement;
    depth++;
  }

  Logger.error('No scrollable container found');
  return null;
}

/**
 * Scroll upward and wait for messages to load
 */
async function scrollToTop(container, onProgress) {
  return new Promise((resolve, reject) => {
    let iteration = 0;
    let previousMessageCount = 0;
    let stableCount = 0;
    let previousScrollTop = container.scrollTop;

    const scrollInterval = setInterval(async () => {
      iteration++;

      // Check max iterations
      if (iteration > SCROLL_CONFIG.maxIterations) {
        clearInterval(scrollInterval);
        Logger.warn('Reached maximum iterations');
        resolve({
          success: true,
          reason: 'max_iterations',
          iterations: iteration
        });
        return;
      }

      // Get current position
      const currentScrollTop = container.scrollTop;
      const currentMessageCount = document.querySelectorAll('[data-tid="channel-replies-pane-message"]').length;

      Logger.debug(`Iteration ${iteration}: scrollTop=${currentScrollTop}, messages=${currentMessageCount}`);

      // Report progress
      if (onProgress) {
        onProgress({
          iteration,
          scrollTop: currentScrollTop,
          messageCount: currentMessageCount,
          action: 'scrolling'
        });
      }

      // Check if we reached the top
      if (currentScrollTop <= SCROLL_CONFIG.topThreshold) {
        // Wait a bit more to ensure all loaded
        await sleep(SCROLL_CONFIG.scrollDelay * 2);

        const finalMessageCount = document.querySelectorAll('[data-tid="channel-replies-pane-message"]').length;

        clearInterval(scrollInterval);
        Logger.info(`Reached the top. Total messages: ${finalMessageCount}`);
        resolve({
          success: true,
          reason: 'reached_top',
          messageCount: finalMessageCount,
          iterations: iteration
        });
        return;
      }

      // Check if messages stopped loading (stable for 3 iterations)
      if (currentMessageCount === previousMessageCount) {
        stableCount++;
        if (stableCount >= 3 && currentScrollTop === previousScrollTop) {
          clearInterval(scrollInterval);
          Logger.info(`Messages stable. Total: ${currentMessageCount}`);
          resolve({
            success: true,
            reason: 'stable',
            messageCount: currentMessageCount,
            iterations: iteration
          });
          return;
        }
      } else {
        stableCount = 0;
        previousMessageCount = currentMessageCount;
      }

      previousScrollTop = currentScrollTop;

      // Scroll upward
      container.scrollTop = Math.max(0, container.scrollTop - SCROLL_CONFIG.scrollStep);

    }, SCROLL_CONFIG.scrollDelay);
  });
}

/**
 * Complete export with auto-scroll
 */
async function exportCompleteConversation(onProgress) {
  const result = {
    success: false,
    messageCount: 0,
    iterations: 0,
    error: null,
    data: null
  };

  try {
    // 1. Find scrollable container
    if (onProgress) onProgress({ stage: 'finding_container', message: 'Finding container...' });

    const container = findScrollContainer();
    if (!container) {
      throw new Error('Scrollable container not found');
    }

    Logger.info('Container found:', {
      scrollHeight: container.scrollHeight,
      clientHeight: container.clientHeight,
      scrollTop: container.scrollTop
    });

    // 2. Count initial messages
    const initialMessageCount = document.querySelectorAll('[data-tid="channel-replies-pane-message"]').length;
    Logger.info(`Initial visible messages: ${initialMessageCount}`);

    if (onProgress) {
      onProgress({
        stage: 'initial_count',
        message: `Visible messages: ${initialMessageCount}`,
        messageCount: initialMessageCount
      });
    }

    // 3. Auto-scroll upward
    if (onProgress) onProgress({ stage: 'scrolling', message: 'Loading old messages...' });

    const scrollResult = await scrollToTop(container, (progress) => {
      if (onProgress) {
        onProgress({
          stage: 'scrolling',
          message: `Loading... ${progress.messageCount} messages`,
          messageCount: progress.messageCount,
          iteration: progress.iteration
        });
      }
    });

    result.iterations = scrollResult.iterations;
    Logger.info('Auto-scroll completed:', scrollResult);

    // 4. Wait a bit more to ensure rendering
    if (onProgress) onProgress({ stage: 'waiting', message: 'Finalizing load...' });
    await sleep(1500);

    // 5. Extract all messages
    if (onProgress) onProgress({ stage: 'extracting', message: 'Extracting messages...' });

    const threadTitle = extractThreadTitle();
    const messages = extractAllMessages();

    result.messageCount = messages.length;
    Logger.info(`Extracted ${messages.length} messages in total`);

    // 6. Format
    if (onProgress) onProgress({ stage: 'formatting', message: 'Generating file...' });

    const exportData = createExport(messages, threadTitle);

    result.success = true;
    result.data = exportData.content;
    result.messageCount = exportData.messageCount;

    if (onProgress) {
      onProgress({
        stage: 'complete',
        message: `✅ ${result.messageCount} messages exported`,
        messageCount: result.messageCount
      });
    }

  } catch (error) {
    Logger.error('Error in complete export:', error);
    result.error = error.message;

    if (onProgress) {
      onProgress({
        stage: 'error',
        message: `Error: ${error.message}`
      });
    }
  }

  return result;
}

/**
 * Cancel auto-scroll in progress
 */
let scrollCancelled = false;

function cancelScroll() {
  scrollCancelled = true;
  Logger.info('Auto-scroll cancelled by user');
}

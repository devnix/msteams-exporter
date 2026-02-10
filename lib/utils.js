// General utilities for the extension

/**
 * Escape HTML to prevent XSS
 */
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

/**
 * Format date to IRC format: YYYY-MM-DD HH:MM:SS
 */
function formatDate(date) {
  const pad = (n) => n.toString().padStart(2, '0');

  const year = date.getFullYear();
  const month = pad(date.getMonth() + 1);
  const day = pad(date.getDate());
  const hours = pad(date.getHours());
  const minutes = pad(date.getMinutes());
  const seconds = pad(date.getSeconds());

  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

/**
 * Parse Teams timestamp to Date
 * Teams can use various formats, we try to detect automatically
 */
function parseTeamsTimestamp(timestampStr) {
  // Try to parse ISO timestamp
  let date = new Date(timestampStr);

  if (isNaN(date.getTime())) {
    // Try other common formats
    // Examples: "12:30", "Yesterday at 3:45 PM", "Mon 12:30"
    // For now, use current date as fallback
    console.warn('Could not parse timestamp:', timestampStr);
    date = new Date();
  }

  return date;
}

/**
 * Clean text from unwanted elements
 */
function cleanText(text) {
  if (!text) return '';

  return text
    .trim()
    .replace(/\s+/g, ' ')  // Normalize spaces
    .replace(/[\r\n]+/g, ' ');  // Remove line breaks
}

/**
 * Extract text from an element including children
 */
function getTextContent(element) {
  if (!element) return '';

  // Clone to not modify the original
  const clone = element.cloneNode(true);

  // Remove unwanted elements (scripts, styles, etc.)
  const unwanted = clone.querySelectorAll('script, style, noscript');
  unwanted.forEach(el => el.remove());

  return cleanText(clone.textContent);
}

/**
 * Debounce to avoid excessive calls
 */
function debounce(func, wait) {
  let timeout;
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}

/**
 * Wait for a specified time (promise)
 */
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Generate unique ID for elements
 */
function generateId() {
  return `msg-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * Logger with levels
 */
const Logger = {
  debug: (...args) => console.log('[Teams Exporter Debug]', ...args),
  info: (...args) => console.info('[Teams Exporter]', ...args),
  warn: (...args) => console.warn('[Teams Exporter]', ...args),
  error: (...args) => console.error('[Teams Exporter]', ...args)
};

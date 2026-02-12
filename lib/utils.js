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
function parseTeamsTimestamp(timestamp) {
  if (!timestamp) {
    Logger.warn('Empty timestamp provided');
    return new Date();
  }

  // If it's already a number (Unix timestamp in milliseconds), use it directly
  if (typeof timestamp === 'number') {
    return new Date(timestamp);
  }

  // If it's a string that looks like a number, parse it as Unix timestamp
  const numericTimestamp = parseInt(timestamp);
  if (!isNaN(numericTimestamp) && numericTimestamp.toString() === timestamp) {
    return new Date(numericTimestamp);
  }

  // Try to parse ISO timestamp or other standard formats
  let date = new Date(timestamp);

  if (!isNaN(date.getTime())) {
    return date;
  }

  // If all parsing fails, log a warning and use current time as last resort
  Logger.warn('Could not parse timestamp, using current time:', timestamp);
  return new Date();
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

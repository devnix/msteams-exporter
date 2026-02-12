// Message formatter to IRC style

/**
 * Parse reaction text to emoji
 * Converts "1 Like reaction with light skin tone." to "👍"
 */
function parseReactionToEmoji(reactionText) {
  const text = reactionText.toLowerCase();

  // Map Spanish/English reaction descriptions to emojis
  const emojiMap = {
    'me gusta': '👍',
    'like': '👍',
    'love': '❤️',
    'corazón': '❤️',
    'amor': '❤️',
    'carcajada': '😂',
    'laugh': '😂',
    'sorpresa': '😮',
    'surprise': '😮',
    'wow': '😮',
    'triste': '😢',
    'sad': '😢',
    'enfadado': '😠',
    'angry': '😠',
    'celebrar': '🎉',
    'celebrate': '🎉',
    'party': '🎉',
    'aplaudir': '👏',
    'clap': '👏'
  };

  // Try to find matching emoji
  for (const [key, emoji] of Object.entries(emojiMap)) {
    if (text.includes(key)) {
      // Extract count if present
      const countMatch = reactionText.match(/(\d+)\s+(reacci[oó]n|reaction)/);
      const count = countMatch ? parseInt(countMatch[1]) : 1;

      // Repeat emoji based on count
      return Array(count).fill(emoji).join('');
    }
  }

  // Fallback: try to extract emoji from text
  const emojiMatch = reactionText.match(/[\u{1F300}-\u{1F9FF}]/u);
  if (emojiMatch) {
    return emojiMatch[0];
  }

  // Last resort: return generic reaction
  return '👍';
}

/**
 * Add @ prefix to mentions in text
 * Detects user names and adds @ prefix
 */
function formatMentions(text) {
  if (!text) return text;

  // Pattern for names: Capitalized words (First Last)
  // Matches patterns like "John Smith Doe" but not "De" or "en"
  const namePattern = /\b([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+)+)\b/g;

  let formatted = text;
  const matches = [...text.matchAll(namePattern)];

  // Process matches in reverse to maintain positions
  for (let i = matches.length - 1; i >= 0; i--) {
    const match = matches[i];
    const name = match[1];
    const position = match.index;

    // Don't add @ if it already has it
    if (position > 0 && formatted[position - 1] === '@') {
      continue;
    }

    // Check if it's likely a name (has at least 2 words)
    const words = name.split(/\s+/);
    if (words.length >= 2) {
      formatted = formatted.substring(0, position) + '@' + formatted.substring(position);
    }
  }

  return formatted;
}

/**
 * Format individual message to IRC format
 * @param {object} message - Message object
 * @param {string|null} lastTimestamp - Last known timestamp (for grouping)
 * @returns {object} - {formatted: string, timestamp: string}
 */
function formatMessage(message, lastTimestamp = null) {
  const lines = [];
  let currentTimestamp = null;

  try {
    // Parse timestamp
    let timestampDisplay;
    if (message.timestamp) {
      const date = parseTeamsTimestamp(message.timestamp);
      currentTimestamp = formatDate(date);
      timestampDisplay = currentTimestamp;
    } else if (lastTimestamp) {
      // Use previous timestamp if available
      currentTimestamp = lastTimestamp;
      timestampDisplay = lastTimestamp;
    } else {
      // No timestamp available - use whitespace to maintain alignment
      // Format is [YYYY-MM-DD HH:MM:SS] = 21 chars including brackets
      timestampDisplay = ' '.repeat(19); // 19 spaces (timestamp without brackets)
    }

    const timestampPrefix = message.timestamp || lastTimestamp ? `[${timestampDisplay}]` : `[${timestampDisplay}]`;

    // Format author
    const author = message.author || null;

    // Subject of message (if exists)
    if (message.subject) {
      lines.push(`${timestampPrefix} === ${message.subject} ===`);
    }

    // Main message - always use <author> format
    let content = message.content || '';

    // If content has multiple paragraphs, format each one
    if (content.includes('\n\n')) {
      const paragraphs = content.split('\n\n');
      paragraphs.forEach((para, idx) => {
        if (idx === 0) {
          lines.push(`${timestampPrefix} <${author}> ${para}`);
        } else {
          lines.push(`${timestampPrefix}   ${para}`);
        }
      });
    } else if (content.trim()) {
      lines.push(`${timestampPrefix} <${author}> ${content}`);
    }

    // Mark if edited
    if (message.isEdited && content && !content.includes('[EDITED]')) {
      lines[lines.length - 1] += ' [EDITED]';
    }

    // Attachments
    if (message.attachments && message.attachments.length > 0) {
      message.attachments.forEach(attachment => {
        const displayAuthor = author || 'SYSTEM';
        lines.push(`${timestampPrefix} <${displayAuthor}> [ATTACHMENT: ${attachment.name}]`);
      });
    }

    // Reactions - parse to emojis only
    if (message.reactions && message.reactions.length > 0) {
      const emojis = message.reactions
        .map(r => {
          if (r.emojis && r.emojis.length > 0) {
            // Use first emoji variant from DOM and repeat by count
            // (DOM doesn't provide breakdown for skin tone variations)
            const emoji = r.emojis[0];
            const count = r.count || 1;
            return Array(count).fill(emoji).join('');
          } else if (r.emoji) {
            // Backward compatibility: single emoji
            return Array(r.count || 1).fill(r.emoji).join('');
          } else if (r.text) {
            // Fallback: parse text if emoji not extracted from DOM
            return parseReactionToEmoji(r.text);
          }
          return '';
        })
        .filter(e => e)
        .join(' ');

      if (emojis) {
        lines.push(`${timestampPrefix} ${emojis}`);
      }
    }

  } catch (error) {
    Logger.error('Error formatting message:', error);
    lines.push(`[ERROR] Could not format message: ${error.message}`);
  }

  return {
    formatted: lines.join('\n'),
    timestamp: currentTimestamp
  };
}

/**
 * Format complete message list
 */
function formatMessages(messages, options = {}) {
  const lines = [];

  // Options
  const includeHeader = options.includeHeader !== false;
  const threadTitle = options.threadTitle || 'Teams Conversation';

  // Header
  if (includeHeader) {
    lines.push('='.repeat(80));
    lines.push(`MS Teams - ${threadTitle}`);
    lines.push(`Exported: ${formatDate(new Date())}`);
    lines.push(`Total messages: ${messages.length}`);
    lines.push('Format: IRC (optimized for LLMs)');
    lines.push('='.repeat(80));
    lines.push('');
  }

  // Track last known author and timestamp for consecutive messages
  let lastAuthor = 'Unknown';
  let lastTimestamp = null;

  // Messages
  messages.forEach((message, index) => {
    // If message has author, update last known author
    if (message.author && message.author.trim()) {
      lastAuthor = message.author;
    } else {
      // Use last known author for messages without author
      message.author = lastAuthor;
    }

    // Format message with last known timestamp
    const result = formatMessage(message, lastTimestamp);

    // Debug: log if message is being skipped
    if (!result.formatted.trim()) {
      Logger.warn(`Message ${index} skipped - no content:`, {
        author: message.author,
        timestamp: message.timestamp,
        content: message.content,
        subject: message.subject,
        attachments: message.attachments?.length || 0,
        reactions: message.reactions?.length || 0
      });
    }

    if (result.formatted.trim()) {  // Only add non-empty messages
      lines.push(result.formatted);
    }

    // Update last timestamp if this message had one
    if (result.timestamp) {
      lastTimestamp = result.timestamp;
    }

    // Separator between messages (optional)
    if (options.addSeparator && index < messages.length - 1) {
      lines.push('');
    }
  });

  // Footer
  if (includeHeader) {
    lines.push('');
    lines.push('='.repeat(80));
    lines.push('End of conversation');
    lines.push('='.repeat(80));
  }

  return lines.join('\n');
}

/**
 * Create export file
 */
function createExport(messages, threadTitle) {
  const content = formatMessages(messages, {
    includeHeader: true,
    threadTitle: threadTitle,
    addSeparator: false
  });

  return {
    content: content,
    messageCount: messages.length,
    size: new Blob([content]).size,
    timestamp: new Date().toISOString()
  };
}

// Message extractor from Teams DOM

/**
 * Selectors for Microsoft Teams
 * Verified on 2026-02-10 with Teams Web
 */
const SELECTORS = {
  // Individual message - MAIN SELECTOR
  messageItem: '[data-tid="channel-replies-pane-message"]',

  // Thread/conversation title
  threadTitle: 'h2[id^="subject-line"]',

  // Message author (span inside message)
  author: 'span.fui-StyledText',

  // Message timestamp
  timestamp: 'time',

  // Message content
  messageBody: '[data-tid="message-body"]',

  // Content paragraphs
  contentParagraphs: 'p',

  // Message header (contains avatar and metadata)
  messageHeader: '[data-tid="post-message-header-avatar"]',

  // Attachments and images
  attachmentImages: 'img[role="button"]',

  // Reactions
  reactions: '[data-tid="channel-message-reaction-summary"]'
};

/**
 * Extract thread title
 */
function extractThreadTitle() {
  // Find h2 with id starting with "subject-line"
  const titleElement = document.querySelector(SELECTORS.threadTitle);

  if (titleElement) {
    return cleanText(titleElement.textContent);
  }

  // Fallback: find any h2 in message area
  const h2Elements = document.querySelectorAll('h2');
  for (const h2 of h2Elements) {
    const text = cleanText(h2.textContent);
    if (text && text.length > 5) {
      return text;
    }
  }

  return 'Untitled conversation';
}

/**
 * Extract information from an individual message
 */
function extractMessage(messageElement) {
  const message = {
    id: generateId(),
    author: null,
    timestamp: null,
    content: null,
    subject: null,
    attachments: [],
    reactions: [],
    isEdited: false,
    isSystemMessage: false,
    raw: {}
  };

  try {
    // Extract author - first span.fui-StyledText containing a name
    const authorSpans = messageElement.querySelectorAll(SELECTORS.author);
    for (const span of authorSpans) {
      const text = cleanText(span.textContent);
      // Author is typically a name (has spaces or is long)
      // Filter out timestamps, buttons, etc.
      if (text && (text.includes(' ') || text.length > 3) &&
          !text.includes(':') && !text.match(/^\d+:\d+/) &&
          !text.match(/^(Yesterday|Today|Tomorrow)/i)) {
        message.author = text;
        message.raw.author = span.outerHTML.substring(0, 200);
        break;
      }
    }

    // Extract timestamp - Teams stores it in the id attribute as "timestamp-{unix_ms}"
    const timestampEl = messageElement.querySelector(SELECTORS.timestamp);
    if (timestampEl) {
      // Extract timestamp from id attribute (format: "timestamp-1770721260027")
      const id = timestampEl.getAttribute('id');
      if (id && id.startsWith('timestamp-')) {
        const timestampMs = parseInt(id.replace('timestamp-', ''));
        if (!isNaN(timestampMs)) {
          // Store as Unix timestamp in milliseconds
          message.timestamp = timestampMs;
        }
      }

      // Fallback to datetime attribute or text content if id parsing fails
      if (!message.timestamp) {
        const datetime = timestampEl.getAttribute('datetime');
        message.timestamp = datetime || cleanText(timestampEl.textContent);
      }

      message.raw.timestamp = timestampEl.outerHTML.substring(0, 200);
    } else {
      Logger.debug('No timestamp element found for message');
    }

    // Extract subject (h2 if exists in this message)
    const subjectEl = messageElement.querySelector('h2');
    if (subjectEl) {
      message.subject = cleanText(subjectEl.textContent);
    }

    // Extract message content
    const messageBodyEl = messageElement.querySelector(SELECTORS.messageBody);
    if (messageBodyEl) {
      // Clone the element to modify it
      const bodyClone = messageBodyEl.cloneNode(true);

      // Find all mentions and mark them with a special delimiter
      // Teams splits full names into multiple mention elements (e.g., "John Smith Doe")
      const mentions = bodyClone.querySelectorAll('[itemtype="http://schema.skype.com/Mention"]');
      mentions.forEach(mention => {
        const text = mention.textContent.trim();
        if (text) {
          // Use special markers to identify mentions in the text
          mention.textContent = `§§${text}§§`;
        }
      });

      // Extract content paragraphs
      const paragraphs = bodyClone.querySelectorAll(SELECTORS.contentParagraphs);

      if (paragraphs.length > 0) {
        message.content = Array.from(paragraphs)
          .map(p => cleanText(p.textContent))
          .filter(text => text.length > 0)
          .join('\n\n');
      } else {
        // Fallback: take all text from body
        message.content = cleanText(bodyClone.textContent);
      }

      // Post-process: combine consecutive mentions into @<Full Name> format
      // Pattern: §§Word1§§ §§Word2§§ §§Word3§§ -> @<Word1 Word2 Word3>
      message.content = message.content.replace(/§§([^§]+)§§(?:\s+§§([^§]+)§§)*/g, (match) => {
        // Extract all mention parts from the match
        const parts = match.match(/§§([^§]+)§§/g).map(m => m.replace(/§§/g, ''));

        if (parts.length === 1) {
          // Single mention: @Word
          return `@${parts[0]}`;
        } else {
          // Multiple consecutive mentions (full name): @<Word1 Word2 Word3>
          return `@<${parts.join(' ')}>`;
        }
      });

      message.raw.content = messageBodyEl.outerHTML.substring(0, 500);

      // Detect if edited
      const fullText = messageBodyEl.textContent.toLowerCase();
      message.isEdited = fullText.includes('edited') || fullText.includes('editado');
    }

    // Extract attachments (images, files)
    const images = messageElement.querySelectorAll(SELECTORS.attachmentImages);
    images.forEach(img => {
      const alt = img.getAttribute('alt') || 'image';
      const src = img.getAttribute('src') || '';
      // Ignore profile avatars
      if (!src.includes('profilepicture') && !src.includes('avatar')) {
        message.attachments.push({
          name: alt,
          url: src,
          type: 'image'
        });
      }
    });

    // Extract reactions
    const reactionsEl = messageElement.querySelector(SELECTORS.reactions);
    if (reactionsEl) {
      message.reactions = extractReactions(reactionsEl);
      message.raw.reactions = reactionsEl.outerHTML.substring(0, 200);
    }

    // Detect system message
    message.isSystemMessage = !message.author || message.author === '';

  } catch (error) {
    Logger.error('Error extracting message:', error);
    message.error = error.message;
  }

  return message;
}

/**
 * Extract reactions from an element
 */
function extractReactions(reactionsElement) {
  const reactions = [];

  try {
    // Find all buttons with reactions (data-tid="diverse-reaction-pill-button")
    const reactionButtons = reactionsElement.querySelectorAll('button[data-tid="diverse-reaction-pill-button"]');

    reactionButtons.forEach(button => {
      try {
        // Find ALL img elements with alt attribute (can be multiple for skin tone variations)
        const images = button.querySelectorAll('img[alt]');

        if (images.length > 0) {
          // Extract count from button text (appears in a span)
          const buttonText = button.textContent.trim();
          const countMatch = buttonText.match(/(\d+)/);
          const count = countMatch ? parseInt(countMatch[1]) : images.length;

          // Collect all unique emojis from this button
          const emojis = Array.from(images)
            .map(img => img.alt.trim())
            .filter(alt => alt); // Remove empty

          if (emojis.length > 0) {
            reactions.push({
              emojis: emojis,  // Array of all emoji variants
              count: count
            });
          }
        }
      } catch (err) {
        Logger.warn('Error extracting individual reaction:', err);
      }
    });

    // Fallback: if no buttons found, try to extract from text
    if (reactions.length === 0) {
      const text = reactionsElement.textContent.trim();
      if (text) {
        reactions.push({
          text: text,
          raw: true
        });
      }
    }

  } catch (error) {
    Logger.warn('Error extracting reactions:', error);
  }

  return reactions;
}

/**
 * Extract all visible messages in the DOM
 */
function extractAllMessages() {
  const messages = [];

  try {
    // Find all messages using the main selector
    const messageElements = document.querySelectorAll(SELECTORS.messageItem);

    if (!messageElements || messageElements.length === 0) {
      Logger.warn('No messages found');
      return messages;
    }

    Logger.info(`Found ${messageElements.length} messages`);

    // Extract information from each message
    messageElements.forEach((el, index) => {
      try {
        const message = extractMessage(el);
        message.index = index;
        messages.push(message);
      } catch (error) {
        Logger.error(`Error in message ${index}:`, error);
      }
    });

  } catch (error) {
    Logger.error('Error extracting messages:', error);
  }

  return messages;
}

/**
 * Debug: Analyze DOM structure
 */
function debugDOM() {
  const info = {
    url: window.location.href,
    title: document.title,
    threadTitle: null,
    selectors: {},
    messageCount: 0,
    sampleMessages: []
  };

  try {
    // Extract thread title
    info.threadTitle = extractThreadTitle();

    // Verify main selectors
    info.selectors.messageItems = document.querySelectorAll(SELECTORS.messageItem).length;
    info.selectors.threadTitle = document.querySelectorAll(SELECTORS.threadTitle).length;
    info.selectors.timestamps = document.querySelectorAll(SELECTORS.timestamp).length;
    info.selectors.messageBodies = document.querySelectorAll(SELECTORS.messageBody).length;

    // Try to extract messages
    const messages = extractAllMessages();
    info.messageCount = messages.length;

    if (messages.length > 0) {
      // Include first 3 messages as sample
      info.sampleMessages = messages.slice(0, 3).map(msg => ({
        index: msg.index,
        author: msg.author,
        timestamp: msg.timestamp,
        subject: msg.subject,
        contentPreview: msg.content ? msg.content.substring(0, 100) : null,
        attachmentCount: msg.attachments.length,
        hasReactions: msg.reactions.length > 0
      }));
    }

  } catch (error) {
    info.error = error.message;
    Logger.error('Debug error:', error);
  }

  return info;
}

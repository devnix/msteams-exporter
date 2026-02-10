# MS Teams Chat Exporter - Project Guide

## Project Description

Chrome extension to export Microsoft Teams Web conversations to IRC-style plain text format, optimized for LLM (Large Language Model) ingestion.

### Main Objective
Enable fast and structured export of Teams conversations for:
- LLM analysis
- Historical archiving in readable format
- Conversational data processing

### Technology Stack
- **Platform**: Chrome Extension (Manifest V3)
- **Target**: Microsoft Teams Web (https://teams.cloud.microsoft/)
- **Output format**: IRC-style plain text
- **Encoding**: UTF-8

## Extension Architecture

### File Structure
```
msteams-exporter/
├── manifest.json              # Extension configuration (Manifest V3)
├── CLAUDE.md                  # This file - project instructions
├── PLAN.md                    # Detailed development plan
├── README.md                  # User documentation
├── popup/
│   ├── popup.html            # Extension user interface
│   ├── popup.css             # Interface styles
│   └── popup.js              # UI logic and coordination
├── content/
│   └── content.js            # Script injected into Teams - main extractor
├── background/
│   └── background.js         # Service worker - download handling
├── lib/
│   ├── formatter.js          # Message formatting to IRC style
│   ├── extractor.js          # DOM data extraction
│   └── utils.js              # General utilities
└── icons/
    ├── icon16.png
    ├── icon48.png
    └── icon128.png
```

### Main Components

#### 1. Content Script (content.js)
- Injected into teams.cloud.microsoft.com pages
- Responsibilities:
  - Detect when user is in a conversation
  - Extract messages from DOM
  - Handle scroll to load older messages
  - Communicate with popup via chrome.runtime

#### 2. Popup (popup.html/js)
- Simple user interface
- Features:
  - "Export Current Conversation" button
  - Date range selector (optional)
  - Progress indicator
  - Exported message counter
  - Download button

#### 3. Background Service Worker (background.js)
- Handles .txt file download
- Manages extension permissions and state

#### 4. Formatter (lib/formatter.js)
- Converts extracted data to IRC format
- Handles special cases (edits, replies, etc.)

## IRC Output Format

### Base Format
```
[YYYY-MM-DD HH:MM:SS] <Username> Text message
[YYYY-MM-DD HH:MM:SS] <Username> → Reply to previous message
[YYYY-MM-DD HH:MM:SS] * Username performed an action
```

### Special Cases

#### Messages with Attachments
```
[YYYY-MM-DD HH:MM:SS] <Username> Message
[YYYY-MM-DD HH:MM:SS] <Username> [ATTACHMENT: filename.pdf]
```

#### Edited Messages
```
[YYYY-MM-DD HH:MM:SS] <Username> Edited message [EDITED]
```

#### Reactions
```
[YYYY-MM-DD HH:MM:SS] <Username> Message
[YYYY-MM-DD HH:MM:SS] * Reactions: 👍(3) ❤️(2) 😄(1)
```

#### System Messages
```
[YYYY-MM-DD HH:MM:SS] *** User A joined the conversation
[YYYY-MM-DD HH:MM:SS] *** User B changed chat title to "New Title"
```

#### Reply Threads
```
[YYYY-MM-DD HH:MM:SS] <User A> Main message
[YYYY-MM-DD HH:MM:SS] <User B> → Thread reply
[YYYY-MM-DD HH:MM:SS] <User C> → → Nested reply
```

## Microsoft Teams Technical Information

### Teams Web Architecture
- Framework: React with virtualized components
- Rendering: Virtual scrolling for optimization
- State: Redux/similar for state management
- Communication: WebSocket + REST API

### Technical Challenges

#### 1. List Virtualization
Teams only renders messages visible in the viewport. To export long conversations:
- Implement auto-scroll upward
- Wait for older messages to load
- Detect when conversation beginning is reached

#### 2. Lazy Loading
Messages load dynamically:
- Monitor DOM changes (MutationObserver)
- Identify "loading more messages" indicator
- Detect oldest available message

#### 3. Dynamic DOM Structure
Teams uses generated class names (CSS-in-JS):
- Prefer data-* attributes when available
- Use [role] and [aria-*] as stable selectors
- Have fallbacks for broken selectors

### DOM Selectors (To Research and Document)

**IMPORTANT**: The following selectors are hypothetical and must be verified during debugging session:

```javascript
// Selector template to discover
const SELECTORS = {
  // Main conversation container
  conversationContainer: '[data-tid="conversation-pane"]', // TO VERIFY

  // Message list
  messageList: '[role="list"][aria-label*="messages"]', // TO VERIFY
  messageItems: '[role="listitem"]', // TO VERIFY

  // Message components
  messageAuthor: '[data-tid="message-author"]', // TO VERIFY
  messageTime: 'time, [data-tid="message-timestamp"]', // TO VERIFY
  messageContent: '[data-tid="message-body-content"]', // TO VERIFY
  messageEdited: '[aria-label*="edited"]', // TO VERIFY

  // Attachments
  attachments: '[data-tid="attachment"]', // TO VERIFY
  attachmentName: '[data-tid="attachment-name"]', // TO VERIFY

  // Reactions
  reactions: '[data-tid="message-reactions"]', // TO VERIFY

  // Threads
  threadReplies: '[data-tid="thread-replies"]', // TO VERIFY

  // Loading indicators
  loadingIndicator: '[data-tid="loading-messages"]', // TO VERIFY
  scrollToTop: '[data-tid="scroll-to-top"]', // TO VERIFY
};
```

### Teams API (Advanced Alternative)

If GraphQL or REST calls are detected during debugging:
```javascript
// Possible endpoints to monitor in Network tab
// https://teams.microsoft.com/api/csa/*/conversations
// https://graph.microsoft.com/v1.0/chats/{id}/messages
// WebSocket: wss://teams.microsoft.com/...
```

## Development Strategy

### Phase 1: Research (Current State)
1. ✅ Create project structure
2. ✅ Open Teams in Chrome DevTools
3. ⏳ Log in and navigate to conversation
4. ⏳ Inspect DOM and document real selectors
5. ⏳ Analyze Network tab to identify APIs
6. ⏳ Update SELECTORS in this document

### Phase 2: Extractor Prototype
1. Create basic extractor function in DevTools console
2. Test single message extraction
3. Expand to full list of visible messages
4. Implement date/time parsing
5. Validate IRC format output

### Phase 3: Scroll and Loading Handling
1. Implement auto-scroll upward
2. Detect new message loading
3. Prevent duplicates
4. Detect end of conversation
5. Implement safety timeout

### Phase 4: Extension Development
1. Create manifest.json with required permissions
2. Implement basic popup
3. Implement content script with extractor
4. Integrate formatter
5. Implement file download

### Phase 5: Refinement
1. Add robust error handling
2. Implement progress indicator
3. Add filters (date, user)
4. Optimize for large conversations (>1000 messages)
5. Testing in different scenarios

## Instructions for Claude Code

### When working on this project:

1. **DOM Analysis**
   - Always verify selectors in current Teams session
   - Update this document with correct selectors
   - Document variations found

2. **Incremental Development**
   - Start with minimum viable functionality
   - Test each component in isolation in console
   - Don't assume DOM structure, always verify

3. **Real-Time Testing**
   - Use Chrome DevTools to test inline JavaScript
   - Validate that selectors work before writing code
   - Capture screenshots of structure when useful

4. **Output Format**
   - Maintain compatibility with specified IRC format
   - Prioritize human readability
   - Optimize for LLM parsing (consistent structure)

5. **Error Handling**
   - Graceful degradation if data is missing
   - Clear logging for debugging
   - Don't fail entire export if one message has issues

6. **Chrome Permissions**
   - Request minimum necessary permissions
   - Document why each permission is needed
   - Follow Manifest V3 best practices

### Useful Commands During Development

```bash
# Development in project directory
cd ~/dev/devnix/msteams-exporter

# Quick testing structure
# 1. Open DevTools in Teams
# 2. Copy/paste extractor code in console
# 3. Execute and verify output
# 4. Iterate until it works
# 5. Move to extension files

# To load extension in Chrome:
# chrome://extensions -> Developer Mode -> Load Unpacked -> select directory
```

### Immediate Next Steps

1. **Log in to Teams** (using Chrome DevTools)
2. **Navigate to a test conversation**
3. **Execute DOM inspection:**
   ```javascript
   // Code to execute in DevTools console
   // Identify message container
   document.querySelectorAll('[role="list"]');

   // Identify individual messages
   document.querySelectorAll('[role="listitem"]');

   // Examine message structure
   const firstMsg = document.querySelector('[role="listitem"]');
   console.log(firstMsg.innerHTML);
   ```

4. **Document findings** in this file (SELECTORS section)
5. **Develop extractor prototype** in console
6. **Validate output format**

## Important Notes

- **Only works with Teams Web**, not desktop application
- **Requires login** - extension doesn't handle authentication
- **Limited by virtualization** - very long conversations may take time
- **No access to deleted messages** - only what's in DOM
- **Dependent on Teams UI** - may break with updates

## Resources and References

- [Chrome Extension Manifest V3](https://developer.chrome.com/docs/extensions/mv3/)
- [Microsoft Graph API - Teams Messages](https://learn.microsoft.com/en-us/graph/api/chat-list-messages)
- [Teams Developer Documentation](https://learn.microsoft.com/en-us/microsoftteams/platform/)

## Maintenance

This project should be updated when:
- Microsoft changes Teams DOM structure
- Chrome updates Manifest V3 with breaking changes
- Bugs are discovered in export
- New features are required (filters, formats, etc.)

---

**Last updated**: 2026-02-10
**Status**: Phase 1 - Research in progress
**Next action**: Log in to Teams and document DOM selectors

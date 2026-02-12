# Changelog - MS Teams Chat Exporter

## [0.1.1] - 2026-02-12

### 🐛 Bugfixes

- **Fixed timestamp extraction**: All messages were showing the same timestamp (current time) instead of their actual message time
  - Now extracts Unix timestamps from `<time>` element's `id` attribute (`timestamp-{unix_ms}`)
  - Properly handles numeric timestamps in milliseconds
  - Fixes parsing fallback that was defaulting to `new Date()`

### ✨ Improvements

- **Message grouping**: Consecutive messages from the same author now reuse the previous timestamp for better visual grouping
  - Messages without timestamps use whitespace padding to maintain alignment
  - Improves readability in IRC-style output

### 🔧 Technical Changes

- Added debug logging for skipped messages and extraction issues
- Improved `parseTeamsTimestamp()` to handle numeric Unix timestamps
- Updated `formatMessage()` to track and reuse last known timestamp

---

## [0.1.0] - 2026-02-10

### Initial Release

#### Features

- ✅ **Two export modes**:
  - 📥 **Export Visible**: Fast, exports currently visible messages only
  - 🔄 **Complete Export**: Auto-scrolls to load and export full conversation history
- ✅ Message extraction from Microsoft Teams Web
- ✅ IRC-style format optimized for LLM analysis
- ✅ Extraction of:
  - Author name
  - Timestamp
  - Message content (with multiple paragraphs)
  - Message subject/title
  - Reactions (as emojis with counts)
  - Attachments and images
  - Multi-word name mentions (`@<Full Name>` format)
- ✅ Debug button to verify DOM selectors
- ✅ Download as .txt file

#### Auto-Scroll System

- Automatic detection of scrollable container
- Progressive scroll algorithm (500px every 800ms)
- Intelligent detection of conversation start
- Safety limit of 200 maximum iterations
- Duplicate message prevention

#### Verified DOM Selectors

```javascript
const SELECTORS = {
  messageItem: '[data-tid="channel-replies-pane-message"]',
  threadTitle: 'h2[id^="subject-line"]',
  author: 'span.fui-StyledText',
  timestamp: 'time',
  messageBody: '[data-tid="message-body"]',
  reactions: '[data-tid="channel-message-reaction-summary"]'
};
```

#### Files

- `manifest.json` - Chrome Extension Manifest V3
- `popup/popup.{html,css,js}` - User interface
- `content/content.js` - Script injected into Teams
- `background/background.js` - Service worker
- `lib/utils.js` - General utilities
- `lib/extractor.js` - DOM extraction logic
- `lib/formatter.js` - IRC format conversion
- `lib/scroller.js` - Auto-scroll functionality
- `icons/` - Extension icons

#### Documentation

- `README.md` - User documentation
- `CLAUDE.md` - Project instructions for LLMs
- `TESTING.md` - Testing guide

---

## Conventions

### Versioning

We follow [Semantic Versioning](https://semver.org/):

- **MAJOR** (X.0.0): Breaking changes
- **MINOR** (0.X.0): New backwards-compatible features
- **PATCH** (0.0.X): Bug fixes

### Change Types

- ✨ **New Features**: New features
- 🐛 **Bugfix**: Error corrections
- 🔧 **Technical Changes**: Refactoring, optimizations
- 📝 **Documentation**: Docs changes
- 🎨 **UI/UX**: Interface improvements
- ⚡ **Performance**: Speed optimizations
- 🔒 **Security**: Security fixes

---

**Maintained by**: devnix
**Project**: MS Teams Chat Exporter
**License**: MIT

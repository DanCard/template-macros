# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **Google Apps Script** project that provides document formatting automation for Google Docs. The entire project consists of a single file (`current.js`) containing custom macros accessible via a "Custom Macros" menu in Google Docs.

## Deployment

This is a Google Apps Script project—there is no traditional build system.

**Document:** https://docs.google.com/document/d/15XVAEj4ljdHIuutK0BDrtY9MHm0uYiyPRD_wQcDQxqw/edit?tab=t.0

**Script Editor:** https://script.google.com/u/1/home/projects/1OcF3lfdrcZR3yFjMS1RN9hbWCjrhtrG9pFSKwBy5v0uF4AdHLrCyl9Ju/edit

Code is deployed via:
- Google Workspace Script Editor UI
- `clasp push` (if using clasp CLI)

Changes require saving in the script editor before testing.

## Architecture

### Entry Point
- `onOpen()` - Creates the "Custom Macros" menu and auto-formats new documents with date/time

### Core Functions

| Function | Purpose |
|----------|---------|
| `formatAndSetTitle()` | Main title formatter with responsive font sizing (6-36pt) based on page width |
| `formatLinks()` | Cleans URLs, removes underlines, applies special formatting for lmarena.ai, aistudio.google, Wikipedia |
| `formatCode()` | Applies code styling (Spectral Light, 12pt, 0.7 line spacing) to selection |
| `formatCurrentTable()` | Formats table at cursor (borders, padding, font size, alignment) |

### Key Constants
- `PAGE_WIDTH_POINTS = 502` - Standard printable area width (trial and error value)
- `CHAR_WIDTH_FACTOR = 0.42` - Character width multiplier for Spectral Light font
- `MAX_TITLE_SIZE = 36`, `MIN_FONT_SIZE = 6`

### Title Format
Expects first line format: `"Title - Date"` or `"Title? Date"`. The title portion gets responsive sizing while date stays at 12pt.

## Known Quirks

### ListItem Bug Workaround
Google Docs API sometimes fails to apply attributes to ListItem elements. The workaround (see `applyFormatting()` at line 451) "nudges" the element by setting a single attribute first to detach it from inherited List styles:
```javascript
if (element.getType() === DocumentApp.ElementType.LIST_ITEM) {
  element.setSpacingBefore(0); // Force element to have its own attributes
}
element.setAttributes(attributes);
```

### Alternative Implementations
`resizeTitle2()` is an older version of the title resizing algorithm with slightly different constants (CHAR_WIDTH_FACTOR = 0.44, MAX_TITLE_SIZE = 30). The primary function is `formatAndSetTitle()`.

## APIs Used

- `DocumentApp` - Document manipulation (Body, Paragraph, Table, Text)
- `Session` - Timezone info
- `Utilities` - Date formatting

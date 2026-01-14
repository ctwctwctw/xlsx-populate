# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

xlsx-populate is an Excel XLSX parser/generator for Node.js and browsers. It provides jQuery/d3-style method chaining and preserves existing workbook features and styles during read/write operations.

## Commands

### Testing
```bash
npm test                    # Run unit tests (Jasmine)
npm run e2e-generate       # Run e2e generation tests (Windows only, requires Excel)
npm run e2e-parse          # Run e2e parsing tests

# Run specific unit test file
npx jasmine test/unit/Cell.spec.js JASMINE_CONFIG_PATH=test/unit/jasmine.json
```

### Building
```bash
gulp build                 # Full build: docs → browser → lint → unit → tests
gulp browser              # Build browser bundles only
gulp lint                 # Run ESLint on lib/ directory
gulp docs                 # Generate README.md from JSDoc comments
gulp                      # Watch mode for development
```

## Architecture

### Class Hierarchy

The library follows a hierarchical object model mirroring Excel's structure:

```
XlsxPopulate (API facade)
└── Workbook (container, I/O)
    └── Sheet (worksheet)
        ├── Row → Cell
        ├── Column
        └── Range (multi-cell operations)
            └── Cell (value, formula, style)
                └── Style (font, fill, border, alignment)
```

### Key Files

| File | Purpose |
|------|---------|
| `lib/XlsxPopulate.js` | Main API entry point |
| `lib/Workbook.js` | Workbook container, file I/O, sheet management |
| `lib/Sheet.js` | Worksheet operations (largest file, 1,555 LOC) |
| `lib/Cell.js` | Individual cell manipulation |
| `lib/Range.js` | Batch operations on cell ranges |
| `lib/Style.js` | Cell styling system |
| `lib/StyleSheet.js` | Style management |
| `lib/Encryptor.js` | ZIP encryption/decryption |
| `lib/addressConverter.js` | A1 notation ↔ row/column conversion |
| `lib/xmlq.js` | XML DOM query utilities |

### Data Flow

1. **Loading**: File → JSZip → XML parsing → Object model
2. **Manipulation**: Modify objects via chainable API
3. **Saving**: Object model → XML building → JSZip → File

### Key Patterns

**Method Overloading (ArgHandler)**: Used throughout for jQuery-style getter/setter methods:
```javascript
method() {
    return new ArgHandler('Method')
        .case(() => { /* getter */ })
        .case('string', arg => { /* setter */ })
        .handle(arguments);
}
```

**Chainable API**: Nearly all operations return `this` for method chaining.

**Promise-Based Async**: All I/O operations use `*Async()` naming and return Promises.

## Code Style

- ES6 with strict mode required
- 4-space indentation
- camelCase for variables/functions, PascalCase for classes
- Arrow functions preferred
- JSDoc comments for all public methods
- No console.log/warn/error

## Test Structure

- **Unit tests**: `test/unit/` - Jasmine specs, uses proxyquire for mocking
- **E2E parse tests**: `test/e2e-parse/` - Parse real Excel files
- **E2E generate tests**: `test/e2e-generate/` - Generate and validate via C# (Windows)
- **Test helpers**: `test/helpers/` - Shared test utilities

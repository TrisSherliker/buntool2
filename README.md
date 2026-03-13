# BunTool

**Automatically create court bundles in seconds.**

Software to take in PDFs and produce perfect English Court bundles.

BunTool2 is the privacy-focussed update to the previous [BunTool](https://github.com/TrisSherliker/buntool). It is a JavaScript library for creating professional PDF court bundles. It processes PDF files and generates merged outputs with:

- Automatic index/table of contents
- Hyperlinked page numbers
- PDF bookmarks/outlines
- Consistent page numbering
- Section breaks

Bundles comply with English court requirements but work well for any context requiring organised PDF compilation.

## Easiest Way to Use

It's hosted at [buntool.co.uk](https://buntool.co.uk) - no installation required, software runs locally on your machine.

## Privacy

BunTool runs entirely in your browser. Your documents are **never uploaded** to any third party server by BunTool - all processing happens locally on your device.

## Usage

If you want to run the library locally without using buntool.co.uk, import the `processTheBundle` function and `Config` as follows:

```javascript
import { processTheBundle } from './buntoolMain.js';
import Config from './buntoolConfig.js';

const config = new Config();
const filesMap = new Map(); // filename -> File object
const indexData = [
  { filename: 'doc1.pdf', title: 'Claim Form', date: '2024-01-15', pageCount: 3 },
  { filename: 'doc2.pdf', title: 'Defence', date: '2024-02-20', pageCount: 5 },
];

const pdfBytes = await processTheBundle(filesMap, indexData, config);
```

Other functions may need importing depending on your needs, but they are documented within the code.

## Files

| File | Purpose |
|------|---------|
| `buntoolMain.js` | Main entry point and bundle orchestration |
| `buntoolFunctions.js` | PDF manipulation functions (merge, paginate, etc.) |
| `buntoolConfig.js` | Configuration class for bundle options |
| `buntoolRestore.js` | Unpack/restore existing bundles - stub in development|

## AI statement

BunTool is hand-coded without the use of AI to generate code. AI was used to design the website frontend (and glue code) at buntool.co.uk, and to review for bugs and generate simple docstrings. 

An exception is the file `buntoolRestore.js`, which was created using a long and winding prompt journey via Claue Code.

## License

Copyright © Tristan Sherliker 2026

Licensed under the Mozilla Public License, version 2.0.


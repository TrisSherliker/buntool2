/**
 * buntoolRestore.js
 *
 * Functions for unpacking and restoring previously-created BunTool bundle PDFs.
 * Extracts embedded metadata, splits bundle into individual documents, and restores configuration.
 */

import * as mupdf from 'https://cdn.jsdelivr.net/npm/mupdf@1.3.6/dist/mupdf.js';
import { PDFDocument } from 'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/+esm';

/**
 * Extracts BunTool metadata from a bundle PDF's hidden annotation.
 * Searches for a FreeText annotation containing "BundleIndexData" on the first page.
 *
 * @param {Uint8Array} pdfBytes - The bundle PDF as a Uint8Array
 * @returns {Array|null} The parsed bundle index metadata array, or null if not found
 */
export function extractBundleMetadata(pdfBytes) {
  try {
    const pdfCopy = new Uint8Array(pdfBytes);
    let doc = mupdf.Document.openDocument(pdfCopy, "application/pdf");

    // Entries are stored in the hidden annotation below.
    // info:BundleIndex contains only config (no entries) to stay under mupdf's ~500-char limit.
    // Read from hidden annotation (all bundles):
    const firstPage = doc.loadPage(0);
    const annotations = firstPage.getAnnotations();

    for (const annot of annotations) {
      const contents = annot.getContents();
      if (typeof contents === 'string' && contents.includes("BundleIndexData:")) {
        // Extract JSON from "BundleIndexData: [...]" or "BundleIndexData: {...}" format
        // Look for either '[' or '{' as the start of JSON
        const bracketIdx = contents.indexOf('[');
        const braceIdx = contents.indexOf('{');

        // Use whichever comes first (and exists)
        let startIdx = -1;
        let isArray = false;
        if (bracketIdx !== -1 && (braceIdx === -1 || bracketIdx < braceIdx)) {
          startIdx = bracketIdx;
          isArray = true;
        } else if (braceIdx !== -1) {
          startIdx = braceIdx;
          isArray = false;
        }

        if (startIdx === -1) continue;

        // Find the matching closing bracket/brace
        let depth = 0;
        let endIdx = -1;
        const openChar = isArray ? '[' : '{';
        const closeChar = isArray ? ']' : '}';

        for (let i = startIdx; i < contents.length; i++) {
          if (contents[i] === openChar) depth++;
          if (contents[i] === closeChar) {
            depth--;
            if (depth === 0) {
              endIdx = i + 1;
              break;
            }
          }
        }

        if (endIdx === -1) {
          console.warn('Could not find matching closing bracket/brace for JSON in annotation');
          continue;
        }

        const jsonString = contents.substring(startIdx, endIdx);
        const parsed = JSON.parse(jsonString);
        return Array.isArray(parsed) ? parsed : [parsed];
      }
    }

    console.warn('No BunTool metadata found in PDF');
    return null;
  } catch (error) {
    console.error('Error extracting bundle metadata:', error);
    return null;
  }
}

/**
 * Splits a bundle PDF into individual documents based on metadata.
 * Uses pdf-lib to extract page ranges for each document.
 *
 * @param {Uint8Array} bundleBytes - The bundle PDF as a Uint8Array
 * @param {Array} metadata - The bundle index metadata array
 * @returns {Promise<Map<string, Uint8Array>>} Map of filename → PDF bytes for each extracted document
 */
export async function splitBundlePdf(bundleBytes, metadata) {
  try {
    // Validate metadata is an array
    if (!Array.isArray(metadata)) {
      console.error('splitBundlePdf received non-array metadata:', typeof metadata, metadata);
      throw new Error(`Invalid metadata: expected array, got ${typeof metadata}`);
    }

    // Load the bundle PDF with pdf-lib
    const bundlePdf = await PDFDocument.load(bundleBytes);
    const totalPages = bundlePdf.getPageCount();

    console.log(`Splitting bundle PDF (${totalPages} pages) into ${metadata.length} items...`);
    console.log('Metadata entries:', metadata);

    // Filter out section breaks - we only split actual documents
    // Section breaks have section: true (and filename: null)
    // Regular documents have section: false and a valid filename
    const documentEntries = metadata.filter(entry => {
      // Only exclude entries that are explicitly marked as sections
      return entry.section !== true;
    });

    if (documentEntries.length === 0) {
      console.warn('No document entries found in metadata');
      return new Map();
    }

    // Calculate TOC length: first document's page number - 1
    // (e.g., if first doc starts at page 3, TOC is pages 1-2, so tocLength = 2)
    const tocLength = documentEntries[0].page - 1;
    console.log(`  TOC length: ${tocLength} pages (first doc starts at page ${documentEntries[0].page})`);

    const extractedFiles = new Map();

    for (let i = 0; i < documentEntries.length; i++) {
      const entry = documentEntries[i];
      const nextEntry = documentEntries[i + 1];

      // Skip entries with invalid filename or page
      if (!entry.filename || entry.page === null || entry.page === undefined) {
        console.warn(`Skipping entry with invalid filename or page:`, entry);
        continue;
      }

      // Calculate page range in bundle (0-indexed)
      // entry.page is 1-indexed bundle page number
      const bundleStartPage = entry.page - 1;
      const bundleEndPage = nextEntry ? nextEntry.page - 1 : totalPages;

      // Create a new PDF for this document
      const docPdf = await PDFDocument.create();

      // Copy pages from bundle to new document
      const pageIndices = [];
      for (let p = bundleStartPage; p < bundleEndPage; p++) {
        pageIndices.push(p);
      }

      const copiedPages = await docPdf.copyPages(bundlePdf, pageIndices);
      copiedPages.forEach(page => docPdf.addPage(page));

      // Save as Uint8Array
      let pdfBytes = await docPdf.save();

      // Remove page numbering (async operation)
      pdfBytes = await removePageNumbering(pdfBytes);

      // Store with filename from metadata
      extractedFiles.set(entry.filename, pdfBytes);

      console.log(`  ✓ Extracted: ${entry.filename} (bundle pages ${bundleStartPage + 1}-${bundleEndPage}, ${bundleEndPage - bundleStartPage} pages)`);
    }

    console.log(`✓ Successfully split bundle into ${extractedFiles.size} documents`);
    return extractedFiles;

  } catch (error) {
    console.error('Error splitting bundle PDF:', error);
    throw new Error(`Failed to split bundle: ${error.message}`);
  }
}

// Unique RGB colour applied to buntool page number footers by pdf-lib.
// In the content stream this appears as: 0.072 0.021 0.073 rg
const BUNTOOL_COLOUR_RE = /0\.072\s+0\.021\s+0\.073\s+rg/;
// Matches the enclosing q...Q graphics state block containing the unique colour.
// pdf-lib wraps each drawText call in its own self-contained q/Q block (no nesting),
// so the non-greedy match is safe.
const FOOTER_BLOCK_RE = /q\b[\s\S]*?0\.072\s+0\.021\s+0\.073\s+rg[\s\S]*?Q\b[ \t]*\n?/g;

/**
 * Removes buntool page number footers from a PDF by targeting the unique colour
 * (0.072 0.021 0.073) used when drawing footer text via pdf-lib. Operates directly
 * on mupdf content streams — no redaction or over-drawing.
 *
 * @param {Uint8Array} pdfBytes - The PDF as a Uint8Array
 * @returns {Promise<Uint8Array>} The PDF with page number footers removed
 */
async function removePageNumbering(pdfBytes) {
  try {
    const pdfCopy = new Uint8Array(pdfBytes);
    const doc = mupdf.Document.openDocument(pdfCopy, "application/pdf");
    const pageCount = doc.countPages();
    let pagesModified = 0;

    for (let i = 0; i < pageCount; i++) {
      const pageDict = doc.findPage(i);
      const contentsRef = pageDict.get('Contents');
      if (!contentsRef || contentsRef.isNull()) continue;

      // Contents can be a single stream or an array of indirect stream refs.
      // Do NOT call .resolve() — the indirect ref retains isStream()=true,
      // but resolve() returns only the dictionary and loses stream access.
      const streamRefs = contentsRef.isArray()
        ? Array.from({ length: contentsRef.length }, (_, j) => contentsRef.get(j))
        : [contentsRef];

      let pageModified = false;

      for (const streamRef of streamRefs) {
        if (!streamRef.isStream()) continue;
        const text = streamRef.readStream().asString();
        if (!BUNTOOL_COLOUR_RE.test(text)) continue;
        FOOTER_BLOCK_RE.lastIndex = 0;
        const cleaned = text.replace(FOOTER_BLOCK_RE, '');
        if (cleaned === text) continue;
        streamRef.writeStream(cleaned);
        pageModified = true;
      }

      if (pageModified) pagesModified++;
    }

    if (pagesModified === 0) return pdfBytes;

    const saved = doc.saveToBuffer("incremental");
    console.log(`Removed page numbers from ${pagesModified}/${pageCount} pages`);
    return saved.asUint8Array().slice(); // .slice() copies out of WASM heap before it's reallocated

  } catch (error) {
    console.error('Error removing page numbering:', error);
    return pdfBytes;
  }
}

/**
 * Parses configuration from PDF metadata fields.
 * Extracts bundle title, project name, and confidential flag.
 *
 * @param {Uint8Array} pdfBytes - The bundle PDF as a Uint8Array
 * @returns {Object} Config object with heading, index, page, and outline options
 */
const DEFAULT_CONFIG = {
  heading: { claimNumber: "", bundleTitle: "", projectName: "", confidential: false },
  pageNumbering: { footerFont: "sansSerif", alignment: "centre", numberingStyle: "PageX", footerPrefix: "" },
  index: { fontFace: "sansSerif", dateStyle: "DD Mon. YYYY", outlineItemStyle: "plain" },
  pageOptions: { printableBundle: false },
};

export function parseConfigFromMetadata(pdfBytes) {
  try {
    const pdfCopy = new Uint8Array(pdfBytes);
    let doc = mupdf.Document.openDocument(pdfCopy, "application/pdf");

    // Primary path: full config embedded in bundle index (v2+ bundles).
    // Wrapped in its own try/catch so a truncated/corrupt JSON falls through to the fallback below.
    try {
      const bundleIndexStr = doc.getMetaData("info:BundleIndex");
      if (bundleIndexStr) {
        const bundleData = JSON.parse(bundleIndexStr);
        if (bundleData && bundleData.version === 2 && bundleData.config) {
          return bundleData.config;
        }
      }
    } catch (e) {
      console.warn('Could not parse info:BundleIndex (truncated or malformed), falling back to standard fields:', e.message);
    }

    // Fallback: reconstruct from individual metadata fields (older bundles)
    const title = doc.getMetaData("Title") || "";
    const subject = doc.getMetaData("Subject") || "";
    const keywords = doc.getMetaData("Keywords") || "";
    const isConfidential = title.startsWith("CONFIDENTIAL ");
    const bundleTitle = isConfidential ? title.substring("CONFIDENTIAL ".length) : title;

    return {
      ...DEFAULT_CONFIG,
      heading: { claimNumber: keywords, bundleTitle, projectName: subject, confidential: isConfidential },
    };

  } catch (error) {
    console.error('Error parsing config from metadata:', error);
    return DEFAULT_CONFIG;
  }
}

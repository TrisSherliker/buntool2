import * as cantoopdfLib from 'https://cdn.jsdelivr.net/npm/@cantoo/pdf-lib@2.3.2/+esm'
import fontkit from 'https://cdn.jsdelivr.net/npm/@pdf-lib/fontkit@1.1.0/+esm'
import * as mupdf from 'https://cdn.jsdelivr.net/npm/mupdf@1.3.6/dist/mupdf.js'
import jsPDF from 'https://cdn.jsdelivr.net/npm/jspdf@3.0.1/+esm'
import jspdfAutotable from 'https://cdn.jsdelivr.net/npm/jspdf-autotable@5.0.2/+esm'
import * as docx from "https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js"
import Config from './buntoolConfig.js';
import { validFonts } from './buntoolConfig.js';

const autoTable = jspdfAutotable;
const pdflib = cantoopdfLib;

/**
 * Font configuration mapping for embedding fonts in PDFs.
 * Fonts are either 'standard' or loaded from a URL. 
 */
const FONT_CONFIG = {
    helvetica: { standard: pdflib.StandardFonts.Helvetica },
    times: { standard: pdflib.StandardFonts.TimesRoman },
    courier: { standard: pdflib.StandardFonts.Courier },
    serif: { url: '/fonts/serif/NotoSerif-Regular.ttf' },
    sansSerif: { url: '/fonts/sans/static/PlusJakartaSans-Regular.ttf' },
    monospaced: { url: '/fonts/mono/UbuntuMono-Regular.ttf' },
    traditional: { url: '/fonts/trad/EBGaramond-VariableFont_wght.ttf' },
};


async function loadFontBytes(pdfDoc, fontName) {
  const fontConfig = FONT_CONFIG[fontName] || FONT_CONFIG['helvetica'];
  if (fontConfig.standard) {
    return await pdfDoc.embedFont(fontConfig.standard);
  } else if (fontConfig.url) {
    const fontBytes = await fetch(fontConfig.url).then(res => res.arrayBuffer());
    return await pdfDoc.embedFont(fontBytes);
  }
}

/*******************************************************
 *              Inner PDF Functions                    *
 *******************************************************/

/**
 * Merges multiple PDF files according to the order specified in the TOC entries.
 * @param {Array<Object>} indexData - Array of TOC entry objects containing filename and metadata (and section headers)
 * @param {Map<string, File>} filesMap - Map of filenames to File objects
 * @param {Config} config - Configuration object containing pageOptions.printableBundle flag
 * @returns {Promise<Uint8Array>} The merged PDF as a Uint8Array
 * @throws {Error} If a file is not found in filesMap or if PDF processing fails
 */
export async function mergePdfsByTOC(tocEntries, filesMap, config) {
  let mergedPdf = await pdflib.PDFDocument.create();
  const printable = config.getOption('pageOptions.printableBundle');
  console.log(`Starting PDF merge. tocEntries: `, tocEntries);
  for (const entry of tocEntries) {
    console.log(`Processing entry: '${entry.title}' (filename: '${entry.filename}')`);
    if (entry.sectionBreak) {
      // section headers don't correspond to files, so we skip them in the merging process
      console.log(`Skipping section header '${entry.title}' in PDF merging`);
    } 
    else {
      let pdfFile;
      try {
        pdfFile = filesMap.get(entry.filename);
        if (!pdfFile) {
          throw new Error(`File not found in filesMap: ${entry.filename}`);
        }
        const pdfBytes = new Uint8Array(await pdfFile.arrayBuffer());
        const inputPdf = await pdflib.PDFDocument.load(pdfBytes);
        const copiedPages = await mergedPdf.copyPages(inputPdf, inputPdf.getPageIndices());
        copiedPages.forEach((page) => mergedPdf.addPage(page));

        // Add blank page if printable mode is enabled and document has odd pages
        // (for printing double-sided allowing tab insertion)
        if (printable) {
          const pageCount = inputPdf.getPageCount();
          if (pageCount % 2 === 1) {
            console.log(`Adding blank page after '${entry.filename}' (${pageCount} pages)`);
            const blankPageBytes = await makeBlankPage();
            const blankPdf = await pdflib.PDFDocument.load(blankPageBytes);
            const [blankPage] = await mergedPdf.copyPages(blankPdf, [0]);
            mergedPdf.addPage(blankPage);
        }
      }

    } catch (error) {
      console.error(`[ERROR] Processing error for file:' ${entry.filename}: `, error);
      throw error;
      }
    }
  }
  const mergedPdfBytes = await mergedPdf.save();
  console.log(`PDFs merged successfully by index`);
  return mergedPdfBytes;
}

/**
 * Creates a single truly blank PDF page with no content.
 * Used for printable bundle mode to ensure proper double-sided alignment.
 * @returns {Promise<Uint8Array>} A single-page blank PDF as a Uint8Array
 */
export async function makeBlankPage() {
  const blankPdf = await pdflib.PDFDocument.create();
  blankPdf.addPage([595.28, 841.89]); // A4 size in points
  return blankPdf.save();
}

/**
 * Creates a single blank PDF page with "This page intentionally left blank" text.
 * @returns {Promise<Uint8Array>} A single-page blank PDF as a Uint8Array
 */
export async function makeIntentionallyBlankPage () {
  const blankPage = await pdflib.PDFDocument.create();
  const page = blankPage.addPage([595.28, 841.89]); // A4 size in points
  page.drawText('This page intentionally left blank.', {
    x: 50,
    y: 400,
    size: 12,
    color: pdflib.rgb(0, 0, 0),
  });
  return blankPage.save();
}

/**
 * Merges two PDF documents into a single PDF.
 * @param {Uint8Array} pdfAbytes - First PDF document as Uint8Array
 * @param {Uint8Array} pdfBbytes - Second PDF document as Uint8Array
 * @returns {Promise<Uint8Array>} The merged PDF as a Uint8Array
 */
export async function mergeTwoPdfs(pdfAbytes, pdfBbytes) {
  //the docs
  const mergedPdf = await pdflib.PDFDocument.create();
  const pdfA = await pdflib.PDFDocument.load(pdfAbytes);
  const pdfB = await pdflib.PDFDocument.load(pdfBbytes);
  // handle A
  const copiedPagesA = await mergedPdf.copyPages(pdfA, pdfA.getPageIndices());
  copiedPagesA.forEach((page) => mergedPdf.addPage(page));
  // handle B
  const copiedPagesB = await mergedPdf.copyPages(pdfB, pdfB.getPageIndices());
  copiedPagesB.forEach((page) => mergedPdf.addPage(page));
  // do the thing
  const mergedPdfBytes = await mergedPdf.save();
  console.log(`Two PDFs merged successfully`);
  return mergedPdfBytes;
}

/**
 * Counts the number of pages in a PDF file.
 * @param {File} file - The PDF file to count pages from
 * @returns {Promise<number>} The number of pages in the PDF
 */
export async function countPdfPages(file) {
  const pdfBytes = await file.arrayBuffer();
  const pdfDoc = await pdflib.PDFDocument.load(pdfBytes);
  return pdfDoc.getPageCount();
}

/**
 * Copies a single page from a source PDF document to a destination PDF document.
 * Uses muPDF API for grafting objects between documents.
 * @param {Object} dstDoc - Destination muPDF document object
 * @param {Object} srcDoc - Source muPDF document object
 * @param {number} pageNumber - Zero-indexed page number to copy
 * @param {Object} dstFromSrc - muPDF graft map for object copying
 */
export function copyPage(dstDoc, srcDoc, pageNumber, dstFromSrc) {
  const srcPage = srcDoc.findPage(pageNumber)
  const dstPage = dstDoc.newDictionary()
  dstPage.put("Type", dstDoc.newName("Page"))
  if (srcPage.get("MediaBox"))
    dstPage.put("MediaBox", dstFromSrc.graftObject(srcPage.get("MediaBox")))
  if (srcPage.get("Rotate"))
    dstPage.put("Rotate", dstFromSrc.graftObject(srcPage.get("Rotate")))
  if (srcPage.get("Resources"))
    dstPage.put("Resources", dstFromSrc.graftObject(srcPage.get("Resources")))
  if (srcPage.get("Contents"))
    dstPage.put("Contents", dstFromSrc.graftObject(srcPage.get("Contents")))
  dstDoc.insertPage(-1, dstDoc.addObject(dstPage))
}

/**
 * Copies all pages from a source PDF document to a destination PDF document.
 * Uses muPDF API for grafting objects between documents.
 * @param {Object} dstDoc - Destination muPDF document object
 * @param {Object} srcDoc - Source muPDF document object
 */
export function copyAllPages(dstDoc, srcDoc) {
  const dstFromSrc = dstDoc.newGraftMap()
  const n = srcDoc.countPages()
  for (let k = 0; k < n; ++k)
    copyPage(dstDoc, srcDoc, k, dstFromSrc)
}

/**
 * Adds page numbering to each page of a PDF document based on config.
 * @param {Uint8Array} pdfDocBytes - The PDF document as a Uint8Array
 * @param {Config} config - Configuration object containing page numbering options
 * @returns {Promise<Uint8Array>} The PDF with page numbers added as a Uint8Array
 */
export async function addPageNumberingToPdf(pdfDocBytes, config) {
  /* This adds a footer to each pdf page, containing the
  * page number with marking (in a configured style), preceded by
  * a prefix if specified.
  * Strategy: use PDF-lib to add a text label to each page.
  */
  const footerLabelText = config.getOption('pageNumbering.footerPrefix')
  const footerAlignment = config.getOption('pageNumbering.alignment')
  const pageNumberingStyle = config.getOption('pageNumbering.numberingStyle')
  const footerFont = config.getOption('pageNumbering.footerFont')
  
  if (pageNumberingStyle === "None") {
    console.log(`No page numbering applied`);
    return pdfDocBytes;
  }
  
  //setup the gubbins
  const pdfDoc = await pdflib.PDFDocument.load(pdfDocBytes);
  let textLabelFont = await pdfDoc.embedFont(pdflib.StandardFonts.Helvetica); 
  let fontBytes = [ ];
  const pages = pdfDoc.getPages();
  pdfDoc.registerFontkit(fontkit);

  textLabelFont = await loadFontBytes(pdfDoc, footerFont);

  //Measurements and sizes
  let textLabelSize = 18
  let totalPageCount = pages.length
  const widestDummyNumber = '8'.repeat(totalPageCount.toString().length); // how wide could the page numbers go? 8 is a big glyph

  // The longest theoreical label is footerLabelText + a big number:
  const labelFormats = {
    'PageX': `Page ${widestDummyNumber}`,
    'PageXofY': `Page ${widestDummyNumber} of ${widestDummyNumber}`,
    'X': `${widestDummyNumber}`,
    'XofY': `${widestDummyNumber} of ${widestDummyNumber}`,
    'XslashY': `${widestDummyNumber}/${widestDummyNumber}`
  };
  const longestLabel = `${footerLabelText} ${labelFormats[pageNumberingStyle] || labelFormats['PageX']}`;
  let maxLabelWidth = textLabelFont.widthOfTextAtSize(longestLabel, textLabelSize)
  let maxLabelHeight = textLabelFont.heightAtSize(textLabelSize)

  //if maxlabelwidth is wider than half the width of a standard a4 portrait page, try decreasing font sizes until it fits:
  const a4Width = 595.28; // A4 width in points
  if (maxLabelWidth > (2 * a4Width / 3)) {
    while (maxLabelWidth > (2 * a4Width / 3)) {
      textLabelSize -= 1;
      maxLabelWidth = textLabelFont.widthOfTextAtSize(longestLabel, textLabelSize);
      maxLabelHeight = textLabelFont.heightAtSize(textLabelSize);
    }
  }

  for (const [pageIdx, thisPage] of pages.entries()) {
    // Construct footer text
    const footerTextFormats = {
      'PageX': `Page ${pageIdx + 1}`,
      'PageXofY': `Page ${pageIdx + 1} of ${totalPageCount}`,
      'X': `${pageIdx + 1}`,
      'XofY': `${pageIdx + 1} of ${totalPageCount}`,
      'XslashY': `${pageIdx + 1}/${totalPageCount}`
    };
    const baseFooterText = footerTextFormats[pageNumberingStyle] || footerTextFormats['PageX'];
    //add zero-width spaces for later searchability, plus any footerLabelText prefix
    const footerText = `\u200B\u200B${footerLabelText ? `${footerLabelText} ` : ''}${baseFooterText}`;

    //alignment calcs
    const marginSidePadding = 30;
    const { width, height } = thisPage.getSize();
    let leftEdgeOfLabel;
    if (footerAlignment === "left") {
      leftEdgeOfLabel = marginSidePadding;
    } else if (footerAlignment === "right") {
      leftEdgeOfLabel = width - maxLabelWidth - marginSidePadding;
    } else if (footerAlignment === "center" || footerAlignment === "centre") {
      const actualLabelWidth = textLabelFont.widthOfTextAtSize(footerText, textLabelSize)
      leftEdgeOfLabel = ((width - actualLabelWidth) / 2);
    } else {
      leftEdgeOfLabel = width - maxLabelWidth - 5;
    }

    //apply text
    thisPage.drawText(footerText, {
      x: leftEdgeOfLabel,
      y: maxLabelHeight,
      size: textLabelSize,
      font: textLabelFont,
      color: pdflib.rgb(0.072, 0.021, 0.073), //unique black
    });
  }
  const pdfOutputBytes = new Uint8Array(await pdfDoc.save());
  
  console.log(`Payload paginated successfully`);
  return pdfOutputBytes;
}

/**
 * Adds clickable hyperlinks to TOC entries that navigate to their corresponding pages.
 * @param {Uint8Array} pdfBytes - The PDF document as a Uint8Array
 * @param {Array<Object>} tocTableRowCoordinates - Array of row coordinate objects with position and dimensions
 * @param {Array<Object>} tocEntries - Array of TOC entry objects containing page references
 * @returns {Uint8Array} The PDF with hyperlinks added as a Uint8Array
 */
export function addHyperlinks(pdfBytes, tocTableRowCoordinates, tocEntries) {
  const pts = (72 / 25.4); //jspdf outputs mm on creation, but mupdf uses pts
  const rowsByPage = groupRowsByPage(tocTableRowCoordinates);

  // Assume inputPdf is a Uint8Array containing PDF data
  // const buffer = Buffer.from(pdfBytes); // Convert Uint8Array to Buffer
  const pdfCopy = new Uint8Array(pdfBytes);
  let doc = mupdf.Document.openDocument(pdfCopy, "application/pdf");

  for (const [pageNumber, rows] of Object.entries(rowsByPage)) {
    const page = doc.loadPage(pageNumber - 1);
    for (const row of rows) {
      const { x, y, width, height, tabNumber} = row;
      const tocEntry = tabNumber 
        ? tocEntries.find(entry => entry.tabNumber === tabNumber) : null //blank for section beaks, no hyperlink needed
      if (!tocEntry) continue; // skip if no matching TOC entry is found
      const destinationPageNumber = (tocEntry.actualStartPage || tocEntry.thisPage) - 1; // mupdf pages are 0-indexed

      page.createLink(
        [x * pts, y * pts, x * pts + width * pts, y * pts + height * pts],
        doc.formatLinkURI(
          {
            type: "XYZ",
            zoom: 100,
            page: destinationPageNumber
          }
        )
      );
    }
    page.update();
  }
  console.log(`Hyperlinks added`);
  const outputPdf = doc.saveToBuffer("incremental").asUint8Array()
  return outputPdf;
}

/**
 * Formats a TOC entry title according to the config outline item style.
 * @param {Object} entry - TOC entry object containing title, date, and page information
 * @param {Config} config - Configuration object containing outline item style preference
 * @returns {string} Formatted outline item text
 */
export function formatOutlineItem(entry, config) {
  const style = config.getOption('index.outlineItemStyle');
  const title = entry.title;
  const date = entry.date;
  const page = entry.actualStartPage ? entry.actualStartPage : entry.thisPage; // fallback to thisPage if actualStartPage is not set (e.g. for section breaks)

  switch (style) {
    case 'withPage':
      return `${title} - pg. ${page}`;

    case 'withDate':
      return date
        ? `${title} (${date})`
        : title;

    case 'withDateandPage':
      return date
        ? `${title} - (${date}) - pg ${page}`
        : `${title}  -pg. ${page}`;

    case 'plain':
    default:
      return title;
  }
}

/**
 * Adds PDF outline (bookmark) items for navigation.
 * Creates an index entry and individual bookmarks for each TOC entry.
 * @param {Uint8Array} pdfBytes - The PDF document as a Uint8Array
 * @param {Array<Object>} tocEntries - Array of TOC entry objects
 * @param {Config} config - Configuration object
 * @returns {Uint8Array} The PDF with outline items added as a Uint8Array
 */
export function addOutlineItems(pdfBytes, tocEntries, config) {

  const pdfCopy = new Uint8Array(pdfBytes);  // TODO: muPDF seems to be clearing and then trying to re-use buffers. Use copy as a temporary fix, but it consumes memory. 
  
  let doc = mupdf.Document.openDocument(pdfCopy, "application/pdf");

  const outlineIterator = doc.outlineIterator();
  // find how many digits in the largest tab number for padding
  const maxTabNumber = Math.max(...tocEntries.map(entry => entry.tabNumber));
  const maxTabNumberLength = maxTabNumber.toString().length;

  // outline item for index
  outlineIterator.insert({
    title: `[${"0".toString().padStart(maxTabNumberLength, '0')}] Index`,
    open: true,
    uri: doc.formatLinkURI({
      page: 0,
      type: "XYZ",
      zoom: 100
    })
  });

  // outline item for each document

  
  tocEntries.forEach(entry => {
    const formattedTitle = formatOutlineItem(entry, config);
    const outlinePage = (entry.actualStartPage || entry.thisPage) - 1; // mupdf pages are 0-indexed
    if (entry.sectionBreak) {
        outlineIterator.insert({
        title: `${formattedTitle}`,
        open: true,
        uri: doc.formatLinkURI({
          page: outlinePage,
          type: "XYZ",
          zoom: 100
        })
      });
    } else {
      outlineIterator.insert({
        title: `[${entry.tabNumber.toString().padStart(maxTabNumberLength, '0')}] ${formattedTitle}`,
        open: true,
        uri: doc.formatLinkURI({
          page: outlinePage,
          type: "XYZ",
          zoom: 100
        })
      });
    }
  });

  //pdfOutputBytes = doc.save();
  console.log(`Outline items added`);
  const outputPdf = doc.saveToBuffer("incremental").asUint8Array()
  return outputPdf;
}

/**
 * Sets PDF metadata including title, subject, producer, and custom bundle index data.
 * @param {Uint8Array} pdfBytes - The PDF document as a Uint8Array
 * @param {Array<Object>} tocEntries - Array of TOC entry objects to store as metadata
 * @param {Config} config - Configuration object containing heading and project information
 * @returns {Uint8Array} The PDF with metadata set as a Uint8Array
 */
export function setMetadata(pdfBytes, tocEntries, config) {
  // const buffer = Buffer.from(pdfBytes); // Convert Uint8Array to Buffer
  const pdfCopy = new Uint8Array(pdfBytes);
  let doc = mupdf.Document.openDocument(pdfCopy, "application/pdf");

  doc.setMetaData("Producer", "BunTool (https://buntool.co.uk)");
  doc.setMetaData("Creator", "BunTool (https://buntool.co.uk)");
  doc.setMetaData(
    "Title",
    config.getOption('heading.confidential')
      ? `CONFIDENTIAL ${config.getOption('heading.bundleTitle')}`
      : config.getOption('heading.bundleTitle')
  );
  doc.setMetaData(
    "Subject",
    config.getOption('heading.projectName')
      ? config.getOption('heading.projectName')
      : ""
  );
  doc.setMetaData(
    "Keywords",
    config.getOption('heading.claimNumber')
      ? config.getOption('heading.claimNumber')
      : ""
  );

  // add custom document metadata field "Bundle Index" which stores tocEntries object:
  const buntoolIndexMetadata = tocEntries.map(entry => ({
    // new index property for ordering (based on position within tocEntries):
    index:  tocEntries.indexOf(entry),
    tab: entry.sectionBreak ? null : entry.tabNumber,
    title: entry.title,
    date: entry.sectionBreak ? null : entry.date,
    section: entry.sectionBreak ? true : false,
    // Use actualStartPage (includes TOC offset) instead of thisPage
    page: entry.sectionBreak ? null : (entry.actualStartPage || entry.thisPage),
    // make new filename to avoid betraying data:
    filename: entry.sectionBreak ? null : `${entry.tabNumber}. ${entry.title} (${entry.date}).pdf`
  }));
  // Store only config in info:BundleIndex (entries are in the annotation below).
  // mupdf getMetaData truncates at ~500 chars; config alone is ~290 chars and fits safely.
  doc.setMetaData("info:BundleIndex", JSON.stringify({
    version: 2,
    config: {
      heading: {
        claimNumber: config.getOption('heading.claimNumber') || '',
        bundleTitle: config.getOption('heading.bundleTitle') || '',
        projectName: config.getOption('heading.projectName') || '',
        confidential: config.getOption('heading.confidential') || false,
      },
      pageNumbering: {
        footerFont: config.getOption('pageNumbering.footerFont') || 'sansSerif',
        alignment: config.getOption('pageNumbering.alignment') || 'centre',
        numberingStyle: config.getOption('pageNumbering.numberingStyle') || 'PageX',
        footerPrefix: config.getOption('pageNumbering.footerPrefix') || '',
      },
      index: {
        fontFace: config.getOption('index.fontFace') || 'sansSerif',
        dateStyle: config.getOption('index.dateStyle') || 'DD Mon. YYYY',
        outlineItemStyle: config.getOption('index.outlineItemStyle') || 'plain',
      },
      pageOptions: {
        printableBundle: config.getOption('pageOptions.printableBundle') ?? false,
      },
    },
  }));

  // add invisibile annotation to first page which stores buntoolIndex as metadata (the annot itself is empty):  
  const firstPage = doc.loadPage(0);
  const metadataAnnotation = firstPage.createAnnotation("FreeText")
  metadataAnnotation.setContents(`BundleIndexData: ${JSON.stringify(buntoolIndexMetadata)}`);
  metadataAnnotation.setRect([0, 0, 0, 0]); // set to zero size
  metadataAnnotation.setOpacity(0) // set to transparent
  metadataAnnotation.setFlags(2) // set to hidden
  metadataAnnotation.setHiddenForEditing(true)


  // pdfOutputBytes = doc.save();
  console.log(`Metadata added`);
  const outputPdf = doc.saveToBuffer("incremental").asUint8Array()
  return outputPdf;
}

/*******************************************************
 *            Table of Contents Generator              *
 *******************************************************/

/**
 * Creates table of contents entries from index data.
 * Processes input documents and section headings, calculating page numbers and tab numbers.
 * @param {Array<Object>} indexData - Array of index entry objects with filename, title, date, pageCount, and secti|| !e.sectionMarkeronMarker
 * @param {Config} config - Configuration object containing pageOptions.printableBundle flag
 * @returns {Promise<Array<Object>>} Array of TOC entry objects with tab numbers, titles, dates, page references, and blankPageAfter flag
 */
/**
 * parse the user's input index data and the provided pdfs
 * to create a table of contents (TOC) entries.
 * input an array of objects filename, title, date, section
 * output an object (mapping) of input documents and section headings
 * each element to contain:
 *   tab number or section number
 *   title,
 *   date,
 *   first page number
 *   filename
**/
 export async function createTocEntries(indexData, config) {
  let tocEntries = [];
  let pdfPageCountTracker = 0;
  let tabNumberTracker = 0;
  let sectionNumberTracker = 0;
  let sectionBeginPage= 0;

  for (const [index, entry] of indexData.entries()) {
    if (entry.sectionMarker === 1) { // section marker test
      sectionNumberTracker++;
      
      // Check if there are any file entries (sectionMarker === 0) after this point
      const hasFilesAfter = indexData.slice(index + 1).some(e => e.sectionMarker === 0  || !e.sectionMarker);
      sectionBeginPage = hasFilesAfter ? pdfPageCountTracker + 1 : pdfPageCountTracker;
      tocEntries.push({
        tabNumber: ``,
        sectionBreak:`Section ${sectionNumberTracker}`,
        title: entry.title,
        date: null,
        thisPage: `${sectionBeginPage}`,
        filename: null
      });
    } else { // for files
      tabNumberTracker++;
      const willAddBlankPage = config.getOption('pageOptions.printableBundle') && (entry.pageCount % 2 === 1);
      
      tocEntries.push({
        tabNumber: tabNumberTracker,
        sectionBreak: null,
        title: entry.title,
        date: entry.date,
        thisPage: pdfPageCountTracker + 1,
        filename: entry.filename,
        blankPageAfter: willAddBlankPage
      });
      
      pdfPageCountTracker += entry.pageCount;
      if (willAddBlankPage) {
        pdfPageCountTracker += 1; // Account for blank page
      }
    }
  }
  console.log(`TOC entries created: `, tocEntries);
  return tocEntries;
}

/**
 * Formats a date string according to the specified style.
 * @param {string} entryDate - Date string in YYYY-MM-DD format
 * @param {string} style - Desired date format style (e.g., "YYYY-MM-DD", "DD-MM-YYYY", "DD Mon. YYYY", etc.)
 * @returns {string} Formatted date string, or empty string if style is "None"
 */
function formatDate(entryDate, style) {

  const monthFullName = [
    "January", 
    "February", 
    "March",    
    "April",
    "May",     
    "June",     
    "July",     
    "August",
    "September",
    "October", 
    "November", 
    "December"
  ];
  const monthShortName = [
    "Jan.", 
    "Feb", 
    "Mar",  
    "Apr",
    "May",   
    "Jun",   
    "Jul",   
    "Aug",
    "Sep",   
    "Oct",   
    "Nov",   
    "Dec"
  ];

  if (!entryDate || entryDate === '') {
    return '';
  }
  
  const [y, m, d] = entryDate.split("-");
  const year = y;
  const monthNumber = Number(m) - 1;
  const day = d.padStart(2, "0");
  switch (style) {
    case "YYYY-MM-DD":
      return `${year}-${m}-${d}`;
    case "DD-MM-YYYY":
      return `${day}-${m}-${year}`;
    case "MM/DD/YYYY":
      return `${m}/${day}/${year}`;
    case "DD Mon. YYYY":
      return `${day} ${monthShortName[monthNumber]} ${year}`;
    case "DD Month YYYY":
      return `${day} ${monthFullName[monthNumber]} ${year}`;
    case "Mon. DD, YYYY":
      return `${monthShortName[monthNumber]} ${day}, ${year}`;
    case "Month DD, YYYY":
      return `${monthFullName[monthNumber]} ${day}, ${year}`;
    case "None":
      return "";
    default:
      return entryDate;
  }
}

/**
 * Generate a PDF with a title, project name and a table that can span multiple pages.
 * Thits is a long function which operates on a single pdf document, and so is self-contained rather than being split into sub-functions.
 * @param {string} title - The title to display at the top of the PDF
 * @param {string} project - The project name to display below the title
 * @param {Array<Array>} tocEntries - Array of arrays containing table data (first array is used as header)
 * @param {Object} options - Configuration options for the PDF
 * @param {Object} options.font - Font configuration
 * @param {string} options.font.family - Font family (default: 'helvetica')
 * @param {number} options.font.sizeTitle - Font size for title (default: 16)
 * @param {number} options.font.sizeProject - Font size for project name (default: 14)
 * @param {number} options.font.sizeTable - Font size for table content (default: 10)
 * @param {Object} options.color - Color configuration
 * @param {string} options.color.headerFill - Background color for header row (default: '#f8f8f8')
 * @param {string} options.color.headerText - Text color for header row (default: '#000000')
 * @param {string} options.color.text - Main text color (default: '#000000')
 * @param {Object} options.table - Table configuration
 * @param {boolean} options.table.showBorders - Whether to show table borders (default: true)
 * @param {number} options.table.cellPadding - Cell padding in mm (default: 3)
 * @param {number} options.table.lineHeight - Line height multiplier (default: 1.2)
 * @param {Object} options.margins - Page margin configuration in mm
 * @param {number} options.margins.top - Top margin (default: 20)
 * @param {number} options.margins.right - Right margin (default: 15)
 * @param {number} options.margins.bottom - Bottom margin (default: 20)
 * @param {number} options.margins.left - Left margin (default: 15)
 * @returns {jsPDF} - The generated PDF document object
 */
export async function makeTocPages(tocEntries, options = {}, config, expectedTocLength = 1) {

  const title = config.getOption('heading.bundleTitle');
  const project = config.getOption('heading.projectName');
  const claimNumber = config.getOption('heading.claimNumber');
  const dateStyle = config.getOption('index.dateStyle');

  // First, add formattedDate property
  tocEntries.forEach(entry => {
    if (entry.date && !entry.formattedDate) {
      entry.formattedDate = formatDate(entry.date, dateStyle);
    }
  });
  

    // For now, toc fine tuning is via an internal config 
    // intended for future development
    // TODO
    const tocInternalConfig = {
      font: {
        family: options.font?.family || 'helvetica',
        sizeTitle: options.font?.sizeTitle || 20,
        sizeProject: options.font?.sizeProject || 14,
        sizeClaimNumber: options.font?.sizeClaimNumber || 12,
        sizeTable: options.font?.sizeTable || 12
      },
      color: {
        headerFill: options.color?.headerFill || [200, 200, 200],
        headerText: options.color?.headerText || [0, 0, 0],
        text: options.color?.text || 0,
      },
      table: {
        showBorders: options.table?.showBorders !== undefined ? options.table.showBorders : true,
        cellPadding: options.table?.cellPadding || 3,
        lineHeight: options.table?.lineHeight || 1.2
      },
      margins: {
        top: options.margins?.top || 20,
        right: options.margins?.right || 25,
        bottom: options.margins?.bottom || 20,
        left: options.margins?.left || 25,
        parPadding: 9
      }
    };

    // Create new PDF document with A4 size
    const doc = new jsPDF({
      orientation: 'portrait',
      unit: 'mm',
      format: 'a4'
    });

    // Set font for the document
    let fontForIndexBytes = [];
    let fontForTitleBytes = [];
    let fontForIndex = 'helvetica';
    let fontForTitle = 'helvetica';

    if (!validFonts.includes(config.getOption('index.fontFace'))) {
      console.warn(`[WARNING] Invalid fontFace option '${config.getOption('index.fontFace')}'. Reverting to 'sansSerif'.`);
      config.updateOptions({ index: { fontFace: 'sansSerif' } });
    }

    switch (config.getOption('index.fontFace')) {
      
      case "serif":
        //Get and set main font:
        fontForIndexBytes = await fetch('/fonts/serif/NotoSerif-Regular.ttf').then(res => { if (!res.ok) throw new Error(`Font fetch failed: ${res.url} (${res.status})`); return res.arrayBuffer(); });
        const base64SerifFont = btoa(
          new Uint8Array(fontForIndexBytes).reduce((s,b)=> s+String.fromCharCode(b), '')
        );
        doc.addFileToVFS('NotoSerif.ttf', base64SerifFont);
        doc.addFont('NotoSerif.ttf', 'NotoSerif', 'normal');
        fontForIndex = 'NotoSerif';

        //Get and set title font:
        fontForTitleBytes = await fetch('/fonts/serif/NotoSerif-Bold.ttf').then(res => { if (!res.ok) throw new Error(`Font fetch failed: ${res.url} (${res.status})`); return res.arrayBuffer(); });
        const base64SerifTitleFont = btoa(
          new Uint8Array(fontForTitleBytes).reduce((s,b)=> s+String.fromCharCode(b), '')
        );
        doc.addFileToVFS('NotoSerifBold.ttf', base64SerifTitleFont);
        doc.addFont('NotoSerifBold.ttf', 'NotoSerifBold', 'bold');
        fontForTitle = 'NotoSerifBold';

        //set font sizes for serif:
        tocInternalConfig.font.sizeClaimNumber = 16;
        tocInternalConfig.font.sizeTitle = 24;
        tocInternalConfig.font.sizeProject = 20;
        //set table font size:
        tocInternalConfig.font.sizeTable = 12;
        break;

      case "sansSerif":
          fontForIndexBytes = await fetch('/fonts/sans/static/PlusJakartaSans-Regular.ttf').then(res => { if (!res.ok) throw new Error(`Font fetch failed: ${res.url} (${res.status})`); return res.arrayBuffer(); });
          const base64SansFont = btoa(
            new Uint8Array(fontForIndexBytes).reduce((s,b)=> s+String.fromCharCode(b), '')
          );
          doc.addFileToVFS('PlusJakartaSans.ttf', base64SansFont);
          doc.addFont('PlusJakartaSans.ttf', 'PlusJakartaSans', 'normal');
          fontForIndex = 'PlusJakartaSans';

          fontForTitleBytes = await fetch('/fonts/sans/static/PlusJakartaSans-Bold.ttf').then(res => { if (!res.ok) throw new Error(`Font fetch failed: ${res.url} (${res.status})`); return res.arrayBuffer(); });
          const base64SansTitleFont = btoa(
            new Uint8Array(fontForTitleBytes).reduce((s,b)=> s+String.fromCharCode(b), '')
          );
          doc.addFileToVFS('PlusJakartaSansBold.ttf', base64SansTitleFont);
          doc.addFont('PlusJakartaSansBold.ttf', 'PlusJakartaSansBold', 'bold');
          fontForTitle = 'PlusJakartaSansBold';

          //set font sizes for sans serif:
          tocInternalConfig.font.sizeClaimNumber = 16;
          tocInternalConfig.font.sizeTitle = 22;
          tocInternalConfig.font.sizeProject = 18;

          //set table font size:
          tocInternalConfig.font.sizeTable = 12;
        break;

      case "monospaced":
        //Get and set main font:
        fontForIndexBytes = await fetch('/fonts/mono/UbuntuMono-Regular.ttf').then(res => { if (!res.ok) throw new Error(`Font fetch failed: ${res.url} (${res.status})`); return res.arrayBuffer(); });
        const base64MonoFont = btoa(
          new Uint8Array(fontForIndexBytes).reduce((s,b)=> s+String.fromCharCode(b), '')
        );
        doc.addFileToVFS('UnbuntuMono.ttf', base64MonoFont);
        doc.addFont('UnbuntuMono.ttf', 'UnbuntuMono', 'normal');
        fontForIndex = 'UnbuntuMono';

        //Get and set title font:
        fontForTitleBytes = await fetch('/fonts/mono/UbuntuMono-Bold.ttf').then(res => { if (!res.ok) throw new Error(`Font fetch failed: ${res.url} (${res.status})`); return res.arrayBuffer(); });
        const base64MonoTitleFont = btoa(
          new Uint8Array(fontForTitleBytes).reduce((s,b)=> s+String.fromCharCode(b), '')
        );
        doc.addFileToVFS('UnbuntuMonoBold.ttf', base64MonoTitleFont);
        doc.addFont('UnbuntuMonoBold.ttf', 'UnbuntuMonoBold', 'bold');
        fontForTitle = 'UnbuntuMonoBold';
        //set font sizes for mono:
        tocInternalConfig.font.sizeClaimNumber = 16;
        tocInternalConfig.font.sizeTitle = 24;
        tocInternalConfig.font.sizeProject = 20;
        //set table font size:
        tocInternalConfig.font.sizeTable = 13;
        break;
      
      
        case "traditional":
        //Get and set main font:
        fontForIndexBytes = await fetch('/fonts/trad/static/EBGaramond-Regular.ttf').then(res => { if (!res.ok) throw new Error(`Font fetch failed: ${res.url} (${res.status})`); return res.arrayBuffer(); });
        const base64TradFont = btoa(
          new Uint8Array(fontForIndexBytes).reduce((s,b)=> s+String.fromCharCode(b), '')
        );
        doc.addFileToVFS('EBGaramond.ttf', base64TradFont);
        doc.addFont('EBGaramond.ttf', 'EBGaramond', 'normal');
        fontForIndex = 'EBGaramond';

        //Get and set title font:
        fontForTitleBytes = await fetch('/fonts/trad/static/EBGaramond-Bold.ttf').then(res => { if (!res.ok) throw new Error(`Font fetch failed: ${res.url} (${res.status})`); return res.arrayBuffer(); });
        const base64TradTitleFont = btoa(
          new Uint8Array(fontForTitleBytes).reduce((s,b)=> s+String.fromCharCode(b), '')
        );
        doc.addFileToVFS('EBGaramondBold.ttf', base64TradTitleFont);
        doc.addFont('EBGaramondBold.ttf', 'EBGaramondBold', 'bold');
        fontForTitle = 'EBGaramondBold';
        //set font sizes for serif:
        tocInternalConfig.font.sizeClaimNumber = 16;
        tocInternalConfig.font.sizeTitle = 24;
        tocInternalConfig.font.sizeProject = 20;
        //set table font size:
        tocInternalConfig.font.sizeTable = 13;
        break;

        default:
        fontForIndexBytes = await fetch('/fonts/sans/static/PlusJakartaSans-Regular.ttf').then(res => { if (!res.ok) throw new Error(`Font fetch failed: ${res.url} (${res.status})`); return res.arrayBuffer(); });
        const base64DefaultFont = btoa(
          new Uint8Array(fontForIndexBytes).reduce((s,b)=> s+String.fromCharCode(b), '')
        );
        doc.addFileToVFS('PlusJakartaSans.ttf', base64DefaultFont);
        doc.addFont('PlusJakartaSans.ttf', 'PlusJakartaSans', 'normal');
        fontForIndex = 'PlusJakartaSans';
        break;
    }

    doc.setFont(fontForIndex);

    // Get page dimensions
    const pageWidth = doc.internal.pageSize.getWidth();
    const border = { top: 0, right: 0, bottom: 1, left: 0 };

    // Add Claim No, right-aligned at the top: 
    doc.setFontSize(tocInternalConfig.font.sizeClaimNumber);
    doc.setTextColor(tocInternalConfig.color.text);
    doc.text(
      claimNumber,
      pageWidth - tocInternalConfig.margins.right,
      tocInternalConfig.margins.top,
      { maxWidth: pageWidth * 0.8, align: 'right' }
    );

    //measure claim number height for positioning of next element, with padding = padding:
    const claimNumberHeight = doc.getTextDimensions(claimNumber, {
      maxWidth: pageWidth * 0.8,
      align: 'right',
    }).h;
    const projectNameYOffset = tocInternalConfig.margins.top + claimNumberHeight + tocInternalConfig.margins.parPadding-5;

    // Add project name
    doc.setFontSize(tocInternalConfig.font.sizeProject);
    doc.text(
      project,
      (pageWidth) / 2,
      projectNameYOffset,
      { maxWidth: pageWidth * 0.9, align: 'center' }
    );

    //measure dims for positioning of next element
    const projectNameDimensions = doc.getTextDimensions(project, {
      maxWidth: pageWidth * 0.9,
      align: 'center',
    });

    const titleYOffset = projectNameYOffset + projectNameDimensions.h + tocInternalConfig.margins.parPadding;

    // Add bundle title with or without confidential laabel
    doc.setFontSize(tocInternalConfig.font.sizeTitle)
    doc.setFont(fontForTitle, 'bold'); //setfotstyle deprecated
    let titleDimensions = {};
    if (config.getOption('heading.confidential')) { //if confidential, some measuring is needed since the red text must be separately rendered. Solution: write all in black, then overwrite the confidential part in red:
      const confiPlusTitle = `CONFIDENTIAL ${config.getOption('heading.bundleTitle')}`;
      const confiPlusTitleDimensions = doc.getTextDimensions(confiPlusTitle, {
        maxWidth: pageWidth * 0.7,
        align: 'center'
      });
      //first write the full line in black:
      doc.setTextColor(tocInternalConfig.color.text);
      doc.text(
        confiPlusTitle,
        pageWidth / 2,
        titleYOffset,
        { maxWidth: pageWidth * 0.7, align: 'center' }
      );

      //Need to get x coordinate of the start of the title, which can vary. So split to strings and measure the width of the first part:
      const linesOfTitle = doc.splitTextToSize(confiPlusTitle, pageWidth * 0.7);
      const firstLineOfTitle = linesOfTitle[0];
      const widthOfFirstLine = doc.getTextDimensions(firstLineOfTitle, {
        maxWidth: pageWidth * 0.7,
        align: 'center'
      }).w;

      const startxOfTitle = (pageWidth - widthOfFirstLine) / 2;

      //now apply red text, overwriting the black:
      doc.setTextColor(210, 43, 43); // Set text color to red
      const confidentialLabel = "CONFIDENTIAL";
      doc.text(
        confidentialLabel,
        startxOfTitle,
        titleYOffset,
        { maxWidth: pageWidth * 0.7, align: 'left' }
      );
      //meaasure title height and width for positioning of next element
      titleDimensions = confiPlusTitleDimensions;

    } else { //if not confidential, just use the title
      doc.setTextColor(tocInternalConfig.color.text); // Reset text color to default
      doc.text(
        title,
        pageWidth / 2,
        titleYOffset,
        { maxWidth: pageWidth * 0.7, align: 'center' }
      );
      //meaasure title height and width for positioning of next element
      titleDimensions = doc.getTextDimensions(title,
        {
          maxWidth: pageWidth * 0.7,
          align: 'center'
        });
    }

    // add tramlines: 
    // width = title width
    // first line positioned abovet title: titleYOffset-5, 
    // second line goes under the title: titleYOffset + titleDimensions.h + 5
    // The -5 and +5 in the x coordinates just extend the lines beyond the title a little
    doc.setLineWidth(0.3);
    doc.setDrawColor(0, 0, 0);
    doc.line(
      ((pageWidth - titleDimensions.w) / 2) - 5,
      titleYOffset - tocInternalConfig.margins.parPadding,
      ((pageWidth + titleDimensions.w) / 2) + 5,
      titleYOffset - tocInternalConfig.margins.parPadding
    );
    doc.line(
      ((pageWidth - titleDimensions.w) / 2) - 5,
      titleYOffset + titleDimensions.h - 3,
      ((pageWidth + titleDimensions.w) / 2) + 5,
      titleYOffset + titleDimensions.h - 3
    );

    // Now move on to set up the table of entries:
    const indexTableYOffset = titleYOffset + titleDimensions.h + tocInternalConfig.margins.parPadding;
    
    // Prepare table data
    // Set actualStartPage on original tocEntries so addHyperlinks/addOutlineItems can use them
    for (const entry of tocEntries) {
      entry.actualStartPage = Number(entry.thisPage) + expectedTocLength;
    }

    const body = tocEntries.map(({ filename, ...rest }) => rest); // Remove the filename field from the tocEntries
    for (const entry of body) { // Clear page number for section breaks in table display
      if (entry.sectionBreak) {
        entry.thisPage = '';
        entry.actualStartPage = '';
      }
    }

    //define autotable content by reference to headers
    const headers = {
      tabNumber: 'Tab',
      title: 'Title',
      formattedDate: 'Date',
      actualStartPage: 'Page'
    }

    const tableWidthSetting = pageWidth - tocInternalConfig.margins.left - tocInternalConfig.margins.right;
    const rowCoordinates = [];

    // Configure autoTable
    autoTable(doc, {
      head: [headers],
      body: body,
      startY: indexTableYOffset, // Start below the title and project name
      margin: {
        top: tocInternalConfig.margins.top,
        right: tocInternalConfig.margins.right,
        bottom: tocInternalConfig.margins.bottom,
        left: tocInternalConfig.margins.left
      },
      styles: {
        fontSize: tocInternalConfig.font.sizeTable,
        cellPadding: tocInternalConfig.table.cellPadding,
        // lineColor: tocInternalConfig.table.showBorders ? 40 : false,
        // lineWidth: tocInternalConfig.table.showBorders ? 0.1 : 0,
        lineWidth: border,
        font: fontForIndex,
        textColor: tocInternalConfig.color.text,
        lineHeight: tocInternalConfig.table.lineHeight
      },
      headStyles: {
        fillColor: tocInternalConfig.color.headerFill,
        textColor: tocInternalConfig.color.headerText,
        font: fontForTitle,
      },
      alternateRowStyles: {
        fillColor: [255, 255, 255]
      },
      // Customize colums
      columnStyles: {
        0: { halign: 'right' }, // Right-align the first column (index 0)
        3: { halign: 'right' }, // Right-align the last column (index 3)  
      },
      tableWidth: tableWidthSetting,
      
      //Shading for section breaks
      didParseCell: (data) => {
        if (data.section === "body" && data.row.raw.sectionBreak) {
          data.cell.styles.fillColor = [225, 225, 225];
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.font = fontForTitle;
          data.cell.styles.halign = 'left';
        }
      },

      // Handle page breaks automatically
      willDrawPage: function (data) {
        // Reset font and colors for each page
        doc.setFont(tocInternalConfig.font.family);
        doc.setTextColor(tocInternalConfig.color.text);
      },

      //Wonderfully, jsPDF autotable reports what it did so the coords can be used later:
      didDrawCell: (data) => {
        // Check if this is the first cell in the row (push once per row)
        if (data.section === "body" && data.column.index === 0) {
          const rowInfo = {
            rowNumber: data.row.index + 1, // Row number (1-based index)
            tabNumber: data.cell.raw, // Tab number from the cell
            x: data.cell.x, // X-coords of the row
            y: data.cell.y, // Y-coords of the row
            width: tableWidthSetting,  // Width of the row (=entire table width)
            height: data.row.height, // Height of the row
            pageNumber: data.pageNumber, // Page number where the row is located
            sectionMarker: tocEntries[data.row.index].sectionBreak ? true : false // Whether this row is a section break
          };
          rowCoordinates.push(rowInfo);
        }
      },
    });
    console.log(`drew autotable with row coordinates: `, rowCoordinates);
    const docBytes = doc.output('arraybuffer'); // Get the PDF as an ArrayBuffer
    const uint8Array = new Uint8Array(docBytes); // Convert ArrayBuffer to Uint8Array
    return [uint8Array, rowCoordinates];
  }

/** 
 * Generates dummy TOC pages to determine how many pages the TOC will take up.
 * This is necessary to calculate the correct page numbers for the actual TOC entries.
 * @param {Array<Object>} tocEntries - Array of TOC entry objects
 * @param {Object} options - Configuration options for the PDF
 * @param {Object} config - Configuration object containing heading and index options
 * @returns {Promise<number>} The number of pages the TOC will occupy
 */
export async function makeDummyTocPages (tocEntries, options = {}, config) {
  let dummyTocPdf, _;
  let dummyTocEntries = tocEntries;
  [dummyTocPdf, _] = await makeTocPages(dummyTocEntries, options, config, 1);

  const doc = mupdf.Document.openDocument(dummyTocPdf, "application/pdf");
  const pagecount = doc.countPages();
  return pagecount;
}

/**
 * Groups table row coordinates by page number for hyperlink creation.
 * @param {Array<Object>} rows - Array of row objects with pageNumber property
 * @returns {Object} Object mapping page numbers to arrays of row coordinates
 */
export function groupRowsByPage(rows) {
  return rows.reduce((acc, row) => {
    if (!acc[row.pageNumber]) acc[row.pageNumber] = [];
    acc[row.pageNumber].push(row);
    return acc;
  }, {});
}



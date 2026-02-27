import Config from './buntoolConfig.js';
import {
  createTocEntries,
  addPageNumberingToPdf,
  makeDummyTocPages,
  makeTocPages,
  mergeTwoPdfs,
  mergePdfsByTOC,
  addHyperlinks,
  addOutlineItems,
  setMetadata,
  getBundleIndexMetadata,
} from './buntoolFunctions.js';

/**
 * Function to process the bundle of PDFs according to the provided configuration.
 * @param {*} filesMap 
 * @param {*} indexData 
 * @param {*} config
 * @returns {Uint8Array} The processed payload PDF as a Uint8Array.
 */
export async function processTheBundle(filesMap, indexData, config){

  if (!filesMap || filesMap.size === 0) {
    throw new Error('Error: No files provided');
  }

  if (!indexData || indexData.length === 0) {
    throw new Error('Error: No index data provided');
  }

  if (!config) {
    throw new Error('Error: No configuration provided');
  }

  let payloadPdf = new Uint8Array();

  // TODO: 
  // This internal toc Options config is sketched out but - 
  // for now - is hardcoded defaults. It's designed for 
  // future expansion. Used in the makeTocPages function.
  const tocOptions = {
    font: {
      family: 'helvetica',
      sizeTitle: 18,
      sizeProject: 14,
      sizeTable: 11
    },
    color: {
      headerFill: [200, 200, 200],
      //  headerText: 150,
      //  text: 250
    },
    table: {
      showBorders: true,
      cellPadding: 2,
      lineHeight: 1.3
    },
    margins: {
      top: 25,
      right: 22,
      bottom: 25,
      left: 22
    }
  };


  console.log('[1/11] Validating configuration structure...');
  try //validate structure with method from buntoolConfig
  {
    config.validateStructure();
  } catch (error) {
      console.error(`[ERROR] Config structure validation error: `, error.message);
      throw error;
  }
  console.log('[1/11]...done')

  console.log('[2/11] Validating configuration options...');
  try { //validate options with method from buntoolConfig
    config.validateOptions();
  } catch (error) {
      console.error(`[ERROR] Config validation error: `, error.message);
      throw error;
  }
  console.log('[2/11]...done')

  console.log('[3/11] Creating TOC entries...');
  let tocEntries;
  try {
    tocEntries = await createTocEntries(indexData, config);
    console.log('[3/11]...done')
  } catch (error) {
    console.error(`[ERROR] Failed to create TOC entries: `, error.message);
    throw error;
  }

  console.log('[4/11] Generating dummy TOC pages...');
  let expectedLengthOfToc = 0;
  try {
    expectedLengthOfToc = await makeDummyTocPages(tocEntries, tocOptions, config);
    console.log(`[4/11]...done - dummy TOC PDF length: ${expectedLengthOfToc} pages`)
  } catch (error) {
    console.error(`[ERROR] Failed to generate dummy TOC pages: `, error.message);
    throw error;
  }

  console.log('[5/11] Generating TOC pages...');
  let tocPdf, tocTableRowCoordinates;
  try {
    [tocPdf, tocTableRowCoordinates] = await makeTocPages(tocEntries, tocOptions, config, expectedLengthOfToc);
    console.log(`[5/11]...done - TOC PDF size: ${tocPdf?.length || 0} bytes`)
  } catch (error) {
    console.error(`[ERROR] Failed to generate TOC pages: `, error.message);
    throw error;
  }

// PDF HANDLING:
  console.log('[6/11] Merging input PDFs...');
  try {
    payloadPdf = await mergePdfsByTOC(tocEntries, filesMap, config);
    console.log(`[6/11]...done - Merged PDF size: ${payloadPdf?.length || 0} bytes`)
  } catch (error) {
    console.error(`[ERROR] Failed to merge input PDFs: `, error.message);
    throw error;
  } 

  console.log('Clearing file references from memory.');
  filesMap.clear(); // Clear the map to free up memory

  console.log('[7/11] Merging TOC with content PDF...');
  try {
    payloadPdf = await mergeTwoPdfs(tocPdf, payloadPdf);
    console.log(`[7/11]...done - Combined PDF size: ${payloadPdf?.length || 0} bytes`)
  } catch (error) {
    console.error(`[ERROR] Failed to merge TOC with content: `, error.message);
    throw error;
  }

  console.log('[8/11] Adding page numbering...');
  try {
    payloadPdf = await addPageNumberingToPdf(payloadPdf, config);
    console.log(`[8/11]...done - PDF size: ${payloadPdf?.length || 0} bytes`)
  } catch (error) {
    console.error(`[ERROR] Failed to add page numbering: `, error.message);
    throw error;
  }

  console.log('[9/11] Adding hyperlinks to TOC entries...');
  try {
    payloadPdf = await addHyperlinks(payloadPdf, tocTableRowCoordinates, tocEntries);
    console.log(`[9/11]...done - PDF size: ${payloadPdf?.length || 0} bytes`)
  } catch (error) {
    console.error(`[ERROR] Failed to add hyperlinks: `, error.message);
    throw error;
  }

  console.log('[10/11] Adding outline items and metadata...');
  try {
    payloadPdf = await addOutlineItems(payloadPdf, tocEntries, config);
    console.log(`[10/11]...done - PDF size: ${payloadPdf?.length || 0} bytes`)
  } catch (error) {
    console.error(`[ERROR] Failed to add outline items: `, error.message);
    throw error;
  }

  console.log('[11/11] Setting PDF metadata...');
  try {
    payloadPdf = await setMetadata(payloadPdf, tocEntries, config);
    console.log(`[11/11]...done - Final PDF size: ${payloadPdf?.length || 0} bytes`)
  } catch (error) {
    console.error(`[ERROR] Failed to set metadata: `, error.message);
    throw error;
  }

  //This is a stump of metadata recovery function
  // getBundleIndexMetadata(payloadPdf);  // DISABLED: muPDF clears the buffer, making payloadPdf empty

  console.log(`✓ Bundle processing complete! Returning PDF of size: ${payloadPdf?.length || 0} bytes`);

  // // Save the final PDF to a file
  // fs.writeFileSync(path.join(outputDirPath,"output.pdf"), payloadPdf);
  // console.log(`PDF saved to ${outputDirPath}`);
  return payloadPdf;
}
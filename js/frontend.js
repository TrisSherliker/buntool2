let processTheBundle;
let countPdfPages;
let chrono;
let draggedRow = null;

import Config from './buntoolConfig.js';
console.log('frontend.js loaded');

const fileInput = document.getElementById('file-input');
const fileTable = document.getElementById('file-table');
const fileTableBody = document.getElementById('file-table-body');
const form = document.getElementById('upload-form');
const csvOutput = document.getElementById('csv-output');
const addSectionBreakBtn = document.getElementById('add-section-break-btn');
const clearAllRowsBtn = document.getElementById('clear-all-rows-btn');
const indexData = [];

// Globals for inputs, files and config:
const filesMap = new Map(); // filename -> File
const frontendInputData = {}; // filename -> { title, date, pages }
const config = new Config();
window.config = config; // Expose config as global


/***********************************
 *  Event Listeners and Handlers   *
 ***********************************/

window.addEventListener('DOMContentLoaded', () => {
  import('./buntoolFunctions.js').then(m => countPdfPages = m.countPdfPages);
  import('./buntoolMain.js').then(m => processTheBundle = m.default ?? m.processTheBundle);
  import('https://esm.sh/chrono-node').then(m => chrono = m);
});

// Drag and Drop Handlers
function handleDragStart(e) {
  draggedRow = this;
  e.dataTransfer.effectAllowed = 'move';
  e.dataTransfer.setData('text/html', this.innerHTML);
  this.style.opacity = '0.4';
}

function handleDragOver(e) {
  if (e.preventDefault) {
    e.preventDefault();
  }
  e.dataTransfer.dropEffect = 'move';
  return false;
}

function handleDrop(e) {
  if (e.stopPropagation) {
    e.stopPropagation();
  }
  if (draggedRow !== this) {
    const allRows = Array.from(fileTableBody.querySelectorAll('tr'));
    const draggedIndex = allRows.indexOf(draggedRow);
    const targetIndex = allRows.indexOf(this);

    if (draggedIndex < targetIndex) {
      this.parentNode.insertBefore(draggedRow, this.nextSibling);
    } else {
      this.parentNode.insertBefore(draggedRow, this);
    }
  }
  return false;
}

function handleDragEnd(e) {
  this.style.opacity = '1';
}

fileInput.addEventListener('change', async (e) => {

  const files = Array.from(e.target.files);

  // Show table if we have files (existing or new)
  if (files.length > 0) {
    fileTable.style.display = 'block';
  }

  // Process each new file
  for (const file of files){
    // Skip if file already exists
    if (filesMap.has(file.name)) {
      console.log(`File ${file.name} already in bundle, skipping`);
      continue;
    }

    filesMap.set(file.name, file);
    const prettyTitle = prettifyTitle(file.name);
    const dateParseObj = await parseDateFromFilename(prettyTitle); // returns .date (as date obj), .name (stripped of date)
    const displayTitle = stripDoubleChars(dateParseObj.name);
    if (!countPdfPages){
      ({countPdfPages} = await import('./buntoolFunctions.js'));
    }
    let pageCount = await countPdfPages(file);
    frontendInputData[file.name] = { title: displayTitle, date: dateParseObj.date, pageCount: pageCount };

    const row = document.createElement('tr');
    row.draggable = true;
    row.dataset.filename = file.name;
    row.classList.add('hover:bg-gray-50', 'transition');
    row.innerHTML = `
      <td class="px-2 py-3 cursor-move">
        <svg class="w-5 h-5 text-gray-400" fill="currentColor" viewBox="0 0 20 20">
          <path d="M10 3a1 1 0 01.707.293l3 3a1 1 0 01-1.414 1.414L10 5.414 7.707 7.707a1 1 0 01-1.414-1.414l3-3A1 1 0 0110 3zM10 17a1 1 0 01-.707-.293l-3-3a1 1 0 011.414-1.414L10 14.586l2.293-2.293a1 1 0 011.414 1.414l-3 3A1 1 0 0110 17z"/>
        </svg>
      </td>
      <td class="px-6 py-3 text-sm text-gray-900">${file.name}</td>
      <td class="px-6 py-3">
        <input type="text" class="title-input w-full px-2 py-1 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-blue-500 focus:border-transparent" data-filename="${file.name}" value="${displayTitle}" />
      </td>
      <td class="px-6 py-3">
        <input type="date" class="date-input w-full px-2 py-1 border border-gray-300 rounded text-sm focus:ring-2 focus:ring-blue-500 focus:border-transparent" data-filename="${file.name}" value="${dateParseObj.date || ''}" />
      </td>
      <td class="px-6 py-3 text-sm text-gray-900">${pageCount}</td>
      <td class="px-6 py-3">
        <button type="button" class="delete-row-btn text-red-600 hover:text-red-800 transition" data-filename="${file.name}" title="Delete row">
          ❌
        </button>
      </td>
    `;

    // Add drag event listeners
    row.addEventListener('dragstart', handleDragStart);
    row.addEventListener('dragover', handleDragOver);
    row.addEventListener('drop', handleDrop);
    row.addEventListener('dragend', handleDragEnd);

    fileTableBody.appendChild(row);
  };

  // Reset file input so same file can be selected again if needed
  fileInput.value = '';
});

fileTableBody.addEventListener('input', (e) => {
  const target = e.target;
  if (target.classList.contains('title-input')) {
    const filename = target.getAttribute('data-filename');
    frontendInputData[filename].title = target.value;
  }
  if (target.classList.contains('date-input')) {
    const filename = target.getAttribute('data-filename');
    frontendInputData[filename].date = target.value;
  }
});

// Handle delete row button clicks
fileTableBody.addEventListener('click', (e) => {
  if (e.target.classList.contains('delete-row-btn')) {
    const filename = e.target.getAttribute('data-filename');

    // Remove from filesMap
    filesMap.delete(filename);

    // Remove from frontendInputData
    delete frontendInputData[filename];

    // Remove row from DOM
    const row = e.target.closest('tr');
    row.remove();

    // Hide table if no rows remain
    if (fileTableBody.querySelectorAll('tr').length === 0) {
      fileTable.style.display = 'none';
    }
  }

  // Handle section break deletion
  if (e.target.classList.contains('delete-section-break-btn')) {
    const row = e.target.closest('tr');
    row.remove();

    // Hide table if no rows remain
    if (fileTableBody.querySelectorAll('tr').length === 0) {
      fileTable.style.display = 'none';
    }
  }
});

// Handle "Clear All Rows" button
clearAllRowsBtn?.addEventListener('click', () => {
  if (confirm('Are you sure you want to clear all documents and section breaks?')) {
    filesMap.clear();
    Object.keys(frontendInputData).forEach(key => delete frontendInputData[key]);
    fileTableBody.innerHTML = '';
    fileTable.style.display = 'none';
  }
});

// Handle "Add Section Break" button
addSectionBreakBtn?.addEventListener('click', () => {
  const sectionBreakRow = document.createElement('tr');
  sectionBreakRow.draggable = true;
  sectionBreakRow.classList.add('section-break-row', 'bg-blue-50', 'border-t-2', 'border-blue-300', 'hover:bg-blue-100', 'transition');
  sectionBreakRow.dataset.sectionBreak = 'true';
  sectionBreakRow.innerHTML = `
    <td class="px-2 py-3 cursor-move">
      <svg class="w-5 h-5 text-blue-400" fill="currentColor" viewBox="0 0 20 20">
        <path d="M10 3a1 1 0 01.707.293l3 3a1 1 0 01-1.414 1.414L10 5.414 7.707 7.707a1 1 0 01-1.414-1.414l3-3A1 1 0 0110 3zM10 17a1 1 0 01-.707-.293l-3-3a1 1 0 011.414-1.414L10 14.586l2.293-2.293a1 1 0 011.414 1.414l-3 3A1 1 0 0110 17z"/>
      </svg>
    </td>
    <td colspan="4" class="px-6 py-3 text-center">
      <input type="text" class="section-break-title w-full px-3 py-1 border border-blue-300 rounded bg-white text-blue-700 font-semibold text-center focus:ring-2 focus:ring-blue-500 focus:border-transparent" value="— SECTION BREAK —" placeholder="Section name..."/>
    </td>
    <td class="px-6 py-3">
      <button type="button" class="delete-section-break-btn text-red-600 hover:text-red-800 transition" title="Delete section break">
        ❌
      </button>
    </td>
  `;

  // Add drag event listeners
  sectionBreakRow.addEventListener('dragstart', handleDragStart);
  sectionBreakRow.addEventListener('dragover', handleDragOver);
  sectionBreakRow.addEventListener('drop', handleDrop);
  sectionBreakRow.addEventListener('dragend', handleDragEnd);

  // Add to end of table
  fileTableBody.appendChild(sectionBreakRow);

  // Show table if hidden
  fileTable.style.display = 'block';
});

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  //dynamic (lazy) load the main module
  if (!processTheBundle) {
    ({ processTheBundle } = await import('./buntoolMain.js'));
  };

  // Gather config options from the form
  const configOptions = {
    heading: {
      claimNumber: stripUnsuitableChars(document.getElementById('config-claimNumber').value),
      bundleTitle: stripUnsuitableChars(document.getElementById('config-bundleTitle').value),
      projectName: stripUnsuitableChars(document.getElementById('config-projectName').value),
      confidential: document.getElementById('config-confidential').checked,
    },
    pageNumbering: {
      footerFont: document.getElementById('config-footerFont').value,
      alignment: document.getElementById('config-alignment').value,
      numberingStyle: document.getElementById('config-numberingStyle').value,
      footerPrefix: stripUnsuitableChars(document.getElementById('config-footerPrefix').value),
    },
    index: {
      fontFace: document.getElementById('config-fontFace').value,
      dateStyle: document.getElementById('config-dateStyle').value,
      outlineItemStyle: document.getElementById('config-outlineItemStyle').value,
    },
    pageOptions: {
      printableBundle: document.getElementById('config-printableBundle').value === 'true',
    }
  };
  
  config.updateOptions(configOptions);
  console.log('Config pushed:',JSON.stringify(config));

  // Build indexData array in table order (including section breaks)
  indexData.length = 0; // Clear any previous indexData for repeat uses in same ssn
  const rows = fileTableBody.querySelectorAll('tr');
  rows.forEach(row => {
    // Check if this is a section break
    if (row.dataset.sectionBreak === 'true') {
      const sectionTitleInput = row.querySelector('.section-break-title');
      const sectionTitle = sectionTitleInput ? sectionTitleInput.value : '—';
      indexData.push({
        sectionMarker: 1,  // Indicates section break
        title: sectionTitle
      });
    } else {
      // Regular document row - get filename from second column (first is drag handle)
      const filenameTd = row.querySelectorAll('td')[1];
      if (filenameTd) {
        const filename = filenameTd.textContent.trim();
        if (frontendInputData[filename]) {
          indexData.push({
            filename,
            title: frontendInputData[filename].title,
            date: frontendInputData[filename].date,
            pageCount: frontendInputData[filename].pageCount,
            sectionMarker: 0
          });
        }
      }
    }
  });

  try {
    const pdfBytes = await processTheBundle(filesMap, indexData, config);

    // Validate that we got valid PDF bytes
    if (!pdfBytes || !(pdfBytes instanceof Uint8Array) || pdfBytes.length === 0) {
      throw new Error('Bundle processing returned invalid or empty PDF data');
    }

    const returnedBlob = new Blob([pdfBytes], { type: 'application/pdf' });
    const url = URL.createObjectURL(returnedBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'burnBundleTest.pdf';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 0);
  } catch (error) {
    console.error('[FRONTEND ERROR] Bundle generation failed:', error);
    alert(`Failed to generate bundle:\n\n${error.message}\n\nCheck the browser console for more details.`);
  }
  
});


/***********************************
 *       Frontend Functions        *
***********************************/


async function parseDateFromFilename(filename) {
  let matchedDate = null;
  let filenameWithoutDate = filename;

  // Check for filenames that start with YYYY-MM-DD or DD-MM-YYYY
  const yearFirstDateRegex = /[\[\(]{0,1}(1\d{3}|20\d{2})[-._]?(0[1-9]|1[0-2])[-._]?(0[1-9]|[12][0-9]|3[01])[\]\)]{0,1}/;
  const yearLastDateRegex = /[\[\(]{0,1}(0[1-9]|[12][0-9]|3[01])[-._]?(0[1-9]|1[0-2])[-._]?(1\d{3}|20\d{2})[\]\)]{0,1}/;

  const yearFirstMatch = filename.match(yearFirstDateRegex);
  if (yearFirstMatch) {
    const [fullMatch, year, month, day] = yearFirstMatch;
    const parsedDate = new Date(`${year}-${month}-${day}T00:00:00Z`);
    matchedDate = parsedDate.toISOString().split('T')[0];
    filenameWithoutDate = filenameWithoutDate.replace(fullMatch, '').replace(/^[\s-_]+|[\s-_]+$/g, '');
    return { date: matchedDate, name: filenameWithoutDate };
  }

  const yearLastMatch = filename.match(yearLastDateRegex);
  if (yearLastMatch) {
    const [fullMatch, day, month, year] = yearLastMatch;
    const parsedDate = new Date(`${year}-${month}-${day}T00:00:00Z`);
    matchedDate = parsedDate.toISOString().split('T')[0];
    filenameWithoutDate = filenameWithoutDate.replace(fullMatch, '').replace(/^[\s-_]+|[\s-_]+$/g, '');
    return { date: matchedDate, name: filenameWithoutDate };
  }

  // Fall back to chrono-node for natural language processing
  let chronoParsedResult = [];
  if (typeof chrono !== 'undefined') {
    console.log('filename being parsed:', filename);
    chronoParsedResult = chrono.parse(filename);
  }
  if (chronoParsedResult.length > 0) {
    const parsedDate = chronoParsedResult[0].start.date();
    matchedDate = parsedDate.toISOString().split('T')[0];
    const matchedInputText = chronoParsedResult[0].text;
    console.log('matchedInputText:', matchedInputText);
    console.log('matchedDate:', matchedDate);
    filenameWithoutDate = filenameWithoutDate.replace(matchedInputText, '').replace(/^[\s-_]+|[\s-_]+$/g, '');
    console.log('filenameWithoutDate:', filenameWithoutDate);
    return { date: matchedDate, name: filenameWithoutDate };
  }

  return { date: null, name: filenameWithoutDate };
}

function prettifyTitle(title) {
  // trim off file extension: 
  title = title.replace(/\.[a-zA-Z0-9]{1,4}$/, '');
  // Replace multiple underscores with a single space
  title = title.replace(/_+/g, ' ');
  // Remove any character that is not a word character, space, or punctuation:
  title = title.replace(/[^\p{L}\p{N}\p{P}\p{S}\p{Z}]/gu, ''); // Unicode-aware regex: L is letter, N is number, P is punctuation, S is symbol, Z is separator
  // if any double spaces, underscores or hyphens which might result from the above:
  title = stripDoubleChars(title);
  return title.trim();
}


function stripDoubleChars(str) {
  // Replace multiple spaces, underscores, stops or hyphens with a single space
  str = str.replace(/[_\s\-.,\\/]+/g, ' ');
  return str.trim();
}

function stripUnsuitableChars(input) {
  return input
    // 1) strip out all emoji / pictographic codepoints
    .replace(/\p{Extended_Pictographic}/gu, '')
    // 2) strip out control characters and anything not in these Unicode categories:
    //    L = Letter, N = Number, P = Punctuation, S = Symbol, Z = Separator
    .replace(/[^\p{L}\p{N}\p{P}\p{S}\p{Z}]/gu, '')
    // 3) collapse multiple spaces/tabs/newlines to a single space
    .replace(/\s+/g, ' ')
    .trim();
}

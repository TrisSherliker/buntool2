// Moveable Tables Demo for Tabulator v6.3+
// This file demonstrates two tables with movable rows and row handles as per Tabulator documentation

document.addEventListener('DOMContentLoaded', function() {
  // Example 0: Nested table with moveable rows and rowHeader
  var dummyData = [
    {id:1, SectionTitle:"Section 1", documents:[
      {filename:"doc1.pdf", title:"Document 1", date:"2010-01-01", pageCount:10},
      {filename:"doc2.pdf", title:"Document 2", date:"2011-01-02", pageCount:20},
      {filename:"doc3.pdf", title:"Document 3", date:"2012-01-03", pageCount:30}
    ]},
    {id:2, SectionTitle:"Section 2", documents:[
      {filename:"doc4.pdf", title:"Document 4", date:"2013-01-04", pageCount:40},
      {filename:"doc5.pdf", title:"Document 5", date:"2014-01-05", pageCount:50},
      {filename:"doc6.pdf", title:"Document 6", date:"2015-01-06", pageCount:60}
    ]},
    {id:3, SectionTitle:"Section 3", documents:[
      {filename:"doc7.pdf", title:"Document 7", date:"2016-01-07", pageCount:70},
      {filename:"doc8.pdf", title:"Document 8", date:"2017-01-08", pageCount:80},
      {filename:"doc9.pdf", title:"Document 9", date:"2018-01-09", pageCount:90}
    ]}
  ];

  var subTableInstances = [];

  var table0 = new Tabulator("#example-table0", {
    layout: "fitColumns",
    movableRows: true,
    rowHeader: {headerSort:false, resizable: false, minWidth:30, width:30, rowHandle:true, formatter:"handle"},
    height: "100%",
    data: dummyData,
    columns: [
      {title: "Section", field: "SectionTitle"}
    ],
    rowFormatter: function(row) {
      var sectionId = row.getData().id;
      var holderElement = document.createElement("div");
      var tableElement = document.createElement("div");
      // Assign a unique id to each subtable div
      var subTableId = `subtable-section-${sectionId}`;
      tableElement.id = subTableId;
      holderElement.style.padding = "10px 30px 10px 31px";
      tableElement.style.border = "1px solid #ccc";
      holderElement.appendChild(tableElement);
      row.getElement().appendChild(holderElement);
      // Create the subtable instance
      var subTable = new Tabulator(tableElement, {
        layout: "fitColumns",
        movableRows: true,
        rowHeader: {headerSort:false, resizable: false, minWidth:30, width:30, rowHandle:true, formatter:"handle"},
        data: row.getData().documents,
        columns: [
          {title: "Filename", field: "filename"},
          {title: "Title", field: "title"},
          {title: "Date", field: "date"},
          {title: "Page Count", field: "pageCount"}
        ]
      });
      // Store the subtable instance and selector
      subTableInstances.push({id: subTableId, instance: subTable});
    }
  });

  // After all rows are formatted, connect the subtables for row moving
  setTimeout(function() {
    var selectors = subTableInstances.map(t => `#${t.id}`);
    subTableInstances.forEach(t => {
      t.instance.setOptions({
        movableRowsConnectedTables: selectors.filter(sel => sel !== `#${t.id}`)
      });
    });
  }, 500);

  // Example 1: Minimal moveable rows with rowHeader
  var table1 = new Tabulator("#example-table3", {
    movableRows: true,
    rowHeader: {headerSort:false, resizable: false, minWidth:30, width:30, rowHandle:true, formatter:"handle"},
    columns:[
      {title:"Name", field:"name"},
      {title:"Progress", field:"progress"},
    ],
    data:[
      {id:1, name:"Alpha", progress:10},
      {id:2, name:"Bravo", progress:20},
      {id:3, name:"Charlie", progress:30},
    ]
  });

  // Example 2: More complex table with moveable rows and rowHeader
  var table2 = new Tabulator("#example-table2", {
    height:"311px",
    movableRows:true,
    rowHeader: {headerSort:false, resizable: false, minWidth:30, width:30, rowHandle:true, formatter:"handle"},
    columns:[
      {title:"Name", field:"name", width:150},
      {title:"Progress", field:"progress", formatter:"progress", sorter:"number"},
      {title:"Gender", field:"gender"},
      {title:"Rating", field:"rating", formatter:"star", formatterParams:{stars:6}, hozAlign:"center", width:120},
      {title:"Favourite Color", field:"col"},
      {title:"Date Of Birth", field:"dob", hozAlign:"center", sorter:"date"},
      {title:"Driver", field:"car", hozAlign:"center", formatter:"tickCross"},
    ],
    data:[
      {id:1, name:"John", progress:45, gender:"Male", rating:4, col:"Red", dob:"1990-01-01", car:true},
      {id:2, name:"Jane", progress:78, gender:"Female", rating:5, col:"Blue", dob:"1985-05-15", car:false},
      {id:3, name:"Steve", progress:60, gender:"Male", rating:3, col:"Green", dob:"1992-07-20", car:true},
      {id:4, name:"Mary", progress:90, gender:"Female", rating:6, col:"Yellow", dob:"1988-03-10", car:true},
      {id:5, name:"Paul", progress:30, gender:"Male", rating:2, col:"Purple", dob:"1995-11-25", car:false},
    ]
  });

  // --- File Upload and Flat Table Demo with Tabulator ---
  var fileData = [];

  // Create upload input and add section button
  var controlsDiv = document.createElement('div');
  controlsDiv.style.marginBottom = '1em';
  controlsDiv.innerHTML = `
    <input type="file" id="demo-upload-input" multiple accept="application/pdf" style="margin-right:1em;">
    <button id="add-section-header-btn" type="button">Add Section Header</button>
  `;
  document.body.insertBefore(controlsDiv, document.getElementById('example-table'));

  // Tabulator table instance
  var table = new Tabulator("#example-table", {
    layout: "fitColumns",
    movableRows: true,
    rowHeader: {headerSort:false, resizable: false, minWidth:30, width:30, rowHandle:true, formatter:"handle"},
    columns: [
      {title: "Filename", field: "filename", editor: false},
      {title: "Title", field: "title", editor: "input", placeholder: "", cellEdited: function(cell) { cell.getRow().update({title: cell.getValue()}); }},
      {title: "Date", field: "date", editor: "date", editorParams: {elementAttributes: {placeholder: "YYYY-MM-DD"}}},
      {title: "Pages", field: "pageCount", editor: false, hozAlign: "center"}
    ],
    data: fileData,
    height: 400
  });

  // File upload handler
  document.getElementById('demo-upload-input').addEventListener('change', async function(e) {
    const files = Array.from(e.target.files);
    for (const file of files) {
      // Use PDF.js to get page count
      let pageCount = '';
      try {
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        pageCount = pdf.numPages;
      } catch (err) {
        pageCount = '?';
      }
      fileData.push({
        id: Date.now() + Math.random(),
        filename: file.name,
        title: file.name.replace(/\.[^.]+$/, ''),
        date: '',
        pageCount: pageCount
      });
    }
    table.replaceData(fileData);
    this.value = '';
  });

  // Add Section Header button handler
  document.getElementById('add-section-header-btn').addEventListener('click', function() {
    fileData.unshift({
      id: Date.now() + Math.random(),
      filename: '',
      title: '',
      date: '',
      pageCount: '',
      isSection: true
    });
    table.replaceData(fileData);
    // Set placeholder for the top row's title cell
    setTimeout(() => {
      var row = table.getRowFromPosition(0, true);
      if(row) {
        var cell = row.getCell('title');
        if(cell) {
          cell.getElement().querySelector('input')?.setAttribute('placeholder', 'enter section name');
        }
      }
    }, 100);
  });

  // Custom row formatter for section header styling
  table.setRowFormatter(function(row) {
    var data = row.getData();
    if (data.isSection) {
      row.getElement().style.background = '#f3f6fa';
      row.getElement().style.fontWeight = 'bold';
      row.getElement().style.fontStyle = 'italic';
      row.getElement().style.color = '#4a6a8a';
      // Hide filename, date, pageCount cells
      row.getCell('filename')?.getElement().innerHTML = '';
      row.getCell('date')?.getElement().innerHTML = '';
      row.getCell('pageCount')?.getElement().innerHTML = '';
    }
  });
});

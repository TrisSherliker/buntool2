document.addEventListener('DOMContentLoaded', function() {
  var fileData = [];

  // Tabulator table instance
  var table = new Tabulator("#tabulator-table", {
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

  document.getElementById('tabulator-upload-input').addEventListener('change', async function(e) {
    const files = Array.from(e.target.files);
    for (const file of files) {
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

  document.getElementById('tabulator-add-section-btn').addEventListener('click', function() {
    fileData.unshift({
      id: Date.now() + Math.random(),
      filename: '',
      title: '',
      date: '',
      pageCount: '',
      isSection: true
    });
    table.replaceData(fileData);
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

  table.setRowFormatter(function(row) {
    var data = row.getData();
    if (data.isSection) {
      row.getElement().style.background = '#f3f6fa';
      row.getElement().style.fontWeight = 'bold';
      row.getElement().style.fontStyle = 'italic';
      row.getElement().style.color = '#4a6a8a';
      row.getCell('filename')?.getElement().innerHTML = '';
      row.getCell('date')?.getElement().innerHTML = '';
      row.getCell('pageCount')?.getElement().innerHTML = '';
    }
  });
});
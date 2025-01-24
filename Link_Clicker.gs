function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Link Tools')
      .addItem('Open Selected Links', 'openSelectedLinks')
      .addToUi();
}

function openSelectedLinks() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  var richText = range.getRichTextValues();
  var links = [];

  // Loop through all cells in selection
  for (var row = 0; row < richText.length; row++) {
    for (var col = 0; col < richText[0].length; col++) {
      var cell = richText[row][col];
      // Get the link for the entire cell
      var url = cell.getLinkUrl();
      if (url) {
        links.push(url);
      } else {
        // Try getting links from individual runs within the cell
        var runs = cell.getRuns();
        for (var i = 0; i < runs.length; i++) {
          var runUrl = runs[i].getLinkUrl();
          if (runUrl) {
            links.push(runUrl);
          }
        }
      }
    }
  }

  if (links.length === 0) {
    SpreadsheetApp.getUi().alert('Please select cells that contain clickable links (cells that turn blue and are clickable)');
    return;
  }

  var htmlOutput = HtmlService
    .createHtmlOutput('<script>var links = ' + JSON.stringify(links) + ';' +
    'links.forEach(function(link) { window.open(link, "_blank"); });</script>')
    .setWidth(100)
    .setHeight(50);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Opening ' + links.length + ' links...');
}

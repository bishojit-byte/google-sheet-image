/**
 * @OnlyCurrentDoc
 * Google Apps Script for Google Sheets: Bulk Insert Images into Cells (Horizontally / Vertically)
 * GitHub Repository: https://github.com/
 * Local Path: E:\Applications\Google Sheet Appscripts
 */

/**
 * Creates custom menu in Google Sheets UI on document open.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡ Image Inserter')
    .addItem('Bulk Insert Images...', 'showImageInserterDialog')
    .addToUi();
}

/**
 * Displays the HTML modal dialog for bulk image insertion.
 */
function showImageInserterDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(540)
    .setHeight(620)
    .setTitle('Bulk Insert Images into Cells');
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Bulk Insert Images into Cells');
}

/**
 * Processes list of image URLs and inserts them into cells starting from the active cell.
 * 
 * @param {Object} options Configuration options passed from the dialog HTML.
 * @param {string[]} options.urls Array of image URLs to insert.
 * @param {string} options.direction 'horizontal' or 'vertical'.
 * @param {string} options.mode 'cellImage' (native Apps Script in-cell image) or 'formula' (=IMAGE("url")).
 * @param {boolean} options.autoResize Option to automatically adjust cell width/height for better visibility.
 * @return {Object} Result payload with count and status message.
 */
function processImageUrls(options) {
  try {
    if (!options || !options.urls || !Array.isArray(options.urls) || options.urls.length === 0) {
      return { success: false, error: 'No valid image URLs provided.' };
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const activeRange = sheet.getActiveCell();
    if (!activeRange) {
      return { success: false, error: 'No active cell selected. Please click on a starting cell in your sheet.' };
    }

    const startRow = activeRange.getRow();
    const startCol = activeRange.getColumn();

    const urls = options.urls;
    const direction = options.direction || 'vertical'; // 'horizontal' or 'vertical'
    const mode = options.mode || 'cellImage'; // 'cellImage' or 'formula'
    const autoResize = options.autoResize !== false;

    let insertedCount = 0;

    for (let i = 0; i < urls.length; i++) {
      const url = urls[i].trim();
      if (!url) continue;

      let targetRow, targetCol;
      if (direction === 'horizontal') {
        targetRow = startRow;
        targetCol = startCol + i;
      } else { // vertical
        targetRow = startRow + i;
        targetCol = startCol;
      }

      const cell = sheet.getRange(targetRow, targetCol);

      if (mode === 'cellImage') {
        try {
          const image = SpreadsheetApp.newCellImage()
            .setSourceUrl(url)
            .setAltTextDescription('Inserted via Bulk Image Inserter')
            .build();
          cell.setValue(image);
          insertedCount++;
        } catch (err) {
          // Fallback to formula if native CellImage fails for specific URL format
          cell.setValue(`=IMAGE("${url}")`);
          insertedCount++;
        }
      } else {
        // Formula mode (=IMAGE("url"))
        cell.setValue(`=IMAGE("${url}")`);
        insertedCount++;
      }
    }

    // Optionally adjust row heights and column widths so images display clearly
    if (autoResize && insertedCount > 0) {
      if (direction === 'horizontal') {
        for (let i = 0; i < insertedCount; i++) {
          const col = startCol + i;
          if (sheet.getColumnWidth(col) < 120) {
            sheet.setColumnWidth(col, 120);
          }
        }
        sheet.setRowHeight(startRow, Math.max(sheet.getRowHeight(startRow), 100));
      } else {
        const col = startCol;
        if (sheet.getColumnWidth(col) < 120) {
          sheet.setColumnWidth(col, 120);
        }
        for (let i = 0; i < insertedCount; i++) {
          const row = startRow + i;
          sheet.setRowHeight(row, Math.max(sheet.getRowHeight(row), 100));
        }
      }
    }

    const startCellName = getCellName(startRow, startCol);
    return {
      success: true,
      insertedCount: insertedCount,
      startCell: startCellName,
      direction: direction,
      message: `Successfully inserted ${insertedCount} image(s) ${direction}ly starting at cell ${startCellName}.`
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Helper to convert row and column indices into standard A1 cell notation (e.g. 1, 1 -> A1).
 */
function getCellName(row, col) {
  let temp, letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(65 + temp) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter + row;
}

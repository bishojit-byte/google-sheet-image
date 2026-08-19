/**
 * @OnlyCurrentDoc
 * Google Apps Script for Google Sheets: Bulk Insert Images into Cells (Horizontally / Vertically)
 * GitHub Repository: https://github.com/bishojit-byte/google-sheet-image
 * Local Path: E:\Applications\Google Sheet Appscripts
 */

/**
 * Folder name used in Google Drive to store uploaded images.
 */
const IMAGE_FOLDER_NAME = 'Google Sheet Images';

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
    .setHeight(720)
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
    return insertImageUrls(options.urls, options.direction, options.mode, options.autoResize);
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Uploads images selected from the local computer to Google Drive, then inserts them into cells.
 *
 * @param {Object} options Configuration options passed from the dialog HTML.
 * @param {Object[]} options.files Array of files with { name, mimeType, dataUrl } (dataUrl is base64).
 * @param {string} options.direction 'horizontal' or 'vertical'.
 * @param {string} options.mode 'cellImage' or 'formula'.
 * @param {boolean} options.autoResize Option to automatically adjust cell width/height.
 * @return {Object} Result payload with count and status message.
 */
function processUploadedImages(options) {
  try {
    const files = (options && options.files) || [];
    if (!Array.isArray(files) || files.length === 0) {
      return { success: false, error: 'No images selected to upload.' };
    }

    const folder = getOrCreateImageFolder();
    const urls = [];
    let uploadedCount = 0;

    for (const f of files) {
      if (!f || !f.dataUrl || typeof f.dataUrl !== 'string') continue;
      const commaIndex = f.dataUrl.indexOf(',');
      if (commaIndex === -1) continue;

      const base64 = f.dataUrl.substring(commaIndex + 1);
      const mimeType = f.mimeType || 'image/png';
      const rawName = f.name || 'uploaded_image.png';

      let blob;
      try {
        blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, rawName);
      } catch (err) {
        continue;
      }

      // Use a unique file name to prevent collisions in Google Drive.
      const ext = rawName.replace(/\.[^.]+$/, '').slice(0, 80);
      const uniqueName = `${ext}_${Date.now()}_${uploadedCount}.${mimeType.split('/')[1] || 'png'}`;

      const created = folder.createFile(uniqueName, blob);
      created.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      urls.push(thumbnailUrlFor(created.getId()));
      uploadedCount++;
    }

    if (uploadedCount === 0) {
      return { success: false, error: 'Could not upload any image. Please try again.' };
    }

    const result = insertImageUrls(urls, options.direction, options.mode, options.autoResize);
    result.uploadedFiles = uploadedCount;
    result.folderName = IMAGE_FOLDER_NAME;
    result.message = `Uploaded ${uploadedCount} image(s) to Drive folder "${IMAGE_FOLDER_NAME}". ${result.message}`;
    return result;
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Finds an existing "Google Sheet Images" folder or creates a new one.
 *
 * @return {Folder} The Google Drive folder used to store uploaded images.
 */
function getOrCreateImageFolder() {
  const iter = DriveApp.getFoldersByName(IMAGE_FOLDER_NAME);
  if (iter.hasNext()) {
    return iter.next();
  }
  return DriveApp.createFolder(IMAGE_FOLDER_NAME);
}

/**
 * Builds a direct embeddable thumbnail URL for a Google Drive file id.
 *
 * @param {string} fileId The Google Drive file id.
 * @return {string} A thumbnail URL usable with CellImage or =IMAGE().
 */
function thumbnailUrlFor(fileId) {
  return 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1000';
}

/**
 * Core insertion logic shared by URL and upload flows.
 *
 * @param {string[]} urls Array of image URLs to insert.
 * @param {string} direction 'horizontal' or 'vertical'.
 * @param {string} mode 'cellImage' or 'formula'.
 * @param {boolean} autoResize Option to automatically adjust cell width/height.
 * @return {Object} Result payload with count and status message.
 */
function insertImageUrls(urls, direction, mode, autoResize) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeRange = sheet.getActiveCell();
  if (!activeRange) {
    return { success: false, error: 'No active cell selected. Please click on a starting cell in your sheet.' };
  }

  const startRow = activeRange.getRow();
  const startCol = activeRange.getColumn();

  const dir = direction || 'vertical'; // 'horizontal' or 'vertical'
  const insertMode = mode || 'cellImage'; // 'cellImage' or 'formula'
  const resize = autoResize !== false;

  let insertedCount = 0;

  for (let i = 0; i < urls.length; i++) {
    const url = urls[i].trim();
    if (!url) continue;

    let targetRow, targetCol;
    if (dir === 'horizontal') {
      targetRow = startRow;
      targetCol = startCol + i;
    } else { // vertical
      targetRow = startRow + i;
      targetCol = startCol;
    }

    const cell = sheet.getRange(targetRow, targetCol);

    if (insertMode === 'cellImage') {
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
  if (resize && insertedCount > 0) {
    if (dir === 'horizontal') {
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
    direction: dir,
    message: `Successfully inserted ${insertedCount} image(s) ${dir}ly starting at cell ${startCellName}.`
  };
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
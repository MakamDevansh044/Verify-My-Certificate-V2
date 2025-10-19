/**
 * Certificate Generator - Google Apps Script
 *
 * Bound to a Google Sheet (Certificates_Data by default).
 *
 * Features:
 * - Read rows from sheet and generate personalized certificates from a Slides template
 * - Replace {{tags}} (e.g., {{Name}}, {{Event}}, {{Date}}, {{Issuer}}, {{Certificate_ID}})
 * - Insert QR code (replace a shape containing {{QR}} in the Slide template)
 * - Export as PDF, set sharing to anyone with link, write Verification_Link back to Sheet
 * - Menu: Certificate Generator -> Generate Certificates | Generate Selected | Re-Generate Certificate
 * - Optional web app endpoint: doGet(e) for /verify? id or name
 *
 * Setup:
 * - Create a sheet named "Certificates_Data" with headers:
 *   Name | Event | Date | Issuer | Certificate_ID | Verification_Link | Status
 * - Create a sheet named "Config" with key/value rows:
 *   TEMPLATE_SLIDE_ID | <slides-file-id>
 *   OUTPUT_FOLDER_ID  | <drive-folder-id>    (folder where PDFs will be stored)
 *   WEB_APP_URL       | <optional: public web app url used for QR codes>
 *
 * - In the Slides template include the placeholder tags ({{Name}}, etc.) and include a text box with "{{QR}}"
 *   to indicate where the QR image should be placed.
 *
 * - Paste this script into the Script Editor (Extensions > Apps Script) and save.
 * - Run "initConfigIfMissing" once to scaffold the Config sheet (and set template/folder IDs).
 * - Deploy web app (optional) and add its URL to the Config sheet if desired.
 */

const CONFIG_SHEET_NAME = 'Config';
const DATA_SHEET_NAME = 'Certificates_Data';
const DEFAULT_OUTPUT_FOLDER_NAME = 'Generated_Certificates';
const REQUIRED_HEADERS = ['Name', 'Event', 'Date', 'Issuer', 'Certificate_ID', 'Verification_Link', 'Status'];

/* ---------- onOpen: Add custom menu ---------- */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Certificate Generator')
    .addItem('Generate Certificates', 'generateCertificates')
    .addItem('Generate Selected', 'generateSelected')
    .addItem('Re-Generate Certificate (Active Row)', 'regenerateActiveRow')
    .addToUi();
}

/* ---------- Initialize Config sheet if missing ---------- */
function initConfigIfMissing() {
  const ss = SpreadsheetApp.getActive();
  let config = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!config) {
    config = ss.insertSheet(CONFIG_SHEET_NAME);
    config.getRange(1,1,4,2).setValues([
      ['TEMPLATE_SLIDE_ID', 'PASTE_YOUR_SLIDE_ID_HERE'],
      ['OUTPUT_FOLDER_ID', 'OPTIONAL: PASTE_FOLDER_ID_HERE (or leave blank to auto-create)'],
      ['WEB_APP_URL', 'OPTIONAL: PASTE_DEPLOYED_WEBAPP_URL_HERE'],
      ['OUTPUT_FOLDER_NAME', DEFAULT_OUTPUT_FOLDER_NAME]
    ]);
    SpreadsheetApp.getUi().alert('Config sheet created. Open the "Config" sheet and paste your TEMPLATE_SLIDE_ID (Slides file ID). Optionally set OUTPUT_FOLDER_ID or leave blank to auto-create.');
  } else {
    SpreadsheetApp.getUi().alert('Config sheet already exists. Make sure TEMPLATE_SLIDE_ID is set.');
  }
}

/* ---------- Read config values into an object ---------- */
function getConfig() {
  const ss = SpreadsheetApp.getActive();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) throw new Error('Config sheet missing. Run initConfigIfMissing() or create a sheet named "Config".');
  const values = configSheet.getRange(1,1,configSheet.getLastRow(),2).getValues();
  const cfg = {};
  values.forEach(row => {
    const key = String(row[0]).trim();
    const value = row[1] ? String(row[1]).trim() : '';
    if (key) cfg[key] = value;
  });
  // ensure output folder name
  if (!cfg['OUTPUT_FOLDER_NAME']) cfg['OUTPUT_FOLDER_NAME'] = DEFAULT_OUTPUT_FOLDER_NAME;
  return cfg;
}

/* ---------- Utility: Get data sheet and headers ---------- */
function getDataSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(DATA_SHEET_NAME);
  if (!sheet) {
    // create and add headers
    sheet = ss.insertSheet(DATA_SHEET_NAME);
    sheet.getRange(1,1,1,REQUIRED_HEADERS.length).setValues([REQUIRED_HEADERS]);
    SpreadsheetApp.getUi().alert('Data sheet created with default headers. Fill in certificate rows and re-run.');
  }
  return sheet;
}

function getHeaderMap(sheet) {
  const headerRow = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  const map = {};
  headerRow.forEach((h,i) => { map[String(h).trim()] = i + 1; }); // 1-indexed column
  return map;
}

/* ---------- Generate all pending certificates ---------- */
function generateCertificates() {
  const sheet = getDataSheet();
  const headerMap = getHeaderMap(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No data rows found in ' + DATA_SHEET_NAME);
    return;
  }
  const dataRange = sheet.getRange(2,1,lastRow-1,sheet.getLastColumn());
  const data = dataRange.getValues();
  const cfg = getConfig();
  const outFolder = ensureOutputFolder(cfg);
  let processed = 0;
  for (let i = 0; i < data.length; i++) {
    const rowIndex = i + 2;
    const row = data[i];
    const verificationLink = getCellValue(row, headerMap['Verification_Link']);
    if (verificationLink) {
      // skip existing, mark as Skipped
      sheet.getRange(rowIndex, headerMap['Status']).setValue('Skipped (exists)');
      continue;
    }
    try {
      generateCertificateForRow(sheet, rowIndex, headerMap, cfg, outFolder);
      processed++;
      // small pause to be polite to APIs
      Utilities.sleep(300);
    } catch (e) {
      Logger.log('Error generating for row ' + rowIndex + ': ' + e);
      sheet.getRange(rowIndex, headerMap['Status']).setValue('Error: ' + e.message);
    }
  }
  SpreadsheetApp.getUi().alert('Generation complete. Processed: ' + processed + ' rows.');
}

/* ---------- Generate for currently selected rows ---------- */
function generateSelected() {
  const sheet = getDataSheet();
  const headerMap = getHeaderMap(sheet);
  const selection = sheet.getActiveRange();
  if (!selection) {
    SpreadsheetApp.getUi().alert('Select one or more rows to generate.');
    return;
  }
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  if (startRow === 1) {
    SpreadsheetApp.getUi().alert('Please select data rows (not the header).');
    return;
  }
  const cfg = getConfig();
  const outFolder = ensureOutputFolder(cfg);
  let processed = 0;
  for (let r = startRow; r < startRow + numRows; r++) {
    try {
      generateCertificateForRow(sheet, r, headerMap, cfg, outFolder);
      processed++;
      Utilities.sleep(300);
    } catch (e) {
      Logger.log('Error generating for row ' + r + ': ' + e);
      sheet.getRange(r, headerMap['Status']).setValue('Error: ' + e.message);
    }
  }
  SpreadsheetApp.getUi().alert('Selected generation complete. Processed: ' + processed + ' rows.');
}

/* ---------- Re-generate for active row (force update) ---------- */
function regenerateActiveRow() {
  const sheet = getDataSheet();
  const headerMap = getHeaderMap(sheet);
  const active = sheet.getActiveRange();
  if (!active) {
    SpreadsheetApp.getUi().alert('Place your cursor in the row you want to re-generate.');
    return;
  }
  const row = active.getRow();
  if (row === 1) {
    SpreadsheetApp.getUi().alert('Select a data row (not header).');
    return;
  }
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert('Re-generate certificate?', 'This will overwrite the existing certificate link for row ' + row + '. Continue?', ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  const cfg = getConfig();
  const outFolder = ensureOutputFolder(cfg);
  try {
    generateCertificateForRow(sheet, row, headerMap, cfg, outFolder, {force: true});
    ui.alert('Re-generated for row ' + row);
  } catch (e) {
    Logger.log(e);
    ui.alert('Error: ' + e.message);
  }
}

/* ---------- Core: Generate a single certificate for the given row ---------- */
/**
 * options:
 *   force: boolean - if true, will regenerate even if Verification_Link already exists
 */
function generateCertificateForRow(sheet, rowIndex, headerMap, cfg, outFolder, options) {
  options = options || {};
  const rowValues = sheet.getRange(rowIndex,1,1,sheet.getLastColumn()).getValues()[0];
  // Helper to read values from row by header
  const get = (name) => getCellValue(rowValues, headerMap[name]);
  // Ensure Certificate_ID
  let certId = get('Certificate_ID') || '';
  if (!certId) {
    certId = generateCertificateId();
    sheet.getRange(rowIndex, headerMap['Certificate_ID']).setValue(certId);
  }
  // If link exists and not forced, skip
  const existingLink = get('Verification_Link') || '';
  if (existingLink && !options.force) {
    sheet.getRange(rowIndex, headerMap['Status']).setValue('Skipped (exists)');
    return;
  }
  // Template slide id from Config
  const templateId = cfg['TEMPLATE_SLIDE_ID'];
  if (!templateId || templateId.indexOf('PASTE') === 0) {
    throw new Error('TEMPLATE_SLIDE_ID is missing or placeholder in Config sheet.');
  }
  // Decide verification target URL for QR:
  // prefer WEB_APP_URL (if set) with ?id=<certId>, otherwise will use the PDF link after upload
  const webAppUrl = cfg['WEB_APP_URL'] || '';
  // Make a copy of template Slides in the output folder (temporary), personalize it
  const templateFile = DriveApp.getFileById(templateId);
  const nameSanitized = (get('Name') || 'participant').replace(/[\/\\:\*\?"<>\|]/g, '');
  const copyName = 'Certificate - ' + nameSanitized + ' - ' + certId;
  const tempCopy = templateFile.makeCopy(copyName, outFolder);
  const presentation = SlidesApp.openById(tempCopy.getId());
  // Replace all placeholders for every header (if present)
  // We sanitize empty values to blank strings
  const headerKeys = Object.keys(headerMap);
  headerKeys.forEach(h => {
    const placeholder = '{{' + h + '}}';
    const value = get(h) || '';
    try {
      presentation.replaceAllText(placeholder, String(value));
    } catch (e) {
      // continue if placeholder not found
    }
  });
  // Ensure {{Certificate_ID}} is replaced too
  presentation.replaceAllText('{{Certificate_ID}}', certId);
  // Insert QR code: find shape containing '{{QR}}'
  let verificationTarget = '';
  if (webAppUrl) {
    // prefer pointing QR to web app verify endpoint
    // ensure webAppUrl has no trailing slash
    const base = webAppUrl.replace(/\/+$/, '');
    verificationTarget = base + '?id=' + encodeURIComponent(certId);
  }
  // We'll export PDF first to create permanent PDF link if needed for verificationTarget
  // But for QR, if no webAppUrl we will set QR to pdf link after upload.
  // So: prepare QR target now if webAppUrl exists, else set placeholder and insert proper QR after pdf is created.
  const qrUsedDirectly = !!verificationTarget;
  if (qrUsedDirectly) {
    insertQrIntoPresentation(presentation, verificationTarget);
  } else {
    // Insert a temporary placeholder text like (QR_HERE) or remove the text box but we'll later insert QR
    // We'll attempt to find {{QR}} shape and remove it to make space for image later
    removeQrPlaceholder(presentation);
  }
  presentation.saveAndClose();
  // Export to PDF and store in output folder
  const pdfBlob = tempCopy.getAs(MimeType.PDF).setName(copyName + '.pdf');
  const pdfFile = outFolder.createFile(pdfBlob);
  // Make PDF public (anyone with link can view)
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const pdfUrl = pdfFile.getUrl();
  // If we didn't have webAppUrl, set verificationTarget to pdfUrl so QR points to PDF
  if (!qrUsedDirectly) {
    // reopen presentation to insert QR pointing to pdfUrl at same spot
    const pres = SlidesApp.openById(tempCopy.getId());
    insertQrIntoPresentation(pres, pdfUrl);
    pres.saveAndClose();
    // Re-create PDF to include the QR (replace file)
    // Remove previous pdf file
    pdfFile.setTrashed(true);
    const pdfBlob2 = tempCopy.getAs(MimeType.PDF).setName(copyName + '.pdf');
    const pdfFile2 = outFolder.createFile(pdfBlob2);
    pdfFile2.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    // Update pdfFile reference and url
    pdfFile = pdfFile2;
    const newPdfUrl = pdfFile2.getUrl();
    // final verification link becomes newPdfUrl
    // update variable
    verificationTarget = newPdfUrl;
  } else {
    // If webAppUrl was used, still set Verification_Link to pdfUrl (direct link to certificate)
    verificationTarget = verificationTarget; // unchanged
  }
  // Clean up temp copy of Slides to avoid clutter
  try {
    tempCopy.setTrashed(true);
  } catch (e) {
    // ignore
  }
  // Write back to sheet: Verification_Link = PDF url, Status = Generated
  sheet.getRange(rowIndex, headerMap['Verification_Link']).setValue(pdfFile.getUrl());
  sheet.getRange(rowIndex, headerMap['Status']).setValue('Generated');
  // Optionally write a timestamp column if desired. We'll not change other columns.
  return {
    certificate_id: certId,
    pdf_url: pdfFile.getUrl(),
    verification_target: verificationTarget
  };
}

/* ---------- Helpers ---------- */

function getCellValue(rowArray, colIndex) {
  if (!colIndex || colIndex < 1) return '';
  const value = rowArray[colIndex - 1];
  if (value === null || value === undefined) return '';
  return value;
}

function generateCertificateId() {
  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const rand = Math.random().toString(36).substring(2, 8).toUpperCase();
  return 'CERT-' + ts + '-' + rand;
}

/* Ensure output folder exists; use OUTPUT_FOLDER_ID from config if present, otherwise create folder in Drive root */
function ensureOutputFolder(cfg) {
  if (cfg['OUTPUT_FOLDER_ID']) {
    try {
      return DriveApp.getFolderById(cfg['OUTPUT_FOLDER_ID']);
    } catch (e) {
      // fallthrough to create by name
    }
  }
  const folderName = cfg['OUTPUT_FOLDER_NAME'] || DEFAULT_OUTPUT_FOLDER_NAME;
  // Look for existing folder by name in root
  const existing = DriveApp.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();
  return DriveApp.createFolder(folderName);
}

/* Insert QR image into presentation at the position of the shape containing '{{QR}}' */
function insertQrIntoPresentation(presentation, url) {
  const slides = presentation.getSlides();
  const qrBlob = createQrBlob(url);
  let inserted = false;
  for (let s = 0; s < slides.length; s++) {
    const slide = slides[s];
    const pageElements = slide.getPageElements();
    for (let p = 0; p < pageElements.length; p++) {
      const el = pageElements[p];
      // Only shapes have getShape and can have text
      try {
        if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = el.asShape();
          const textRange = shape.getText();
          if (!textRange) continue;
          const txt = textRange.asString();
          if (txt && txt.indexOf('{{QR}}') !== -1) {
            const left = shape.getLeft();
            const top = shape.getTop();
            const width = shape.getWidth();
            const height = shape.getHeight();
            // Insert image and remove shape
            slide.insertImage(qrBlob, left, top, width, height);
            shape.remove();
            inserted = true;
            break;
          }
        }
      } catch (e) {
        // ignore shapes without text
      }
    }
    if (inserted) break;
  }
  // If no placeholder found, insert in bottom-right corner of first slide
  if (!inserted && slides.length > 0) {
    const first = slides[0];
    const pageWidth = first.getPageWidth ? first.getPageWidth() : 960;
    const pageHeight = first.getPageHeight ? first.getPageHeight() : 540;
    const size = 150;
    const left = pageWidth - size - 40;
    const top = pageHeight - size - 40;
    first.insertImage(qrBlob, left, top, size, size);
  }
}

/* Remove any {{QR}} placeholder shape (used before we insert QR after creating PDF) */
function removeQrPlaceholder(presentation) {
  const slides = presentation.getSlides();
  for (let s = 0; s < slides.length; s++) {
    const slide = slides[s];
    const pageElements = slide.getPageElements();
    for (let p = pageElements.length - 1; p >= 0; p--) {
      const el = pageElements[p];
      try {
        if (el.getPageElementType() === SlidesApp.PageElementType.SHAPE) {
          const shape = el.asShape();
          const txt = shape.getText().asString();
          if (txt && txt.indexOf('{{QR}}') !== -1) {
            shape.remove();
          }
        }
      } catch (e) {
        // ignore
      }
    }
  }
}

/* Create a QR code blob using Google Chart API (UrlFetchApp) */
function createQrBlob(data, size) {
  size = size || 300;
  const url = 'https://chart.googleapis.com/chart?cht=qr&chs=' + size + 'x' + size + '&chl=' + encodeURIComponent(data) + '&chld=L|1';
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  if (res.getResponseCode && res.getResponseCode() === 200) {
    const blob = res.getBlob();
    blob.setName('qr.png');
    return blob;
  } else {
    throw new Error('Unable to generate QR code (HTTP ' + (res.getResponseCode ? res.getResponseCode() : 'error') + ')');
  }
}

/* ---------- Optional Web App: verification endpoint ---------- */
/**
 * Deploy this script as a web app (Execute as: Me, Who has access: Anyone).
 * Example usage:
 *  - https://script.google.com/macros/s/XXXXX/exec?id=CERT-20250101...
 *  - &format=json for JSON response
 *
 * Searches the Certificates_Data sheet for Certificate_ID or Name and returns simple HTML/JSON.
 */
function doGet(e) {
  const params = e.parameter || {};
  const id = params.id || '';
  const name = params.name || '';
  const format = (params.format || 'html').toLowerCase();
  const sheet = getDataSheet();
  const headerMap = getHeaderMap(sheet);

  // gather all rows
  const data = sheet.getRange(2,1,Math.max(sheet.getLastRow()-1,0), sheet.getLastColumn()).getValues();
  let found = null;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const certId = getCellValue(row, headerMap['Certificate_ID']);
    const rowName = getCellValue(row, headerMap['Name']);
    if (id && certId && certId === id) {
      found = toRowObject(row, headerMap);
      break;
    }
    if (!id && name && rowName && String(rowName).toLowerCase() === String(name).toLowerCase()) {
      found = toRowObject(row, headerMap);
      break;
    }
  }
  if (format === 'json') {
    const out = ContentService.createTextOutput();
    out.setMimeType(ContentService.MimeType.JSON);
    if (found) {
      out.setContent(JSON.stringify({valid: true, data: found}));
    } else {
      out.setContent(JSON.stringify({valid: false, data: null}));
    }
    return out;
  } else {
    // HTML response
    let html = '<html><head><meta charset="utf-8"><title>Certificate Verification</title></head><body style="font-family:Arial, sans-serif;">';
    html += '<h2>Certificate Verification</h2>';
    if (found) {
      html += '<p><strong>Status:</strong> Valid</p>';
      html += '<ul>';
      html += '<li><strong>Name:</strong> ' + escapeHtml(found['Name'] || '') + '</li>';
      html += '<li><strong>Event:</strong> ' + escapeHtml(found['Event'] || '') + '</li>';
      html += '<li><strong>Date:</strong> ' + escapeHtml(found['Date'] || '') + '</li>';
      html += '<li><strong>Issuer:</strong> ' + escapeHtml(found['Issuer'] || '') + '</li>';
      html += '<li><strong>Certificate ID:</strong> ' + escapeHtml(found['Certificate_ID'] || '') + '</li>';
      if (found['Verification_Link']) {
        html += '<li><a href="' + found['Verification_Link'] + '" target="_blank">View Certificate (PDF)</a></li>';
      }
      html += '</ul>';
    } else {
      html += '<p><strong>Status:</strong> Invalid certificate or not found.</p>';
    }
    html += '</body></html>';
    return HtmlService.createHtmlOutput(html);
  }
}

function toRowObject(row, headerMap) {
  const obj = {};
  for (const key in headerMap) {
    obj[key] = getCellValue(row, headerMap[key]);
  }
  return obj;
}

/* Simple HTML escape */
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
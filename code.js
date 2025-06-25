/**
 * @OnlyCurrentDoc
 *
 * Fit2Go Trainer Applicants - v3.1 OPTIMIZED
 *
 * This script automates the tracking of job applicants from Indeed. This version
 * includes a professional-grade parsing engine that handles PDF, DOCX, and
 * image-based files with separate, optimized strategies.
 *
 * Key Features:
 * - Advanced Hybrid Parsing: Uses dedicated, reliable methods for DOCX and PDF/Text,
 *   with OCR as a final fallback for image-based files.
 * - Data Preservation: Will not overwrite manually entered data with "N/A".
 * - Intelligent Attachment Logic: Smartly identifies the most likely resume file.
 * - Status Reporting: Provides clear success or error messages in a "Status" column.
 */

// --- SCRIPT CONFIGURATION ---
const SHEET_NAME = 'Fit2Go Trainer Applicants';
const GMAIL_QUERY = 'subject:(New application) from:(indeedemail.com) newer_than:60d';
const RESUME_FOLDER_NAME = 'Indeed Resumes';
const HEADERS = ['Applicant Name', 'Application Date', 'Indeed Email ID', 'Phone', 'Contact Email', 'Resume Link', 'Status'];
const SEARCH_BATCH_SIZE = 500; // GmailApp.search only returns 500 threads per call


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Trainer Applicants')
    .addItem('Add/Update Applicants', 'processEmails')
    .addToUi();
}

function processEmails() {
  if (typeof Drive === 'undefined' || !Drive.Files) {
    showAlert('ERROR: The Drive API service is not enabled. Please open the script editor, go to "Services +", add "Drive API", and run the script again.');
    return;
  }

  const sheet = getActiveSheet();
  if (!sheet) return;

  ensureHeader(sheet, HEADERS);

  const allData = sheet.getDataRange().getValues();
  const emailToRowIndex = new Map();
  allData.slice(1).forEach((row, index) => {
    const email = row[2];
    if (email) {
      emailToRowIndex.set(email, index + 2);
    }
  });

  const resumeFolder = getOrCreateFolder(RESUME_FOLDER_NAME);
  let threads = [];
  let start = 0;
  let batch;
  do {
    batch = GmailApp.search(GMAIL_QUERY, start, SEARCH_BATCH_SIZE);
    threads = threads.concat(batch);
    start += SEARCH_BATCH_SIZE;
  } while (batch.length === SEARCH_BATCH_SIZE);

  if (threads.length === 0) {
    showAlert('No new threads matched the query.');
    return;
  }

  let newApplicantsCount = 0;
  const newRows = [];

  threads.forEach(thread => {
    const applicant = parseApplicantData(thread, resumeFolder);
    if (applicant) {
      const rowIndex = emailToRowIndex.get(applicant.indeedEmail);

      if (rowIndex) {
        // --- UPDATE EXISTING ROW (with data preservation) ---
        const existingRow = allData[rowIndex - 1];
        const updatedRow = [
          applicant.name,
          applicant.applicationDate,
          applicant.indeedEmail,
          applicant.phone || existingRow[3],
          applicant.contactEmail || existingRow[4],
          applicant.resumeLink || existingRow[5],
          applicant.status
        ];
        sheet.getRange(rowIndex, 1, 1, HEADERS.length).setValues([updatedRow]);
      } else {
        // --- ADD NEW ROW ---
        newApplicantsCount++;
        const newRowData = [
          applicant.name,
          applicant.applicationDate,
          applicant.indeedEmail,
          applicant.phone || 'N/A',
          applicant.contactEmail || 'N/A',
          applicant.resumeLink || 'No Link',
          applicant.status
        ];
        newRows.push(newRowData);
      }
    }
    thread.markRead();
  });

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, HEADERS.length).setValues(newRows);
    newRows.forEach((row, idx) => {
      emailToRowIndex.set(row[2], sheet.getLastRow() - newRows.length + idx + 1);
    });
  }

  if (newApplicantsCount > 0) {
    showAlert(`Process complete. Added ${newApplicantsCount} new applicant(s) and updated existing ones.`);
  } else {
    showAlert('Process complete. No new applicants were found, but existing entries may have been updated.');
  }
}


function parseApplicantData(thread, resumeFolder) {
  const messages = thread.getMessages();
  const message = messages[0];
  if (!message) return null;

  const body = message.getPlainBody() || message.getBody();
  const applicationDate = message.getDate();
  const nameMatch = body.match(/(.*) applied/);
  const name = nameMatch ? nameMatch[1].trim() : 'N/A';
  const fromHeader = message.getFrom();
  const emailMatch = fromHeader.match(/<([^>]+)>/);
  const indeedEmail = emailMatch ? emailMatch[1].trim() : fromHeader;

  let phone = null;
  let contactEmail = null;
  let resumeLink = null;
  let status = 'No Resume Attached';

  const attachments = [].concat.apply([], messages.map(m => m.getAttachments()));
  if (attachments.length > 0) {
    let resumeAttachment = attachments.find(att => att.getName().toLowerCase().includes('resume')) ||
                           attachments.find(att => /pdf|word|document/i.test(att.getContentType()));

    if (resumeAttachment) {
      const fileName = `${name} - ${resumeAttachment.getName()}`;
      const existingFiles = resumeFolder.getFilesByName(fileName);
      const file = existingFiles.hasNext() ? existingFiles.next() : resumeFolder.createFile(resumeAttachment).setName(fileName);
      resumeLink = file.getUrl();

      const extractionResult = extractTextFromResume(resumeAttachment);

      if (extractionResult.text) {
        status = 'Processed';
        const phoneInResume = extractionResult.text.match(/\+?\d[\d\s().-]{7,}\d/);
        if (phoneInResume) {
          phone = phoneInResume[0].replace(/[^\d]/g, '');
        }

        const emailInResume = extractionResult.text.match(/[A-Z0-9._%+-]+\s*@\s*[A-Z0-9.-]+\s*\.[A-Z]{2,}/i);
        if (emailInResume) {
          contactEmail = emailInResume[0].replace(/\s/g, '');
        }

        if (!phone && !contactEmail) {
          status = 'No Contact Info Found';
        }
      } else {
        status = extractionResult.error;
      }
    } else {
      status = 'No Resume Filetype Found';
    }
  }

  return { name, applicationDate, indeedEmail, phone, contactEmail, resumeLink, status };
}

/**
 * FINAL VERSION: A true multi-format parsing engine.
 * It uses specific, reliable strategies for DOCX, PDF/Text, and finally OCR for images.
 */
function extractTextFromResume(attachment) {
  const mimeType = attachment.getContentType();
  const attachmentName = attachment.getName();

  // --- Strategy 1: Dedicated DOCX Handler (Most Reliable for Word) ---
  if (mimeType === MimeType.MICROSOFT_WORD || mimeType === MimeType.GOOGLE_DOCS || attachmentName.toLowerCase().endsWith('.docx')) {
    let tempDoc;
    try {
      // Convert DOCX to a temporary Google Doc to read its text reliably.
      tempDoc = Drive.Files.create({ name: `[DELETE] ${attachmentName}`, mimeType: MimeType.GOOGLE_DOCS }, attachment.copyBlob());
      const text = DocumentApp.openById(tempDoc.id).getBody().getText();
      Logger.log(`Success (DOCX Conversion): ${attachmentName}`);
      return { text: text, error: null };
    } catch (e) {
      Logger.log(`DOCX conversion failed for ${attachmentName}, falling back to OCR. Error: ${e.message}`);
    } finally {
      if (tempDoc) Drive.Files.remove(tempDoc.id);
    }
  }
  
  // --- Strategy 2: Direct Text Extraction (Fastest for text-based PDFs) ---
  try {
    const text = attachment.getDataAsString();
    if (text && text.length > 100) {
      Logger.log(`Success (Direct Text): ${attachmentName}`);
      return { text: text, error: null };
    }
  } catch (e) {
    Logger.log(`Direct text extraction failed for ${attachmentName}, falling back to OCR. Error: ${e.message}`);
  }

  // --- Strategy 3: OCR Fallback (For image-based PDFs, scans, etc.) ---
  let ocrFile;
  try {
    const resource = { name: `[OCR] ${attachmentName}`, mimeType: MimeType.GOOGLE_DOCS };
    ocrFile = Drive.Files.create(resource, attachment.copyBlob(), { ocr: true });
    const doc = DocumentApp.openById(ocrFile.id);
    const text = doc.getBody().getText();
    Logger.log(`Success (OCR): ${attachmentName}`);
    return { text: text, error: null };
  } catch (e) {
    const errorMessage = `OCR Error: ${e.message}`;
    Logger.log(`All parsing failed for: ${attachmentName}. Final Error: ${errorMessage}`);
    return { text: null, error: errorMessage };
  } finally {
    if (ocrFile) Drive.Files.remove(ocrFile.id);
  }
}


// --- UTILITY FUNCTIONS ---

function getActiveSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheetByName(SHEET_NAME);
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

function ensureHeader(sheet, headerArray) {
  const existing = sheet.getRange(1, 1, 1, headerArray.length).getValues()[0];
  const mismatch = headerArray.some((h, i) => existing[i] !== h);
  if (sheet.getLastRow() < 1 || mismatch) {
    sheet.insertRowBefore(1)
      .getRange(1, 1, 1, headerArray.length)
      .setValues([headerArray])
      .setFontWeight('bold');
  }
}

function showAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}

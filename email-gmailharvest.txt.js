function saveEmailsBodiesAndAttachments() {
    // ---- CONFIG ----
    var addresses = [
      "slkunitz@aol.com",
      "figaro1226@gmail.com",
      "sharon.kunitz@aol.com"
    ];
    var batchSize = 50; // threads per batch
    var folderId = "1nqXp_SGGHpQIiKF-FKLfiFF5SeJxPrtd"; // Drive folder
    var contentFilter = ""; // optional content filter, e.g., "handbook"
    var startDateStr = ""; // optional start date filter: format 'yyyy/MM/dd'
    var endDateStr = "";   // optional end date filter: format 'yyyy/MM/dd'
  var savePlainText = true; // save message.getPlainBody() as .txt
  var saveHtmlBody = true;  // save message.getBody() as .html (full HTML)
  var saveRawEml = true;    // save raw RFC822 .eml via Advanced Gmail API
  var previewChars = 10000;  // preview length for sheet logging
    // ---------------
  
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var logSheet = getOrCreateLogSheet_();
    var folder = DriveApp.getFolderById(folderId);
  
    // Get last logged date from sheet (assumes dates are in column E)
    var lastRow = sheet.getLastRow();
    var lastDate;
    if (lastRow > 1) { // assuming row 1 is header
      lastDate = sheet.getRange("E" + lastRow).getValue();
    } else {
      lastDate = new Date(0); // Jan 1, 1970 if sheet empty
    }
    var tz = Session.getScriptTimeZone();
    var formattedLastDate = Utilities.formatDate(lastDate, tz, 'yyyy/MM/dd HH:mm:ss');
    Logger.log('LastRow=' + lastRow + ' LastDate=' + lastDate + ' Formatted=' + formattedLastDate);
    log_(logSheet, 'Start', '', 'lastRow=' + lastRow, 'lastDate=' + lastDate, 'formatted=' + formattedLastDate);
    if (contentFilter) {
      log_(logSheet, 'Filter', '', 'contentFilter=' + contentFilter, '', '');
    }
  
    // Single combined search: any thread containing any of the addresses (from, to, or cc)
    var start = 0;
    var threads;
    var afterDateStr = Utilities.formatDate(lastDate, tz, 'yyyy/MM/dd');
    var participantClauses = [];
    addresses.forEach(function(a) {
      participantClauses.push('from:' + a);
      participantClauses.push('to:' + a);
      participantClauses.push('cc:' + a);
    });
    var participantsQuery = '(' + participantClauses.join(' OR ') + ')';
    // Build date clause: prefer explicit start/end if provided; otherwise use lastDate-based 'after:'
    var dateClauses = [];
    if (startDateStr) dateClauses.push('after:' + startDateStr);
    if (endDateStr) dateClauses.push('before:' + endDateStr);
    if (dateClauses.length === 0) dateClauses.push('after:' + afterDateStr);

    // Log which date filter is used
    if (startDateStr || endDateStr) {
      log_(logSheet, 'DateRange', 'ANY', 'start=' + (startDateStr || ''), 'end=' + (endDateStr || ''), '');
    } else {
      log_(logSheet, 'DateRange', 'ANY', 'after=' + afterDateStr, '', 'from lastDate');
    }

    var query = participantsQuery + ' ' + dateClauses.join(' ') + (contentFilter ? (' ' + contentFilter) : '');
    Logger.log('Combined query: ' + query);
    log_(logSheet, 'Query', 'ANY', query, '', '');
    var minMsgDate = null;
    var maxMsgDate = null;
    var totalThreads = 0;
    var totalMessages = 0;

    do {
      threads = GmailApp.search(query, start, batchSize);
      totalThreads += threads.length;
      Logger.log('Batch start=' + start + ' threads=' + threads.length);
      log_(logSheet, 'Batch', 'ANY', 'start=' + start, 'threads=' + threads.length, '');

      threads.forEach(function(thread) {
        var messages = thread.getMessages();

        messages.forEach(function(message) {
          var date = message.getDate();
          totalMessages++;
          if (!minMsgDate || date < minMsgDate) minMsgDate = date;
          if (!maxMsgDate || date > maxMsgDate) maxMsgDate = date;
          var from = message.getFrom();
          var to = message.getTo();
          var subject = message.getSubject();
          var body = message.getPlainBody();

          // Apply content filter if provided (checks subject and body)
          if (contentFilter) {
            var needle = contentFilter.toLowerCase();
            var s = (subject || '').toLowerCase();
            var b = (body || '').toLowerCase();
            var h = (html || '').toLowerCase();
            if (s.indexOf(needle) === -1 && b.indexOf(needle) === -1 && h.indexOf(needle) === -1) {
              return; // skip this message
            }
          }

          // Safe subject for filename
          var safeSubject = subject.replace(/[^\w\s]/g, "_").substring(0,50);
          var baseName = date.toISOString() + " - " + safeSubject;

          // Save plain text body
          if (savePlainText) {
            var bodyFileName = baseName + ".txt";
            var bodyFile = folder.createFile(bodyFileName, body, MimeType.PLAIN_TEXT);
            sheet.appendRow([
              new Date(),
              from,
              to,
              subject,
              date,
              "BodyPlain",
              bodyFile.getUrl(),
              "",
              (body || '').substring(0, previewChars)
            ]);
          }

          // Save HTML body
          if (saveHtmlBody && html) {
            var htmlFileName = baseName + ".html";
            var htmlFile = folder.createFile(htmlFileName, html, MimeType.HTML);
            sheet.appendRow([
              new Date(),
              from,
              to,
              subject,
              date,
              "BodyHtml",
              htmlFile.getUrl(),
              "",
              (body || '').substring(0, previewChars)
            ]);
          }

          // Save raw EML (Advanced Gmail service required)
          if (saveRawEml) {
            try {
              var messageId = message.getId();
              var rawResponse = Gmail.Users.Messages.get('me', messageId, {format: 'raw'});
              if (rawResponse && rawResponse.raw) {
                var rawBytes = Utilities.base64DecodeWebSafe(rawResponse.raw);
                var emlBlob = Utilities.newBlob(rawBytes, 'message/rfc822', baseName + '.eml');
                var emlFile = folder.createFile(emlBlob);
                sheet.appendRow([
                  new Date(),
                  from,
                  to,
                  subject,
                  date,
                  "RawEML",
                  emlFile.getUrl(),
                  "",
                  ""
                ]);
              }
            } catch (rawErr) {
              log_(logSheet, 'RawEMLError', '', String(rawErr), '', '');
            }
          }

          // Save attachments
          var attachments = message.getAttachments();
          attachments.forEach(function(att) {
            var fileName = date.toISOString() + " - " + att.getName();
            var file = folder.createFile(att.copyBlob()).setName(fileName);

            // Log attachment metadata
            sheet.appendRow([
              new Date(),
              from,
              to,
              subject,
              date,
              "Attachment",
              file.getUrl(),
              att.getName(),
              "" // no preview for attachments
            ]);
          });
        });
      });

      start += batchSize;
    } while (threads.length === batchSize);
    var minStr = minMsgDate ? Utilities.formatDate(minMsgDate, tz, 'yyyy/MM/dd HH:mm:ss') : '';
    var maxStr = maxMsgDate ? Utilities.formatDate(maxMsgDate, tz, 'yyyy/MM/dd HH:mm:ss') : '';
    Logger.log('Summary threads=' + totalThreads + ' messages=' + totalMessages + ' range=' + minStr + ' .. ' + maxStr);
    log_(logSheet, 'Summary', 'ANY', 'threads=' + totalThreads, 'messages=' + totalMessages, 'range=' + minStr + ' .. ' + maxStr);
  
    Logger.log("Email fetch complete.");
    log_(logSheet, 'Complete', '', '', '', '');
  }

// Ensure a dedicated log sheet exists
function getOrCreateLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('EmailScriptLogs');
  if (!sheet) {
    sheet = ss.insertSheet('EmailScriptLogs');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Timestamp','Event','Address','Detail1','Detail2','Detail3']);
  }
  return sheet;
}

// Append a single log row
function log_(logSheet, event, address, d1, d2, d3) {
  try {
    logSheet.appendRow([new Date(), event, address || '', d1 || '', d2 || '', d3 || '']);
  } catch (e) {
    Logger.log('log_ error: ' + e);
  }
}
  
// When a value is typed into column A, perform a Gmail search across
// the configured addresses and write results starting from column B.
// The original search term is copied down column A for each result row.
function onEdit(e) {
  try {
    var range = e && e.range;
    if (!range) return;
    if (range.getColumn() !== 1) return; // only react to edits in column A
    var sheet = range.getSheet();
    var term = String(range.getValue()).trim();
    if (!term) return;
    performGmailSearchIntoSheet_(sheet, range.getRow(), term);
  } catch (err) {
    Logger.log('onEdit error: ' + err);
  }
}

// Helper to search Gmail and write results into the sheet, starting at row startRow.
// Columns:
// A: search term (repeated)
// B: from address
// C: date
// D: subject
// E: snippet
// F: thread link
function performGmailSearchIntoSheet_(sheet, startRow, term) {
  // Keep this list in sync with the main function's addresses
  var addresses = [
    "slkunitz@aol.com",
    "figaro1226@gmail.com",
    "sharon.kunitz@aol.com"
  ];

  var maxThreads = 100; // upper bound to keep writes manageable
  var maxMessages = 200; // absolute cap on messages written

  // Build a single Gmail query for all addresses
  var addressQuery = '(' + addresses.map(function(a){ return 'from:' + a; }).join(' OR ') + ')';
  var query = addressQuery + ' ' + term;

  var threads = GmailApp.search(query, 0, maxThreads);

  var addressSet = {};
  addresses.forEach(function(a){ addressSet[a.toLowerCase()] = true; });

  var rows = [];
  threads.forEach(function(thread){
    var threadUrl = '';
    try { threadUrl = thread.getPermalink(); } catch (e) { threadUrl = ''; }
    var messages = thread.getMessages();
    messages.forEach(function(message){
      var from = String(message.getFrom() || '').toLowerCase();
      // Extract just the email from formats like "Name <email@domain>"
      var match = from.match(/<([^>]+)>/);
      var fromEmail = (match && match[1]) ? match[1] : from.replace(/(^\s+|\s+$)/g, '');
      if (!addressSet[fromEmail]) return;

      var subject = message.getSubject() || '';
      var body = message.getPlainBody() || '';
      var termLower = term.toLowerCase();
      if (subject.toLowerCase().indexOf(termLower) === -1 && body.toLowerCase().indexOf(termLower) === -1) {
        // If the term doesn't appear in subject/body, skip this message (thread may match due to other messages)
        return;
      }

      var date = message.getDate();
      var snippet = body.replace(/\s+/g, ' ').substring(0, 200);
      rows.push([term, fromEmail, date, subject, snippet, threadUrl]);
    });
  });

  // Sort results by date descending
  rows.sort(function(a, b){ return b[2] - a[2]; });
  if (rows.length > maxMessages) rows = rows.slice(0, maxMessages);

  if (rows.length === 0) {
    // If no results, just keep the term on the edited row and clear B:F
    sheet.getRange(startRow, 2, 1, 5).clearContent();
    return;
  }

  // Write values starting at column A (to replicate term) through F
  var numRows = rows.length;
  var writeRange = sheet.getRange(startRow, 1, numRows, 6);
  writeRange.setValues(rows);
}
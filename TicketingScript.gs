/**
 * @OnlyCurrentDoc
 *
 * Automated IT Ticketing System for Slack & Google Sheets
 *
 * This script connects to a Google Sheet that is populated by a Slack Workflow.
 * It performs the following actions:
 * 1. Fetches the full text of a Slack message from a permalink.
 * 2. Determines the user's location based on the Slack channel ID.
 * 3. Categorizes the ticket by searching the message for keywords.
 * 4. Assigns a unique, formatted Ticket ID.
 * 5. Formats rows with colors based on location.
 * 6. Periodically removes duplicate entries.
 * 7. Sends a daily summary report of open tickets to a specified Slack channel.
 *
 * @version 1.2 - Forever Edition
 * 1.2 Changelong: Updated CreateOrReplaceDailyTriggers function.
 * @author Tyler Wong/Tendril Studio
 * 
 */

// =================================================================================
// --- ⚙️ CONFIGURATION - IMPORTANT: UPDATE THESE VALUES! ⚙️ ---
// =================================================================================

// --- Sheet Configuration ---
// The exact name of the sheet (tab) in your Google Spreadsheet where tickets are logged.
const TARGET_SHEET_NAME = 'IT Tickets';

// --- Column Index Configuration (1-indexed) ---
// These numbers represent the column position (A=1, B=2, C=3, etc.).
// MAKE SURE THESE MATCH YOUR GOOGLE SHEET'S LAYOUT EXACTLY.
const MESSAGE_LINK_COLUMN_INDEX = 4; // Column D: Where Slack Workflow puts the message permalink.
const MESSAGE_COLUMN_INDEX = 3;      // Column C: Where the fetched Slack message text will be written.
const PLATFORM_COLUMN_INDEX = 5;     // Column E: Where the location/platform will be written.
const CATEGORY_COLUMN_INDEX = 6;     // Column F: Where the auto-detected category will be written.
const STATUS_COLUMN_INDEX = 7;       // Column G: Where the ticket status (e.g., "Open") is located.
const TICKET_ID_COLUMN_INDEX = 10;   // Column J: Where the unique Ticket ID will be generated.

// --- Data & Formatting Configuration ---
const HEADER_ROW_COUNT = 1;          // Number of header rows to skip in the sheet.
const COLOR_LOCATION_1 = "#D9EDF7";  // A readable pastel blue for the first location.
const COLOR_LOCATION_2 = "#DFF0D8";  // A readable light pastel green for the second location.

// --- Slack Integration Configuration ---
// The Channel ID for the daily status report.
// To get a channel ID, right-click its name in Slack -> Copy Link. The ID starts with 'C' or 'G'.
const REPORTING_CHANNEL_ID = 'C0123ABCDE'; // <--- ⚠️ UPDATE THIS

// --- Location Mapping (Channel ID -> Location Name) ---
// Maps a Slack Channel ID to a human-readable location name. This is used for coloring rows
// and for generating the Ticket ID prefix. Add all relevant channels here.
const CHANNEL_TO_LOCATION_MAP = {
  'C0123456789': 'Toronto', // <--- ⚠️ UPDATE THESE with your Channel IDs and Location Names
  'C9876543210': 'Sao Paulo',
  'C1112223334': 'New York'
};

// --- Keyword-Based Category Definitions ---
// Define categories and the keywords that trigger them.
// `type: "exact"` is case-sensitive.
// `type: "or"` is case-insensitive.
const CATEGORY_DEFINITIONS = [
  { name: "Hardware", searches: [ { type: "or", phrases: ["macbook", "laptop", "workstation", "gpu", "cpu", "ram", "monitor", "keyboard", "mouse", "peripheral", "display", "power supply", "ups", "kext", "driver", "firmware", "bios", "boot", "crash", "crashing", "blue screen", "temperature", "reboot", "restart", "network", "internet"] } ] },
  { name: "Software", searches: [ { type: "or", phrases: ["houdini", "maxon", "autodesk", "maya", "nuke", "redshift", "arnold", "v-ray", "octane", "plugin", "script", "install", "uninstall", "bug", "error", "version", "render engine", "adobe", "premiere", "photoshop", "after effects", "unity", "unreal", "update", "c4d", "cinema"] } ] },
  { name: "Render Farm", searches: [ { type: "or", phrases: ["deadline", "farm", "render", "job", "queue", "submit", "slave", "worker", "repository", "pulse", "monitoring", "rendering issue", "frame", "render manager"] } ] },
  { name: "Licensing", searches: [ { type: "or", phrases: ["license", "licensing", "licence", "activation", "dongle", "rlm", "flexlm", "floating", "node-locked"] } ] },
  { name: "Data Management", searches: [ { type: "or", phrases: ["dropbox", "storage", "sync", "nas", "san", "server", "share", "backup", "files", "restore", "archive", "data loss", "corrupt", "permission", "access", "volume", "project folder"] } ] },
  { name: "Remote Access", searches: [ { type: "or", phrases: ["parsec", "remote box", "remote access", "rdp", "vnc", "teamviewer", "anydesk", "teradici", "pcoip", "latency", "lag", "disconnect", "streaming"] } ] },
  { name: "Account Management", searches: [ { type: "or", phrases: ["login", "vpn", "password", "2fa", "mfa", "sso", "active directory", "ad", "user account", "onboarding", "offboarding", "access denied", "timecards"] } ] },
  { name: "Spam/Phishing", searches: [ { type: "or", phrases: ["phishing", "spam", "virus", "malware", "scam", "suspicious email", "security alert"] } ] }
];


// =================================================================================
// --- 🛑 DO NOT EDIT BELOW THIS LINE UNLESS YOU KNOW WHAT YOU ARE DOING 🛑 ---
// =================================================================================

// --- Global Variables ---
// Fetches the Slack Bot Token from Script Properties for security.
const SLACK_BOT_TOKEN = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');

// --- Trigger Management ---

/**
 * Main trigger function that runs when the spreadsheet is edited.
 * This is the primary entry point for real-time processing.
 * @param {Object} e The event object from the on-edit trigger.
 */
function onSheetChange(e) {
  if (!e || !e.source) {
    Logger.log('onSheetChange was run manually or without a proper event object. Exiting.');
    return;
  }
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== TARGET_SHEET_NAME) return;
  // We process on EDIT, INSERT_ROW, or OTHER to catch various ways data can arrive.
  if (e.changeType === 'EDIT' || e.changeType === 'INSERT_ROW' || e.changeType === 'OTHER') {
    Logger.log(`onSheetChange triggered: Type=${e.changeType}. Processing new links.`);
    processSlackMessageLinks();
  }
}

/**
 * A scheduled function to clean up duplicates and process any missed rows.
 * This ensures data integrity and catches anything the real-time trigger missed.
 */
function removeDuplicatesAndProcess() {
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0=Sunday, 6=Saturday
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    Logger.log('Skipping scheduled processing on weekend.');
    return;
  }
  Logger.log("Starting scheduled duplicate check...");
  removeDuplicateRows();
  Logger.log("Duplicate check finished. Starting main processing...");
  processSlackMessageLinks();
}

/**
 * Scans the sheet for rows with the same Slack message link and removes older duplicates.
 */
function removeDuplicateRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) {
    Logger.log(`Error in removeDuplicateRows: Sheet '${TARGET_SHEET_NAME}' not found.`);
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW_COUNT) return;

  const dataRange = sheet.getRange(HEADER_ROW_COUNT + 1, MESSAGE_LINK_COLUMN_INDEX, lastRow - HEADER_ROW_COUNT, 1);
  const values = dataRange.getValues();
  const formulas = dataRange.getFormulas();
  
  const seenLinks = new Set();
  const rowsToDelete = [];

  // Iterate over all rows to find duplicate links.
  formulas.forEach((row, index) => {
    let link = '';
    const formula = row[0];
    const value = values[index][0];

    // Extract URL from HYPERLINK formula or use the plain text value.
    if (formula.startsWith('=HYPERLINK(')) {
      const match = formula.match(/=HYPERLINK\("([^"]+)"/);
      if (match && match[1]) link = match[1];
    } else {
      link = value;
    }
    
    link = link.toString().trim();

    if (link !== '') {
      if (seenLinks.has(link)) {
        // If we've seen this link before, mark the row for deletion.
        rowsToDelete.push(HEADER_ROW_COUNT + 1 + index);
      } else {
        seenLinks.add(link);
      }
    }
  });

  // Delete marked rows, starting from the bottom to avoid shifting indices.
  if (rowsToDelete.length > 0) {
    Logger.log(`Found ${rowsToDelete.length} duplicate rows to delete.`);
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      const rowNum = rowsToDelete[i];
      sheet.deleteRow(rowNum);
      Logger.log(`Deleted duplicate row: ${rowNum}`);
    }
  } else {
    Logger.log("No duplicate rows found.");
  }
}

// --- Main Processing Function ---

/**
 * The core function that iterates through the sheet and processes unprocessed rows.
 */
function processSlackMessageLinks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) {
    Logger.log(`Error: Sheet '${TARGET_SHEET_NAME}' not found.`);
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW_COUNT) {
    Logger.log("No data rows found to process.");
    return;
  }

  // Get the entire data range to minimize API calls.
  const dataRange = sheet.getRange(HEADER_ROW_COUNT + 1, 1, lastRow - HEADER_ROW_COUNT, sheet.getLastColumn());
  const values = dataRange.getValues();
  const formulas = dataRange.getFormulas();
  let rowsUpdated = 0;

  values.forEach((row, rowIndex) => {
    const actualRowNumber = HEADER_ROW_COUNT + 1 + rowIndex;
    
    // Get values from the current row using our column index constants.
    const messageLink = row[MESSAGE_LINK_COLUMN_INDEX - 1];
    const currentMessageLinkFormula = formulas[rowIndex][MESSAGE_LINK_COLUMN_INDEX - 1];
    let messageTextInSheet = row[MESSAGE_COLUMN_INDEX - 1];
    let platformInSheet = row[PLATFORM_COLUMN_INDEX - 1];
    let categoryInSheet = row[CATEGORY_COLUMN_INDEX - 1];
    let ticketIdInSheet = row[TICKET_ID_COLUMN_INDEX - 1];
    let rowNeedsProcessingThisPass = false;

    // A row needs processing if it has a link but is missing key data.
    if (messageLink && messageLink.toString().trim() !== '' && (!messageTextInSheet || !platformInSheet || !categoryInSheet)) {
      rowNeedsProcessingThisPass = true;
    }

    if (rowNeedsProcessingThisPass) {
      Logger.log(`Initiating processing for row ${actualRowNumber}.`);
      try {
        // 1. Fetch Slack Message Text if it's missing.
        if (!messageTextInSheet) {
          const fetchedDetails = getMessageDetailsFromSlackLink(messageLink);
          if (fetchedDetails && fetchedDetails.text) {
            sheet.getRange(actualRowNumber, MESSAGE_COLUMN_INDEX).setValue(fetchedDetails.text);
            messageTextInSheet = fetchedDetails.text; // Update local variable for subsequent steps
            Logger.log(`Populated message for row ${actualRowNumber}.`);
          } else {
            sheet.getRange(actualRowNumber, MESSAGE_COLUMN_INDEX).setValue("Failed to retrieve message.");
            Logger.log(`Failed to retrieve message for row ${actualRowNumber}.`);
            messageTextInSheet = " "; // Prevent re-processing this failed row
          }
        }
        
        // 2. Determine Platform/Location if it's missing.
        if (!platformInSheet && messageTextInSheet.toString().trim() !== '') {
          const channelIdFromLink = parseSlackPermalink(messageLink)?.channelId;
          if (channelIdFromLink) {
            const location = determineLocation(channelIdFromLink);
            if (location) {
              sheet.getRange(actualRowNumber, PLATFORM_COLUMN_INDEX).setValue(location);
              platformInSheet = location; // Update local variable
              Logger.log(`Populated Platform for row ${actualRowNumber}: ${location}`);
            }
          }
        }
        
        // 3. Find Category if it's missing.
        if (!categoryInSheet && messageTextInSheet.toString().trim() !== '') {
          const foundCategories = findCategoriesInMessage(messageTextInSheet.toString());
          if (foundCategories.length > 0) {
            sheet.getRange(actualRowNumber, CATEGORY_COLUMN_INDEX).setValue(foundCategories.join(', '));
            categoryInSheet = foundCategories.join(', '); // Update local variable
            Logger.log(`Populated Category for row ${actualRowNumber}: ${foundCategories.join(', ')}`);
          }
        }
      } catch (error) {
        Logger.log(`Script error during processing for row ${actualRowNumber}: ${error.toString()}`);
      }
    }

    // Convert plain text link to a clickable HYPERLINK formula.
    if (messageLink && messageLink.toString().trim() !== '' && !currentMessageLinkFormula.startsWith('=HYPERLINK(')) {
      const hyperlinkFormula = `=HYPERLINK("${messageLink.toString().trim()}", "View Message")`;
      sheet.getRange(actualRowNumber, MESSAGE_LINK_COLUMN_INDEX).setFormula(hyperlinkFormula);
      rowNeedsProcessingThisPass = true;
    }

    // Set row background color based on platform/location.
    if (platformInSheet) {
      const rowRange = sheet.getRange(actualRowNumber, 1, 1, sheet.getLastColumn());
      const lowerCasePlatform = platformInSheet.toString().toLowerCase();
      // Match against the location names defined in CHANNEL_TO_LOCATION_MAP
      if (lowerCasePlatform.includes('toronto')) rowRange.setBackground(COLOR_LOCATION_1);
      else if (lowerCasePlatform.includes('sao paulo')) rowRange.setBackground(COLOR_LOCATION_2);
    }

    // Assign a Ticket ID if one doesn't exist and the platform has been identified.
    if (!ticketIdInSheet && platformInSheet) {
      const year = new Date().getFullYear().toString().slice(-2);
      const platformPrefix = platformInSheet.toString().substring(0, 3).toUpperCase();
      const newTicketId = `${platformPrefix}-${year}-${actualRowNumber}`;
      sheet.getRange(actualRowNumber, TICKET_ID_COLUMN_INDEX).setValue(newTicketId);
      Logger.log(`Assigned new Ticket ID for row ${actualRowNumber}: ${newTicketId}`);
      rowNeedsProcessingThisPass = true;
    }

    if (rowNeedsProcessingThisPass) rowsUpdated++;
  });

  if (rowsUpdated > 0) Logger.log(`Finished processing. Total rows with updates: ${rowsUpdated}.`);
  else Logger.log("Finished processing. No new row data needed updating.");
}

// --- Helper Functions ---

/**
 * Determines the location name from a channel ID using the configuration map.
 * @param {string} channelId The Slack channel ID.
 * @return {string|null} The location name or null if not found.
 */
function determineLocation(channelId) {
  return CHANNEL_TO_LOCATION_MAP[channelId] || null;
}

/**
 * Searches message text for keywords defined in the CATEGORY_DEFINITIONS.
 * @param {string} messageText The text of the Slack message.
 * @return {string[]} An array of found category names.
 */
function findCategoriesInMessage(messageText) {
  const foundCategories = new Set();
  const lowerCaseMessage = messageText.toLowerCase();
  
  CATEGORY_DEFINITIONS.forEach(categoryDef => {
    for (const search of categoryDef.searches) {
      let isMatch = false;
      if (search.type === "exact") {
        // Case-sensitive match using word boundaries
        isMatch = search.phrases.some(phrase => new RegExp(`\\b${phrase.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`).test(messageText));
      } else if (search.type === "or") {
        // Case-insensitive match using word boundaries
        isMatch = search.phrases.some(phrase => new RegExp(`\\b${phrase.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i').test(lowerCaseMessage));
      }
      
      if (isMatch) {
        foundCategories.add(categoryDef.name);
        break; // Move to the next category once one match is found for this category
      }
    }
  });
  return Array.from(foundCategories);
}

// --- Slack API & Permalink Parsing ---

/**
 * Fetches message details from the Slack API using a permalink.
 * @param {string} permalink The full URL to the Slack message.
 * @return {Object|null} An object with message details or null on failure.
 */
function getMessageDetailsFromSlackLink(permalink) {
  if (!SLACK_BOT_TOKEN) {
    Logger.log("SLACK_BOT_TOKEN is not set in Script Properties. Cannot fetch message.");
    return null;
  }
  const parsed = parseSlackPermalink(permalink);
  if (!parsed) {
    Logger.log(`Invalid Slack message link format: ${permalink}`);
    return null;
  }
  
  const { channelId, messageTs } = parsed;
  const slackApiUrl = `https://slack.com/api/conversations.history?channel=${channelId}&latest=${messageTs}&inclusive=true&limit=1`;
  const options = {
    'method': 'get',
    'headers': { 'Authorization': `Bearer ${SLACK_BOT_TOKEN}` },
    'muteHttpExceptions': true // Prevents script from stopping on HTTP errors (e.g., 404)
  };

  try {
    const response = UrlFetchApp.fetch(slackApiUrl, options);
    const jsonResponse = JSON.parse(response.getContentText());
    if (jsonResponse.ok && jsonResponse.messages && jsonResponse.messages.length > 0) {
      const message = jsonResponse.messages[0];
      return { text: message.text, user: message.user, channel: channelId };
    } else {
      Logger.log(`Slack API Error for link ${permalink}: ${jsonResponse.error || 'Message not found'}`);
      return null;
    }
  } catch (e) {
    Logger.log(`Exception fetching Slack message for link ${permalink}: ${e.toString()}`);
    return null;
  }
}

/**
 * Parses a Slack permalink to extract the channel ID and message timestamp (ts).
 * @param {string} permalink The full URL to the Slack message.
 * @return {Object|null} An object with {channelId, messageTs} or null if invalid.
 */
function parseSlackPermalink(permalink) {
  if (typeof permalink !== 'string') return null;
  // Regex to find channel ID (starts with C, G, D, or R) and message timestamp (p-prefixed)
  const match = permalink.match(/archives\/([CGRUD][A-Z0-9]+)\/p(\d{16})/);
  if (match && match.length === 3) {
    const channelId = match[1];
    const messageTs = `${match[2].substring(0, 10)}.${match[2].substring(10)}`;
    return { channelId, messageTs };
  }
  return null;
}

// --- Daily Status Report Functions ---

/**
 * Sends a message to a specified Slack channel.
 * @param {string} channelId The ID of the channel to post in.
 * @param {string} text The message text to send. Supports Slack mrkdwn.
 */
function sendSlackMessage(channelId, text) {
  if (!SLACK_BOT_TOKEN) {
    Logger.log("SLACK_BOT_TOKEN is not set. Cannot send report.");
    return;
  }
  const payload = {
    channel: channelId,
    text: text,
    mrkdwn: true
  };
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': `Bearer ${SLACK_BOT_TOKEN}` },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', options);
    const jsonResponse = JSON.parse(response.getContentText());
    if (jsonResponse.ok) {
      Logger.log(`Successfully sent message to channel ${channelId}.`);
    } else {
      Logger.log(`Failed to send message. Slack API Error: ${jsonResponse.error}`);
    }
  } catch (e) {
    Logger.log(`Exception while sending message: ${e.toString()}`);
  }
}

/**
 * Compiles and sends the daily status report on open tickets.
 */
function sendDailyStatusReport() {
  const today = new Date();
  const dayOfWeek = today.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) { // 0=Sun, 6=Sat
    Logger.log('Skipping daily status report on weekend.');
    return;
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) {
    Logger.log(`Daily Report Error: Sheet '${TARGET_SHEET_NAME}' not found.`);
    return;
  }

  const sheetUrl = spreadsheet.getUrl(); 
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW_COUNT) {
    sendSlackMessage(REPORTING_CHANNEL_ID, "*IT Tickets Status Report*: There are currently no open tickets. Happy days!");
    return;
  }

  // Count rows where the status column is "Open".
  const statusValues = sheet.getRange(HEADER_ROW_COUNT + 1, STATUS_COLUMN_INDEX, lastRow - HEADER_ROW_COUNT, 1).getValues();
  const openTicketsCount = statusValues.filter(row => row[0] && row[0].toString().toLowerCase() === 'open').length;

  let message = openTicketsCount > 0 
    ? `*IT Tickets Status Report*: There are *${openTicketsCount}* open tickets as of today. Please review the <${sheetUrl}|tracker> for any outstanding issues.`
    : "*IT Tickets Status Report*: There are currently no open tickets. Happy days!";

  Logger.log(`Sending daily report: ${message}`);
  sendSlackMessage(REPORTING_CHANNEL_ID, message);
}

// --- Initial Setup Function ---

/**
 * !! RUN THIS FUNCTION ONCE MANUALLY TO SET UP THE DAILY TRIGGERS !!
 * Deletes any old triggers for this project and creates new ones for the
 * daily report and duplicate check to run every weekday.
 */
function createOrReplaceDailyTriggers() {
  // --- Delete all old triggers to prevent duplicates ---
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    const handlerFunction = trigger.getHandlerFunction();
    // Check for all functions that might have been triggered previously.
    if (['sendDailyStatusReport', 'removeDuplicatesAndProcess', 'createOrReplaceDailyTriggers'].includes(handlerFunction)) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`Deleted existing trigger for ${handlerFunction}.`);
    }
  }
    // --- Create trigger for duplicate check and processing (runs before report) ---
  ScriptApp.newTrigger('removeDuplicatesAndProcess')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(25) // Run at ~9:25 AM ET
    .inTimezone('America/Toronto')
    .create();
  Logger.log('Created new trigger for removeDuplicatesAndProcess.');

  // --- Create trigger for the daily status report ---
  ScriptApp.newTrigger('sendDailyStatusReport')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(30) // Run at ~9:30 AM ET
    .inTimezone('America/Toronto')
    .create();
  Logger.log('Created new trigger for sendDailyStatusReport.');

  SpreadsheetApp.getUi().alert('✅ Success! New daily triggers have been set for 9:30 AM ET (Weekdays only).');
}

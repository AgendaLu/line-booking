// Configuration
const CONFIG = {
  LINE_CHANNEL_ACCESS_TOKEN: 'YOUR_CHANNEL_ACCESS_TOKEN',
  LINE_CHANNEL_SECRET: 'YOUR_CHANNEL_SECRET',
  RAW_SHEET_ID: 'YOUR_RAW_SHEET_ID',
  SUMMARY_SHEET_ID: 'YOUR_SUMMARY_SHEET_ID',
  RAW_SHEET_NAME: 'Messages',
  SUMMARY_SHEET_NAME: 'Bookings',
  MAX_SEATS: 10,
  CLEAR_TIME: 1,
  VALID_INCREMENTS: new Set([
    '+1', '+2', '+3', '+4', 
    '加一', '加二', '加三', '加四',
    '-1', '-2', '-3', '-4',
    '減一', '減二', '減三', '減四'
  ]),
  CHECK_KEYWORD: 'check',
  NAME_LIST_KEYWORD: 'name'
};

// Main webhook handler
function doPost(e) {
  Logger.log('Received webhook: ' + JSON.stringify(e));
  
  if (!e || (!e.postData && !e.parameter)) {
    logError('No data received in webhook');
    return createResponse('No data received', 400);
  }
  
  try {
    const webhookData = getWebhookData(e);
    if (!webhookData) {
      return createResponse('Invalid data format', 400);
    }
    
    webhookData.events.forEach(event => {
      if (event.type === 'message' && event.message.type === 'text') {
        const messageText = event.message.text.trim().toLowerCase();
        
        if (messageText === CONFIG.CHECK_KEYWORD) {
          handleStatusCheck(event);
        } else if (messageText === CONFIG.NAME_LIST_KEYWORD) {
          handleNameList(event);
        } else if (isValidBookingMessage(messageText)) {
          handleBooking(event);
        }
      }
    });
    
    return createResponse('OK', 200);
  } catch (error) {
    logError('Error in doPost: ' + error.toString());
    return createResponse('Internal server error', 500);
  }
}

// Message processing
function isValidBookingMessage(message) {
  return CONFIG.VALID_INCREMENTS.has(message);
}

function handleBooking(event) {
  try {
    const currentTotal = getCurrentBookingTotal();
    const userCurrentBookings = getCurrentUserBookings(event.source.userId);
    const increment = parseIncrement(event.message.text);
    
    // Prepare booking info with initial increment
    const bookingInfo = {
      timestamp: new Date(event.timestamp),
      groupId: event.source.groupId || 'N/A',
      userId: event.source.userId || 'N/A',
      userName: getUserName(event.source.userId),
      increment: increment,
      messageId: event.message.id
    };

    // Check if message is already logged
    if (isMessageAlreadyLogged(bookingInfo.messageId)) {
      Logger.log('Duplicate message detected, skipping: ' + bookingInfo.messageId);
      return;
    }

    // Handle cancellations
    if (increment < 0) {
      const cancelAmount = Math.abs(increment);
      if (cancelAmount > userCurrentBookings) {
        bookingInfo.increment = 0; // Mark as invalid cancellation
        logRawMessage(bookingInfo);
        replyMessage(event.replyToken, `無法取消預約。您目前只預訂了 ${userCurrentBookings} 位。`);
        return;
      }
    } 
    // Handle new bookings
    else {
      // Check total seats limit
      if (currentTotal + increment > CONFIG.MAX_SEATS) {
        bookingInfo.increment = 0; // Mark as invalid booking
        logRawMessage(bookingInfo);
        replyMessage(event.replyToken, `抱歉，目前剩餘座位不足。現有預訂人數: ${currentTotal}，最大座位數: ${CONFIG.MAX_SEATS}`);
        return;
      }
      
      // Check per-user seats limit
      const userNewTotal = userCurrentBookings + increment;
      if (userNewTotal > CONFIG.MAX_SEATS_PER_USER) {
        bookingInfo.increment = 0; // Mark as exceeding user limit
        logRawMessage(bookingInfo);
        replyMessage(event.replyToken, 
          `您已報名 ${userCurrentBookings} 位，每人報名限制${CONFIG.MAX_SEATS_PER_USER} 位。`);
        return;
      }
    }
    
    // If we get here, the booking is valid
    logRawMessage(bookingInfo);
    recalculateBookingSummary();
    
    // Get new total after recalculation
    const newTotal = getCurrentBookingTotal();
    replyMessage(event.replyToken, `已記錄${increment > 0 ? '預訂' : '取消'} ${Math.abs(increment)} 位。目前總計: ${newTotal} 位`);
  } catch (error) {
    logError('Error handling booking: ' + error.toString());
  }
}

function handleNameList(event) {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SUMMARY_SHEET_ID).getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    // Skip header row and filter out rows with zero or empty bookings
    const bookings = data.slice(1)
      .filter(row => row[4] && row[4] > 0)
      .map(row => ({
        name: row[3],
        seats: row[4]
      }))
      .sort((a, b) => b.seats - a.seats); // Sort by number of seats, descending
    
    if (bookings.length === 0) {
      replyMessage(event.replyToken, '目前還沒有人預訂');
      return;
    }
    
    const totalSeats = bookings.reduce((sum, booking) => sum + booking.seats, 0);
    const nameList = bookings
      .map(booking => `${booking.name}: ${booking.seats} 位`)
      .join('\n');
    
    const message = `預訂名單：\n${nameList}\n\n總計：${totalSeats} 位`;
    replyMessage(event.replyToken, message);
  } catch (error) {
    logError('Error handling name list request: ' + error.toString());
  }
}

function handleStatusCheck(event) {
  try {
    const currentTotal = getCurrentBookingTotal();
    const remaining = CONFIG.MAX_SEATS - currentTotal;
    const message = `目前預訂狀況:\n總預訂人數: ${currentTotal} 位\n剩餘座位: ${remaining} 位`;
    replyMessage(event.replyToken, message);
  } catch (error) {
    logError('Error handling status check: ' + error.toString());
  }
}

// Spreadsheet operations
function logRawMessage(bookingInfo) {
  const sheet = SpreadsheetApp.openById(CONFIG.RAW_SHEET_ID).getSheetByName(CONFIG.RAW_SHEET_NAME);
  
  // Check if message already exists
  if (isMessageAlreadyLogged(bookingInfo.messageId)) {
    Logger.log('Duplicate message detected, skipping: ' + bookingInfo.messageId);
    return;
  }
  
  const rowData = [
    bookingInfo.timestamp,
    bookingInfo.groupId,
    bookingInfo.userId,
    bookingInfo.userName,
    bookingInfo.increment,
    bookingInfo.messageId
  ];
  
  sheet.appendRow(rowData);
}

function isMessageAlreadyLogged(messageId) {
  const sheet = SpreadsheetApp.openById(CONFIG.RAW_SHEET_ID).getSheetByName(CONFIG.RAW_SHEET_NAME);
  
  // Get the message ID column (assuming it's the last column, index 5)
  const messageIdColumn = sheet.getRange('F:F').getValues();
  
  // Check if messageId exists in the column
  return messageIdColumn.flat().includes(messageId);
}

function getCurrentUserBookings(userId) {
  const sheet = SpreadsheetApp.openById(CONFIG.SUMMARY_SHEET_ID).getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // Find user's row (userId is in column 2, index 2)
  const userRow = data.find(row => row[2] === userId);
  return userRow ? (parseInt(userRow[4]) || 0) : 0;
}

function recalculateBookingSummary() {
  try {
    // Get raw messages
    const rawSheet = SpreadsheetApp.openById(CONFIG.RAW_SHEET_ID).getSheetByName(CONFIG.RAW_SHEET_NAME);
    const rawData = rawSheet.getDataRange().getValues();
    const headers = rawData[0];
    
    // Create column index map
    const colIndex = {
      timestamp: headers.indexOf('Timestamp'),
      groupId: headers.indexOf('Group ID'),
      userId: headers.indexOf('User ID'),
      userName: headers.indexOf('User Name'),
      increment: headers.indexOf('Increment')
    };
    
    // Skip header row and process all messages
    const bookings = new Map(); // Map to store userId -> booking info
    
    // Process each message and aggregate by user
    rawData.slice(1).forEach(row => {
      const userId = row[colIndex.userId];
      const increment = parseInt(row[colIndex.increment]) || 0;
      
      if (bookings.has(userId)) {
        const existing = bookings.get(userId);
        existing.total += increment;
        if (row[colIndex.timestamp] > existing.timestamp) {
          existing.timestamp = row[colIndex.timestamp];
        }
      } else {
        bookings.set(userId, {
          timestamp: row[colIndex.timestamp],
          groupId: row[colIndex.groupId],
          userName: row[colIndex.userName],
          total: increment
        });
      }
    });
    
    // Clear and update summary sheet
    const summarySheet = SpreadsheetApp.openById(CONFIG.SUMMARY_SHEET_ID).getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
    summarySheet.clearContents();
    
    // Add headers
    const summaryHeaders = ['Last Update', 'Group ID', 'User ID', 'User Name', 'Total Seats'];
    summarySheet.getRange(1, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
    
    // Convert Map to array and sort by timestamp
    const summaryData = Array.from(bookings.entries())
      .map(([userId, info]) => [
        info.timestamp,
        info.groupId,
        userId,
        info.userName,
        info.total
      ])
      .filter(row => row[4] > 0) // Only include rows with positive total
      .sort((a, b) => b[0] - a[0]); // Sort by timestamp descending
    
    // Write data if we have any
    if (summaryData.length > 0) {
      summarySheet.getRange(2, 1, summaryData.length, summaryHeaders.length)
        .setValues(summaryData);
    }
    
    // Reapply formatting
    summarySheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
    summarySheet.autoResizeColumns(1, summaryHeaders.length);
    
  } catch (error) {
    logError('Error recalculating booking summary: ' + error.toString());
    throw error; // Re-throw to handle in calling function
  }
}

function getCurrentBookingTotal() {
  const sheet = SpreadsheetApp.openById(CONFIG.SUMMARY_SHEET_ID).getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // Skip header row and sum the booking column (index 4)
  return data.slice(1).reduce((total, row) => total + (parseInt(row[4]) || 0), 0);
}

// LINE API functions
function replyMessage(replyToken, message) {
  const url = 'https://api.line.me/v2/bot/message/reply';
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + CONFIG.LINE_CHANNEL_ACCESS_TOKEN
    },
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: [{
        type: 'text',
        text: message
      }]
    })
  };
  
  UrlFetchApp.fetch(url, options);
}

function getUserName(userId) {
  if (!userId) return 'Unknown';
  
  try {
    const url = `https://api.line.me/v2/bot/profile/${userId}`;
    const options = {
      headers: {
        'Authorization': 'Bearer ' + CONFIG.LINE_CHANNEL_ACCESS_TOKEN
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      const profile = JSON.parse(response.getContentText());
      return profile.displayName || 'Unknown';
    }
    return 'Error fetching name';
  } catch (error) {
    logError('Error fetching user profile: ' + error.toString());
    return 'Error';
  }
}

// Utility functions
function parseIncrement(message) {
  const numberMap = {
    '+1': 1, '加一': 1,
    '+2': 2, '加二': 2,
    '+3': 3, '加三': 3,
    '+4': 4, '加四': 4,
    '-1': -1, '減一': -1,
    '-2': -2, '減二': -2,
    '-3': -3, '減三': -3,
    '-4': -4, '減四': -4
  };
  return numberMap[message] || 0;
}

function getWebhookData(e) {
  try {
    if (e.postData && e.postData.contents) {
      return JSON.parse(e.postData.contents);
    }
    if (e.parameter && e.parameter.data) {
      return JSON.parse(e.parameter.data);
    }
    return null;
  } catch (error) {
    logError('Error parsing webhook data: ' + error.toString());
    return null;
  }
}

function createResponse(message, code) {
  return ContentService
    .createTextOutput(JSON.stringify({ message: message }))
    .setMimeType(ContentService.MimeType.JSON)
    .setStatusCode(code);
}

function logError(message) {
  Logger.log('ERROR: ' + message);
}

// Setup functions
function setupSheets() {
  // Setup raw messages sheet
  const rawSheet = SpreadsheetApp.openById(CONFIG.RAW_SHEET_ID).getSheetByName(CONFIG.RAW_SHEET_NAME);
  const rawHeaders = ['Timestamp', 'Group ID', 'User ID', 'User Name', 'Increment', 'Message ID'];
  setupSheet(rawSheet, rawHeaders);
  
  // Setup booking summary sheet
  const summarySheet = SpreadsheetApp.openById(CONFIG.SUMMARY_SHEET_ID).getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
  const summaryHeaders = ['Last Update', 'Group ID', 'User ID', 'User Name', 'Total Seats'];
  setupSheet(summarySheet, summaryHeaders);
}

function setupSheet(sheet, headers) {
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sheet.autoResizeColumns(1, headers.length);
}

function setupDailyTrigger() {
  // Remove any existing triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyCleanup') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new trigger for specified time
  ScriptApp.newTrigger('dailyCleanup')
    .timeBased()
    .atHour(CONFIG.CLEAR_TIME)
    .everyDays(1)
    .create();
    
  Logger.log('Daily cleanup trigger set for ' + CONFIG.CLEAR_TIME + ':00');
}

function dailyCleanup() {
  try {
    Logger.log('Starting daily cleanup at ' + new Date());
    
    // Clear both sheets
    clearSheets();
    
    // Recalculate summary (in case there are any remaining valid bookings)
    recalculateBookingSummary();
    
    Logger.log('Daily cleanup completed');
  } catch (error) {
    logError('Error in daily cleanup: ' + error.toString());
  }
}

function clearSheets() {
  // Clear raw messages sheet
  const rawSheet = SpreadsheetApp.openById(CONFIG.RAW_SHEET_ID).getSheetByName(CONFIG.RAW_SHEET_NAME);
  const rawHeaders = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getValues()[0];
  rawSheet.clearContents();
  rawSheet.getRange(1, 1, 1, rawHeaders.length).setValues([rawHeaders]);
  
  // Clear summary sheet
  const summarySheet = SpreadsheetApp.openById(CONFIG.SUMMARY_SHEET_ID).getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
  const summaryHeaders = summarySheet.getRange(1, 1, 1, summarySheet.getLastColumn()).getValues()[0];
  summarySheet.clearContents();
  summarySheet.getRange(1, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);
  
  Logger.log('Sheets cleared, headers preserved');
}

function doGet() {
  return createResponse('Booking bot is active', 200);
}
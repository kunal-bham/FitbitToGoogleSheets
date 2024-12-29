const CLIENT_ID = 'YOUR_FITBIT_CLIENT_ID';
const CLIENT_SECRET = 'YOUR_FITBIT_CLIENT_SECRET';
const SPREADSHEET_ID = 'YOUR_GOOGLE_SHEET_ID';
const CALENDAR_ID = 'YOUR_CALENDAR_ID';
const SHEET_NAME = 'Fitbit Data';

/**
 * Writes data to Google Sheet and creates calendar event
 */
function writeToSheet(rowData) {
  // Open spreadsheet and get sheet
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }

  // Get next empty row
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const targetRow = lastRow + 1;

  // Write data to sheet
  sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);

  // Create calendar event for the day
  const calendar = CalendarApp.getCalendarById(CALENDAR_ID) || CalendarApp.getDefaultCalendar();
  const eventDate = new Date(rowData[0]);
  const eventTitle = `Health Summary for ${Utilities.formatDate(eventDate, 'America/Chicago', 'MM/dd')}`;
  
  // Format event description with key metrics
  const description = [
    `${rowData[5]} Hours Slept`,
    `${Math.round(rowData[1])} Steps`,
    `${rowData[3]} Bed Time`,
    `${rowData[4]} Wake-up Time`
  ].join('<br/>');

  // Create all-day event
  calendar.createAllDayEvent(eventTitle, eventDate, {
    description: description
  });
}

/**
 * Main function to fetch Fitbit data and write to Google Sheets
 */
function getFitbitData() {
  const service = getFitbitService();
  if (!service.hasAccess()) {
    Logger.log('Authorization required: ' + service.getAuthorizationUrl());
    return;
  }

  // Check if we're using a custom date
  const customDate = PropertiesService.getUserProperties().getProperty('customDate');
  const now = new Date();
  
  // Get health data from 2 days ago relative to the target date
  let targetDate;
  if (customDate) {
    targetDate = new Date(customDate);
    Logger.log('Using custom date:', customDate);
  } else {
    targetDate = now;
  }
  
  const dayBeforeYesterday = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate() - 2);
  const healthDateString = Utilities.formatDate(dayBeforeYesterday, 'America/Chicago', 'yyyy-MM-dd');
  Logger.log(`Fetching health data for: ${healthDateString}`);

  try {
    // Fetch data from Fitbit API
    const endpoints = {
      activities: `/activities/date/${healthDateString}.json`,
      heartRate: `/activities/heart/date/${healthDateString}/1d.json`,
      hrv: `/hrv/date/${healthDateString}.json`,
      temperature: `/temp/skin/date/${healthDateString}.json`,
      spo2: `/spo2/date/${healthDateString}.json`,
      breathingRate: `/br/date/${healthDateString}.json`
    };

    const data = {};
    for (const [key, endpoint] of Object.entries(endpoints)) {
      try {
        data[key] = fetchFitbitEndpoint(service, endpoint);
        Logger.log(`Raw ${key} data for ${healthDateString}: ` + JSON.stringify(data[key], null, 2));
      } catch (error) {
        Logger.log(`Error fetching ${key}: ${error.toString()}`);
        data[key] = {};
      }
    }

    // Get sleep data from yesterday relative to the target date
    const yesterday = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate() - 1);
    const sleepDateString = Utilities.formatDate(yesterday, 'America/Chicago', 'yyyy-MM-dd');
    Logger.log(`Fetching sleep data for: ${sleepDateString}`);
    
    try {
      data.sleep = fetchFitbitEndpoint(service, `/sleep/date/${sleepDateString}.json`);
      Logger.log(`Raw sleep data for ${sleepDateString}: ` + JSON.stringify(data.sleep, null, 2));
    } catch (error) {
      Logger.log(`Error fetching sleep data for ${sleepDateString}: ${error.toString()}`);
      data.sleep = {};
    }

    // Extract activity metrics
    const steps = data.activities?.summary?.steps;
    const restingHR = data.heartRate['activities-heart']?.[0]?.value?.restingHeartRate;
    
    // Extract activity minutes from the API response
    const activityMinutes = {
      veryActive: data.activities?.summary?.veryActiveMinutes || 0,
      fairlyActive: data.activities?.summary?.fairlyActiveMinutes || 0,
      lightlyActive: data.activities?.summary?.lightlyActiveMinutes || 0,
      sedentary: data.activities?.summary?.sedentaryMinutes || 0
    };

    Logger.log('Activity Minutes:', JSON.stringify(activityMinutes, null, 2));

    // Extract health metrics with detailed logging
    let breathingRate = null;
    let hrv = null;
    let skinTemp = null;
    let spo2 = null;

    try {
      // Breathing Rate - using confirmed working path
      if (data.breathingRate?.br?.[0]?.value?.breathingRate) {
        breathingRate = data.breathingRate.br[0].value.breathingRate;
        Logger.log('Breathing Rate:', breathingRate);
      }

      // HRV
      if (data.hrv?.hrv?.[0]?.value?.dailyRmssd) {
        hrv = data.hrv.hrv[0].value.dailyRmssd;
        Logger.log('HRV:', hrv);
      }

      // Skin Temperature
      if (data.temperature?.tempSkin?.[0]?.value?.nightlyRelative) {
        skinTemp = data.temperature.tempSkin[0].value.nightlyRelative;
        Logger.log('Skin Temp:', skinTemp);
      }

      // SpO2
      if (data.spo2?.value?.avg) {
        spo2 = data.spo2.value.avg;
        Logger.log('SpO2:', spo2);
      }
    } catch (error) {
      Logger.log('Error extracting health metrics:', error.toString());
    }

    // Process sleep data
    let sleepMetrics = null;
    let summaryMetrics = null;
    let bedTime = null;
    let wakeTime = null;
    let totalWakeMinutes = 0;

    if (data.sleep?.sleep?.length > 0) {
      // Get main sleep session
      const mainSleep = data.sleep.sleep.find(s => s.isMainSleep) || 
                       data.sleep.sleep.reduce((a, b) => (a.timeInBed > b.timeInBed ? a : b));

      // Extract and format bed time and wake time
      if (mainSleep.startTime) {
        const bedDateTime = new Date(mainSleep.startTime);
        bedTime = Utilities.formatDate(bedDateTime, 'America/Chicago', 'h:mm a');
        Logger.log(`Bed time: ${bedTime}`);
      }
      
      if (mainSleep.endTime) {
        const wakeDateTime = new Date(mainSleep.endTime);
        wakeTime = Utilities.formatDate(wakeDateTime, 'America/Chicago', 'h:mm a');
        Logger.log(`Wake time: ${wakeTime}`);
      }

      sleepMetrics = {
        timeInBed: mainSleep.timeInBed,
        totalMinutesAsleep: mainSleep.minutesAsleep,
        totalSleepRecords: data.sleep.summary.totalSleepRecords
      };

      // Calculate total wake time by counting all minutes with value 2 or 3
      if (mainSleep.minuteData) {
        totalWakeMinutes = mainSleep.minuteData.reduce((count, minute) => {
          const isWake = minute.value === "2" || minute.value === "3";
          return count + (isWake ? 1 : 0);
        }, 0);
      }

      // Extract sleep stages from summary if available
      if (data.sleep.summary.stages) {
        const stages = data.sleep.summary.stages;
        summaryMetrics = {
          light: stages.light || 0,
          deep: stages.deep || 0,
          rem: stages.rem || 0,
          wake: totalWakeMinutes
        };
      }
    }

    // Prepare row data with all metrics
    const rowData = [
      healthDateString,
      steps,
      restingHR,
      bedTime,
      wakeTime,
      sleepMetrics?.totalMinutesAsleep ? (sleepMetrics.totalMinutesAsleep / 60).toFixed(2) : null,
      totalWakeMinutes,
      summaryMetrics?.light || 0,
      summaryMetrics?.deep || 0,
      summaryMetrics?.rem || 0,
      breathingRate,
      hrv,
      skinTemp,
      spo2,
      activityMinutes.veryActive,
      activityMinutes.fairlyActive,
      activityMinutes.lightlyActive,
      activityMinutes.sedentary
    ];

    writeToSheet(rowData);
    Logger.log('Successfully wrote data to sheet');

  } catch (error) {
    Logger.log('Error in getFitbitData:', error.toString());
    throw error;
  }
}

/**
 * Fetches data from Fitbit API endpoint with retry logic
 */
function fetchFitbitEndpoint(service, endpoint, retryCount = 0) {
  const MAX_RETRIES = 3;
  const RETRY_DELAY_MS = 3000;
  const RATE_LIMIT_DELAY_MS = 5000;

  try {
    Utilities.sleep(1000);

    const response = UrlFetchApp.fetch(
      `https://api.fitbit.com/1/user/-${endpoint}`,
      {
        headers: {
          Authorization: `Bearer ${service.getAccessToken()}`,
          'Accept-Language': 'en_US'
        },
        muteHttpExceptions: true
      }
    );

    const responseCode = response.getResponseCode();
    const contentText = response.getContentText();

    if (responseCode === 429) {
      Logger.log(`Rate limited on ${endpoint}, waiting ${RATE_LIMIT_DELAY_MS/1000} seconds...`);
      Utilities.sleep(RATE_LIMIT_DELAY_MS);
      
      if (retryCount < MAX_RETRIES) {
        return fetchFitbitEndpoint(service, endpoint, retryCount + 1);
      }
    }

    if (responseCode !== 200) {
      if (retryCount < MAX_RETRIES) {
        Logger.log(`Error ${responseCode} on ${endpoint}, retrying in ${RETRY_DELAY_MS/1000} seconds...`);
        Utilities.sleep(RETRY_DELAY_MS);
        return fetchFitbitEndpoint(service, endpoint, retryCount + 1);
      }
      throw new Error(`Failed after ${MAX_RETRIES} retries: ${contentText}`);
    }

    return JSON.parse(contentText);
  } catch (error) {
    Logger.log(`Error fetching ${endpoint}:`, error.toString());
    throw error;
  }
}

/**
 * Sets up OAuth2 service for Fitbit
 */
function getFitbitService() {
  return OAuth2.createService('Fitbit')
    .setAuthorizationBaseUrl('https://www.fitbit.com/oauth2/authorize')
    .setTokenUrl('https://api.fitbit.com/oauth2/token')
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope([
      'activity',
      'sleep',
      'heartrate',
      'temperature',
      'oxygen_saturation',
      'cardio_fitness',
      'respiratory_rate'
    ].join(' '))
    .setTokenHeaders({
      Authorization: `Basic ${Utilities.base64Encode(CLIENT_ID + ':' + CLIENT_SECRET)}`
    });
}

/**
 * Handles OAuth callback
 */
function authCallback(request) {
  const service = getFitbitService();
  const authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  }
  return HtmlService.createHtmlOutput('Failed to authorize.');
}

/**
 * Creates a daily trigger to run the script
 */
function createDailyTrigger() {
  // Remove existing triggers
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'getFitbitData') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new trigger to run at 1 AM daily
  ScriptApp.newTrigger('getFitbitData')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
}

/**
 * Forces reauthorization by clearing current tokens
 */
function clearAndReauthorize() {
  const service = getFitbitService();
  service.reset();
  Logger.log('Authorization URL: ' + service.getAuthorizationUrl());
}

/**
 * Fetches data for the past two weeks
 */
function getTwoWeeksData() {
  Logger.log('\n=== FETCHING TWO WEEKS OF DATA ===');
  
  // First, ensure headers are set up
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
  }
  
  // Clear existing content
  sheet.clear();
  
  // Set up headers
  const headers = [
    'Date', 'Steps', 'Resting HR', 'Bed Time', 'Wake Time', 'Total Sleep Time', 
    'Wake Time', 'Light Sleep', 'Deep Sleep', 'REM Sleep', 'Breathing Rate', 
    'Heart Rate Variability', 'Skin Temperature', 'Oxygen Saturation',
    'Very Active Minutes', 'Fairly Active Minutes', 'Lightly Active Minutes', 'Sedentary Minutes'
  ];

  // Add headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  
  // Calculate date range (past 14 days)
  const endDate = new Date();
  endDate.setDate(endDate.getDate() - 1);  // Yesterday
  const startDate = new Date(endDate);
  startDate.setDate(startDate.getDate() - 13);  // 14 days before yesterday

  Logger.log(`Fetching data from ${startDate.toISOString().split('T')[0]} to ${endDate.toISOString().split('T')[0]}\n`);

  // Process each day
  const currentDate = new Date(startDate);
  while (currentDate <= endDate) {
    const dateString = Utilities.formatDate(currentDate, 'America/Chicago', 'yyyy-MM-dd');
    Logger.log(`\n=== Processing ${dateString} ===`);

    try {
      // Store the date in PropertiesService for getFitbitData to use
      PropertiesService.getUserProperties().setProperty('customDate', dateString);

      // Call getFitbitData for this date
      getFitbitData();

      // Add significant delay between days to avoid rate limits
      Logger.log(`Waiting 10 seconds before processing next day...`);
      Utilities.sleep(10000);  // 10 seconds between days
    } catch (error) {
      Logger.log(`Error processing ${dateString}: ${error.toString()}`);
      // Add extra delay after an error
      Utilities.sleep(15000);  // 15 seconds after error
    }

    // Move to next day
    currentDate.setDate(currentDate.getDate() + 1);
  }

  // Clean up - remove custom date property
  PropertiesService.getUserProperties().deleteProperty('customDate');

  Logger.log('\n=== COMPLETED FETCHING TWO WEEKS OF DATA ===');
} 
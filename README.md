# FitbitToGoogleSheets

A Google Apps Script project that automatically fetches Fitbit health data and logs it to Google Sheets. The script collects comprehensive health metrics including:

- Daily activity (steps)
- Heart rate metrics (resting HR, HRV)
- Sleep data (duration, stages, bed/wake times)
- Advanced health metrics (breathing rate, SpO2, skin temperature)

## Setup

1. Create a new Google Apps Script project
2. Copy the code from `src/Code.js` into your project
3. Set up a Fitbit Developer Account and create an app to get your API credentials
4. Configure the following constants in the script:
   - `CLIENT_ID`: Your Fitbit API Client ID
   - `CLIENT_SECRET`: Your Fitbit API Client Secret
   - `SPREADSHEET_ID`: ID of your Google Sheet
   - `SHEET_NAME`: Name for the data sheet (default: 'Fitbit Data')

## Features

- Daily automatic data collection
- Comprehensive sleep metrics
- Advanced health indicators
- Rate limit handling
- Error retry logic
- Automatic sheet formatting
- Historical data backfill option

## Usage

1. Run the `getFitbitService()` function and authorize the app
2. Use `createDailyTrigger()` to set up automatic daily data collection
3. For historical data, use `getTwoWeeksData()` to fetch the past 14 days

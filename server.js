const express = require('express');
const axios = require('axios');
const cheerio = require('cheerio');
const { google } = require('googleapis');
const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');
const readline = require('readline');

const app = express();
const port = 3000;

// Google Sheets API setup
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_PATH = 'token.json';

let sheets = null;

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
  if (err) {
    console.error('Error loading client secret file:', err);
    return;
  }

  let credentials;
  try {
    credentials = JSON.parse(content);
  } catch (error) {
    console.error('Error parsing credentials.json:', error);
    return;
  }

  authorize(credentials, (authClient) => {
    sheets = google.sheets({ version: 'v4', auth: authClient });
  });
});

function authorize(credentials, callback) {
  const { client_secret, client_id, redirect_uris } = credentials.web;

  if (!client_secret || !client_id || !redirect_uris) {
    console.error('Missing required fields in credentials.json');
    return;
  }

  const oAuth2Client = new google.auth.OAuth2(client_id, client_secret, redirect_uris[0]);

  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) {
      return getAccessToken(oAuth2Client, callback);
    }
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

function getAccessToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error retrieving access token', err);
      oAuth2Client.setCredentials(token);
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

app.get('/scrape-and-upload', async (req, res) => {
  try {
    const { data } = await axios.get('https://en.wikipedia.org/wiki/List_of_FIFA_World_Cup_finals');
    const $ = cheerio.load(data);

    // Find the correct table
    const table = $('caption:contains("List of FIFA World Cup finals")').closest('table');
    const rows = table.find('tbody tr');
    const extractedData = [];

    rows.each((index, row) => {
      if (index < 10) {
        const cells = $(row).find('td');
        const year = $(cells[0]).text().trim();
        const winner = $(cells[1]).text().trim();
        const score = $(cells[2]).text().trim().replace(/\[\d+\]/g, '');
        const runner = $(cells[3]).text().trim();
        extractedData.push([year, winner, score, runner]);
      }
    });

    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.aoa_to_sheet([
      ['YEAR', 'WINNER', 'SCORE', 'RUNNER'],
      ...extractedData,
    ]);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const filePath = path.join(__dirname, 'FIFA_World_Cup_Finals.xlsx');
    xlsx.writeFile(workbook, filePath);

    const spreadsheetId = '1_ih8SM_rn28dhM76SEr3fqbJdSNypky-54pzAKyovCo'; // Replace with your actual Spreadsheet ID
    const range = 'Sheet1!A1'; // Replace with your actual sheet name and cell range
    const valueInputOption = 'RAW';

    const resource = {
      values: extractedData,
    };

    sheets.spreadsheets.values.append(
      {
        spreadsheetId,
        range,
        valueInputOption,
        resource,
      },
      (err, result) => {
        if (err) {
          console.error('The API returned an error: ' + err);
          return res.status(500).send('Error appending data to Google Sheets');
        }

        res.status(200).send('Data appended successfully and Excel file created');
      }
    );
  } catch (error) {
    console.error('Error scraping data:', error);
    res.status(500).send('Error scraping data');
  }
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}/`);
});

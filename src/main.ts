

const { app, BrowserWindow } = require('electron');
import * as path from 'path';
import express from 'express';
import bodyParser from 'body-parser';
const mssql = require('mssql')
import * as exceljs from 'exceljs'


const createWindow = () => {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 1200,
    height: 1000,
    webPreferences: {
      nodeIntegration: true,
      preload: path.join(__dirname, 'preload.js')

    }
  })

  // and load the index.html of the app.
  mainWindow.loadFile('./build/index.html');

  // Open the DevTools.
  // mainWindow.webContents.openDevTools()
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.

// Create an express app
const ExpressApp = express();

// Start the server
ExpressApp.listen(3000, () => {
  console.log('Server listening on port 3000');
});

// Wait until the app is ready
app.whenReady().then(() => {
  // Create the main window
  createWindow();

  // Quit when all windows are closed
  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
      app.quit();

    }
  });
});

// Parse application/x-www-form-urlencoded
ExpressApp.use(bodyParser.urlencoded({ extended: false }));

// Parse application/json
ExpressApp.use(bodyParser.json());

// Handle form submission
ExpressApp.post('/submit', (req: any, res: any) => {
  const stack = req.body.stack;
  const startDate = req.body.start_date;
  const endDate = req.body.end_date;
  const user = req.body.user;

  // Handle the form data as needed
  console.log(`Stack: ${stack}, Start Date: ${startDate}, End Date: ${endDate}`);

  var baseDBURL = '.database.ap-southeast-2.rds.amazonaws.com'; // Change to your common database url

  var serverURL = stack + baseDBURL;


  const QueryConditionArray: string[] = ['']; // Insert any specific WHERE condition variables to iterate through , can be a table or other item

// Adding workbook and worksheet
const workbook = new exceljs.Workbook();
const worksheet = workbook.addWorksheet('Results');

async function runQueryLoop() {
  const dbConfig = {
    user: 'username',
    password: 'password',
    server: serverURL,
    database: 'databasename',
    options: {
      encrypt: false
    }
  }

  await mssql.connect(dbConfig);

  for (let i = 0; i < QueryConditionArray.length; i++) {

    var SQL = ` // Your SQL query goes here
    ${QueryConditionArray[i]}`;

    const result = await mssql.query(SQL);

    // Clear the worksheet
    for (let j = worksheet.rowCount; j >= 1; j--) {
      worksheet.spliceRows(j, 1);
    }

    // Add the new rows at the beginning of the worksheet
    result.recordset.forEach((record:any, index:any ) => {
      worksheet.addRow(Object.values(record));
    });

      // Save the workbook
      const outputFile = path.join('C:', 'Users' ,`${user}`, 'Documents', `${QueryConditionArray[i]}` + ' ' + 'ExampleReport.xlsx');
      await workbook.xlsx.writeFile(outputFile);
  // End of For Loop
  
   // Clear the worksheet
   for (let j = worksheet.rowCount; j >= 1; j--) {
    worksheet.spliceRows(j, 1);
  }

}
await mssql.close();
// End of Async
}

runQueryLoop();
const exportPath = path.join('C:', 'Users' , user, 'Documents');
// Read index.html file and store it in a variable
const exportLocation = exportPath.toString();

// Add banner HTML code to the response
const bannerHTML = `<div class="banner">Generating Reports in: <span id="export-location">${exportLocation}</span></div>`;
// Send a response back to the client
res.send(bannerHTML);
});

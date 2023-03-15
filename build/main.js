"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const { app, BrowserWindow } = require('electron');
const path = __importStar(require("path"));
const express_1 = __importDefault(require("express"));
const body_parser_1 = __importDefault(require("body-parser"));
const mssql = require('mssql');
const exceljs = __importStar(require("exceljs"));
const createWindow = () => {
    // Create the browser window.
    const mainWindow = new BrowserWindow({
        width: 1200,
        height: 1000,
        webPreferences: {
            nodeIntegration: true,
            preload: path.join(__dirname, 'preload.js')
        }
    });
    // and load the index.html of the app.
    mainWindow.loadFile('./build/index.html');
    // Open the DevTools.
    // mainWindow.webContents.openDevTools()
};
// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
// Create an express app
const ExpressApp = (0, express_1.default)();
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
ExpressApp.use(body_parser_1.default.urlencoded({ extended: false }));
// Parse application/json
ExpressApp.use(body_parser_1.default.json());
// Handle form submission
ExpressApp.post('/submit', (req, res) => {
    const stack = req.body.stack;
    const startDate = req.body.start_date;
    const endDate = req.body.end_date;
    const user = req.body.user;
    // Handle the form data as needed
    console.log(`Stack: ${stack}, Start Date: ${startDate}, End Date: ${endDate}`);
    var baseDBURL = '.database.ap-southeast-2.rds.amazonaws.com'; // Change to your common database url
    var serverURL = stack + baseDBURL;
    const QueryConditionArray = ['']; // Insert any specific WHERE condition variables to iterate through , can be a table or other item
    // Adding workbook and worksheet
    const workbook = new exceljs.Workbook();
    const worksheet = workbook.addWorksheet('Results');
    function runQueryLoop() {
        return __awaiter(this, void 0, void 0, function* () {
            const dbConfig = {
                user: 'username',
                password: 'password',
                server: serverURL,
                database: 'databasename',
                options: {
                    encrypt: false
                }
            };
            yield mssql.connect(dbConfig);
            for (let i = 0; i < QueryConditionArray.length; i++) {
                var SQL = ` // Your SQL query goes here
    ${QueryConditionArray[i]}`;
                const result = yield mssql.query(SQL);
                // Clear the worksheet
                for (let j = worksheet.rowCount; j >= 1; j--) {
                    worksheet.spliceRows(j, 1);
                }
                // Add the new rows at the beginning of the worksheet
                result.recordset.forEach((record, index) => {
                    worksheet.addRow(Object.values(record));
                });
                // Save the workbook
                const outputFile = path.join('C:', 'Users', `${user}`, 'Documents', `${QueryConditionArray[i]}` + ' ' + 'ExampleReport.xlsx');
                yield workbook.xlsx.writeFile(outputFile);
                // End of For Loop
                // Clear the worksheet
                for (let j = worksheet.rowCount; j >= 1; j--) {
                    worksheet.spliceRows(j, 1);
                }
            }
            yield mssql.close();
            // End of Async
        });
    }
    runQueryLoop();
    const exportPath = path.join('C:', 'Users', user, 'Documents');
    // Read index.html file and store it in a variable
    const exportLocation = exportPath.toString();
    // Add banner HTML code to the response
    const bannerHTML = `<div class="banner">Generating Reports in: <span id="export-location">${exportLocation}</span></div>`;
    // Send a response back to the client
    res.send(bannerHTML);
});

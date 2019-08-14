require('dotenv').config();

const Joi = require('@hapi/joi');
const winston = require('winston');
const express = require('express');
const Datastore = require('nedb');
const moment = require('moment');
const { Google } = require('./googleapi/googleapi.js');
const app = express();
const bodyParser = require('body-parser');
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

const db = new Datastore();

const LOGS_ID = process.env.LOGS_DIRECTORY_ID;

const google = new Google({
  credentials: './credentials.json',
  tokenPath: './token.json',
  scopes: ['https://www.googleapis.com/xauth/drive.file', 'https://www.googleapis.com/auth/spreadsheets']
});

const logger = winston.createLogger({
  level: 'error',
  format: winston.format.json(),
  transports: [
    new winston.transports.File({ filename: 'logs/error.log', level: 'error' })
  ]
});

/**
 *
 * @param {String} message
 */
function errHandler (message) {
  logger.error({
    message: message
  });
}

(async () => {
  await google.authorize().then((msg) => {
    console.log(msg);
  }).catch(err => {
    errHandler(err);
  });
  start();
})();

async function start () {
  if (verifyLogsFolder().catch(err => { errHandler(err); })) {
    const folderArr = await saveLogKindName().catch(err => {
      errHandler(err);
    });
    checkSpreadsheetAndSaveToDB(folderArr).then((msg) => {
      console.log(msg);
    }).catch(err => {
      errHandler(err);
    });
  }
  function verifyLogsFolder () {
    return new Promise((resolve, reject) => {
      google.listFiles({
        pageSize: 100
      })
        .then(files => {
          if (!files) {
            logger.error('No Folder Found');
          } else {
            if (checkForLogsFolder(files)) {
              resolve(true);
            } else {
              throw new Error('Logs Folder Not Found');
            }
          }
        })
        .catch(err => {
          reject(err);
        });
    });
  }

  function saveLogKindName () {
    return new Promise(async (resolve, reject) => {
      const folders = await google.listFiles({ pageSize: 100, q: `'${LOGS_ID}' in parents` }).catch(err => { reject(err); });
      if (folders) {
        const folderArr = folders.map(folder => {
          return { name: folder.name, id: folder.id, type: 'folder', parent: LOGS_ID };
        });
        db.insert(folderArr, function (err) {
          if (err) {
            reject(err);
          } else {
            console.log('All Folder Saved To db');
            resolve(folderArr);
          }
        });
      }
    });
  }

  /**
   * @param {Array} folderArr
   */
  function checkSpreadsheetAndSaveToDB (folderArr) {
    let promises = [];
    promises.push(new Promise(async (resolve, reject) => {
      for (const folder of folderArr) {
        const files = await google.listFiles({ pageSize: 100, q: `'${folder.id}' in parents` }).catch(err => { reject(err); });
        if (files) {
          if (checkForCurrentDateSpreadsheet(files)) {
            for (const file of files) {
              saveCurrentSpreadsheetDBAndResize(file, folder).then(() => {
                resolve('SpreadsheetID Saved');
              }).catch(err => {
                reject(err);
              });
            }
          } else {
            const lastSpreadsheet = await findDB({ id: files[0].id }).catch(err => {
              errHandler(err);
            });
            const spreadsheet = await google.createSpreadsheet({ fileName: `${moment().format('YYYY-MM')}`, parents: [folder.id] })
              .catch(err => {
                reject(err);
              });
            const resFromAddingSheet = await google.addSheet(spreadsheet.data.id, lastSpreadsheet[0].sheets).catch(err => {
              errHandler(err);
            });
            await google.deleteSheet(spreadsheet.data.id, '0').catch(err => { errHandler(err); });
            await google.addRow(spreadsheet.data.id, resFromAddingSheet.data.replies[0].addSheet.properties.title, { values: [['current_date', 'unix_time_to_date', 'unix_time', 'sub_sub_kind_name', 'log_text']] })
              .catch(err => { errHandler(err); });
            saveAllFilesInLogKindName(folder).then(msg => {
              resolve(msg);
            }).catch(err => {
              reject(err);
            });
          }
        }
      }
    }));
    return Promise.all(promises);
  }
}

app.post('/', async (req, res) => {
  const logKindName = req.body.log_kind_name;
  const subKindName = req.body.sub_kind_name;
  const subSubKindName = req.body.sub_sub_kind_name;
  const logText = req.body.log_text;
  const unixTime = req.body.unix_time;

  const schema = Joi.object().keys({
    logKindName: Joi.string().max(30).required(),
    subKindName: Joi.string().max(30).required(),
    subSubKindName: Joi.string().max(30).required(),
    logText: Joi.string().max(150).required(),
    unixTime: Joi.string().max(11).required()
  });

  let result = Joi.validate({ logKindName: logKindName, subKindName: subKindName, subSubKindName: subSubKindName, logText: logText, unixTime: unixTime }, schema);

  if (result.error === null) {
    const unixTimeToDate = moment.unix(unixTime).utc().format('ddd MMM DD YYYY HH:mm:ss') + ' GMT+07:00 (Indochina Time)';
    const currentDate = moment().format('ddd MMM DD YYYY HH:mm:ss') + ' GMT+07:00 (Indochina Time)';
    const folder = await findDB({ name: logKindName }).catch(err => {
      errHandler(err);
    });
    if (folder.length) {
      const spreadsheets = await findDB({ parent: folder[0].id }).catch(err => { errHandler(err); });
      const boolean = checkForCurrentDateSpreadsheet(spreadsheets);
      if (boolean) {
        const sheets = await findDB({ id: boolean.id, sheets: { $elemMatch: { title: subKindName } } }).catch(err => { errHandler(err); });
        if (sheets.length) {
          await google.addRow(boolean.id, subKindName, { values: [[currentDate, unixTimeToDate, unixTime, subSubKindName, logText]] })
            .catch(err => { errHandler(err); });
        } else {
          const resFromAddingSheet = await google.addSheet(boolean.id, subKindName)
            .catch(err => {
              errHandler(err);
            });
          db.update({ id: resFromAddingSheet.data.spreadsheetId }, { $push: { sheets: { sheetId: resFromAddingSheet.data.replies[0].addSheet.properties.sheetId, title: resFromAddingSheet.data.replies[0].addSheet.properties.title, rowLength: 1 } } }, {}, function (err) {
            if (err) {
              errHandler(err);
            }
          });
          google.resizeColumn(boolean.id, resFromAddingSheet.data.replies[0].addSheet.properties.sheetId, 0, 2, 350, 2, 2, 150, 3, 5, 400).catch(() => {
            errHandler('Failed to resize column');
          });
          await google.addRow(boolean.id, subKindName, { values: [['current_date', 'unix_time_to_date', 'unix_time', 'sub_sub_kind_name', 'log_text'], [currentDate, unixTimeToDate, unixTime, subSubKindName, logText]] })
            .catch(err => { errHandler(err); });
        }
      } else {
        const resFromCreatingSpreadsheet = await google.createSpreadsheet({ fileName: `${moment().format('YYYY-MM')}`, parents: [folder[0].id] })
          .catch(err => { errHandler(err); });
        await google.renameSheet(resFromCreatingSpreadsheet.data.id, 0, subKindName).catch(err => { errHandler(err); });
        await google.addRow(resFromCreatingSpreadsheet.data.id, subKindName, { values: [['current_date', 'unix_time_to_date', 'unix_time', 'sub_sub_kind_name', 'log_text'], [currentDate, unixTimeToDate, unixTime, subSubKindName, logText]] })
          .catch(err => { errHandler(err); });
        saveCurrentSpreadsheetDBAndResize(resFromCreatingSpreadsheet.data, folder[0]).then((msg) => {
          console.log(msg);
        }).catch(err => { errHandler(err); });
      }
    } else {
      const folderCreated = await google.createFolder({ folderName: logKindName, parents: [LOGS_ID] }).catch(err => {
        errHandler(err);
      });
      const folderArr = { name: folderCreated.data.name, id: folderCreated.data.id, type: 'folder', parent: LOGS_ID };
      db.insert(folderArr, function (err) {
        if (err) {
          errHandler(err);
        } else {
          console.log('Log-kind-name Saved To db');
        }
      });
      const resFromCreatingSpreadsheet = await google.createSpreadsheet({ fileName: `${moment().format('YYYY-MM')}`, parents: [folderCreated.data.id] })
        .catch(err => { errHandler(err); });
      await google.renameSheet(resFromCreatingSpreadsheet.data.id, 0, subKindName).catch(err => { errHandler(err); });
      await google.addRow(resFromCreatingSpreadsheet.data.id, subKindName, { values: [['current_date', 'unix_time_to_date', 'unix_time', 'sub_sub_kind_name', 'log_text'], [currentDate, unixTimeToDate, unixTime, subSubKindName, logText]] })
        .catch(err => { errHandler(err); });
      saveCurrentSpreadsheetDBAndResize(resFromCreatingSpreadsheet.data, folderCreated.data).then((msg) => {
        console.log(msg);
      }).catch(err => { errHandler(err); });
    }
  } else {
    errHandler('Invalid Input');
    res.send(result.error);
  }
  res.end('done');
});

app.listen('3000');

// Other Function
/**
 *
 * @param {Object} query
 */
function findDB (query) {
  return new Promise((resolve, reject) => {
    db.find(query, function (err, data) {
      if (err) {
        reject(err);
      } else {
        resolve(data);
      }
    });
  });
}

/**
 *
 * @param {Object} files
 */
function checkForLogsFolder (files) {
  for (const file of files) {
    if (file.id === process.env.LOGS_DIRECTORY_ID) {
      return true;
    }
  }
  return false;
}

/**
 * @param {Array} files
 */
function checkForCurrentDateSpreadsheet (files) {
  const currentDateSheet = files.filter(file => {
    return file.name === moment().format('YYYY-MM');
  })[0];
  if (currentDateSheet) {
    return currentDateSheet;
  } else {
    return false;
  }
}

/**
 *
 * @param {Object} folder
 */
function saveAllFilesInLogKindName (folder) {
  return new Promise(async (resolve, reject) => {
    const files = await google.listFiles({ pageSize: 100, q: `'${folder.id}' in parents` })
      .catch(err => { reject(err); });
    for (const file of files) {
      saveCurrentSpreadsheetDBAndResize(file, folder).then((msg) => {
        resolve(msg);
      }).catch(err => { reject(err); });
    }
  });
}

/**
 * @param {Array} files
 * @param {Object} folder
 */
function saveCurrentSpreadsheetDBAndResize (file, folder) {
  return new Promise(async (resolve, reject) => {
    const resFromGettingSpreadsheet = await google.getSpreadsheet(file.id)
      .catch(err => {
        reject(err);
      });
    const sheetToAdd = resFromGettingSpreadsheet.data.sheets.map(sheet => {
      google.resizeColumn(file.id, sheet.properties.sheetId, 0, 2, 350, 2, 2, 150, 3, 5, 400).catch(err => {
        logger.error('Failed to resize column:', err);
      });
      return { sheetId: sheet.properties.sheetId, title: sheet.properties.title, rowLength: sheet.data[0].rowData.length - 1 };
    });

    db.insert({ name: file.name, id: file.id, type: 'Spreadsheet', sheets: sheetToAdd, parent: folder.id }, function (err) {
      if (err) {
        reject(err);
      } else {
        resolve('Current Year-Month Spreadsheet Created and Saved to DB');
      }
    });
  });
}

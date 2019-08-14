const { google } = require('googleapis');
const readline = require('readline');
const fs = require('fs');
// fs with promise
function readFile (path) {
  return new Promise(function (resolve, reject) {
    fs.readFile(path, (err, output) => {
      if (err) {
        reject(err);
      } else {
        resolve(output);
      }
    });
  });
}

function writeFile (path, data) {
  return new Promise(function (resolve, reject) {
    fs.writeFile(path, data, (err) => {
      if (err) {
        reject(err);
      } else {
        resolve(true);
      }
    });
  });
}

function accessFile (path) {
  return new Promise(function (resolve) {
    fs.access(path, fs.constants.F_OK, (err) => {
      if (err) {
        resolve(false);
      } else {
        resolve(true);
      }
    });
  });
}

class Google {
  /**
   * @param {credentials} String
   * @param {tokenPath} String
   * @param {scopes} String
   */
  constructor ({
    credentials,
    tokenPath,
    scopes
  }) {
    this.credentialsPath = credentials;
    this.tokenPath = tokenPath;
    this.scopes = scopes;
    this.credentialsJSON = null;
    this.tokenJSON = null;
    this.oAuth2Client = null;
    this.drive = null;
  }
  authorize () {
    return new Promise(async (resolve, reject) => {
      await readFile(this.credentialsPath).then(output => {
        this.credentialsJSON = JSON.parse(output);
      }).catch((err) => {
        reject(err);
      });
      const {
        client_id,
        client_secret,
        redirect_uris
      } = this.credentialsJSON.installed;
      this.oAuth2Client = new google.auth.OAuth2(
        client_id, client_secret, redirect_uris[0]);
      await accessFile(this.tokenPath).then(async (found) => {
        if (found) {
          await this.setToken().then(() => {
            resolve('Token set');
          }).catch(err => {
            reject(err);
          });
        } else {
          this.getNewAccessToken().then(() => {
            resolve('New Token Saved');
          }).catch(err => {
            reject(err);
          });
        }
      });
    });
  }

  setToken () {
    return new Promise(async (resolve, reject) => {
      await readFile(this.tokenPath).then(token => {
        this.tokenJSON = JSON.parse(token);
        this.oAuth2Client.setCredentials(this.tokenJSON);
        this.drive = google.drive({ version: 'v3', auth: this.oAuth2Client });
        this.sheet = google.sheets({ version: 'v4', auth: this.oAuth2Client });
        // <- you set an api property for our class here
        // for example if you want to add drive api : this.drive = google.drive({ version: 'v3', auth: this.oAuth2Client });
        resolve(true);
      }).catch((err) => {
        reject(err);
      });
    });
  }

  getNewAccessToken () {
    return new Promise((resolve, reject) => {
      const authUrl = this.oAuth2Client.generateAuthUrl({
        access_type: 'offline',
        scope: this.scopes
      });
      console.log('Authorize this app by visiting this url:', authUrl);
      const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
      });
      rl.question('Enter the code from that page here: ', (code) => {
        rl.close();
        this.oAuth2Client.getToken(code, async (err, token) => {
          if (err) {
            return console.log('Error retrieving access token', err);
          } else {
            this.oAuth2Client.setCredentials(token);
            await writeFile(this.tokenPath, JSON.stringify(token)).then(() => {
              this.drive = google.drive({ version: 'v3', auth: this.oAuth2Client });
              this.sheet = google.sheets({ version: 'v4', auth: this.oAuth2Client });
              // <- you set an api property for our class here
              // for example if you want to add drive api : this.drive = google.drive({ version: 'v3', auth: this.oAuth2Client });
              resolve(true);
            }).catch((err) => {
              reject(err);
            });
          }
        });
      });
    });
  }
  /**
   * @param {Object} param
   */
  listFiles (param) {
    return new Promise((resolve, reject) => {
      this.drive.files.list(param, (err, res) => {
        if (err) {
          reject(err);
        } else {
          const files = res.data.files;
          if (files.length) {
            resolve(files);
          } else {
            resolve(false);
          }
        }
      });
    });
  }

  /**
   * @param {String} fileName
   * @param {Array} parents
   */
  createSpreadsheet ({ fileName, parents }) {
    return new Promise((resolve, reject) => {
      this.drive.files
        .create({
          requestBody: {
            name: fileName,
            mimeType: 'application/vnd.google-apps.spreadsheet',
            parents: parents
          }
        })
        .then((res) => {
          resolve(res);
        })
        .catch(err => {
          reject(err);
        });
    });
  }

  /**
   * @param {String} folderName
   * @param {Array} parents
   */
  createFolder ({ folderName, parents }) {
    return new Promise((resolve, reject) => {
      this.drive.files
        .create({
          requestBody: {
            name: folderName,
            mimeType: 'application/vnd.google-apps.folder',
            parents: parents
          }
        })
        .then((res) => {
          resolve(res);
        })
        .catch(err => reject(err));
    });
  }

  /**
   *
   * @param {String} SpreadsheetID
   * @param {String} title
   * @param {Array} row
   */
  addRow (SpreadsheetID, title, row) {
    return new Promise((resolve, reject) => {
      var request = {
        spreadsheetId: SpreadsheetID,
        range: `${title}!A1`,
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        resource: row,
        auth: this.oAuth2Client
      };

      this.sheet.spreadsheets.values.append(request, function (err) {
        if (err) {
          reject(err);
        } else {
          resolve('ok');
        }
      });
    });
  }

  /**
   *
   * @param {String} spreadsheetId
   */
  getSpreadsheet (spreadsheetId) {
    return new Promise((resolve, reject) => {
      var request = {
        spreadsheetId: spreadsheetId,
        includeGridData: true,
        auth: this.oAuth2Client
      };

      this.sheet.spreadsheets.get(request, function (err, response) {
        if (err) {
          reject(err);
        } else {
          resolve(response);
        }
      });
    });
  }

  /**
   *
   * @param {String} spreadsheetId
   * @param {String} sheetId
   * @param {Int} startIndex1
   * @param {Int} endIndex1
   * @param {Int} size1
   * @param {Int} startIndex2
   * @param {Int} endIndex2
   * @param {Int} size2
   * @param {Int} startIndex3
   * @param {Int} endIndex3
   * @param {Int} size3
   */
  resizeColumn (spreadsheetId, sheetId, startIndex1, endIndex1, size1, startIndex2, endIndex2, size2, startIndex3, endIndex3, size3) {
    return new Promise((resolve, reject) => {
      var request = {
        // The spreadsheet to apply the updates to.
        spreadsheetId: spreadsheetId, // TODO: Update placeholder value.

        resource: {
          requests: [{
            updateDimensionProperties: {
              range: {
                sheetId: sheetId,
                dimension: 'COLUMNS',
                startIndex: startIndex1,
                endIndex: endIndex1
              },
              properties: {
                pixelSize: size1
              },
              fields: 'pixelSize'
            }
          },
          {
            updateDimensionProperties: {
              range: {
                sheetId: sheetId,
                dimension: 'COLUMNS',
                startIndex: startIndex2,
                endIndex: endIndex2
              },
              properties: {
                pixelSize: size2
              },
              fields: 'pixelSize'
            }
          },
          {
            updateDimensionProperties: {
              range: {
                sheetId: sheetId,
                dimension: 'COLUMNS',
                startIndex: startIndex3,
                endIndex: endIndex3
              },
              properties: {
                pixelSize: size3
              },
              fields: 'pixelSize'
            }
          }] // TODO: Update placeholder value.

          // TODO: Add desired properties to the request body.
        },

        auth: this.oAuth2Client
      };

      this.sheet.spreadsheets.batchUpdate(request, function (err, response) {
        if (err) {
          reject(err);
        } else {
          resolve(response);
        }
      });
    });
  }

  /**
   *
   * @param {String} spreadsheetId
   * @param {String} sheets
   */
  addSheet (spreadsheetId, sheets) {
    let sheetToAdd = sheets.map(sheet => {
      return {
        addSheet: {
          properties: {
            title: sheet.title,
            gridProperties: {
              rowCount: 1000,
              columnCount: 7
            }
          }
        }
      };
    });

    return new Promise((resolve, reject) => {
      var request = {
        // The spreadsheet to apply the updates to.
        spreadsheetId: spreadsheetId, // TODO: Update placeholder value.

        resource: {
          requests: [sheetToAdd]
        },

        auth: this.oAuth2Client
      };

      this.sheet.spreadsheets.batchUpdate(request, function (err, response) {
        if (err) {
          reject(err);
        } else {
          resolve(response);
        }
      });
    });
  }

  /**
   *
   * @param {String} spreadsheetId
   * @param {String} sheetId
   */
  deleteSheet (spreadsheetId, sheetId) {
    return new Promise((resolve, reject) => {
      var request = {
        // The spreadsheet to apply the updates to.
        spreadsheetId: spreadsheetId, // TODO: Update placeholder value.

        resource: {
          requests: [{
            deleteSheet: {
              sheetId: sheetId
            }
          }]
        },

        auth: this.oAuth2Client
      };

      this.sheet.spreadsheets.batchUpdate(request, function (err, response) {
        if (err) {
          reject(err);
        } else {
          resolve(response);
        }
      });
    });
  }

  /**
   *
   * @param {String} spreadsheetId
   * @param {String} sheetId
   * @param {String} newName
   */
  renameSheet (spreadsheetId, sheetId, newName) {
    return new Promise((resolve, reject) => {
      var request = {
        // The spreadsheet to apply the updates to.
        spreadsheetId: spreadsheetId, // TODO: Update placeholder value.

        resource: {
          requests: [{
            updateSheetProperties: {
              properties: {
                sheetId: sheetId,
                title: newName
              },
              fields: 'title'
            }
          }]
        },

        auth: this.oAuth2Client
      };

      this.sheet.spreadsheets.batchUpdate(request, function (err, response) {
        if (err) {
          reject(err);
        } else {
          resolve(response);
        }
      });
    });
  }

  /**
   *
   * @param {String} spreadsheetId
   * @param {String} range
   */
  getRow (spreadsheetId, range) {
    this.sheet.spreadsheets.values.get({
      spreadsheetId,
      range
    }, (err, result) => {
      if (err) {
        // Handle error
        console.log(err);
      } else {
        const numRows = result.values ? result.values.length : 0;
        console.log(`${numRows} rows retrieved.`);
      }
    });
  }

  /**
   * @param {String} id
   */
  deleteFile (id) {
    this.drive.files.delete({ fileId: id }).then(() => {
      console.log('File Deleted');
    }).catch(err => {
      console.log(err);
    });
  }
}

module.exports = { Google };

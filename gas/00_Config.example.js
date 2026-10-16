/** Copy to 00_Config.js and enter your own Google resource IDs. */
const CONFIG = Object.freeze({
  sourceFolderId: 'REPLACE_SOURCEFOLDERID',
  formId: 'REPLACE_FORMID',
  inputFolderIds: Object.freeze({
    excel: 'REPLACE_EXCEL',
    beforeCsv: 'REPLACE_BEFORECSV',
    afterCsv: 'REPLACE_AFTERCSV'
  }),
  excelSpreadsheetId: 'REPLACE_EXCELSPREADSHEETID',
  outputFolderId: 'REPLACE_OUTPUTFOLDERID',
  logFolderId: 'REPLACE_LOGFOLDERID',
  scriptId: 'REPLACE_SCRIPTID',
  formResponseUrl: 'https://docs.google.com/forms/d/e/REPLACE/viewform',
  ownerEmail: 'owner@example.com',
  maxBytes: 10 * 1024 * 1024,
  maxDataRows: 9999,
  maxCells: 500000,
  maxFiles: 200
});

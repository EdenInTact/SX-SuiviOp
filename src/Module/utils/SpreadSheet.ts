import GSpreadsheet = GoogleAppsScript.Spreadsheet;

/**
 * The Spreadsheet ands sheets composing it
 *
 * @property spreadsheet, the active spreadsheet
 * @property sheets, the list of sheets composing the spreadsheet
 * @property dataSheet
 * @property suiviSheet
 *
 *
 * */
export default class SpreadSheet {
    get backlogSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        return this._backlogSheet;
    }

    set backlogSheet(value: GoogleAppsScript.Spreadsheet.Sheet) {
        this._backlogSheet = value;
    }

    // ****************************************************************************************** \\
    //                                           PROPERTIES                                       \\
    // ****************************************************************************************** \\

    private _spreadsheet: GSpreadsheet.Spreadsheet;
    private _sheets: GSpreadsheet.Sheet[] = [];
    private _dataSheet: GSpreadsheet.Sheet;
    private _suiviSheet: GSpreadsheet.Sheet;
    private _backlogSheet: GSpreadsheet.Sheet;

    // ****************************************************************************************** \\
    //                                          CONSTRUCTOR                                       \\
    // ****************************************************************************************** \\

    constructor() {
        this._spreadsheet = SpreadsheetApp.getActive();
        this._spreadsheet.getSheets().forEach(sheet => {
            this._sheets[sheet.getName()] = sheet;
        });
        this._dataSheet = this._sheets["Data"];
        this._suiviSheet = this._sheets["Suivi opérationnel"];
        this._backlogSheet = this._sheets["Product Backlog"];

    }

    // ****************************************************************************************** \\
    //                                             ACCESSORS                                      \\
    // ****************************************************************************************** \\

    get spreadsheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
        return this._spreadsheet;
    }

    set spreadsheet(value: GoogleAppsScript.Spreadsheet.Spreadsheet) {
        this._spreadsheet = value;
    }

    get sheets(): GoogleAppsScript.Spreadsheet.Sheet[] {
        return this._sheets;
    }

    set sheets(value: GoogleAppsScript.Spreadsheet.Sheet[]) {
        this._sheets = value;
    }

    get dataSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        return this._dataSheet;
    }

    set dataSheet(value: GoogleAppsScript.Spreadsheet.Sheet) {
        this._dataSheet = value;
    }

    get suiviSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        return this._suiviSheet;
    }

    set suiviSheet(value: GoogleAppsScript.Spreadsheet.Sheet) {
        this._suiviSheet = value;
    }
}

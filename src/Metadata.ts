import g from '@battis/gas-lighter';
import prefix from './Prefix';
import * as SKY from './SKY';

type Range = {
    row: number;
    column: number;
    numRows: number;
    numColumns: number;
    sheet: string;
};

const LIST = prefix('list');
const RANGE = prefix('range');
const LAST_UPDATED = prefix('lastUpdated');

function rangeToJSON(range: GoogleAppsScript.Spreadsheet.Range): Range {
    if (!range) {
        return null;
    }
    return {
        row: range.getRow(),
        column: range.getColumn(),
        numRows: range.getNumRows(),
        numColumns: range.getNumColumns(),
        sheet: range.getSheet().getName(),
    };
}

function rangeFromJSON(json: Range): GoogleAppsScript.Spreadsheet.Range {
    if (!json) {
        return null;
    }
    const sheet = SpreadsheetApp.getActive().getSheetByName(json.sheet);
    // TODO should this be resilient to sheet name changes?
    return sheet.getRange(json.row, json.column, json.numRows, json.numColumns);
}

const get = (key: string, sheet?: GoogleAppsScript.Spreadsheet.Sheet) => {
    sheet = sheet || SpreadsheetApp.getActive().getActiveSheet();
    return g.SpreadsheetApp.DeveloperMetadata.get(sheet, key);
};

export const getList = (
    sheet?: GoogleAppsScript.Spreadsheet.Sheet
): SKY.School.Lists.Metadata => get(LIST, sheet);
export const getRange = (sheet?: GoogleAppsScript.Spreadsheet.Sheet) =>
    rangeFromJSON(get(RANGE, sheet));
export const getLastUpdated = (sheet?: GoogleAppsScript.Spreadsheet.Sheet) =>
    get(LAST_UPDATED, sheet);

const set = (
    key: string,
    sheet: GoogleAppsScript.Spreadsheet.Sheet = null,
    value: any
) => {
    sheet = sheet || SpreadsheetApp.getActive().getActiveSheet();
    return g.SpreadsheetApp.DeveloperMetadata.set(sheet, key, value);
};

export const setList = (
    list: SKY.School.Lists.Metadata,
    sheet?: GoogleAppsScript.Spreadsheet.Sheet
) => set(LIST, sheet, list);
export const setRange = (
    range: GoogleAppsScript.Spreadsheet.Range,
    sheet?: GoogleAppsScript.Spreadsheet.Sheet
) => set(RANGE, sheet, rangeToJSON(range));
export const setLastUpdated = (
    lastUpdated: Date,
    sheet?: GoogleAppsScript.Spreadsheet.Sheet
) => set(LAST_UPDATED, sheet, lastUpdated.toLocaleString());

const remove = (key: string, sheet?: GoogleAppsScript.Spreadsheet.Sheet) => {
    sheet = sheet || SpreadsheetApp.getActive().getActiveSheet();
    return g.SpreadsheetApp.DeveloperMetadata.remove(sheet, key);
};

export const removeList = remove.bind(null, LIST);
export const removeRange = remove.bind(null, RANGE);
export const removeLastUpdated = remove.bind(null, LAST_UPDATED);

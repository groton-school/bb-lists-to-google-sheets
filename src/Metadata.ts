import * as g from '@battis/gas-lighter';
import prefix from './Prefix';

const LIST = prefix('list');
const RANGE = prefix('range');
const LAST_UPDATED = prefix('lastUpdated');

function get(key: string, sheet: GoogleAppsScript.Spreadsheet.Sheet = null) {
    sheet = sheet || SpreadsheetApp.getActive().getActiveSheet();
    if (sheet) {
        return g.SpreadsheetApp.DeveloperMetadata.get(sheet, key);
    }
    return null;
}

export const getList = get.bind(null, LIST);
export const getRange = get.bind(null, RANGE);
export const getLastUpdated = get.bind(null, LAST_UPDATED);

function set(
    key: string,
    sheet: GoogleAppsScript.Spreadsheet.Sheet = null,
    value: any
) {
    sheet = sheet || SpreadsheetApp.getActive().getActiveSheet();
    if (sheet) {
        return g.SpreadsheetApp.DeveloperMetadata.set(sheet, key, value);
    }
    return false;
}

export const setList = set.bind(null, LIST);
export const setRange = set.bind(null, RANGE);
export const setLastUpdated = set.bind(null, LAST_UPDATED);

function remove(key: string, sheet: GoogleAppsScript.Spreadsheet.Sheet = null) {
    sheet = sheet || SpreadsheetApp.getActive().getActiveSheet();
    if (sheet) {
        return g.SpreadsheetApp.DeveloperMetadata.remove(sheet, key);
    }
    return null;
}

export const removeList = remove.bind(null, LIST);
export const removeRange = remove.bind(null, RANGE);
export const removeLastUpdated = remove.bind(null, LAST_UPDATED);

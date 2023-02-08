import { Terse } from '@battis/gas-lighter';

import { PREFIX } from '../Constants';
import * as State from '../State';

const LIST = `${PREFIX}.list`;
const RANGE = `${PREFIX}.range`;
const LAST_UPDATED = `${PREFIX}.lastUpdated`;

function get(key: string, sheet: GoogleAppsScript.Spreadsheet.Sheet = null) {
    sheet = sheet || State.getSheet();
    if (sheet) {
        return Terse.SpreadsheetApp.DeveloperMetadata.get(sheet, key);
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
    sheet = sheet || State.getSheet();
    if (sheet) {
        return Terse.SpreadsheetApp.DeveloperMetadata.set(sheet, key, value);
    }
    return false;
}

export const setList = set.bind(null, LIST);
export const setRange = set.bind(null, RANGE);
export const setLastUpdated = set.bind(null, LAST_UPDATED);

function remove(key: string, sheet: GoogleAppsScript.Spreadsheet.Sheet = null) {
    sheet = sheet || State.getSheet();
    if (sheet) {
        return Terse.SpreadsheetApp.DeveloperMetadata.remove(sheet, key);
    }
    return null;
}

export const removeList = remove.bind(null, LIST);
export const removeRange = remove.bind(null, RANGE);
export const removeLastUpdated = remove.bind(null, LAST_UPDATED);

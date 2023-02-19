import * as g from '@battis/gas-lighter';
import { PREFIX } from './Constants';
import * as SKY from './SKY';

export enum Intent {
    CreateSpreadsheet = 'create',
    AppendSheet = 'append',
    ReplaceSelection = 'replace',
    UpdateExisting = 'update',
}

const FOLDER = `${PREFIX}.State.folder`;
const SPREADSHEET = `${PREFIX}.State.spreadsheet`;
const SHEET = `${PREFIX}.State.sheet`;
const SELECTION = `${PREFIX}.State.selection`;
const LIST = `${PREFIX}.State.list`;
const INTENT = `${PREFIX}.State.intent`;

// FIXME data and page should be stored in SKY.ServiceManaager, not State
const DATA = `${PREFIX}.State.data`;
const PAGE = `${PREFIX}.State.page`;

export function getFolder() {
    const id = g.PropertiesService.getUserProperty(FOLDER);
    let folder = null;
    if (id) {
        folder = DriveApp.getFolderById(id);
    }
    return folder;
}

export function setFolder(folder: GoogleAppsScript.Drive.Folder) {
    const id = folder && folder.getId();
    return g.PropertiesService.setUserProperty(FOLDER, id);
}

export function getSpreadsheet(): GoogleAppsScript.Spreadsheet.Spreadsheet {
    const id = g.PropertiesService.getUserProperty(SPREADSHEET);
    let spreadsheet = null;
    if (id) {
        spreadsheet = SpreadsheetApp.openById(id);
    } else {
        spreadsheet = SpreadsheetApp.getActive();
        setSpreadsheet(spreadsheet);
    }
    return spreadsheet;
}

export function setSpreadsheet(
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
) {
    const id = spreadsheet && spreadsheet.getId();
    return g.PropertiesService.setUserProperty(SPREADSHEET, id);
}

export function getSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const spreadsheet = getSpreadsheet();
    let sheet = null;
    if (spreadsheet) {
        const name = g.PropertiesService.getUserProperty(SHEET);
        if (name) {
            sheet = spreadsheet.getSheetByName(name);
        } else {
            sheet = spreadsheet.getActiveSheet();
            setSheet(sheet);
        }
    }
    return sheet;
}

export function setSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
    const name = sheet && sheet.getName();
    setSpreadsheet(sheet && sheet.getParent());
    return g.PropertiesService.setUserProperty(SHEET, name);
}

export function getSelection(): GoogleAppsScript.Spreadsheet.Range {
    const sheet = getSheet();
    let selection = null;
    if (sheet) {
        const range = g.PropertiesService.getUserProperty(SELECTION);
        if (range) {
            selection = sheet.getRange(range);
        } else {
            selection = sheet.getActiveRange();
            setSelection(selection);
        }
    }
    return selection;
}

export function setSelection(range: GoogleAppsScript.Spreadsheet.Range) {
    const a1notation = range && range.getA1Notation();
    setSheet(range && range.getSheet());
    return g.PropertiesService.setUserProperty(SELECTION, a1notation);
}

export const getList = g.PropertiesService.getUserProperty.bind(
    null,
    LIST
) as () => SKY.School.Lists.Metadata;
export const setList = g.PropertiesService.setUserProperty.bind(null, LIST);
export const getData = g.PropertiesService.getUserProperty.bind(null, DATA);
export const setData = g.PropertiesService.setUserProperty.bind(null, DATA);

export function appendData(page) {
    const data = getData() || [];
    data.push(...page);
    setData(data);
}

export const getPage = g.PropertiesService.getUserProperty.bind(
    null,
    PAGE
) as () => number;
export const setPage = g.PropertiesService.setUserProperty.bind(null, PAGE);
export const getIntent = g.PropertiesService.getUserProperty.bind(
    null,
    INTENT
) as () => Intent;
export const setIntent = g.PropertiesService.setUserProperty.bind(null, INTENT);

export function reset() {
    g.PropertiesService.deleteUserProperty(FOLDER);
    g.PropertiesService.deleteUserProperty(SPREADSHEET);
    g.PropertiesService.deleteUserProperty(SHEET);
    g.PropertiesService.deleteUserProperty(SELECTION);
    g.PropertiesService.deleteUserProperty(LIST);
    g.PropertiesService.deleteUserProperty(DATA);
    g.PropertiesService.deleteUserProperty(PAGE);
    setIntent(Intent.CreateSpreadsheet);
}

export function update(arg) {
    if (arg) {
        var {
            parameters: { state },
        } = arg;
        if (state) {
            state = JSON.parse(state);
            if (state.folder) {
                setFolder(DriveApp.getFolderById(state.folder));
            }

            var spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
            if (state.spreadsheet) {
                spreadsheet = SpreadsheetApp.openById(state.spreadsheet);
            }
            var sheet: GoogleAppsScript.Spreadsheet.Sheet;
            if (state.sheet) {
                spreadsheet = spreadsheet || getSpreadsheet();
                sheet = spreadsheet.getSheetByName(state.sheet);
            }
            if (state.selection) {
                sheet = sheet || getSheet();
                setSelection(sheet.getRange(state.selection));
            } else if (sheet) {
                setSheet(sheet);
            } else if (spreadsheet) {
                setSpreadsheet(spreadsheet);
            }
            if (state.list) {
                setList(state.list);
            }
            if (state.page) {
                setPage(state.page);
            }
            if (state.data) {
                setData(state.data);
            }
            if (state.intent) {
                setIntent(state.intent);
            }
        }
    }
}

export function toString(): string {
    const folder = getFolder();
    const intent = getIntent();
    const spreadsheet = getSpreadsheet();
    const sheet = getSheet();
    const selection = getSelection();
    const list = getList();
    const page = getPage();
    const data = getData();
    return JSON.stringify(
        {
            folder: folder && folder.getId(),
            intent,
            spreadsheet: spreadsheet && spreadsheet.getId(),
            sheet: sheet && sheet.getName(),
            selection: selection && selection.getA1Notation(),
            list,
            page,
            data,
        },
        null,
        2
    );
}

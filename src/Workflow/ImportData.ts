import g from '@battis/gas-lighter';
import * as Metadata from '../Metadata';
import * as SKY from '../SKY';

export enum Target {
    Update = 'update',
    Selection = 'selection',
    Sheet = 'sheet',
    Spreadsheet = 'spreadsheet',
}

let target: Target = null;
let list: SKY.School.Lists.Metadata = null;
const data: any[][] = [];
let spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = null;
let sheet: GoogleAppsScript.Spreadsheet.Sheet = null;
let range: GoogleAppsScript.Spreadsheet.Range = null;
let prevRange: GoogleAppsScript.Spreadsheet.Range = null;
let progress = null;

export function importData(
    l: SKY.School.Lists.Metadata,
    t: Target,
    thread: string
) {
    list = l;
    target = t;
    progress = g.HtmlService.Element.Progress.bindTo(thread);
    progress.reset();
    progress.setMax(7);
    progress.setStatus('Loading…');
    progress.incrementValue();

    loadData();
    connectToSpreadsheet(data.length, data[0].length);

    progress.setStatus('Writing data…');
    progress.incrementValue();
    range.setValues(data);
    formatColumnHeaders();
    if (target == Target.Sheet || target == Target.Spreadsheet) {
        setNextAvailableSheetName();
    }
    updateMetadata();

    progress.setComplete(
        g.SpreadsheetApp.Dialog.getHtml({
            message: getCompletionMessage(),
        })
    );
}
global.importData = importData;

function loadData() {
    let frame: any[][];
    let page = 1;
    let complete = false;
    do {
        frame = SKY.School.Lists.get(
            list.id,
            SKY.ServiceManager.ResponseFormat.Array,
            page
        ) as any[][];
        if (data.length > 0) {
            frame.shift();
        }
        data.push(...frame);
        if (frame.length >= SKY.PAGE_SIZE) {
            progress.setStatus(`Loaded page ${page++} (${data.length - 1} rows)`);
            progress.incrementValue();
            progress.setMax(progress.getMax() + 1);
        } else {
            complete = true;
        }
    } while (!complete);
}

function connectToSpreadsheet(numRows: number, numColumns: number) {
    progress.setStatus('Connecting to spreadsheet…');
    progress.incrementValue();
    spreadsheet = SpreadsheetApp.getActive();
    switch (target) {
        case Target.Update:
        case Target.Selection:
            sheet = spreadsheet.getActiveSheet();
            if (target == Target.Update) {
                prevRange = Metadata.getRange(sheet);
            } else {
                prevRange = sheet.getActiveRange();
            }
            prevRange.clearContent();
            range = adjustRange(
                {
                    row: prevRange.getRow(),
                    column: prevRange.getColumn(),
                    numRows,
                    numColumns,
                },
                prevRange
            );
            break;
        case Target.Sheet:
            sheet = spreadsheet.insertSheet();
            range = adjustRange(
                {
                    row: 1,
                    column: 1,
                    numRows,
                    numColumns,
                },
                null,
                sheet
            );
            break;
        case Target.Spreadsheet:
            spreadsheet = SpreadsheetApp.create(list.name, numRows, numColumns);
            sheet = spreadsheet.getSheets()[0];
            range = sheet.getRange(1, 1, numRows, numColumns);
    }

    return { target, list, spreadsheet, sheet, range, prevRange };
}

function adjustRange(
    { row, column, numRows, numColumns },
    range = null,
    sheet = null
) {
    if (range) {
        sheet = range.getSheet();
        if (numRows > range.getNumRows()) {
            progress.setStatus('Adding rows…');
            sheet.insertRows(range.getLastRow(), numRows - range.getNumRows());
        }
        if (numColumns > range.getNumColumns()) {
            progress.setStatus('Adding columns…');
            sheet.insertColumns(
                range.getLastColumn(),
                numColumns - range.getNumColumns()
            );
        }
    } else if (sheet) {
        if (numRows < sheet.getMaxRows()) {
            progress.setStatus('Trimming rows…');
            sheet.deleteRows(numRows + 1, sheet.getMaxRows() - numRows);
        }
        if (numColumns < sheet.getMaxColumns()) {
            progress.setStatus('Trimming columns…');
            sheet.deleteColumns(numColumns + 1, sheet.getMaxColumns() - numColumns);
        }
    }
    return sheet.getRange(row, column, numRows, numColumns);
}

function formatColumnHeaders() {
    progress.setStatus('Formatting column headers…');
    progress.incrementValue();
    range.offset(0, 0, 1, range.getNumColumns()).setFontWeight('bold');
    if (target == Target.Sheet || target == Target.Spreadsheet) {
        sheet.setFrozenRows(1);
    }
}

function updateMetadata() {
    progress.setStatus('Updating metadata…');
    progress.incrementValue();
    const timestamp = new Date();
    Metadata.setList(list, sheet);
    Metadata.setRange(range, sheet);
    Metadata.setLastUpdated(timestamp, sheet);
    range
        .offset(0, 0, 1, 1)
        .setNote(`Last updated from "${list.name}" ${timestamp.toLocaleString()}`);
}

function setNextAvailableSheetName() {
    progress.setStatus('Naming sheet…');
    progress.incrementValue();
    const baseName = list.name;
    var name = baseName;
    for (var i = 1; spreadsheet.getSheetByName(name); i++) {
        name = `${baseName} ${i}`; // I don't like this format, but it mirrors Sheets naming conventions
    }
    sheet.setName(name);
}

function getCompletionMessage() {
    switch (target) {
        case Target.Update:
            return `<code>${range.getA1Notation()}</code> has been updated with the latest data from "${list.name
                }" on Blackbaud.${range.getA1Notation() != prevRange.getA1Notation()
                    ? ` (The prior data occupied <code>${prevRange.getA1Notation()}</code>, so the sheet has been expanded rightward and downward to accommodate the increased amount of data.)`
                    : ''
                }`;
        case Target.Selection:
            return `<code>${range.getA1Notation()}</code> has been connected to "${list.name
                }" on Blackbaud and the data has been imported.${range.getA1Notation() != prevRange.getA1Notation()
                    ? ` (The original selection occupied <code>${prevRange.getA1Notation()}</code>, so the sheet has been expanded rightward and downward to accommodate the larger amount of data.)`
                    : ''
                }`;
        case Target.Sheet:
            return `"${list.name
                }" on Blackbaud has been connected to the sheet "${sheet.getName()}" and the data has been imported.`;
        case Target.Spreadsheet:
            return `"${list.name
                }" on Blackbaud has been connected to the sheet of the same name in the spreadsheet "${spreadsheet.getName()}" and the data has been imported.<br/>
                    <a href="${spreadsheet.getUrl()}" target="_blank">Open ${spreadsheet.getName()}get</a>`;
    }
}

import * as g from '@battis/gas-lighter';
import * as Metadata from '../Metadata';
import * as SKY from '../SKY';

export enum Target {
    selection,
    sheet,
    spreadsheet,
    update,
}

function adjustRange(
    { row, column, numRows, numColumns },
    range = null,
    sheet = null
) {
    if (range) {
        sheet = range.getSheet();
        if (numRows > range.getNumRows()) {
            sheet.insertRows(range.getLastRow() + 1, numRows - range.getNumRows());
        }
        if (numColumns > range.getNumColumns()) {
            sheet.insertColumns(
                range.getLastColumn() + 1,
                numColumns - range.getNumColumns()
            );
        }
    } else if (sheet) {
        if (numRows < sheet.getMaxRows()) {
            sheet.deleteRows(numRows + 1, sheet.getMaxRows() - numRows);
        }
        if (numColumns < sheet.getMaxColumns()) {
            sheet.deleteColumns(numColumns + 1, sheet.getMaxColumns() - numColumns);
        }
    }
    return sheet.getRange(row, column, numRows, numColumns);
}

export default (
    list: SKY.School.Lists.Metadata,
    target: Target,
    thread: string
) => {
    const progress = g.HtmlService.Element.Progress.getInstance(thread);
    progress.reset();
    progress.setStatus('Loading…');
    const data: any[][] = [];
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
            progress.setStatus(`Loaded page ${page} (${data.length - 1} rows)`);
            progress.setValue(page++);
            progress.setMax(page);
        } else {
            progress.setStatus('Writing data to sheet');
            complete = true;
        }
    } while (!complete);

    let spreadsheet = SpreadsheetApp.getActive();
    let sheet: GoogleAppsScript.Spreadsheet.Sheet = null;
    let range: GoogleAppsScript.Spreadsheet.Range = null;
    switch (target) {
        case Target.sheet:
            sheet = spreadsheet.insertSheet();
            range = adjustRange(
                {
                    row: 1,
                    column: 1,
                    numRows: data.length,
                    numColumns: data[0].length,
                },
                null,
                sheet
            );
            break;
        case Target.selection:
            sheet = spreadsheet.getActiveSheet();
            range = sheet.getActiveRange();
            range.clearContent();
            range = adjustRange(
                {
                    row: range.getRow(),
                    column: range.getColumn(),
                    numRows: data.length,
                    numColumns: data[0].length,
                },
                range
            );
            break;
        case Target.update:
            const metaRange = Metadata.getRange();
            range = adjustRange(
                {
                    row: metaRange.getRow(),
                    column: metaRange.getColumn(),
                    numRows: data.length,
                    numColumns: data[0].length,
                },
                metaRange
            );
            break;
        case Target.spreadsheet:
        default:
            spreadsheet = SpreadsheetApp.create(
                list.name,
                data.length,
                data[0].length
            );
            sheet = spreadsheet.getSheets()[0];
            range = sheet.getRange(1, 1, data.length, data[0].length);
    }

    range.setValues(data);
    range.offset(0, 0, 1, range.getNumColumns()).setFontWeight('bold');
    const timestamp = new Date();

    Metadata.setList(list, range.getSheet());
    Metadata.setRange(range, range.getSheet());
    Metadata.setLastUpdated(timestamp, range.getSheet());
    range
        .offset(0, 0, 1, 1)
        .setNote(`Last updated from "${list.name}" ${timestamp.toLocaleString()}`);

    let message = 'Complete';
    switch (target) {
        case Target.sheet:
            range.getSheet().setFrozenRows(1);
            const baseName = list.name;
            var name = baseName;
            const spreadsheet = range.getSheet().getParent();
            for (var i = 1; spreadsheet.getSheetByName(name); i++) {
                name = `${baseName} ${i}`; // I don't like this format, but it mirrors Sheets naming conventions
            }
            range.getSheet().setName(name);
            message = `${range.getSheet().getParent().getUrl()}#${range
                .getSheet()
                .getSheetId()}`;
            break;
        case Target.spreadsheet:
            range.getSheet().setFrozenRows(1);
            range.getSheet().setName(list.name);
            message = range.getSheet().getParent().getUrl();
            break;
    }

    progress.setComplete(message);
};

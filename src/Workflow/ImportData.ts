import * as g from '@battis/gas-lighter';
import * as Metadata from '../Metadata';
import * as SKY from '../SKY';

export enum Target {
    Selection = 'selection',
    Sheet = 'sheet',
    Spreadsheet = 'spreadsheet',
    Update = 'update',
}

global.getImportTargetOptions = (list: SKY.School.Lists.Metadata) => {
    const prevList = Metadata.getList();
    const selection = SpreadsheetApp.getActive()
        .getActiveSheet()
        .getActiveRange()
        .getA1Notation();
    let message = `Where would you like the data from "${list.name}" to be imported?`;
    let buttons: g.UI.Dialog.Button[] = [
        { name: 'Add Sheet', value: 'sheet', class: 'action' },
        { name: 'New Spreadsheet', value: 'spreadsheet', class: 'action' },
    ];
    if (prevList) {
        message += ` This sheet is already connected to "${prevList.name}" on Blackbaud, so you need to choose a new sheet or spreadsheet as a new destination for "${list.name}"`;
    } else {
        message += ` If you choose to replace ${selection}, its contents will be erased and, if necessary, additional rows and/or columns will be added to make room for the data from "${list.name}".`;
        buttons.unshift({
            name: `Replace ${selection}`,
            value: 'selection',
            class: 'create',
        });
    }
    return g.SpreadsheetApp.Dialog.getHtml({
        message,
        buttons,
        functionName: 'importReturnTarget',
    });
};

const callImportData: g.UI.Dialog.ResponseHandler = (target) => ({
    functionName: 'importData',
    args: [target],
});
global.importReturnTarget = callImportData;

export function importData(
    list: SKY.School.Lists.Metadata,
    target: Target,
    thread: string
) {
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
        // FIXME multipage updates are failing
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
        case Target.Sheet:
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
        case Target.Selection:
            // FIXME data is not being written to selection
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
        case Target.Update:
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
        case Target.Spreadsheet:
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
        case Target.Sheet:
            range.getSheet().setFrozenRows(1);
            const baseName = list.name;
            var name = baseName;
            const spreadsheet = range.getSheet().getParent();
            for (var i = 1; spreadsheet.getSheetByName(name); i++) {
                name = `${baseName} ${i}`; // I don't like this format, but it mirrors Sheets naming conventions
            }
            range.getSheet().setName(name);
            SpreadsheetApp.getActive().setActiveSheet(range.getSheet());
            message = g.SpreadsheetApp.Dialog.getHtml({
                message: `"${list.name
                    }" on Blackbaud has been connected to the sheet "${range
                        .getSheet()
                        .getName()}" and the data has been imported.`,
            });
            break;
        case Target.Spreadsheet:
            range.getSheet().setFrozenRows(1);
            range.getSheet().setName(list.name);
            message = g.SpreadsheetApp.Dialog.getHtml({
                message: `"${list.name
                    }" on Blackbaud has been connected to the sheet of the same name in the spreadsheet "${range
                        .getSheet()
                        .getParent()
                        .getName()}" and the data has been imported.<br/>
                    <a href="${range
                        .getSheet()
                        .getParent()
                        .getUrl()}" target="_blank">Open ${range
                            .getSheet()
                            .getParent()
                            .getName()}get</a>`,
            });
            break;
    }

    progress.setComplete(message);
}
global.importData = importData;

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

import * as State from '../State';

export * as Metadata from './Metadata';

export function rangeToJSON(range) {
    return {
        row: range.getRow(),
        column: range.getColumn(),
        numRows: range.getNumRows(),
        numColumns: range.getNumColumns(),
        sheet: range.getSheet().getName(),
    };
}

export function rangeFromJSON(json) {
    // FIXME this fallback is unsafe without tracking tab changes!
    const sheet =
        State.getSpreadsheet().getSheetByName(json.sheet) || State.getSheet();
    return sheet.getRange(json.row, json.column, json.numRows, json.numColumns);
}

export function adjustRange(
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

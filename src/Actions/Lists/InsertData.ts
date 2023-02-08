import * as Sheets from '../../Sheets';
import * as SKY from '../../SKY';
import * as State from '../../State';
import { sheetAppendedAction } from '../Sheets/SheetAppended';
import { spreadsheetCreatedAction } from '../Sheets/SpreadsheetCreated';
import { updatedAction } from '../Sheets/Updated';
import { emptyListAction } from './EmptyList';

export function insertDataAction(arg = null) {
    State.update(arg);
    const data = State.getData();

    if (!data || data.length == 0) {
        return emptyListAction(SKY.ServiceManager.getLastResponse());
    }

    var range = null;
    switch (State.getIntent()) {
        case State.Intent.AppendSheet:
            const sheet = State.getSpreadsheet().insertSheet();
            range = Sheets.adjustRange(
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
        case State.Intent.ReplaceSelection:
            State.getSelection().clearContent();
            range = Sheets.adjustRange(
                {
                    row: State.getSelection().getRow(),
                    column: State.getSelection().getColumn(),
                    numRows: data.length,
                    numColumns: data[0].length,
                },
                State.getSelection()
            );
            break;
        case State.Intent.UpdateExisting:
            const metaRange = Sheets.Metadata.getRange();
            range = Sheets.adjustRange(
                {
                    ...metaRange,
                    numRows: data.length,
                    numColumns: data[0].length,
                },
                Sheets.rangeFromJSON(metaRange)
            );
            break;
        case State.Intent.CreateSpreadsheet:
        default:
            State.setSpreadsheet(
                SpreadsheetApp.create(State.getList().name, data.length, data[0].length)
            );
            // FIXME once again not creating in desired folder
            if (State.getFolder()) {
                DriveApp.getFileById(State.getSpreadsheet().getId()).moveTo(
                    State.getFolder()
                );
            }
            State.setSheet(State.getSpreadsheet().getSheets()[0]);
            range = State.getSheet().getRange(1, 1, data.length, data[0].length);
    }

    range.setValues(data);
    range.offset(0, 0, 1, range.getNumColumns()).setFontWeight('bold');
    const timestamp = new Date().toLocaleString();

    Sheets.Metadata.setList(range.getSheet(), State.getList());
    Sheets.Metadata.setRange(range.getSheet(), Sheets.rangeToJSON(range));
    Sheets.Metadata.setLastUpdated(range.getSheet(), timestamp);
    range
        .offset(0, 0, 1, 1)
        .setNote(`Last updated from "${State.getList().name}" ${timestamp}`);

    switch (State.getIntent()) {
        case State.Intent.ReplaceSelection:
            return updatedAction();
        case State.Intent.UpdateExisting:
            return updatedAction();
        case State.Intent.AppendSheet:
            range.getSheet().setFrozenRows(1);
            const baseName = State.getList().name;
            var name = baseName;
            const spreadsheet = range.getSheet().getParent();
            for (var i = 1; spreadsheet.getSheetByName(name); i++) {
                name = `${baseName}${i}`; // I don't like this format, but it mirrors Sheets naming conventions
            }
            range.getSheet().setName(name);
            State.setSheet(range.getSheet());
            // TODO why isn't the appended sheet made active?
            State.getSpreadsheet().setActiveSheet(range.getSheet());
            return sheetAppendedAction();
        case State.Intent.CreateSpreadsheet:
        default:
            range.getSheet().setFrozenRows(1);
            range.getSheet().setName(State.getList().name);
            return spreadsheetCreatedAction();
    }
}
global.action_lists_insertData = insertDataAction;
export default 'action_lists_insertData';

import { Terse } from '@battis/google-apps-script-helpers';
import { PREFIX } from './Constants';
import { List } from './SKY';

export enum Intent {
    CreateSpreadsheet = 'create',
    AppendSheet = 'append',
    ReplaceSelection = 'replace',
    UpdateExisting = 'update',
}

export default class State {
    private static readonly FOLDER = `${PREFIX}.State.folder`;
    private static readonly SPREADSHEET = `${PREFIX}.State.spreadsheet`;
    private static readonly SHEET = `${PREFIX}.State.sheet`;
    private static readonly SELECTION = `${PREFIX}.State.selection`;
    private static readonly LIST = `${PREFIX}.State.list`;
    private static readonly DATA = `${PREFIX}.State.data`;
    private static readonly PAGE = `${PREFIX}.State.page`;
    private static readonly INTENT = `${PREFIX}.State.intent`;

    public static getFolder() {
        return Terse.PropertiesService.getUserProperty(
            State.FOLDER,
            (id) => id && DriveApp.getFolderById(id)
        );
    }

    public static setFolder(folder: GoogleAppsScript.Drive.Folder) {
        if (folder) {
            return Terse.PropertiesService.setUserProperty(
                State.FOLDER,
                folder.getId()
            );
        }
        return Terse.PropertiesService.deleteUserProperty(State.FOLDER);
    }

    public static getSpreadsheet() {
        return Terse.PropertiesService.getUserProperty(
            State.SPREADSHEET,
            (id) => id && SpreadsheetApp.openById(id)
        );
    }

    public static setSpreadsheet(
        spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
    ) {
        if (spreadsheet) {
            return Terse.PropertiesService.setUserProperty(
                State.SPREADSHEET,
                spreadsheet.getId()
            );
        }
        return Terse.PropertiesService.deleteUserProperty(State.SPREADSHEET);
    }

    public static getSheet() {
        const spreadsheet = State.getSpreadsheet();
        if (spreadsheet) {
            return Terse.PropertiesService.getUserProperty(
                State.SHEET,
                (name) => name && spreadsheet.getSheetByName(name)
            );
        }
        return null;
    }

    public static setSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
        if (sheet) {
            State.setSpreadsheet(sheet.getParent());

            return Terse.PropertiesService.setUserProperty(
                State.SHEET,
                sheet.getName()
            );
        }
        return Terse.PropertiesService.deleteUserProperty(State.SHEET);
    }

    public static getSelection() {
        const sheet = State.getSheet();
        if (sheet) {
            return Terse.PropertiesService.getUserProperty(
                State.SELECTION,
                (a1notation) => a1notation && sheet.getRange(a1notation)
            );
        }
        return null;
    }

    public static setSelection(range: GoogleAppsScript.Spreadsheet.Range) {
        if (range) {
            State.setSheet(range.getSheet());
            return Terse.PropertiesService.setUserProperty(
                State.SELECTION,
                range.getA1Notation()
            );
        }
        return Terse.PropertiesService.deleteUserProperty(State.SELECTION);
    }

    public static getList() {
        return Terse.PropertiesService.getUserProperty(State.LIST, JSON.parse);
    }

    public static setList(list: List) {
        if (list) {
            return Terse.PropertiesService.setUserProperty(
                State.LIST,
                JSON.stringify(list)
            );
        }
        return Terse.PropertiesService.deleteUserProperty(State.LIST);
    }

    public static getData() {
        return Terse.PropertiesService.getUserProperty(State.DATA, JSON.parse);
    }

    public static setData(data) {
        if (data) {
            return Terse.PropertiesService.setUserProperty(
                State.DATA,
                JSON.stringify(data)
            );
        }
        return Terse.PropertiesService.deleteUserProperty(State.DATA);
    }

    public static appendData(page) {
        const data = State.getData() || [];
        data.push(...page);
        return Terse.PropertiesService.setUserProperty(
            State.DATA,
            JSON.stringify(data)
        );
    }

    public static getPage() {
        return Terse.PropertiesService.getUserProperty(State.PAGE, parseInt);
    }

    public static setPage(page) {
        return Terse.PropertiesService.setUserProperty(State.PAGE, page);
    }

    public static getIntent() {
        return Terse.PropertiesService.getUserProperty(State.INTENT);
    }

    public static setIntent(intent: Intent) {
        return Terse.PropertiesService.setUserProperty(State.INTENT, intent);
    }

    public static reset() {
        Terse.PropertiesService.deleteUserProperty(State.FOLDER);
        Terse.PropertiesService.deleteUserProperty(State.SPREADSHEET);
        Terse.PropertiesService.deleteUserProperty(State.SHEET);
        Terse.PropertiesService.deleteUserProperty(State.SELECTION);
        Terse.PropertiesService.deleteUserProperty(State.LIST);
        Terse.PropertiesService.deleteUserProperty(State.DATA);
        Terse.PropertiesService.deleteUserProperty(State.PAGE);
        State.setIntent(Intent.CreateSpreadsheet);

        const spreadsheet = SpreadsheetApp.getActive();
        if (spreadsheet) {
            State.setSelection(
                spreadsheet.getActiveSheet().getSelection().getActiveRange()
            );
        }
    }

    public static update({ parameters: { state = null } }) {
        if (state) {
            state = JSON.parse(state);
            if (state.folder) {
                State.setFolder(DriveApp.getFolderById(state.folder));
            }

            var spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
            if (state.spreadsheet) {
                spreadsheet = SpreadsheetApp.openById(state.spreadsheet);
            }
            var sheet: GoogleAppsScript.Spreadsheet.Sheet;
            if (state.sheet) {
                spreadsheet = spreadsheet || State.getSpreadsheet();
                sheet = spreadsheet.getSheetByName(state.sheet);
            }
            if (state.selection) {
                sheet = sheet || State.getSheet();
                State.setSelection(sheet.getRange(state.selection));
            } else if (sheet) {
                State.setSheet(sheet);
            } else if (spreadsheet) {
                State.setSpreadsheet(spreadsheet);
            }
            if (state.list) {
                State.setList(state.list);
            }
            if (state.page) {
                State.setPage(state.page);
            }
            if (state.data) {
                State.setData(state.data);
            }
            if (state.intent) {
                State.setIntent(state.intent);
            }
        }
    }
}
